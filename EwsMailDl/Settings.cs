using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Exchange.WebServices.Data;

namespace EwsMailDl
{
    class Settings
    {
        public ExchangeVersion Version { get; set; }

        public Uri Url { get; set; }

        public string Username { get; set; }

        public string Password { get; set; }

        public int Lifetime
        {
            get { return _lifetime; }
            set { _lifetime = value < 1 ? 1 : value > 30 ? 30 : value; }
        }

        public string SavePath { get; set; }

        public string InputPath { get; set; }

        public string FolderName { get; set; }

        public FolderId FolderId { get; set; }

        public IList<string> SubjectFilters { get; set; }

        public bool Timestamp { get; set; }

        public bool Body { get; set; }

        public DeleteMode Delete { get; set; }

        private int _lifetime = 30;

        public Settings()
        {
            Version = ExchangeVersion.Exchange2010_SP2;
            Url = new Uri("https://localhost/EWS/Exchange.asmx");
            Username = "someone@localhost";
            Password = "T0PS3CR3T";
            SavePath = "";
            InputPath = "";
            FolderName = "Inbox";
            FolderId = null;
            SubjectFilters = new List<string>();
            Timestamp = false;
            Body = false;
            Delete = DeleteMode.HardDelete;
        }

        public void ReadFromArgs(string[] args)
        {
            foreach (var arg in args)
            {
                if (!arg.StartsWith("/"))
                {
                    continue;
                }

                var parts = arg.Split(new char[] { '=' }, 2);

                if (parts.Length != 2)
                {
                    continue;
                }

                string argName = parts[0].Substring(1);
                string argValue = parts[1];

                switch (argName)
                {
                    case "version":
                        Version = (ExchangeVersion)Enum.Parse(typeof(ExchangeVersion), argValue, true);
                        break;

                    case "url":
                        Url = new Uri(argValue);
                        break;

                    case "username":
                        Username = argValue;
                        break;

                    case "password":
                        Password = argValue;
                        break;

                    case "lifetime":
                        Lifetime = Int32.Parse(argValue);
                        break;

                    case "foldername":
                        FolderName = argValue;
                        break;
                    
                    case "folderid":
                        FolderId = new FolderId(argValue);
                        break;

                    case "savepath":
                        SavePath = argValue;
                        break;

                    case "inputpath":
                        InputPath = argValue;
                        break;

                    case "subject":
                        SubjectFilters.Add(argValue);
                        break;

                    case "timestamp":
                        Timestamp = true;
                        break;

                    case "body":
                        Body = true;
                        break;

                    case "delete":
                        Delete = argValue.Equals("soft")
                            ? DeleteMode.SoftDelete
                            : argValue.Equals("move")
                                ? DeleteMode.MoveToDeletedItems
                                : DeleteMode.HardDelete;
                        break;
                }
            }
        }

        public override string ToString()
        {
            var sb = new StringBuilder();

            sb.AppendFormat("Version    : {0}", Version);
            sb.AppendLine();
            sb.AppendFormat("URL        : {0}", Url);
            sb.AppendLine();
            sb.AppendFormat("Username   : {0}", Username);
            sb.AppendLine();
            sb.AppendFormat("Lifetime   : {0}", Lifetime);
            sb.AppendLine();
            sb.AppendFormat("Folder name: {0}", FolderName);
            sb.AppendLine();
            sb.AppendFormat("Folder ID  : {0}", FolderId);
            sb.AppendLine();
            sb.AppendFormat("Save path  : {0}", SavePath);
            sb.AppendLine();
            sb.AppendFormat("Input path : {0}", InputPath);
            sb.AppendLine();
            sb.AppendFormat("Timestamp  : {0}", Timestamp ? "Yes" : "No");
            sb.AppendLine();
            sb.AppendFormat("Body       : {0}", Body ? "Yes" : "No");
            sb.AppendLine();
            sb.AppendFormat("Delete     : {0}", Delete);
            sb.AppendLine();
            sb.AppendFormat("Subject    : {0}", String.Join(" OR ", SubjectFilters));
            
            return sb.ToString();
        }

        public FolderId CreateFolderId()
        {
            if (FolderId != null)
            {
                return FolderId;
            }

            WellKnownFolderName folderName;

            if (Enum.TryParse<WellKnownFolderName>(FolderName, true, out folderName))
            {
                return new FolderId(folderName);
            }

            var results = CreateExchangeService().FindFolders(
                WellKnownFolderName.Root,
                new SearchFilter.IsEqualTo(FolderSchema.DisplayName, FolderName),
                new FolderView(1) { PropertySet = PropertySet.IdOnly, Traversal = FolderTraversal.Deep }
            );

            return results.Folders.Count > 0 ? results.Folders[0].Id : null;
        }

        public SearchFilter CreateSearchFilter()
        {
            var hasAttachments = new SearchFilter.IsEqualTo(ItemSchema.HasAttachments, true);
            var subjectFilters = new SearchFilter.SearchFilterCollection(LogicalOperator.Or);

            foreach (var subjectFilter in SubjectFilters)
            {
                subjectFilters.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, subjectFilter, ContainmentMode.Substring, ComparisonMode.IgnoreCaseAndNonSpacingCharacters));
            }

            if (Body)
            {
                if (subjectFilters.Count == 0)
                {
                    return null;
                }

                return subjectFilters;
            }
            
            return new SearchFilter.SearchFilterCollection(LogicalOperator.And, hasAttachments, subjectFilters);
        }

        public ExchangeService CreateExchangeService()
        {
            return new ExchangeService(Version)
            {
                Credentials = new WebCredentials(Username, Password),
                Url = Url
            };
        }
    }
}
