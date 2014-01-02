using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using Microsoft.Exchange.WebServices.Data;

namespace EwsMailDl
{
    class EmailDownloader
    {
        private EventLog eventLog;

        private List<string> emailIdCache = new List<string>(10);

        private CancellationTokenSource tokenSource;

        private BlockingCollection<ItemId> emailIdQueue;

        private ExchangeService exchangeService;

        private string savePath;

        private IList<string> subjectFilters;

        private bool timestamp;

        public EmailDownloader(EventLog eventLog, BlockingCollection<ItemId> emailIdQueue, CancellationTokenSource tokenSource, Settings settings)
        {
            this.eventLog = eventLog;
            this.tokenSource = tokenSource;
            this.emailIdQueue = emailIdQueue;
            this.exchangeService = settings.CreateExchangeService();
            this.savePath = settings.SavePath;
            this.subjectFilters = settings.SubjectFilters;
            this.timestamp = settings.Timestamp;
        }

        public void Run()
        {
            while (!tokenSource.IsCancellationRequested && !emailIdQueue.IsAddingCompleted)
            {
                ItemId emailId = null;

                try
                {
                    emailId = emailIdQueue.Take(tokenSource.Token);
                }
                catch (OperationCanceledException)
                {
                    if (tokenSource.IsCancellationRequested)
                    {
                        if (Environment.UserInteractive)
                        {
                            Console.WriteLine("Downloader cancelled!");
                        }

                        break;
                    }
                }

                if (emailId == null || emailIdCache.Contains(emailId.UniqueId))
                {
                    if (Environment.UserInteractive)
                    {
                        Console.WriteLine("Ignoring a duplicate e-mail: {0}", emailId);
                    }

                    continue;
                }

                if (emailIdCache.Count == emailIdCache.Capacity)
                {
                    emailIdCache.RemoveAt(0);
                }

                emailIdCache.Add(emailId.UniqueId);

                EmailMessage email = null;

                try
                {
                    email = EmailMessage.Bind(
                        exchangeService,
                        emailId,
                        new PropertySet(
                            EmailMessageSchema.Subject,
                            EmailMessageSchema.Attachments,
                            EmailMessageSchema.DateTimeReceived
                        )
                    );
                }
                catch (Exception) { }

                if (email == null)
                {
                    continue;
                }

                if (MatchEmail(email))
                {
                    if (Environment.UserInteractive)
                    {
                        Console.WriteLine("Processing a new matching e-mail: {0}", email.Subject);
                    }

                    DownloadAndDelete(email);
                }
                else if (Environment.UserInteractive && email != null)
                {
                    Console.WriteLine("Ignoring a not matching e-mail: {0}", email.Subject);
                }
            }
        }

        private bool MatchEmail(EmailMessage email)
        {
            return email.Attachments.Count > 0
                && (subjectFilters.Count == 0 || subjectFilters.Any(phrase => email.Subject.ToLower().Contains(phrase.ToLower())));
        }

        private void DownloadAndDelete(EmailMessage email)
        {
            foreach (var attachment in email.Attachments)
            {
                if (!(attachment is FileAttachment))
                {
                    continue;
                }

                var fileAttachment = attachment as FileAttachment;

                try
                {
                    if (Environment.UserInteractive)
                    {
                        Console.WriteLine("Downloading an attachment: {0}", fileAttachment.Name);
                    }

                    var filePath = CreateFilePath(email.DateTimeReceived, fileAttachment.Name);

                    fileAttachment.Load(filePath);
                    File.SetCreationTime(filePath, email.DateTimeReceived);
                }
                catch (Exception x)
                {
                    HandleException("Failed to download the attachment", x);
                }
            }

            try
            {
                if (Environment.UserInteractive)
                {
                    Console.WriteLine("Deleting the e-mail...");
                }

                email.Delete(DeleteMode.HardDelete);
            }
            catch (Exception x)
            {
                HandleException("Failed to delete the e-mail", x);
            }
        }

        private string CreateFilePath(DateTime dateTime, string fileName)
        {
            if (timestamp)
            {
                fileName = String.Format("{0}@{1}", (dateTime - new DateTime(1970, 1, 1).ToLocalTime()).TotalSeconds, fileName);
            }

            return Path.Combine(savePath, fileName);
        }

        private void HandleException(string prefix, Exception x)
        {
            if (Environment.UserInteractive)
            {
                Console.WriteLine(prefix + ": " + x);
            }
            else
            {
                this.eventLog.WriteEntry(prefix + ": " + x, EventLogEntryType.Warning);
            }
        }
    }
}
