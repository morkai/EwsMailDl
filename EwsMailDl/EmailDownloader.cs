using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using Microsoft.Exchange.WebServices.Data;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using System.Text;

namespace EwsMailDl
{
    class EmailDownloader
    {
        private int emailCounter = 0;

        private EventLog eventLog;

        private List<string> emailIdCache = new List<string>(10);

        private CancellationTokenSource tokenSource;

        private BlockingCollection<ItemId> emailIdQueue;

        private ExchangeService exchangeService;

        private string savePath;

        private IList<string> subjectFilters;

        private bool timestamp;

        private bool body;

        private DeleteMode delete;

        public EmailDownloader(EventLog eventLog, BlockingCollection<ItemId> emailIdQueue, CancellationTokenSource tokenSource, Settings settings)
        {
            this.eventLog = eventLog;
            this.tokenSource = tokenSource;
            this.emailIdQueue = emailIdQueue;
            this.exchangeService = settings.CreateExchangeService();
            this.savePath = settings.SavePath;
            this.subjectFilters = settings.SubjectFilters;
            this.timestamp = settings.Timestamp;
            this.body = settings.Body;
            this.delete = settings.Delete;
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

                if (emailId == null)
                {
                    continue;
                }

                WriteLine("Processing e-mail: {0}", emailId.UniqueId);

                if (emailIdCache.Contains(emailId.UniqueId))
                {
                    WriteLine("\tignored duplicate.", emailId);

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
                            EmailMessageSchema.From,
                            EmailMessageSchema.ToRecipients,
                            EmailMessageSchema.CcRecipients,
                            EmailMessageSchema.BccRecipients,
                            EmailMessageSchema.Body,
                            EmailMessageSchema.Attachments,
                            EmailMessageSchema.DateTimeReceived
                        )
                    );
                }
                catch (Exception x)
                {
                    HandleException("Failed to bind e-mail", x);
                }

                if (email == null)
                {
                    continue;
                }

                if (MatchEmail(email))
                {
                    WriteLine("\tmatched: {0}", email.Subject);

                    DownloadAndDelete(email);

                    WriteLine("\tdone.");
                }
                else if (email != null)
                {
                    Console.WriteLine("\tignored: {0}", email.Subject);
                }
            }
        }

        private void WriteLine(string format, params object[] parameters)
        {
            if (Environment.UserInteractive)
            {
                Console.WriteLine(format, parameters);
            }
        }

        private bool MatchEmail(EmailMessage email)
        {
            return (body || email.Attachments.Count > 0)
                && (subjectFilters.Count == 0 || subjectFilters.Any(phrase => email.Subject.ToLower().Contains(phrase.ToLower())));
        }

        private void DownloadAndDelete(EmailMessage email)
        {
            ++emailCounter;

            var emailId = String.Format("{0}@EMAIL_{1}", GetUnixTimestamp(DateTime.Now), emailCounter);

            if (body)
            {
                CreateEmailDirectory(emailId);
            }

            DownloadAttachments(body ? emailId : null, email);

            if (body)
            {
                DownloadEmail(emailId, email);
            }

            DeleteEmail(email);
        }

        private void CreateEmailDirectory(string emailId)
        {
            if (Environment.UserInteractive)
            {
                Console.WriteLine("\tcreating e-mail directory: {0}", emailId);
            }

            try
            {
                Directory.CreateDirectory(Path.Combine(savePath, emailId));
            }
            catch (Exception x)
            {
                HandleException("Failed to create e-mail directory", x);
            }
        }

        private string CreateFilePath(string emailId, DateTime dateTime, string fileName)
        {
            if (timestamp)
            {
                fileName = String.Format("{0}@{1}", GetUnixTimestamp(dateTime), fileName);
            }

            if (emailId == null)
            {
                return Path.Combine(savePath, fileName);
            }

            return Path.Combine(savePath, emailId, fileName);
        }

        private Int32 GetUnixTimestamp(DateTime dateTime)
        {
            return (Int32)dateTime.Subtract(new DateTime(1970, 1, 1)).TotalSeconds;
        }

        private void DownloadAttachments(string emailId, EmailMessage email)
        {
            if (email.Attachments.Count > 0)
            {
                WriteLine("\tdownloading attachments...");
            }

            foreach (var attachment in email.Attachments)
            {
                if (!(attachment is FileAttachment))
                {
                    continue;
                }

                var fileAttachment = attachment as FileAttachment;

                WriteLine("\t\t{0}", fileAttachment.Name);

                try
                {
                    var filePath = CreateFilePath(emailId, email.DateTimeReceived, fileAttachment.Name);

                    fileAttachment.Load(filePath);
                    File.SetCreationTime(filePath, email.DateTimeReceived);
                }
                catch (Exception x)
                {
                    HandleException("Failed to download attachment", x);
                }
            }
        }

        private void DownloadEmail(string emailId, EmailMessage email)
        {
            WriteLine("\tdownloading e-mail...");

            try
            {
                var sb = new StringBuilder();
                var sw = new StringWriter(sb);

                using (var writer = new JsonTextWriter(sw))
                {
                    writer.Formatting = Formatting.Indented;

                    writer.WriteStartObject();
                    writer.WritePropertyName("id");
                    writer.WriteValue(email.Id.UniqueId.ToString());
                    writer.WritePropertyName("receivedAt");
                    writer.WriteValue(email.DateTimeReceived);
                    writer.WritePropertyName("subject");
                    writer.WriteValue(email.Subject);
                    writer.WritePropertyName("from");
                    writer.WriteValue(email.From.ToString());
                    writer.WritePropertyName("to");
                    writer.WriteStartArray();
                    foreach (var recipient in email.ToRecipients) writer.WriteValue(recipient.ToString());
                    writer.WriteEndArray();
                    writer.WritePropertyName("cc");
                    writer.WriteStartArray();
                    foreach (var recipient in email.CcRecipients) writer.WriteValue(recipient.ToString());
                    writer.WriteEndArray();
                    writer.WritePropertyName("bcc");
                    writer.WriteStartArray();
                    foreach (var recipient in email.BccRecipients) writer.WriteValue(recipient.ToString());
                    writer.WriteEndArray();
                    writer.WritePropertyName("body");
                    writer.WriteValue(email.Body.ToString());
                    writer.WritePropertyName("attachments");
                    writer.WriteStartArray();
                    foreach (var attachment in email.Attachments)
                    {
                        writer.WriteStartObject();
                        writer.WritePropertyName("id");
                        writer.WriteValue(attachment.Id);
                        writer.WritePropertyName("contentId");
                        writer.WriteValue(attachment.ContentId);
                        writer.WritePropertyName("contentLocation");
                        writer.WriteValue(attachment.ContentLocation);
                        writer.WritePropertyName("contentType");
                        writer.WriteValue(attachment.ContentType);
                        writer.WritePropertyName("name");
                        writer.WriteValue(attachment.Name);
                        writer.WritePropertyName("size");
                        writer.WriteValue(attachment.Size);
                        writer.WriteEndObject();
                    }
                    writer.WriteEndArray();
                    writer.WriteEndObject();
                }

                File.WriteAllText(Path.Combine(savePath, emailId, "email.json"), sb.ToString());
            }
            catch (Exception x)
            {
                HandleException("Failed to download e-mail", x);
            }
        }

        private void DeleteEmail(EmailMessage email)
        {
            try
            {
                Console.WriteLine("\tdeleting e-mail...");

                email.Delete(delete);
            }
            catch (Exception x)
            {
                HandleException("Failed to delete e-mail", x);
            }
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
