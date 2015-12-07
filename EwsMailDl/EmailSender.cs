using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using Microsoft.Exchange.WebServices.Data;
using System.Text;
using System.Text.RegularExpressions;

namespace EwsMailDl
{
    class EmailSender
    {
        private List<string> cache = new List<string>(10);

        private BlockingCollection<string> queue = new BlockingCollection<string>();

        private EventLog eventLog;

        private CancellationTokenSource tokenSource;

        private ExchangeService exchangeService;

        private FileSystemWatcher watcher;

        public EmailSender(EventLog eventLog, CancellationTokenSource tokenSource, Settings settings)
        {
            this.eventLog = eventLog;
            this.tokenSource = tokenSource;
            this.exchangeService = settings.CreateExchangeService();

            watcher = new FileSystemWatcher(settings.InputPath);
            watcher.Filter = "*.email";
            watcher.Created += OnEmailFileCreated;
        }

        public void Run()
        {
            foreach (var file in Directory.GetFiles(watcher.Path, watcher.Filter, SearchOption.TopDirectoryOnly))
            {
                queue.Add(Path.Combine(watcher.Path, file));
            }

            watcher.EnableRaisingEvents = true;

            while (!tokenSource.IsCancellationRequested && !queue.IsAddingCompleted)
            {
                string emailFullPath = null;

                try
                {
                    emailFullPath = queue.Take(tokenSource.Token);
                }
                catch (OperationCanceledException)
                {
                    if (tokenSource.IsCancellationRequested)
                    {
                        if (Environment.UserInteractive)
                        {
                            Console.WriteLine("Sender cancelled!");
                        }

                        break;
                    }
                }

                if (emailFullPath == null)
                {
                    continue;
                }

                if (cache.Contains(emailFullPath))
                {
                    if (Environment.UserInteractive)
                    {
                        Console.WriteLine("Ignoring a duplicate e-mail: {0}", emailFullPath);
                    }

                    continue;
                }

                if (cache.Count == cache.Capacity)
                {
                    cache.RemoveAt(0);
                }

                cache.Add(emailFullPath);

                EmailMessage email = null;

                try
                {
                    email = CreateEmailMessage(emailFullPath);
                }
                catch (Exception x)
                {
                    HandleException("Failed to create an e-mail message from " + emailFullPath, x);
                }

                if (email == null)
                {
                    continue;
                }

                Send(email);
                
                try
                {
                    File.Delete(emailFullPath);
                }
                catch (Exception x)
                {
                    HandleException("Failed to delete an e-mail file " + emailFullPath, x);
                }
            }

            watcher.EnableRaisingEvents = false;
            queue.CompleteAdding();
        }

        private void OnEmailFileCreated(object sender, FileSystemEventArgs e)
        {
            if (!queue.IsAddingCompleted && e.ChangeType == WatcherChangeTypes.Created)
            {
                if (Environment.UserInteractive)
                {
                    Console.WriteLine("Queued e-mail to send: {0}", e.FullPath);
                }

                queue.Add(e.FullPath);
            }
        }

        private EmailMessage CreateEmailMessage(string emailFullPath, int retry = 0)
        {
            string rawEmail = null;

            try
            {
                var fileStream = new FileStream(emailFullPath, FileMode.Open, FileAccess.Read, FileShare.None);
                var rawEmailBuffer = new byte[fileStream.Length];

                fileStream.Read(rawEmailBuffer, 0, rawEmailBuffer.Length);
                fileStream.Close();

                rawEmail = Encoding.UTF8.GetString(rawEmailBuffer);
            }
            catch (Exception x)
            {
                if (retry == 3)
                {
                    HandleException("Failed to read the e-mail file", x);
                    queue.Add(emailFullPath);
                }
                else
                {
                    Thread.Sleep(1000);

                    return CreateEmailMessage(emailFullPath, retry + 1);
                }
            }

            if (string.IsNullOrWhiteSpace(rawEmail))
            {
                return null;
            }

            var bodyIndex = rawEmail.IndexOf("Body:");

            if (bodyIndex == -1)
            {
                return null;
            }

            var headers = ParseHeaders(rawEmail.Substring(0, bodyIndex).Trim());
            var newLineIndex = rawEmail.IndexOf('\n', bodyIndex);
            var body = rawEmail.Substring(newLineIndex + 1).Trim();

            if (headers.Count == 0 || string.IsNullOrWhiteSpace(body))
            {
                return null;
            }

            var htmlBody = headers.ContainsKey("html") && (headers["html"].Equals("1") || headers["html"].Equals("true"));

            if (!htmlBody)
            {
                body = string.Join(
                    "<br>",
                    body.Split('\n').Select(line =>
                    {
                        line = Regex.Replace(line, "^[\r\t]+", "");

                        var matches = Regex.Match(line, "^( +)");

                        line = line.Trim();

                        if (matches.Success)
                        {
                            var spaceCount = matches.Groups[1].Value.Length;

                            for (var i = 0; i < spaceCount; ++i)
                            {
                                line = "&nbsp;" + line;
                            }
                        }

                        return line;
                    })
                );
            }

            var email = new EmailMessage(exchangeService);
            email.Subject = "";
            email.Body = new MessageBody(BodyType.HTML, body);

            foreach (var header in headers)
            {
                ApplyHeader(email, header.Key, header.Value);
            }

            if (string.IsNullOrWhiteSpace(email.Subject) || email.ToRecipients.Count == 0)
            {
                return null;
            }

            return email;
        }

        private IDictionary<string, string> ParseHeaders(string rawHeaders)
        {
            var headers = new Dictionary<string, string>();

            if (string.IsNullOrWhiteSpace(rawHeaders))
            {
                return headers;
            }

            foreach (var headerLine in rawHeaders.Split('\n'))
            {
                var parts = headerLine.Split(new char[] { ':' }, 2);

                if (parts.Length != 2)
                {
                    continue;
                }

                var headerName = parts[0].Trim().ToLower();
                var headerValue = parts[1].Trim();

                headers[headerName] = headerValue;
            }

            return headers;
        }

        private void ApplyHeader(EmailMessage email, string header, string rawValue)
        {
            switch (header)
            {
                case "subject":
                    email.Subject = rawValue;
                    break;

                case "importance":
                    email.Importance = rawValue.Equals("low", StringComparison.InvariantCultureIgnoreCase)
                        ? Importance.Low
                        : rawValue.Equals("high", StringComparison.InvariantCultureIgnoreCase)
                            ? Importance.High
                            : Importance.Normal;
                    break;

                case "to":
                case "torecipients":
                    AddEmailAddresses(email.ToRecipients, rawValue);
                    break;

                case "cc":
                case "ccrecipients":
                    AddEmailAddresses(email.CcRecipients, rawValue);
                    break;

                case "bcc":
                case "bccrecipients":
                    AddEmailAddresses(email.BccRecipients, rawValue);
                    break;

                case "from":
                    email.From = CreateEmailAddress(rawValue);
                    break;

                case "replyto":
                case "reply-to":
                    AddEmailAddresses(email.ReplyTo, rawValue);
                    break;
            }
        }

        private EmailAddress CreateEmailAddress(string rawValue)
        {
            var match = Regex.Match(rawValue, @"^(.*?)\s*(?:<(.*?)>)?$");

            if (!match.Success)
            {
                return null;
            }

            var emailAddress = new EmailAddress();

            if (string.IsNullOrWhiteSpace(match.Groups[2].Value))
            {
                emailAddress.Address = match.Groups[1].Value.Trim();
            }
            else if (match.Groups[2].Value.IndexOf('@') == -1)
            {
                emailAddress.Address = match.Groups[1].Value.Trim();
                emailAddress.Name = match.Groups[2].Value.Trim();
            }
            else
            {
                emailAddress.Address = match.Groups[2].Value.Trim();
                emailAddress.Name = match.Groups[1].Value.Trim();
            }

            return emailAddress;
        }

        private void AddEmailAddresses(EmailAddressCollection emailAddresses, string rawValue)
        {
            foreach (var rawEmailAddress in rawValue.Split(new char[] { ',' }))
            {
                var emailAddress = CreateEmailAddress(rawEmailAddress);

                if (emailAddresses != null)
                {
                    emailAddresses.Add(emailAddress);
                }
            }
        }

        private void Send(EmailMessage email)
        {
            try
            {
                if (Environment.UserInteractive)
                {
                    Console.WriteLine("Sending the e-mail '{0}' to '{1}'...", email.Subject, String.Join<EmailAddress>(", ", email.ToRecipients.ToArray<EmailAddress>()));
                }

                email.Send();
            }
            catch (Exception x)
            {
                HandleException("Failed to send the e-mail", x);
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
