using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.Net;
using System.Net.Security;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.ServiceProcess;
using System.Threading;
using Microsoft.Exchange.WebServices.Data;

namespace EwsMailDl
{
    class Program : ServiceBase
    {
        private CancellationTokenSource tokenSource = null;

        private IList<ItemId> newEmailIdList = null;

        private BlockingCollection<ItemId> emailIdQueue = null;

        private EmailDownloader downloader = null;

        private Thread downloaderThread = null;

        private StreamingSubscriptionConnection subConn = null;

        private string[] programArgs = null;

        private Settings settings = null;

        private FolderId folderId = null;

        private bool downloadingOld = false;

        static int Main(string[] args)
        {
            if (Environment.UserInteractive)
            {
                if (args.Length > 0)
                {
                    if (args[0] == "/i")
                    {
                        return InstallService(args);
                    }

                    if (args[0] == "/u")
                    {
                        return UninstallService();
                    }
                }
                
                new Program().StartMonitor(args);
            }
            else
            {
                ServiceBase.Run(new Program(args));
            }

            return 0;
        }

        private static int InstallService(string[] args)
        {
            var service = new Program();

            try
            {
                var installerArgs = new string[args.Length];

                for (var i = 1; i < args.Length; ++i)
                {
                    installerArgs[i - 1] = args[i];
                }

                installerArgs[args.Length - 1] = Assembly.GetExecutingAssembly().Location;

                ManagedInstallerClass.InstallHelper(installerArgs);
            }
            catch (Exception x)
            {
                if (x.InnerException != null && x.InnerException.GetType() == typeof(Win32Exception))
                {
                    Win32Exception wx = (Win32Exception)x.InnerException;

                    Console.WriteLine("Error 0x{0:X}: Service already installed!", wx.ErrorCode);

                    return wx.ErrorCode;
                }
                
                Console.WriteLine(x.ToString());

                return -1;
            }

            return 0;
        }

        private static int UninstallService()
        {
            var service = new Program();

            try
            {
                ManagedInstallerClass.InstallHelper(new string[] { "/u", Assembly.GetExecutingAssembly().Location });
            }
            catch (Exception x)
            {
                if (x.InnerException.GetType() == typeof(Win32Exception))
                {
                    Win32Exception wx = (Win32Exception)x.InnerException;

                    Console.WriteLine("Error 0x{0:X}: Service not installed!", wx.ErrorCode);

                    return wx.ErrorCode;
                }
                else
                {
                    Console.WriteLine(x.ToString());

                    return -1;
                }
            }

            return 0;
        }

        public Program()
        {
            ServiceName = "EwsMailDl";
            EventLog.Log = "Application";
            CanHandlePowerEvent = false;
            CanHandleSessionChangeEvent = false;
            CanPauseAndContinue = false;
            CanShutdown = false;
            CanStop = true;
        }

        public Program(string[] args) : this()
        {
            programArgs = args;
        }

        protected override void OnStart(string[] serviceArgs)
        {
            new Thread(new ThreadStart(() => StartMonitor(serviceArgs))).Start();
        }

        private void StartMonitor(string[] serviceArgs)
        {
            try
            {
                settings = new Settings();

                if (programArgs != null)
                {
                    settings.ReadFromArgs(programArgs);
                }

                if (serviceArgs.Length > 0)
                {
                    settings.ReadFromArgs(serviceArgs);
                }

                folderId = settings.CreateFolderId();
            }
            catch (Exception x)
            {
                HandleException(x);
                return;
            }

            programArgs = null;
            serviceArgs = null;

            if (Environment.UserInteractive)
            {
                Console.WriteLine("EwsMailDl");
                Console.WriteLine("--");
                Console.WriteLine(settings);
                Console.WriteLine("--");
            }

            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;

            newEmailIdList = new List<ItemId>();
            tokenSource = new CancellationTokenSource();
            emailIdQueue = new BlockingCollection<ItemId>();
            downloader = new EmailDownloader(EventLog, emailIdQueue, tokenSource, settings);

            StreamingSubscription sub = null;

            try
            {
                sub = settings.CreateExchangeService().SubscribeToStreamingNotifications(new FolderId[] { folderId }, EventType.NewMail);
            }
            catch (Exception x)
            {
                HandleException(x);
                return;
            }

            subConn = new StreamingSubscriptionConnection(sub.Service, settings.Lifetime);

            subConn.OnNotificationEvent += OnNotificationEvent;
            subConn.OnSubscriptionError += OnSubscriptionError;
            subConn.OnDisconnect += OnDisconnect;

            subConn.AddSubscription(sub);

            OpenSubscription(subConn);

            downloaderThread = new Thread(downloader.Run);

            try
            {
                DownloadAndDeleteOld();

                downloaderThread.Start();

                if (Environment.UserInteractive)
                {
                    downloaderThread.Join();
                }
            }
            catch (Exception x)
            {
                HandleException(x);
            }
        }

        private void DownloadAndDeleteOld(int nextPageOffset = 0)
        {
            downloadingOld = true;

            if (tokenSource.IsCancellationRequested)
            {
                return;
            }

            var exchangeService = settings.CreateExchangeService();
            var searchFilter = settings.CreateSearchFilter();
            var view = new ItemView(24, nextPageOffset)
            {
                PropertySet = PropertySet.IdOnly
            };
            
            view.OrderBy.Add(EmailMessageSchema.DateTimeReceived, SortDirection.Ascending);

            if (Environment.UserInteractive)
            {
                Console.WriteLine("Searching for {0} old e-mails at offset {1}...", view.PageSize, nextPageOffset);
            }

            var results = exchangeService.FindItems(folderId, searchFilter, view);

            if (Environment.UserInteractive)
            {
                Console.Write("...found {0} old e-mails", results.Items.Count);

                if (results.MoreAvailable)
                {
                    Console.WriteLine(" and more are available ({0} total)...", results.TotalCount);
                }
                else
                {
                    Console.WriteLine("...");
                }
            }

            foreach (var item in results.Items)
            {
                if (item is EmailMessage)
                {
                    emailIdQueue.Add(item.Id);
                }
            }

            if (results.NextPageOffset.HasValue)
            {
                DownloadAndDeleteOld(results.NextPageOffset.Value);
            }
            else
            {
                if (newEmailIdList != null)
                {
                    if (newEmailIdList.Count > 0)
                    {
                        foreach (var emailId in newEmailIdList)
                        {
                            emailIdQueue.Add(emailId);
                        }

                        Console.WriteLine("Enqueued {0} new e-mails!", newEmailIdList.Count);
                    }

                    newEmailIdList = null;
                }

                downloadingOld = false;
            }
        }

        private void HandleException(Exception x)
        {
            if (Environment.UserInteractive)
            {
                Console.WriteLine(x);

                Environment.Exit(1);
            }
            else
            {
                throw x;
            }
        }

        protected override void OnStop()
        {
            if (emailIdQueue != null)
            {
                emailIdQueue.CompleteAdding();
            }

            if (tokenSource != null)
            {
                tokenSource.Cancel(false);
            }

            if (subConn != null && subConn.IsOpen)
            {
                try
                {
                    subConn.Close();
                }
                catch (Exception) { }
            }

            try
            {
                if (downloaderThread != null)
                {
                    downloaderThread.Join(1337);
                }
            }
            catch (Exception) { }
        }

        private void OnNotificationEvent(object sender, NotificationEventArgs args)
        {
            if (emailIdQueue.IsAddingCompleted)
            {
                return;
            }

            foreach (NotificationEvent notificationEvent in args.Events)
            {
                if (notificationEvent.EventType == EventType.NewMail && notificationEvent is ItemEvent)
                {
                    var itemId = (notificationEvent as ItemEvent).ItemId;

                    if (newEmailIdList == null)
                    {
                        emailIdQueue.Add(itemId);
                    }
                    else
                    {
                        newEmailIdList.Add(itemId);
                    }
                }
            }
        }

        private void OnSubscriptionError(object sender, SubscriptionErrorEventArgs args)
        {
            if (args.Exception != null)
            {
                if (Environment.UserInteractive)
                {
                    Console.WriteLine(args.Exception.Message);
                }
                else
                {
                    throw args.Exception;
                }
            }
        }

        private void OnDisconnect(object sender, SubscriptionErrorEventArgs args)
        {
            var message = "Subscription disconnected"
                + (args.Exception == null ? " :(" : (": " + args.Exception.Message));

            if (Environment.UserInteractive)
            {
                Console.WriteLine(message);
            }
            else
            {
                EventLog.WriteEntry(message, EventLogEntryType.Warning);
            }

            OpenSubscription(sender as StreamingSubscriptionConnection);

            if (!downloadingOld)
            {
                DownloadAndDeleteOld();
            }
        }

        private void OpenSubscription(StreamingSubscriptionConnection subConn)
        {
            if (tokenSource.IsCancellationRequested)
            {
                return;
            }

            if (Environment.UserInteractive)
            {
                Console.WriteLine("Connecting the subscription...");
            }

            try
            {
                subConn.Open();
            }
            catch (Exception x)
            {
                HandleException(x);
                return;
            }

            if (subConn.IsOpen)
            {
                var message = "Subscription connected :)";

                if (Environment.UserInteractive)
                {
                    Console.WriteLine(message);
                }
                else
                {
                    EventLog.WriteEntry(message, EventLogEntryType.Information);
                }
            }
        }

        private static bool CertificateValidationCallback(
            object sender,
            X509Certificate certificate,
            X509Chain chain,
            SslPolicyErrors sslPolicyErrors)
        {
            if (sslPolicyErrors == SslPolicyErrors.None)
            {
                return true;
            }

            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) == 0)
            {
                return false;
            }

            if (chain != null && chain.ChainStatus != null)
            {
                foreach (X509ChainStatus status in chain.ChainStatus)
                {
                    if ((certificate.Subject == certificate.Issuer) && (status.Status == X509ChainStatusFlags.UntrustedRoot))
                    {
                        continue;
                    }

                    if (status.Status != X509ChainStatusFlags.NoError)
                    {
                        return false;
                    }
                }
            }

            return true;
        }
    }
}
