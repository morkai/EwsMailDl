using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;
using System.Text;
using Microsoft.Win32;

namespace EwsMailDl
{
    [RunInstaller(true)]
    public class ProgramInstaller : Installer
    {
        private ServiceProcessInstaller processInstaller;
        private ServiceInstaller serviceInstaller;

        public ProgramInstaller()
        {
            processInstaller = new ServiceProcessInstaller();
            serviceInstaller = new ServiceInstaller();

            processInstaller.Account = ServiceAccount.LocalSystem;
            serviceInstaller.StartType = ServiceStartMode.Automatic;
            serviceInstaller.ServiceName = "EwsMailDl";
            serviceInstaller.DisplayName = "EwsMailDl";
            serviceInstaller.Description = "Monitors and downloads matching e-mail attachments arriving to the specified Exchange account.";
            
            Installers.Add(serviceInstaller);
            Installers.Add(processInstaller);

            processInstaller.AfterInstall += OnAfterInstall;
        }

        protected void OnAfterInstall(object sender, InstallEventArgs args)
        {
            var cmd = new StringBuilder();

            foreach (string key in Context.Parameters.Keys)
            {
                if (key == "logtoconsole" || key == "assemblypath" || key == "logfile")
                {
                    continue;
                }
                
                cmd.AppendFormat(" /{0}=\"{1}\"", key, Context.Parameters[key]);
            }

            var keyName = "HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\services\\" + serviceInstaller.ServiceName;
            var valueName = "ImagePath";
            var value = Registry.GetValue(keyName, "ImagePath", null);

            if (value != null)
            {
                Registry.SetValue(keyName, valueName, (value as string) + cmd.ToString(), RegistryValueKind.ExpandString);
            }
        }
    }  
}
