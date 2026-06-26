using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Reflection;
using System.Threading.Tasks;

namespace QSBar.Setup
{
    internal static class Program
    {
        private const string ProgID = "QSBar.WpsAddIn";
        private const string CLSID = "{D8A7F4B2-1234-4A32-B8E5-9F1E8A9C82DF}";
        private const string AddInName = "QSBar";
        private const string FriendlyName = "QSBar (COM)";
        private const string Description = "QSBar COM Add-in for Excel and WPS";

        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool isUninstall = args.Contains("/u") || args.Contains("-u") || args.Contains("--uninstall");

            InstallForm form = new InstallForm(isUninstall);

            form.Shown += async (s, e) =>
            {
                if (isUninstall)
                {
                    await form.ExecuteAction(async (report) => await UninstallAsync(report));
                    MessageBox.Show("QSBar has been uninstalled.", "Uninstall Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    await form.ExecuteAction(async (report) => await InstallAsync(report));
                    MessageBox.Show("QSBar has been installed and registered.\nPlease restart Excel or WPS.", "Install Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };

            Application.Run(form);
        }

        static async Task InstallAsync(Action<int, string> report)
        {
            report(10, "Closing Office processes...");
            KillProcesses();

            report(20, "Preparing installation folder...");
            string installDir = GetInstallDir();
            if (!Directory.Exists(installDir)) Directory.CreateDirectory(installDir);

            report(30, "Copying program files...");
            string sourceDir = AppDomain.CurrentDomain.BaseDirectory;
            string dllSource = Path.Combine(sourceDir, "QSBar.dll");
            if (!File.Exists(dllSource))
            {
                dllSource = Path.Combine(sourceDir, "..", "..", "..", "QSBar", "bin", "Debug", "QSBar.dll");
                if (!File.Exists(dllSource))
                {
                    dllSource = Path.Combine(sourceDir, "payload", "QSBar.dll");
                }
            }

            if (!File.Exists(dllSource))
                throw new FileNotFoundException("QSBar.dll was not found. Make sure it is located in the installer folder or the payload folder.");

            string dllDest = Path.Combine(installDir, "QSBar.dll");
            File.Copy(dllSource, dllDest, true);

            foreach (string file in Directory.GetFiles(Path.GetDirectoryName(dllSource), "*.dll"))
            {
                if (Path.GetFileName(file) == "QSBar.dll") continue;
                File.Copy(file, Path.Combine(installDir, Path.GetFileName(file)), true);
            }

            report(50, "Running COM registration (may take a few seconds)...");
            await Task.Run(() => RegisterCom(dllDest));

            report(70, "Configuring Office registry entries...");
            RegisterAddIn();

            report(80, "Configuring WPS whitelist...");
            RegisterWpsWhitelist();

            report(85, "Setting environment variables...");
            Environment.SetEnvironmentVariable("VSTO_LOGALERTS", "1", EnvironmentVariableTarget.User);

            report(90, "Cleaning Office disabled items...");
            ClearResiliency();

            report(100, "Installation complete.");
        }

        static async Task UninstallAsync(Action<int, string> report)
        {
            report(10, "Closing Office processes...");
            KillProcesses();

            report(30, "Unregistering COM components...");
            await Task.Run(() => UnregisterCom());

            report(50, "Removing registry entries...");
            UnregisterAddIn();

            report(70, "Removing WPS whitelist entries...");
            UnregisterWpsWhitelist();

            report(90, "Deleting program files...");
            string installDir = GetInstallDir();
            if (Directory.Exists(installDir))
            {
                try { Directory.Delete(installDir, true); } catch { }
            }

            report(100, "Uninstall complete.");
        }

        static void KillProcesses()
        {
            string[] procs = { "excel", "wps", "et", "wpp" };
            foreach (var name in procs)
            {
                foreach (var p in Process.GetProcessesByName(name))
                {
                    try { p.Kill(); p.WaitForExit(3000); } catch { }
                }
            }
        }

        static string GetInstallDir()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "QSBar");
        }

        static void RegisterCom(string dllPath)
        {
            try
            {
                Assembly asm = Assembly.LoadFrom(dllPath);
                RegistrationServices regSvc = new RegistrationServices();
                if (!regSvc.RegisterAssembly(asm, AssemblyRegistrationFlags.SetCodeBase))
                {
                    throw new Exception("COM registration failed: the assembly may not contain any registrable types.");
                }

                using (RegistryKey clsidKey = Registry.CurrentUser.CreateSubKey(string.Format(@"Software\Classes\CLSID\{0}", CLSID)))
                {
                    clsidKey.SetValue("", ProgID);
                    using (RegistryKey inproc = clsidKey.CreateSubKey("InprocServer32"))
                    {
                        inproc.SetValue("", @"C:\Windows\System32\mscoree.dll");
                        inproc.SetValue("ThreadingModel", "Both");
                        inproc.SetValue("Class", "QSBar.WpsExcelAddIn");
                        inproc.SetValue("Assembly", asm.FullName);
                        inproc.SetValue("RuntimeVersion", "v4.0.30319");
                        inproc.SetValue("CodeBase", new Uri(dllPath).AbsoluteUri);
                    }
                }

                using (RegistryKey progIdKey = Registry.CurrentUser.CreateSubKey(string.Format(@"Software\Classes\{0}\CLSID", ProgID)))
                {
                    progIdKey.SetValue("", CLSID);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Error during COM registration: {0}", ex.Message), ex);
            }
        }

        static void UnregisterCom()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(string.Format(@"Software\Classes\CLSID\{0}", CLSID), false);
                Registry.CurrentUser.DeleteSubKeyTree(string.Format(@"Software\Classes\{0}", ProgID), false);
            }
            catch { }
        }

        static void RegisterAddIn()
        {
            string[] paths = {
                string.Format(@"Software\Microsoft\Office\Excel\Addins\{0}", ProgID),
                string.Format(@"Software\Kingsoft\Office\ET\Addins\{0}", ProgID),
                string.Format(@"Software\Kingsoft\Office\ET\AddinsData\{0}", ProgID),
                string.Format(@"Software\Kingsoft\Office\WPS\Addins\{0}", ProgID)
            };

            foreach (var path in paths)
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(path))
                {
                    key.SetValue("Description", Description);
                    key.SetValue("FriendlyName", FriendlyName);
                    key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                    key.SetValue("CommandLineSafe", 1, RegistryValueKind.DWord);
                    key.DeleteValue("Manifest", false);
                }
            }
        }

        static void UnregisterAddIn()
        {
            string[] paths = {
                string.Format(@"Software\Microsoft\Office\Excel\Addins\{0}", ProgID),
                string.Format(@"Software\Kingsoft\Office\ET\Addins\{0}", ProgID),
                string.Format(@"Software\Kingsoft\Office\ET\AddinsData\{0}", ProgID),
                string.Format(@"Software\Kingsoft\Office\WPS\Addins\{0}", ProgID)
            };

            foreach (var path in paths)
            {
                try { Registry.CurrentUser.DeleteSubKeyTree(path, false); } catch { }
            }
        }

        static void RegisterWpsWhitelist()
        {
            string[] products = { "ET", "WPS", "Common", "6.0" };
            foreach (var prod in products)
            {
                string path = string.Format(@"Software\Kingsoft\Office\{0}\AddinsWL", prod);
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(path))
                {
                    key.SetValue(ProgID, "");
                }
            }
        }

        static void UnregisterWpsWhitelist()
        {
            string[] products = { "ET", "WPS", "Common", "6.0" };
            foreach (var prod in products)
            {
                try
                {
                    using (RegistryKey key = Registry.CurrentUser.OpenSubKey(string.Format(@"Software\Kingsoft\Office\{0}\AddinsWL", prod), true))
                    {
                        if (key != null) key.DeleteValue(ProgID, false);
                    }
                }
                catch { }
            }
        }

        static void ClearResiliency()
        {
            string[] paths = {
                @"Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems",
                @"Software\Microsoft\Office\15.0\Excel\Resiliency\DisabledItems",
                @"Software\Kingsoft\Office\ET\Resiliency\DisabledItems"
            };

            foreach (var path in paths)
            {
                try
                {
                    using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path, true))
                    {
                        if (key != null)
                        {
                            foreach (var val in key.GetValueNames())
                            {
                                key.DeleteValue(val);
                            }
                        }
                    }
                }
                catch { }
            }
        }
    }
}
