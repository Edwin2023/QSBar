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
                    MessageBox.Show("QSBar 插件已成功卸载！", "卸载完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    await form.ExecuteAction(async (report) => await InstallAsync(report));
                    MessageBox.Show("QSBar 插件已成功安装并注册！\n请重新启动 Excel 或 WPS。", "安装完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };

            Application.Run(form);
        }

        static async Task InstallAsync(Action<int, string> report)
        {
            // 1. 结束进程
            report(10, "正在关闭 Office 进程...");
            KillProcesses();

            // 2. 准备安装目录
            report(20, "正在准备安装目录...");
            string installDir = GetInstallDir();
            if (!Directory.Exists(installDir)) Directory.CreateDirectory(installDir);

            // 3. 复制文件
            report(30, "正在复制程序文件...");
            string sourceDir = AppDomain.CurrentDomain.BaseDirectory;
            string dllSource = Path.Combine(sourceDir, "QSBar.dll");
            if (!File.Exists(dllSource))
            {
                // 尝试从项目源码结构中寻找 (适配 VS 调试环境)
                dllSource = Path.Combine(sourceDir, "..", "..", "..", "QSBar", "bin", "Debug", "QSBar.dll");
                if (!File.Exists(dllSource))
                {
                    dllSource = Path.Combine(sourceDir, "payload", "QSBar.dll");
                }
            }

            if (!File.Exists(dllSource))
                throw new FileNotFoundException("找不到 QSBar.dll，请确保它位于安装程序目录或 payload 目录中。");

            string dllDest = Path.Combine(installDir, "QSBar.dll");
            File.Copy(dllSource, dllDest, true);

            // 复制依赖项
            foreach (string file in Directory.GetFiles(Path.GetDirectoryName(dllSource), "*.dll"))
            {
                if (Path.GetFileName(file) == "QSBar.dll") continue;
                File.Copy(file, Path.Combine(installDir, Path.GetFileName(file)), true);
            }

            // 4. COM 注册
            report(50, "正在进行 COM 注册 (可能需要几秒)...");
            await Task.Run(() => RegisterCom(dllDest));

            // 5. Office/WPS 加载项注册表设置
            report(70, "正在配置 Office 注册表...");
            RegisterAddIn();

            // 6. WPS 白名单
            report(80, "正在配置 WPS 白名单...");
            RegisterWpsWhitelist();

            // 7. 设置环境变量
            report(85, "正在设置环境参数...");
            Environment.SetEnvironmentVariable("VSTO_LOGALERTS", "1", EnvironmentVariableTarget.User);

            // 8. 清理 Excel 禁用项
            report(90, "正在清理 Office 禁用项...");
            ClearResiliency();

            report(100, "安装完成！");
        }

        static async Task UninstallAsync(Action<int, string> report)
        {
            report(10, "正在关闭 Office 进程...");
            KillProcesses();

            // 1. 反注册 COM
            report(30, "正在反注册 COM 组件...");
            await Task.Run(() => UnregisterCom());

            // 2. 清除加载项注册表
            report(50, "正在清除注册表设置...");
            UnregisterAddIn();

            // 3. 清除 WPS 白名单
            report(70, "正在清除 WPS 白名单...");
            UnregisterWpsWhitelist();

            // 4. 删除安装目录
            report(90, "正在删除程序文件...");
            string installDir = GetInstallDir();
            if (Directory.Exists(installDir))
            {
                try { Directory.Delete(installDir, true); } catch { }
            }

            report(100, "卸载完成！");
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
            // 使用 RegistrationServices 进行注册 (相当于 RegAsm /codebase)
            // 注意：这需要在管理员权限下运行
            try
            {
                Assembly asm = Assembly.LoadFrom(dllPath);
                RegistrationServices regSvc = new RegistrationServices();
                if (!regSvc.RegisterAssembly(asm, AssemblyRegistrationFlags.SetCodeBase))
                {
                    throw new Exception("COM 注册失败：程序集可能不包含可注册的类。");
                }

                // 额外手动写入 HKCU 注册表，确保 Excel/WPS 可见 (用户级别)
                // 模拟 quick_setup.ps1 中的逻辑
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
                throw new Exception(string.Format("COM 注册过程中出错: {0}", ex.Message), ex);
            }
        }

        static void UnregisterCom()
        {
            // 清理 HKCU 下的自定义 COM 注册
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
                    key.DeleteValue("Manifest", false); // 确保没有旧的 VSTO 残留
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
