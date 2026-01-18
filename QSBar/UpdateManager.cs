using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace QSBar
{
    public class UpdateManager
    {
        // Gitee 仓库配置
        private const string VERSION_URL = "https://gitee.com/kevin137/qsbar/raw/master/version.json";
        private const string UPDATE_LOG_FILE = "updated.txt";

        public class UpdateInfo
        {
            public string Version { get; set; }
            public string DownloadUrl { get; set; }
            public string ChangeLog { get; set; }
        }

        public static async Task CheckForUpdateAsync(bool silent = false)
        {
            try
            {
                // 确保使用 TLS 1.2
                try { ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072; } catch { }

                UpdateInfo info = await GetLatestVersionInfo();
                if (info == null)
                {
                    if (!silent) MessageBox.Show("无法连接到更新服务器，请检查网络连接或 Gitee 地址是否正确。", "更新检测");
                    return;
                }

                // 使用更可靠的方式获取当前 DLL 的版本
                Version currentVersion = typeof(UpdateManager).Assembly.GetName().Version;
                
                if (string.IsNullOrEmpty(info.Version))
                {
                    if (!silent) MessageBox.Show("服务器返回的版本信息格式不正确。", "更新错误");
                    return;
                }

                Version latestVersion = new Version(info.Version);

                if (latestVersion > currentVersion)
                {
                    var result = MessageBox.Show($"检测到新版本: {info.Version}\n\n当前版本: {currentVersion}\n\n更新内容:\n{info.ChangeLog}\n\n是否立即自动更新？", 
                        "发现新版本", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                    if (result == DialogResult.Yes)
                    {
                        await PerformUpdate(info);
                    }
                }
                else
                {
                    if (!silent) MessageBox.Show($"当前已是最新版本 (v{currentVersion})。", "更新检测");
                }
            }
            catch (Exception ex)
            {
                if (!silent) MessageBox.Show($"检查更新时出错: {ex.Message}", "更新错误");
            }
        }

        private static async Task<UpdateInfo> GetLatestVersionInfo()
        {
            try
            {
                using (WebClient client = new WebClient())
                {
                    client.Encoding = System.Text.Encoding.UTF8;
                    // Gitee 必须设置 User-Agent，否则可能会被拦截
                    client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
                    
                    // 强制禁用缓存，确保获取的是最新 JSON
                    client.CachePolicy = new System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore);
                    
                    string json = await client.DownloadStringTaskAsync(VERSION_URL + "?t=" + DateTime.Now.Ticks);
                    return ParseUpdateInfo(json);
                }
            }
            catch (Exception ex)
            {
                // 可以考虑在这里记录日志或抛出更详细的异常
                throw new Exception($"Download version.json failed: {ex.Message}");
            }
        }

        private static UpdateInfo ParseUpdateInfo(string json)
        {
            try
            {
                return Newtonsoft.Json.JsonConvert.DeserializeObject<UpdateInfo>(json);
            }
            catch { return null; }
        }

        private static async Task PerformUpdate(UpdateInfo info)
        {
            string tempFile = Path.Combine(Path.GetTempPath(), "QSBarUpdate.exe");
            try
            {
                using (WebClient client = new WebClient())
                {
                    // 设置 User-Agent 同样重要
                    client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
                    
                    ProgressDialog progress = new ProgressDialog("正在下载更新...");
                    progress.Show();
                    
                    client.DownloadProgressChanged += (s, e) => {
                        progress.UpdateProgress(e.ProgressPercentage);
                    };

                    await client.DownloadFileTaskAsync(info.DownloadUrl, tempFile);
                    progress.Close();
                }

                // 创建更新脚本/批处理，用于替换当前文件并重启
                string currentPath = Assembly.GetExecutingAssembly().Location;
                string currentDir = Path.GetDirectoryName(currentPath);
                string batchFile = Path.Combine(Path.GetTempPath(), "qs_update.bat");

                // 写入更新后的内容标记，以便下次启动显示
                string logPath = Path.Combine(currentDir, UPDATE_LOG_FILE);
                File.WriteAllText(logPath, info.ChangeLog);

                // 批处理逻辑：等待主程序退出 -> 复制文件 -> 删除临时文件 -> 重新启动 Excel (可选) -> 删除自身
                string batchContent = $@"
@echo off
timeout /t 2 /nobreak > nul
copy /y ""{tempFile}"" ""{currentPath}""
del ""{tempFile}""
start """" ""{Process.GetCurrentProcess().MainModule.FileName}""
del ""%~f0""
";
                File.WriteAllText(batchFile, batchContent, System.Text.Encoding.Default);

                Process.Start(new ProcessStartInfo
                {
                    FileName = batchFile,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WindowStyle = ProcessWindowStyle.Hidden
                });

                // 退出当前进程
                Application.Exit();
            }
            catch (Exception)
            {
                MessageBox.Show("更新失败", "更新错误");
            }
        }

        public static void CheckForUpdateResult()
        {
            try
            {
                string currentDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string logPath = Path.Combine(currentDir, UPDATE_LOG_FILE);

                if (File.Exists(logPath))
                {
                    string changeLog = File.ReadAllText(logPath);
                    File.Delete(logPath);

                    MessageBox.Show($"更新成功！\n\n本次更新内容:\n{changeLog}", "更新完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch { }
        }
    }

    public class ProgressDialog : Form
    {
        private ProgressBar pb;
        private Label lbl;

        public ProgressDialog(string title)
        {
            this.Text = title;
            this.Size = new System.Drawing.Size(300, 120);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            lbl = new Label() { Text = "正在准备下载...", Left = 20, Top = 20, Width = 260 };
            pb = new ProgressBar() { Left = 20, Top = 50, Width = 240, Height = 20, Maximum = 100 };

            this.Controls.Add(lbl);
            this.Controls.Add(pb);
        }

        public void UpdateProgress(int percentage)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateProgress(percentage)));
                return;
            }
            pb.Value = percentage;
            lbl.Text = $"已下载: {percentage}%";
        }
    }
}
