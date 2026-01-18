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
        
        // 缓存最新的更新信息
        public static UpdateInfo LatestUpdateInfo { get; private set; }
        public static bool HasNewVersion { get; private set; }

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
                
                LatestUpdateInfo = info;
                Version currentVersion = typeof(UpdateManager).Assembly.GetName().Version;
                
                if (string.IsNullOrEmpty(info.Version))
                {
                    if (!silent) MessageBox.Show("服务器返回的版本信息格式不正确。", "更新错误");
                    return;
                }

                Version latestVersion = new Version(info.Version);
                HasNewVersion = latestVersion > currentVersion;

                if (HasNewVersion)
                {
                    // 无论是否 silent，发现新版本后都刷新 Ribbon 状态
                    // 不再在这里弹窗，统一由 Ribbon 上的按钮触发
                    WpsExcelAddIn.RefreshRibbon();
                    return;
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
            string json = "";
            try
            {
                using (WebClient client = new WebClient())
                {
                    // 禁用代理，防止本地代理软件（如 Clash）未开启导致的连接失败
                    client.Proxy = null;
                    client.Encoding = System.Text.Encoding.UTF8;
                    client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
                    client.CachePolicy = new System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore);
                    
                    json = await client.DownloadStringTaskAsync(VERSION_URL + "?t=" + DateTime.Now.Ticks);
                    
                    // 强力清洗：去除 BOM 字符和前后空白
                    if (!string.IsNullOrEmpty(json))
                    {
                        json = json.Trim().Trim('\uFEFF', '\u200B');
                    }

                    if (string.IsNullOrWhiteSpace(json))
                    {
                        throw new Exception("服务器返回了空内容。");
                    }
                    return ParseUpdateInfo(json);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"获取版本信息失败: {ex.Message}\n\n服务器响应内容: {(json.Length > 100 ? json.Substring(0, 100) : json)}");
            }
        }

        private static UpdateInfo ParseUpdateInfo(string json)
        {
            try
            {
                var info = Newtonsoft.Json.JsonConvert.DeserializeObject<UpdateInfo>(json);
                if (info == null) throw new Exception("JSON 反序列化结果为空。");
                return info;
            }
            catch (Exception ex)
            {
                throw new Exception($"解析 JSON 失败: {ex.Message}\n\n原始 JSON: {json}");
            }
        }

        public static async Task StartUpdateFlow()
        {
            if (LatestUpdateInfo == null) return;

            Version currentVersion = typeof(UpdateManager).Assembly.GetName().Version;
            var result = MessageBox.Show($"检测到新版本: {LatestUpdateInfo.Version}\n" +
                $"当前版本: {currentVersion}\n\n" +
                $"更新内容:\n{LatestUpdateInfo.ChangeLog}\n\n" +
                "更新将尝试自动关闭 Excel/WPS 进程并替换文件。\n是否立即开始？",
                "确认重启更新", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                await PerformUpdate(LatestUpdateInfo);
            }
        }

        private static async Task PerformUpdate(UpdateInfo info)
        {
            string tempFile = Path.Combine(Path.GetTempPath(), "QSBarUpdate.exe");
            try
            {
                using (WebClient client = new WebClient())
                {
                    client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
                    
                    using (ProgressDialog progress = new ProgressDialog("正在下载更新..."))
                    {
                        progress.Show();
                        
                        client.DownloadProgressChanged += (s, e) => {
                            progress.UpdateProgress(e.ProgressPercentage);
                        };

                        await client.DownloadFileTaskAsync(info.DownloadUrl, tempFile);
                    }
                }

                if (!File.Exists(tempFile) || new FileInfo(tempFile).Length == 0)
                {
                    MessageBox.Show("下载文件失败或文件为空，请重试。", "更新错误");
                    return;
                }

                // 检查下载的 DLL 版本是否正确
                try
                {
                    var downloadedVersion = FileVersionInfo.GetVersionInfo(tempFile).FileVersion;
                    Version vDownloaded = new Version(downloadedVersion);
                    Version vLatest = new Version(info.Version);
                    
                    if (vDownloaded < vLatest)
                    {
                        MessageBox.Show($"警告：下载的文件版本 ({downloadedVersion}) 低于目标版本 ({info.Version})。\n这可能是由于 Gitee 缓存或上传文件错误导致的。更新已取消。", "更新校验失败");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    // 如果无法读取版本，记录但不阻断，除非文件确实损坏
                    Debug.WriteLine("Version check failed: " + ex.Message);
                }

                string currentPath = Assembly.GetExecutingAssembly().Location;
                string currentDir = Path.GetDirectoryName(currentPath);
                string batchFile = Path.Combine(Path.GetTempPath(), "qs_update.bat");

                string logPath = Path.Combine(currentDir, UPDATE_LOG_FILE);
                File.WriteAllText(logPath, info.ChangeLog);

                string excelExe = Process.GetCurrentProcess().MainModule.FileName;
                string batchContent = $@"
@echo off
setlocal enabledelayedexpansion
title QSBar Update Script

echo Waiting for Excel/WPS to close...
set /a count=0
:WAIT_LOOP
taskkill /im excel.exe /im wps.exe > nul 2>&1
timeout /t 1 /nobreak > nul
tasklist | findstr /i ""excel.exe wps.exe"" > nul
if %errorlevel% equ 0 (
    set /a count+=1
    if !count! gtr 3 (
        echo Forcing close...
        taskkill /f /im excel.exe /im wps.exe > nul 2>&1
    )
    if !count! lss 10 goto WAIT_LOOP
)

echo Replacing DLL...
copy /y ""{tempFile}"" ""{currentPath}""
if %errorlevel% neq 0 (
    echo Error: Failed to replace DLL.
    pause
    exit
)

if exist ""{tempFile}"" del ""{tempFile}""

echo Restarting Excel...
timeout /t 1 /nobreak > nul
start """" ""{excelExe}""
del ""%~f0""
";
                File.WriteAllText(batchFile, batchContent, System.Text.Encoding.Default);

                Process.Start(new ProcessStartInfo
                {
                    FileName = batchFile,
                    CreateNoWindow = false,
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Normal
                });

                Application.Exit();
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新失败: {ex.Message}", "更新错误");
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
                    // 仅静默删除日志文件，不再弹窗提示
                    File.Delete(logPath);
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
            this.Size = new System.Drawing.Size(400, 150);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.TopMost = true;

            lbl = new Label() { Text = "正在准备下载...", Left = 20, Top = 25, Width = 350, Font = new System.Drawing.Font("微软雅黑", 10) };
            pb = new ProgressBar() { Left = 20, Top = 65, Width = 340, Height = 25, Maximum = 100 };

            this.Controls.Add(lbl);
            this.Controls.Add(pb);
        }

        public void UpdateProgress(int percentage)
        {
            if (this.IsDisposed) return;
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UpdateProgress(percentage)));
                return;
            }
            pb.Value = Math.Min(100, Math.Max(0, percentage));
            lbl.Text = $"已下载: {percentage}%";
        }
    }
}
