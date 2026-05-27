using System;
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
        private const string VERSION_URL_PRIMARY = "https://gitee.com/kevin137/qsbar-release/raw/master/version.json";
        private const string VERSION_URL_FALLBACK = "https://raw.githubusercontent.com/pengkang135/QSBar/master/version.json";
        private const int VERSION_REQUEST_TIMEOUT_MS = 3000;
        private const string UPDATE_LOG_FILE = "updated.txt";
        private const string UPDATE_TRACE_FILE = "QSBar.Update.Trace.log";

        public static UpdateInfo LatestUpdateInfo { get; private set; }
        public static bool HasNewVersion { get; private set; }

        static UpdateManager()
        {
            try { ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072; } catch { }
        }

        public class UpdateInfo
        {
            public string Version { get; set; }
            public string DownloadUrl { get; set; }
            public string ChangeLog { get; set; }
        }

        private static string GetTracePath()
        {
            return Path.Combine(Path.GetTempPath(), UPDATE_TRACE_FILE);
        }

        private static void WriteTrace(string message)
        {
            try
            {
                string line = string.Format("{0:yyyy-MM-dd HH:mm:ss.fff} | pid={1} | tid={2} | {3}{4}",
                    DateTime.Now,
                    Process.GetCurrentProcess().Id,
                    Environment.CurrentManagedThreadId,
                    message,
                    Environment.NewLine);
                File.AppendAllText(GetTracePath(), line);
            }
            catch { }
        }

        public static async Task CheckForUpdateAsync(bool silent = false, bool promptOnNewVersion = false)
        {
            try
            {
                WriteTrace(string.Format("CheckForUpdateAsync start silent={0}", silent));

                // 整个网络操作跑在后台线程，绝不碰 UI 线程
                UpdateInfo info = await Task.Run(() => GetLatestVersionInfo()).ConfigureAwait(false);

                LatestUpdateInfo = info;
                Version currentVersion = typeof(UpdateManager).Assembly.GetName().Version;

                if (string.IsNullOrEmpty(info.Version))
                {
                    LatestUpdateInfo = null;
                    HasNewVersion = false;
                    if (!silent) MessageBox.Show("服务器返回的版本信息格式不正确。", "更新错误");
                    return;
                }

                Version latestVersion = new Version(info.Version);
                HasNewVersion = latestVersion > currentVersion;
                WriteTrace(string.Format("Version current={0} latest={1} hasNew={2}", currentVersion, latestVersion, HasNewVersion));

                if (HasNewVersion)
                {
                    if (promptOnNewVersion)
                    {
                        await StartUpdateFlow();
                    }
                    return;
                }
                else
                {
                    if (!silent) MessageBox.Show(string.Format("当前已是最新版本 (v{0})。", currentVersion), "更新检测");
                }
            }
            catch (Exception ex)
            {
                LatestUpdateInfo = null;
                HasNewVersion = false;
                WriteTrace("CheckForUpdateAsync exception: " + ex);
                if (!silent) MessageBox.Show(string.Format("检查更新时出错: {0}", ex.Message), "更新错误");
            }
        }

        private static UpdateInfo GetLatestVersionInfo()
        {
            string[] urls = { VERSION_URL_PRIMARY, VERSION_URL_FALLBACK };
            Exception lastException = null;
            string json = "";

            foreach (string baseUrl in urls)
            {
                try
                {
                    WriteTrace("GetLatestVersionInfo trying " + baseUrl);
                    string url = baseUrl + "?t=" + DateTime.Now.Ticks;

                    using (WebClient client = new WebClient())
                    {
                        client.Proxy = null;
                        client.Encoding = System.Text.Encoding.UTF8;
                        client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
                        client.CachePolicy = new System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore);

                        var downloadTask = client.DownloadStringTaskAsync(url);
                        if (!downloadTask.Wait(VERSION_REQUEST_TIMEOUT_MS))
                        {
                            try { client.CancelAsync(); } catch { }
                            throw new TimeoutException(string.Format("访问更新服务器超时（>{0}ms）。", VERSION_REQUEST_TIMEOUT_MS));
                        }

                        json = downloadTask.Result;
                    }

                    WriteTrace("GetLatestVersionInfo downloaded json length=" + (json == null ? "null" : json.Length.ToString()));

                    if (!string.IsNullOrEmpty(json))
                    {
                        json = json.Trim().Trim('﻿', '​');
                    }

                    if (string.IsNullOrWhiteSpace(json))
                    {
                        throw new Exception("服务器返回了空内容。");
                    }

                    return ParseUpdateInfo(json);
                }
                catch (Exception ex)
                {
                    WriteTrace(string.Format("GetLatestVersionInfo failed for {0}: {1}", baseUrl, ex.Message));
                    lastException = ex;
                    // 立即尝试下一个 URL
                }
            }

            throw new Exception(string.Format("所有更新源均无法访问。最后错误: {0}", lastException?.Message ?? "未知错误"));
        }

        private static UpdateInfo ParseUpdateInfo(string json)
        {
            var info = Newtonsoft.Json.JsonConvert.DeserializeObject<UpdateInfo>(json);
            if (info == null) throw new Exception("JSON 反序列化结果为空。");
            WriteTrace(string.Format("ParseUpdateInfo ok version={0}", info.Version));
            return info;
        }

        public static async Task StartUpdateFlow()
        {
            if (LatestUpdateInfo == null)
            {
                WriteTrace("StartUpdateFlow skipped because LatestUpdateInfo is null");
                return;
            }

            WriteTrace(string.Format("StartUpdateFlow prompt version={0}", LatestUpdateInfo.Version));

            var result = MessageBox.Show(string.Format("将更新到 v{0}。\n\n请确认已保存好当前文件，是否现在关闭 Excel 并执行更新？", LatestUpdateInfo.Version),
                "QS工具箱", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            WriteTrace("StartUpdateFlow dialog result=" + result);

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
                WriteTrace("PerformUpdate begin downloadUrl=" + info.DownloadUrl);

                // 下载也包在 Task.Run 中
                await Task.Run(() =>
                {
                    using (WebClient client = new WebClient())
                    {
                        client.Proxy = null;
                        client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");

                        using (ProgressDialog progress = new ProgressDialog("正在下载更新..."))
                        {
                            progress.Show();
                            WriteTrace("PerformUpdate progress dialog shown");

                            client.DownloadProgressChanged += (s, e) =>
                            {
                                progress.UpdateProgress(e.ProgressPercentage);
                            };

                            var downloadTask = client.DownloadFileTaskAsync(info.DownloadUrl, tempFile);
                            downloadTask.Wait();
                            WriteTrace("PerformUpdate download finished tempFile=" + tempFile);
                        }
                    }
                }).ConfigureAwait(false);

                if (!File.Exists(tempFile) || new FileInfo(tempFile).Length == 0)
                {
                    WriteTrace("PerformUpdate downloaded file missing or empty");
                    MessageBox.Show("下载文件失败或文件为空，请重试。", "更新错误");
                    return;
                }
                WriteTrace("PerformUpdate downloaded file size=" + new FileInfo(tempFile).Length);

                try
                {
                    var downloadedVersion = FileVersionInfo.GetVersionInfo(tempFile).FileVersion;
                    Version vDownloaded = new Version(downloadedVersion);
                    Version vLatest = new Version(info.Version);

                    if (vDownloaded < vLatest)
                    {
                        WriteTrace(string.Format("PerformUpdate version check failed downloaded={0} latest={1}", downloadedVersion, info.Version));
                        MessageBox.Show(string.Format("警告：下载的文件版本 ({0}) 低于目标版本 ({1})。\n更新已取消。", downloadedVersion, info.Version), "更新校验失败");
                        return;
                    }
                    WriteTrace(string.Format("PerformUpdate version check ok downloaded={0} latest={1}", downloadedVersion, info.Version));
                }
                catch (Exception ex)
                {
                    WriteTrace("PerformUpdate version check exception: " + ex);
                    Debug.WriteLine("Version check failed: " + ex.Message);
                }

                string currentPath = Assembly.GetExecutingAssembly().Location;
                string currentDir = Path.GetDirectoryName(currentPath);
                string batchFile = Path.Combine(Path.GetTempPath(), "qs_update.bat");

                string logPath = Path.Combine(currentDir, UPDATE_LOG_FILE);
                File.WriteAllText(logPath, info.ChangeLog);
                WriteTrace("PerformUpdate wrote changelog path=" + logPath);

                int hostProcessId = Process.GetCurrentProcess().Id;
                string tracePath = GetTracePath().Replace("\"", "\"\"");
                string batchContent = $@"@echo off
set ""TRACE_LOG={tracePath}""
>> ""%TRACE_LOG%"" echo [%date% %time%] batch start host_pid={hostProcessId}
timeout /t 1 /nobreak > nul

set RETRIES=0
:KILL_LOOP
taskkill /f /pid {hostProcessId} > nul 2>&1
>> ""%TRACE_LOG%"" echo [%date% %time%] taskkill attempt retries=%RETRIES%

tasklist /fi ""pid eq {hostProcessId}"" | find ""{hostProcessId}"" > nul
if %errorlevel% equ 0 (
    set /a RETRIES+=1
    >> ""%TRACE_LOG%"" echo [%date% %time%] host still alive retries=%RETRIES%
    if %RETRIES% geq 20 goto INSTALL_NOW
    timeout /t 1 /nobreak > nul
    goto KILL_LOOP
)

:INSTALL_NOW
>> ""%TRACE_LOG%"" echo [%date% %time%] enter install phase
timeout /t 1 /nobreak > nul

start /wait """" ""{tempFile}"" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART
>> ""%TRACE_LOG%"" echo [%date% %time%] installer exitcode=%errorlevel%

if %errorlevel% equ 0 (
    >> ""%TRACE_LOG%"" echo [%date% %time%] relaunch excel
    start """" excel.exe
)

del ""{tempFile}""
>> ""%TRACE_LOG%"" echo [%date% %time%] cleanup done
del ""%~f0""";
                File.WriteAllText(batchFile, batchContent, System.Text.Encoding.Default);
                WriteTrace("PerformUpdate batch written path=" + batchFile);

                Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/c \"" + batchFile + "\"",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WindowStyle = ProcessWindowStyle.Hidden
                });
                WriteTrace("PerformUpdate batch launched");
            }
            catch (Exception ex)
            {
                WriteTrace("PerformUpdate exception: " + ex);
                MessageBox.Show(string.Format("更新失败: {0}", ex.Message), "更新错误");
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
            lbl.Text = string.Format("已下载: {0}%", percentage);
        }
    }
}
