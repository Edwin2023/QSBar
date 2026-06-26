using System;
using System.Net;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.Win32;
using Newtonsoft.Json;

namespace QSBar
{
    public static class AnalyticsHelper
    {
        private const string Endpoint = "https://qsbar-analytics.REPLACE_ME.workers.dev/ping";
        private const string RegistryPath = @"Software\QSBar";
        private const string ClientIdName = "ClientId";
        private static bool _pingSent = false;
        private static readonly object _lock = new object();

        static AnalyticsHelper()
        {
            try { ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072; } catch { }
        }

        public static void Ping()
        {
            lock (_lock)
            {
                if (_pingSent) return;
                _pingSent = true;
            }

            Task.Run(() =>
            {
                try
                {
                    string clientId = GetOrCreateClientId();
                    if (string.IsNullOrEmpty(clientId)) return;

                    var payload = new
                    {
                        client_id = clientId,
                        version = GetVersion(),
                        language = GetLanguage(),
                        app = GetAppName(),
                        os_version = Environment.OSVersion.VersionString
                    };

                    string json = JsonConvert.SerializeObject(payload);

                    using (var client = new WebClient())
                    {
                        client.Proxy = null;
                        client.Encoding = System.Text.Encoding.UTF8;
                        client.Headers.Add("User-Agent", "QSBar/1.0");
                        client.Headers.Add("Content-Type", "application/json");
                        var task = client.UploadStringTaskAsync(Endpoint, "POST", json);
                        if (!task.Wait(5000))
                            client.CancelAsync();
                    }
                }
                catch { }
            });
        }

        private static string GetOrCreateClientId()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RegistryPath))
                {
                    object value = key?.GetValue(ClientIdName);
                    if (value != null)
                        return value.ToString();
                }

                string newId = Guid.NewGuid().ToString("D");
                using (var key = Registry.CurrentUser.CreateSubKey(RegistryPath))
                {
                    key.SetValue(ClientIdName, newId);
                }
                return newId;
            }
            catch { return ""; }
        }

        private static string GetVersion()
        {
            try { return Assembly.GetExecutingAssembly().GetName().Version.ToString(); }
            catch { return "0.0.0.0"; }
        }

        private static string GetLanguage()
        {
            try { return WpsExcelAddIn.UseChineseRibbon ? "zh" : "en"; }
            catch { return "zh"; }
        }

        private static string GetAppName()
        {
            try
            {
                if (WpsExcelAddIn.App != null)
                {
                    string name = WpsExcelAddIn.App.Name ?? "";
                    if (name.Contains("WPS")) return "WPS";
                }
                return "Excel";
            }
            catch { return "Excel"; }
        }
    }
}
