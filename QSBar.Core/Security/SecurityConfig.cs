using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace QSBar.Core.Security;

public static class SecurityConfig
{
    static string GetDefaultPath()
    {
        var baseDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var dir = Path.Combine(baseDir, "QSBar");
        return Path.Combine(dir, "security.json");
    }

    public static void EnsureDefault()
    {
        var path = GetDefaultPath();
        var dir = Path.GetDirectoryName(path) ?? "";
        if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
        if (!File.Exists(path))
        {
            var obj = new SecurityData
            {
                AllowedVersions = new List<string> { "1.0", "2.0" },
                AllowedUsers = new List<string> { Environment.UserName }
            };
            var json = JsonConvert.SerializeObject(obj);
            File.WriteAllText(path, json);
        }
    }

    public static SecurityData Load()
    {
        var path = GetDefaultPath();
        if (!File.Exists(path)) return new SecurityData();
        var json = File.ReadAllText(path);
        return JsonConvert.DeserializeObject<SecurityData>(json) ?? new SecurityData();
    }
}

public class SecurityData
{
    public List<string> AllowedVersions { get; set; } = new List<string>();
    public List<string> AllowedUsers { get; set; } = new List<string>();
}
