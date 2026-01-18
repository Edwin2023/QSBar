using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace QSBar.Core.Formatting;

public static class StyleTemplateConfig
{
    public static string GetDefaultPath()
    {
        var baseDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var dir = Path.Combine(baseDir, "QSBar");
        return Path.Combine(dir, "style.json");
    }

    public static Dictionary<int, RowStyle> GetDefaultTemplate()
    {
        return new Dictionary<int, RowStyle>
        {
            { 0, new RowStyle { ColorIndex = null, Bold = false, BorderLeft = false } },
            { 1, new RowStyle { ColorIndex = 15, Bold = true, BorderLeft = false } },
            { 2, new RowStyle { ColorIndex = 14, Bold = true, BorderLeft = false } },
            { 3, new RowStyle { ColorIndex = 13, Bold = true, BorderLeft = true } },
        };
    }

    public static void EnsureDefault()
    {
        var path = GetDefaultPath();
        try
        {
            var dir = Path.GetDirectoryName(path) ?? "";
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            if (!File.Exists(path))
            {
                var json = JsonConvert.SerializeObject(GetDefaultTemplate());
                File.WriteAllText(path, json);
            }
        }
        catch
        {
        }
    }

    public static void LoadDefault()
    {
        var path = GetDefaultPath();
        if (!File.Exists(path)) return;
        var json = File.ReadAllText(path);

        Dictionary<int, RowStyle> dict = null;
        try
        {
            dict = JsonConvert.DeserializeObject<Dictionary<int, RowStyle>>(json);
        }
        catch
        {
        }

        if (dict == null)
        {
            try
            {
                var dictStr = JsonConvert.DeserializeObject<Dictionary<string, RowStyle>>(json);
                if (dictStr != null)
                {
                    dict = new Dictionary<int, RowStyle>();
                    foreach (var kv in dictStr)
                    {
                        if (int.TryParse(kv.Key, out int key))
                            dict[key] = kv.Value;
                    }
                }
            }
            catch
            {
            }
        }

        dict ??= new Dictionary<int, RowStyle>();
        RowStyleProfiles.ClearAllOverrides();
        foreach (var kv in dict)
            RowStyleProfiles.SetOverride(kv.Key, kv.Value);
    }

    public static void SaveDefaultFromCurrent()
    {
        var path = GetDefaultPath();
        var dir = Path.GetDirectoryName(path) ?? "";
        if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
        var dict = RowStyleProfiles.GetAllOverrides();
        var json = JsonConvert.SerializeObject(dict);
        File.WriteAllText(path, json);
    }

    public static void LoadFrom(string path)
    {
        if (!File.Exists(path)) return;
        var json = File.ReadAllText(path);

        Dictionary<int, RowStyle> dict = null;
        try
        {
            dict = JsonConvert.DeserializeObject<Dictionary<int, RowStyle>>(json);
        }
        catch
        {
        }

        if (dict == null)
        {
            try
            {
                var dictStr = JsonConvert.DeserializeObject<Dictionary<string, RowStyle>>(json);
                if (dictStr != null)
                {
                    dict = new Dictionary<int, RowStyle>();
                    foreach (var kv in dictStr)
                    {
                        if (int.TryParse(kv.Key, out int key))
                            dict[key] = kv.Value;
                    }
                }
            }
            catch
            {
            }
        }

        dict ??= new Dictionary<int, RowStyle>();
        RowStyleProfiles.ClearAllOverrides();
        foreach (var kv in dict)
            RowStyleProfiles.SetOverride(kv.Key, kv.Value);
    }

    public static void SaveTo(string path)
    {
        try
        {
            var dir = Path.GetDirectoryName(path) ?? "";
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            var dict = RowStyleProfiles.GetAllOverrides();
            var json = JsonConvert.SerializeObject(dict);
            File.WriteAllText(path, json);
        }
        catch
        {
        }
    }
}
