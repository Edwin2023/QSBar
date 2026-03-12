using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace QSBar.Core.Export
{
    public static class ExportRecent
    {
        static string GetPath()
        {
            var baseDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var dir = Path.Combine(baseDir, "QSBar");
            return Path.Combine(dir, "export.history.json");
        }

        public static List<string> Load()
        {
            var path = GetPath();
            try
            {
                if (!File.Exists(path)) return new List<string>();
                var json = File.ReadAllText(path);
                return JsonConvert.DeserializeObject<List<string>>(json) ?? new List<string>();
            }
            catch { return new List<string>(); }
        }

        public static void Save(List<string> list)
        {
            var path = GetPath();
            var dir = Path.GetDirectoryName(path) ?? "";
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            var json = JsonConvert.SerializeObject(list);
            File.WriteAllText(path, json);
        }

        public static void Add(string dir)
        {
            var list = Load();
            list.RemoveAll(d => string.Equals(d, dir, StringComparison.OrdinalIgnoreCase));
            list.Insert(0, dir);
            while (list.Count > 10) list.RemoveAt(list.Count - 1);
            Save(list);
        }

        public static string GetLast()
        {
            var list = Load();
            return list.Count > 0 ? list[0] : null;
        }
    }
}
