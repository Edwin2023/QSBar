using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace QSBar.Core.Formatting
{
    public static class StyleTemplateHistory
    {
        static string GetHistoryPath()
        {
            var baseDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var dir = Path.Combine(baseDir, "QSBar");
            return Path.Combine(dir, "style.history.json");
        }

        public static List<string> Load()
        {
            var path = GetHistoryPath();
            try
            {
                if (!File.Exists(path)) return new List<string>();
                var json = File.ReadAllText(path);
                var list = JsonConvert.DeserializeObject<List<string>>(json) ?? new List<string>();
                return list;
            }
            catch
            {
                return new List<string>();
            }
        }

        public static void Save(List<string> list)
        {
            var path = GetHistoryPath();
            var dir = Path.GetDirectoryName(path) ?? "";
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            var json = JsonConvert.SerializeObject(list);
            File.WriteAllText(path, json);
        }

        public static void Add(string path)
        {
            var list = Load();
            list.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
            list.Insert(0, path);
            while (list.Count > 10) list.RemoveAt(list.Count - 1);
            Save(list);
        }
    }
}
