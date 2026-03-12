using System;
using System.IO;

namespace QSBar.Core.Export
{
    public static class ExportPathStrategy
    {
        public static string GetDefaultDirectory()
        {
            var doc = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var dir = Path.Combine(doc, "QSBar");
            dir = Path.Combine(dir, "Reports");
            dir = Path.Combine(dir, DateTime.Now.ToString("yyyy-MM"));
            return dir;
        }

        public static string GetDefaultFilePath(string baseName)
        {
            var dir = GetDefaultDirectory();
            var safeName = string.IsNullOrWhiteSpace(baseName) ? "Report" : baseName.Trim();
            var fileName = string.Format("{0:yyyyMMdd_HHmm}_{1}.xlsx", DateTime.Now, safeName);
            var file = Path.Combine(dir, fileName);
            return file;
        }
    }
}
