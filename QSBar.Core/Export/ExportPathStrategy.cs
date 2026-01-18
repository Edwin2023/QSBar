using System.IO;

namespace QSBar.Core.Export;

public static class ExportPathStrategy
{
    public static string GetDefaultDirectory()
    {
        var doc = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var dir = Path.Combine(doc, "QSBar", "Reports", DateTime.Now.ToString("yyyy-MM"));
        return dir;
    }

    public static string GetDefaultFilePath(string baseName)
    {
        var dir = GetDefaultDirectory();
        var safeName = string.IsNullOrWhiteSpace(baseName) ? "Report" : baseName.Trim();
        var file = Path.Combine(dir, $"{DateTime.Now:yyyyMMdd_HHmm}_{safeName}.xlsx");
        return file;
    }
}
