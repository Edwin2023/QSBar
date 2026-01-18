public static class TextUtil
{
    public static string Normalize(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
        return s.Trim();
    }
}
