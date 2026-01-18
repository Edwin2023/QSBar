namespace QSBar.Core.Formatting;

public static class PresetStyles
{
    public static Dictionary<int, RowStyle> Classic()
    {
        return new Dictionary<int, RowStyle>
        {
            { 0, new RowStyle { ColorIndex = null, Bold = false, BorderLeft = false } },
            { 1, new RowStyle { ColorIndex = 15, Bold = true, BorderLeft = false } },
            { 2, new RowStyle { ColorIndex = 14, Bold = true, BorderLeft = false } },
            { 3, new RowStyle { ColorIndex = 13, Bold = true, BorderLeft = true } },
        };
    }

    public static Dictionary<int, RowStyle> Light()
    {
        return new Dictionary<int, RowStyle>
        {
            { 0, new RowStyle { ColorIndex = null, Bold = false, BorderLeft = false } },
            { 1, new RowStyle { ColorIndex = 36, Bold = true, BorderLeft = false } },
            { 2, new RowStyle { ColorIndex = 35, Bold = true, BorderLeft = false } },
            { 3, new RowStyle { ColorIndex = 34, Bold = true, BorderLeft = true } },
        };
    }

    public static Dictionary<int, RowStyle> PrintFriendly()
    {
        return new Dictionary<int, RowStyle>
        {
            { 0, new RowStyle { ColorIndex = null, Bold = false, BorderLeft = false } },
            { 1, new RowStyle { ColorIndex = null, Bold = true, BorderLeft = false } },
            { 2, new RowStyle { ColorIndex = null, Bold = true, BorderLeft = true } },
            { 3, new RowStyle { ColorIndex = null, Bold = true, BorderLeft = true } },
        };
    }

    public static void Apply(Dictionary<int, RowStyle> dict)
    {
        RowStyleProfiles.ClearAllOverrides();
        foreach (var kv in dict)
            RowStyleProfiles.SetOverride(kv.Key, kv.Value);
    }
}
