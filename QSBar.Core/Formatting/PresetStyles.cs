using System.Collections.Generic;

namespace QSBar.Core.Formatting
{
    public static class PresetStyles
    {
        // OLE 颜色 = R | (G << 8) | (B << 16)
        private static readonly int L1_RGB = 198 | (217 << 8) | (241 << 16);   // #C6D9F1
        private static readonly int L2_RGB = 238 | (242 << 8) | (250 << 16);   // #EEF2FA
        private static readonly int L3_RGB = 251 | (229 << 8) | (214 << 16);   // #FBE5D6

        public static Dictionary<int, RowStyle> Classic()
        {
            return new Dictionary<int, RowStyle>
            {
                { 0, new RowStyle { ColorIndex = null, Bold = false, BorderLeft = false } },
                { 1, new RowStyle { ColorRgb = L1_RGB, Bold = true, BorderLeft = false } },
                { 2, new RowStyle { ColorRgb = L2_RGB, Bold = true, BorderLeft = false } },
                { 3, new RowStyle { ColorRgb = L3_RGB, Bold = true, BorderLeft = true } },
            };
        }

        public static Dictionary<int, RowStyle> Light()
        {
            return new Dictionary<int, RowStyle>
            {
                { 0, new RowStyle { ColorIndex = null, Bold = false, BorderLeft = false } },
                { 1, new RowStyle { ColorRgb = L1_RGB, Bold = true, BorderLeft = false } },
                { 2, new RowStyle { ColorRgb = L2_RGB, Bold = true, BorderLeft = false } },
                { 3, new RowStyle { ColorRgb = L3_RGB, Bold = true, BorderLeft = true } },
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
}
