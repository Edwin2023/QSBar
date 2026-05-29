using System;
using System.Collections.Generic;

namespace QSBar.Core.Formatting
{
    public static class RowStyleProfiles
    {
        private static readonly object Sync = new object();
        private static readonly Dictionary<int, RowStyle> Overrides = new Dictionary<int, RowStyle>();

        public static RowStyle Get(int level)
        {
            lock (Sync)
            {
                RowStyle ov;
                if (Overrides.TryGetValue(level, out ov)) return ov;
                if (level <= 0) return new RowStyle { ColorIndex = null, Bold = false, BorderLeft = false };
                if (level == 1) return new RowStyle { ColorRgb = 198 | (217 << 8) | (241 << 16), Bold = true, BorderLeft = false };
                if (level == 2) return new RowStyle { ColorRgb = 238 | (242 << 8) | (250 << 16), Bold = true, BorderLeft = false };
                return new RowStyle { ColorRgb = 251 | (229 << 8) | (214 << 16), Bold = true, BorderLeft = true };
            }
        }

        public static void SetOverride(int level, RowStyle style)
        {
            lock (Sync)
            {
                Overrides[level] = style;
            }
        }

        public static void ClearOverride(int level)
        {
            lock (Sync)
            {
                if (Overrides.ContainsKey(level)) Overrides.Remove(level);
            }
        }

        public static void ClearAllOverrides()
        {
            lock (Sync)
            {
                Overrides.Clear();
            }
        }

        public static Dictionary<int, RowStyle> GetAllOverrides()
        {
            lock (Sync)
            {
                return new Dictionary<int, RowStyle>(Overrides);
            }
        }

        public static void WithWriteLock(Action action)
        {
            lock (Sync)
            {
                action();
            }
        }
    }
}
