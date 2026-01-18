using System;
using System.Collections.Generic;

namespace QSBar.Core.Formatting;

public static class RowLevelDetector
{
    public static int[] ParseLevels(object[] values)
    {
        var levels = new int[values.Length];
        for (int i = 0; i < values.Length; i++)
        {
            levels[i] = ParseLevel(values[i]);
        }
        return levels;
    }

    public static int[] ParseLevels(object[,] data, int columnIndexZeroBased)
    {
        int rows = data.GetLength(0);
        var levels = new int[rows];
        for (int r = 0; r < rows; r++)
        {
            object v = null;
            int cols = data.GetLength(1);
            if (columnIndexZeroBased >= 0 && columnIndexZeroBased < cols)
                v = data[r, columnIndexZeroBased];
            levels[r] = ParseLevel(v);
        }
        return levels;
    }

    public static List<(int start, int end, int level)> ComputeGroups(int[] levels)
    {
        var groups = new List<(int start, int end, int level)>();
        var stack = new Stack<(int level, int start)>();
        int prev = levels.Length > 0 ? levels[0] : 0;
        for (int i = 1; i < levels.Length; i++)
        {
            int curr = levels[i];
            if (curr > prev)
            {
                stack.Push((curr, i));
            }
            else if (curr < prev)
            {
                while (stack.Count > 0 && stack.Peek().level > curr)
                {
                    var top = stack.Pop();
                    groups.Add((top.start, i - 1, top.level));
                }
            }
            prev = curr;
        }
        while (stack.Count > 0)
        {
            var top = stack.Pop();
            groups.Add((top.start, levels.Length - 1, top.level));
        }
        return groups;
    }

    static int ParseLevel(object v)
    {
        if (v == null) return 0;
        if (v is double d) return (int)d;
        if (v is int i) return i;
        var raw = v.ToString();
        if (string.IsNullOrWhiteSpace(raw)) return 0;
        var sTrim = raw.Trim();
        if (int.TryParse(sTrim, out var n)) return n;
        var s = raw;
        int leadingSpaces = 0;
        int prefixCount = 0;
        for (int idx = 0; idx < s.Length; idx++)
        {
            var ch = s[idx];
            if (ch == ' ') { leadingSpaces++; continue; }
            if (ch == '\t') { leadingSpaces += 2; continue; }
            if (ch == '-' || ch == '>' || ch == '?') { prefixCount++; continue; }
            break;
        }
        int levelFromIndent = leadingSpaces / 2;
        int level = Math.Max(levelFromIndent, prefixCount);
        if (level == 0) return 0;
        return level;
    }
}
