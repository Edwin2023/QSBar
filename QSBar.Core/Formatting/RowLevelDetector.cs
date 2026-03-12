using System;
using System.Collections.Generic;

namespace QSBar.Core.Formatting
{
    public struct GroupInfo
    {
        public int Start;
        public int End;
        public int Level;

        public GroupInfo(int start, int end, int level)
        {
            Start = start;
            End = end;
            Level = level;
        }
    }

    struct StackItem
    {
        public int Level;
        public int Start;

        public StackItem(int level, int start)
        {
            Level = level;
            Start = start;
        }
    }

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

        public static List<GroupInfo> ComputeGroups(int[] levels)
        {
            var groups = new List<GroupInfo>();
            var stack = new Stack<StackItem>();
            int prev = levels.Length > 0 ? levels[0] : 0;
            for (int i = 1; i < levels.Length; i++)
            {
                int curr = levels[i];
                if (curr > prev)
                {
                    stack.Push(new StackItem(curr, i));
                }
                else if (curr < prev)
                {
                    while (stack.Count > 0 && stack.Peek().Level > curr)
                    {
                        var top = stack.Pop();
                        groups.Add(new GroupInfo(top.Start, i - 1, top.Level));
                    }
                }
                prev = curr;
            }
            while (stack.Count > 0)
            {
                var top = stack.Pop();
                groups.Add(new GroupInfo(top.Start, levels.Length - 1, top.Level));
            }
            return groups;
        }

        static int ParseLevel(object v)
        {
            if (v == null) return 0;
            if (v is double) return (int)(double)v;
            if (v is int) return (int)v;
            var raw = v.ToString();
            if (string.IsNullOrWhiteSpace(raw)) return 0;
            var sTrim = raw.Trim();
            int n;
            if (int.TryParse(sTrim, out n)) return n;
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
}
