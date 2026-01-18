using System;

namespace QSBar.Core.Lookup;

public static class LookupFunctions
{
    public static object VLookupExact(object[,] table, object key, int returnColOffset)
    {
        int rows = table.GetLength(0);
        int cols = table.GetLength(1);
        for (int r = 0; r < rows; r++)
        {
            var v = table[r, 0];
            if (EqualsKey(v, key))
            {
                int c = 0 + returnColOffset;
                if (c < 0 || c >= cols) return "";
                return table[r, c] ?? "";
            }
        }
        return "";
    }

    public static object HLookupExact(object[,] table, object key, int returnRowOffset)
    {
        int rows = table.GetLength(0);
        int cols = table.GetLength(1);
        for (int c = 0; c < cols; c++)
        {
            var v = table[0, c];
            if (EqualsKey(v, key))
            {
                int r = 0 + returnRowOffset;
                if (r < 0 || r >= rows) return "";
                return table[r, c] ?? "";
            }
        }
        return "";
    }

    static bool EqualsKey(object a, object b)
    {
        if (a is double da && b is double db) return da == db;
        var sa = a?.ToString()?.Trim() ?? "";
        var sb = b?.ToString()?.Trim() ?? "";
        return string.Equals(sa, sb, System.StringComparison.OrdinalIgnoreCase);
    }

    public static object[,] VLookupExactMulti(object[,] table, object key, int[] returnColOffsets)
    {
        int rows = table.GetLength(0);
        int cols = table.GetLength(1);
        for (int r = 0; r < rows; r++)
        {
            var v = table[r, 0];
            if (EqualsKey(v, key))
            {
                int n = returnColOffsets.Length;
                var result = new object[1, n];
                for (int i = 0; i < n; i++)
                {
                    int c = 0 + returnColOffsets[i];
                    result[0, i] = (c >= 0 && c < cols) ? table[r, c] ?? "" : "";
                }
                return result;
            }
        }
        return new object[1, returnColOffsets.Length];
    }

    public static object[,] JoinByKey(object[,] left, object[,] right, int rightReturnOffset)
    {
        int lrows = left.GetLength(0);
        int lcols = left.GetLength(1);
        int rrows = right.GetLength(0);
        int rcols = right.GetLength(1);
        var result = new object[lrows, 1];
        for (int lr = 0; lr < lrows; lr++)
        {
            var key = left[lr, 0];
            object val = "";
            for (int rr = 0; rr < rrows; rr++)
            {
                var rk = right[rr, 0];
                if (EqualsKey(key, rk))
                {
                    int c = 0 + rightReturnOffset;
                    val = (c >= 0 && c < rcols) ? right[rr, c] ?? "" : "";
                    break;
                }
            }
            result[lr, 0] = val;
        }
        return result;
    }
}
