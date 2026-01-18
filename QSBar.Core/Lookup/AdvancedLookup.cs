using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;

namespace QSBar.Core.Lookup
{
    public static class AdvancedLookup
    {
        public static object SmartLookupFromArrays(string searchValue, object[,] tableValues, object[,] returnValues, double similarity)
        {
            if (tableValues == null || returnValues == null) return "#N/A";

            int tRows = tableValues.GetLength(0);
            int rRows = returnValues.GetLength(0);
            if (tRows <= 0 || rRows <= 0) return "#N/A";

            int rowCount = Math.Min(tRows, rRows);
            int bestMatch = -1;
            double bestScore = 0;

            string normalizedSearch = NormalizeText(searchValue);

            for (int r = 0; r < rowCount; r++)
            {
                var cell = tableValues[r, 0];
                if (cell == null) continue;
                var currentValue = NormalizeText(cell.ToString());
                if (string.IsNullOrEmpty(currentValue)) continue;

                if (string.Equals(currentValue, normalizedSearch, StringComparison.Ordinal))
                {
                    return returnValues[r, 0] ?? "";
                }

                double score = CalculateSmartSimilarity(normalizedSearch, currentValue);
                if (score > bestScore && score >= similarity)
                {
                    bestScore = score;
                    bestMatch = r;
                    if (score > 0.95)
                    {
                        return returnValues[bestMatch, 0] ?? "";
                    }
                }
            }

            if (bestMatch >= 0) return returnValues[bestMatch, 0] ?? "";
            return "#N/A";
        }

        public static object AdvancedLookupFromArrays(string searchValue, object[,] tableValues, object[,] returnValues, int matchType, double similarity)
        {
            if (tableValues == null || returnValues == null) return "#N/A";

            int tRows = tableValues.GetLength(0);
            int rRows = returnValues.GetLength(0);
            if (tRows <= 0 || rRows <= 0) return "#N/A";

            int rowCount = Math.Min(tRows, rRows);
            int bestMatch = -1;
            double bestScore = 0;

            string normalizedSearch = NormalizeText(searchValue);

            for (int r = 0; r < rowCount; r++)
            {
                var cell = tableValues[r, 0];
                if (cell == null) continue;
                var currentValue = NormalizeText(cell.ToString());
                if (string.IsNullOrEmpty(currentValue)) continue;

                double score = CalculateScore(normalizedSearch, currentValue, matchType);
                if (score > bestScore && score >= similarity)
                {
                    bestScore = score;
                    bestMatch = r;
                    if (score >= 1.0)
                    {
                        return returnValues[bestMatch, 0] ?? "";
                    }
                }
            }

            if (bestMatch >= 0) return returnValues[bestMatch, 0] ?? "";
            return "#N/A";
        }

        public static string NormalizeText(string s)
        {
            if (s == null) return "";
            return s.Trim().ToUpperInvariant();
        }

        public static double CalculateSmartSimilarity(string searchValue, string currentValue)
        {
            if (string.Equals(searchValue, currentValue, StringComparison.Ordinal)) return 1;

            if (currentValue.Contains(searchValue))
            {
                return (double)searchValue.Length / Math.Max(1, currentValue.Length);
            }

            if (searchValue.Contains(currentValue))
            {
                return (double)currentValue.Length / Math.Max(1, searchValue.Length) * 0.9;
            }

            var searchKeywords = ExtractKeywords(searchValue);
            var currentKeywords = ExtractKeywords(currentValue);

            double keywordScore = CalculateKeywordSimilarity(searchKeywords, currentKeywords);
            double charScore = CalculateCharSimilarity(searchValue, currentValue);
            return keywordScore * 0.7 + charScore * 0.3;
        }

        public static string[] ExtractKeywords(string text)
        {
            if (string.IsNullOrEmpty(text)) return Array.Empty<string>();

            var clean = text;
            clean = clean.Replace(';', ' ');
            clean = clean.Replace(',', ' ');
            clean = clean.Replace('.', ' ');
            clean = clean.Replace('-', ' ');
            clean = clean.Replace('_', ' ');
            clean = clean.Replace('/', ' ');
            clean = clean.Replace('(', ' ');
            clean = clean.Replace(')', ' ');

            clean = RemoveUnimportantNumbers(clean);

            var stopWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "THE","AND","OR","TO","FROM","WITH","AS","PER","FOR","OF","IN","ON","AT","BY","IS","ARE","WAS","WERE","BE","BEEN","HAVE","HAS","HAD","A","AN","MM","THICK"
            };

            var parts = clean.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return Array.Empty<string>();

            var result = new List<string>(parts.Length);
            for (int i = 0; i < parts.Length; i++)
            {
                var w = parts[i].Trim();
                if (w.Length <= 2) continue;
                if (stopWords.Contains(w)) continue;
                result.Add(w);
            }

            return result.Count == 0 ? Array.Empty<string>() : result.ToArray();
        }

        private static string RemoveUnimportantNumbers(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";

            var result = text;
            try
            {
                result = Regex.Replace(result, "\\d+(\\.\\d+)?\\s*(MM|CM|M|KG|G|%)", "", RegexOptions.IgnoreCase);
                result = Regex.Replace(result, "\\d+(\\.\\d+)?\\s*THICK", "", RegexOptions.IgnoreCase);
                result = Regex.Replace(result, "\\d+(ST|ND|RD|TH)\\s*FLOOR", "FLOOR", RegexOptions.IgnoreCase);
                result = Regex.Replace(result, "\\s{2,}", " ");
            }
            catch { }
            return result.Trim();
        }

        public static double CalculateKeywordSimilarity(string[] keywords1, string[] keywords2)
        {
            if (keywords1 == null || keywords2 == null) return 0;
            if (keywords1.Length == 0 || keywords2.Length == 0) return 0;

            double matchCount = 0;
            for (int i = 0; i < keywords1.Length; i++)
            {
                for (int j = 0; j < keywords2.Length; j++)
                {
                    if (string.Equals(keywords1[i], keywords2[j], StringComparison.OrdinalIgnoreCase))
                    {
                        matchCount += 1;
                        break;
                    }

                    if (keywords1[i].Length > 3 && keywords2[j].Length > 3)
                    {
                        if (keywords1[i].IndexOf(keywords2[j], StringComparison.OrdinalIgnoreCase) >= 0 || 
                            keywords2[j].IndexOf(keywords1[i], StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            matchCount += 0.7;
                            break;
                        }
                    }
                }
            }

            int totalKeywords = keywords1.Length + keywords2.Length;
            if (totalKeywords <= 0) return 0;
            double score = (2 * matchCount) / totalKeywords;
            return score > 1 ? 1 : score;
        }

        public static double CalculateCharSimilarity(string str1, string str2)
        {
            int len1 = str1?.Length ?? 0;
            int len2 = str2?.Length ?? 0;
            if (len1 == 0 && len2 == 0) return 1;
            if (len1 == 0 || len2 == 0) return 0;

            int minLen = len1 < len2 ? len1 : len2;
            int maxLen = len1 > len2 ? len1 : len2;
            if (maxLen > minLen * 2) return 0.1;

            int commonChars = 0;
            for (int i = 0; i < minLen; i++)
            {
                if (char.ToUpperInvariant(str1[i]) == char.ToUpperInvariant(str2[i])) commonChars++;
            }

            return (double)commonChars / maxLen;
        }

        public static double CalculateScore(string searchValue, string currentValue, int matchType)
        {
            switch (matchType)
            {
                case 0: // Exact
                    return string.Equals(searchValue, currentValue, StringComparison.Ordinal) ? 1 : 0;
                case 2: // Regex/Contain (simulated)
                    if (currentValue.Contains(searchValue))
                        return (double)searchValue.Length / Math.Max(1, currentValue.Length);
                    if (searchValue.Contains(currentValue))
                        return (double)currentValue.Length / Math.Max(1, searchValue.Length) * 0.8;
                    return 0;
                case 3: // StartsWith
                    if (currentValue.Length >= searchValue.Length && currentValue.StartsWith(searchValue, StringComparison.Ordinal))
                        return (double)searchValue.Length / Math.Max(1, currentValue.Length);
                    return 0;
                case 4: // EndsWith
                    if (currentValue.Length >= searchValue.Length && currentValue.EndsWith(searchValue, StringComparison.Ordinal))
                        return (double)searchValue.Length / Math.Max(1, currentValue.Length);
                    return 0;
                case 1: // Wildcard/Fuzzy
                    if (string.Equals(searchValue, currentValue, StringComparison.Ordinal)) return 1;
                    if (currentValue.Contains(searchValue))
                        return (double)searchValue.Length / Math.Max(1, currentValue.Length) * 0.9;
                    if (searchValue.Contains(currentValue))
                        return (double)currentValue.Length / Math.Max(1, searchValue.Length) * 0.7;
                    return 0;
                default:
                    return string.Equals(searchValue, currentValue, StringComparison.Ordinal) ? 1 : 0;
            }
        }
    }
}
