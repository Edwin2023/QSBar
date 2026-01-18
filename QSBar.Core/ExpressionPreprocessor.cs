using System.Text;

namespace QSBar.Core;



public static class ExpressionPreprocessor

{

    // Mapping full-width Chinese characters to their half-width equivalents for expression parsing

    static readonly string[] From = { "【", "】", "＋", "－", "×", "÷", "（", "）", "，", "。", "\n", " " };

    static readonly string[] To = { "[", "]", "+", "-", "*", "/", "(", ")", ",", ".", "", "" };



    public static string Preprocess(string input)

    {

        if (string.IsNullOrEmpty(input)) return string.Empty;

        var s = input;

        for (var n = 0; n < From.Length; n++) s = s.Replace(From[n], To[n]);

        s = RemoveBracketSections(s);

        return s;

    }



    static string RemoveBracketSections(string s)

    {

        var sb = new StringBuilder();

        var skip = 0;

        foreach (var ch in s)

        {

            if (ch == '[') { skip++; continue; }

            if (ch == ']') { if (skip > 0) skip--; continue; }

            if (skip == 0) sb.Append(ch);

        }

        return sb.ToString();

    }

}

