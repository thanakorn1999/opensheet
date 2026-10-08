using System.Globalization;
using System.Text;

namespace OpenSheet.Core;

/// <summary>Reads and writes CSV the way Excel does: quoted fields, "" escapes, line breaks inside quotes.</summary>
static class Csv
{
    public static bool IsCsv(string path) => System.IO.Path.GetExtension(path).Equals(".csv", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// File text: UTF-8 (with or without BOM). Anything else is a legacy Excel export, read with the
    /// Windows code page for the user's language (Thai → 874, otherwise Western 1252).
    /// </summary>
    // ponytail: two legacy code pages cover this app's users; add an encoding picker if others show up.
    public static string ReadText(string path)
    {
        var bytes = File.ReadAllBytes(path);
        try
        {
            return new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes).TrimStart('﻿');
        }
        catch (DecoderFallbackException)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            int codePage = CultureInfo.CurrentCulture.TwoLetterISOLanguageName == "th" ? 874 : 1252;
            return Encoding.GetEncoding(codePage).GetString(bytes);
        }
    }

    /// <summary>The separator used in the first line: comma, semicolon (European Excel) or tab, whichever appears most.</summary>
    public static char DetectDelimiter(string text)
    {
        char best = ',';
        int bestCount = 0;
        foreach (var d in new[] { ',', ';', '\t' })
        {
            int count = 0;
            bool quoted = false;
            foreach (var ch in text)
            {
                if (ch == '"') quoted = !quoted;
                else if (!quoted && ch is '\n' or '\r') break;
                else if (!quoted && ch == d) count++;
            }
            if (count > bestCount) (best, bestCount) = (d, count);
        }
        return best;
    }

    public static List<List<string>> Parse(string text, char delimiter)
    {
        var rows = new List<List<string>>();
        var row = new List<string>();
        var field = new StringBuilder();
        bool quoted = false;
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (quoted)
            {
                if (ch != '"') field.Append(ch);
                else if (i + 1 < text.Length && text[i + 1] == '"') field.Append(text[++i]); // "" inside quotes
                else quoted = false;
            }
            else if (ch == '"') quoted = true;
            else if (ch == delimiter)
            {
                row.Add(field.ToString());
                field.Clear();
            }
            else if (ch is '\r' or '\n')
            {
                if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                row.Add(field.ToString());
                field.Clear();
                rows.Add(row);
                row = [];
            }
            else field.Append(ch);
        }
        if (field.Length > 0 || row.Count > 0) // last line without a trailing newline
        {
            row.Add(field.ToString());
            rows.Add(row);
        }
        return rows;
    }

    /// <summary>
    /// Writes rows as CSV in UTF-8 with a BOM. Excel needs the BOM to read Thai and other non-English text correctly.
    /// </summary>
    public static void Write(string path, IEnumerable<IEnumerable<string>> rows, char delimiter)
    {
        var sb = new StringBuilder();
        foreach (var row in rows)
        {
            sb.AppendJoin(delimiter, row.Select(f =>
                f.IndexOfAny([delimiter, '"', '\r', '\n']) >= 0 ? "\"" + f.Replace("\"", "\"\"") + "\"" : f));
            sb.Append("\r\n");
        }
        File.WriteAllText(path, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
    }
}
