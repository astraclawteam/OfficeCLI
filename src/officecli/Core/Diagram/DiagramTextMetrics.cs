// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;

namespace OfficeCli.Core.Diagram;

/// <summary>
/// One deterministic text metric shared by layout, native Office emitters and SVG.
/// It deliberately uses conservative script-aware advances instead of a machine font
/// API so Chinese layout is reproducible on Windows, macOS and Linux.
/// </summary>
internal static class DiagramTextMetrics
{
    internal const double NodeMaxLineCm = 7.0;
    internal const double NodeLineHeightCm = 0.62;
    internal const double EdgeLineHeightCm = 0.46;

    internal static double WidthCm(string text)
    {
        double width = 0;
        foreach (var rune in text.EnumerateRunes())
        {
            if (Rune.IsWhiteSpace(rune)) width += 0.22;
            else if (IsWide(rune)) width += 0.58;
            else width += 0.30;
        }
        return width;
    }

    internal static IReadOnlyList<string> Wrap(string text, double widthCm)
    {
        var lines = new List<string>();
        foreach (var paragraph in text.Replace("\r", "", StringComparison.Ordinal).Split('\n'))
        {
            if (paragraph.Length == 0)
            {
                lines.Add(string.Empty);
                continue;
            }

            var words = paragraph.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var tokens = new List<string>();
            for (var index = 0; index < words.Length; index++)
            {
                // Keep slash-separated product/file-format names intact at the
                // word boundary: "/ PPT" moves as one token, so WPS never renders
                // the visually broken "P" / "PT" seen with rune-greedy wrapping.
                if (words[index] == "/" && index + 1 < words.Length)
                    tokens.Add("/ " + words[++index]);
                else
                    tokens.Add(words[index]);
            }

            var current = string.Empty;
            foreach (var token in tokens)
            {
                var candidate = current.Length == 0 ? token : current + " " + token;
                if (WidthCm(candidate) <= widthCm)
                {
                    current = candidate;
                    continue;
                }
                if (current.Length > 0)
                {
                    lines.Add(current);
                    current = string.Empty;
                }
                if (WidthCm(token) <= widthCm)
                    current = token;
                else
                {
                    var split = SplitLongToken(token, widthCm);
                    lines.AddRange(split.Take(split.Count - 1));
                    current = split[^1];
                }
            }
            if (current.Length > 0) lines.Add(current);
        }
        if (lines.Count == 0) lines.Add(string.Empty);
        return lines;
    }

    private static List<string> SplitLongToken(string token, double widthCm)
    {
        var result = new List<string>();
        var current = new StringBuilder();
        foreach (var rune in token.EnumerateRunes())
        {
            var candidate = current.ToString() + rune;
            if (current.Length > 0 && WidthCm(candidate) > widthCm)
            {
                result.Add(current.ToString());
                current.Clear();
            }
            current.Append(rune.ToString());
        }
        if (current.Length > 0) result.Add(current.ToString());

        // A one-character CJK tail is visually worse than two balanced lines.
        // Move one complete rune from the previous line when doing so remains
        // within the same deterministic width bound.
        if (result.Count >= 2 && result[^1].EnumerateRunes().Count() == 1)
        {
            var previous = result[^2].EnumerateRunes().ToList();
            if (previous.Count > 2)
            {
                var moved = previous[^1].ToString();
                previous.RemoveAt(previous.Count - 1);
                result[^2] = string.Concat(previous.Select(item => item.ToString()));
                result[^1] = moved + result[^1];
            }
        }
        return result;
    }

    internal static string WrappedText(string text, double widthCm) =>
        string.Join('\n', Wrap(text, widthCm));

    internal static (double Width, int Lines) NodeExtent(string text)
    {
        var wrapped = Wrap(text, NodeMaxLineCm);
        var width = wrapped.Count == 0 ? 0 : wrapped.Max(WidthCm);
        return (Math.Min(width, NodeMaxLineCm), Math.Max(1, wrapped.Count));
    }

    internal static EdgeLabel EdgeLabel(string text, double cx, double cy, bool opaque)
    {
        var width = Math.Clamp(WidthCm(text) + 0.5, 1.0, 5.0);
        var lines = Wrap(text, Math.Max(0.8, width - 0.3)).Count;
        return new EdgeLabel
        {
            Text = text, Cx = cx, Cy = cy, Opaque = opaque,
            W = width, H = Math.Max(0.58, lines * EdgeLineHeightCm + 0.18),
        };
    }

    private static bool IsWide(Rune rune) => rune.Value >= 0x2E80
        || rune.Value is >= 0x1100 and <= 0x11FF
        || rune.Value is >= 0xAC00 and <= 0xD7AF
        || rune.Value is >= 0x1F000 and <= 0x1FAFF;
}
