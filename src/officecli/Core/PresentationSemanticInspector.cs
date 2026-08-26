// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

internal static partial class PresentationSemanticInspector
{
    [GeneratedRegex(@"^\s*(?:[•·▪◦‣⁃]|[-–—])\s+", RegexOptions.CultureInvariant)]
    private static partial Regex VisibleBulletPrefix();

    internal static bool HasDuplicateBullet(Drawing.Paragraph paragraph, out string text,
                                            Drawing.ListStyle? listStyle = null,
                                            bool inheritsPlaceholderBullet = false)
    {
        text = string.Concat(paragraph.Descendants<Drawing.Text>().Select(item => item.Text));
        var properties = paragraph.ParagraphProperties;
        if (properties?.ChildElements.Any(element => element.LocalName == "buNone") == true) return false;
        var hasNativeBullet = HasBullet(properties);
        if (!hasNativeBullet && listStyle != null)
        {
            var level = Math.Clamp((int)(properties?.Level?.Value ?? 0), 0, 8) + 1;
            var levelProperties = listStyle.ChildElements.FirstOrDefault(element => element.LocalName == $"lvl{level}pPr");
            hasNativeBullet = HasBullet(levelProperties);
        }
        hasNativeBullet |= inheritsPlaceholderBullet;
        return hasNativeBullet && VisibleBulletPrefix().IsMatch(text);
    }

    internal static IReadOnlyList<string> CrossSuiteFontRisks(Drawing.Theme? theme, bool containsCjkText)
    {
        if (!containsCjkText) return [];
        var scheme = theme?.ThemeElements?.FontScheme;
        var names = new[]
        {
            scheme?.MajorFont?.LatinFont?.Typeface?.Value,
            scheme?.MinorFont?.LatinFont?.Typeface?.Value,
        };
        return names.Where(value => !string.IsNullOrWhiteSpace(value)
                && value is "Aptos" or "Aptos Display" or "Calibri" or "Calibri Light")
            .Select(value => value!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Order(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool HasBullet(OpenXmlElement? properties) => properties?.ChildElements
        .Any(element => element.LocalName is "buChar" or "buAutoNum" or "buBlip") == true;
}
