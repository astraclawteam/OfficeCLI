// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Core;

internal static partial class PresentationSemanticInspector
{
    [GeneratedRegex(@"^\s*(?:[•·▪◦‣⁃]|[-–—])\s+", RegexOptions.CultureInvariant)]
    private static partial Regex VisibleBulletPrefix();

    internal static bool HasDuplicateBullet(Drawing.Paragraph paragraph, out string text)
    {
        text = string.Concat(paragraph.Descendants<Drawing.Text>().Select(item => item.Text));
        var properties = paragraph.ParagraphProperties;
        var hasNativeBullet = properties?.ChildElements.Any(element => element.LocalName is "buChar" or "buAutoNum" or "buBlip") == true;
        return hasNativeBullet && VisibleBulletPrefix().IsMatch(text);
    }
}
