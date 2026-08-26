// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Collections.Generic;
using System.Globalization;

namespace OfficeCli.Core.Diagram;

/// <summary>
/// Format-agnostic visual style for a laid-out diagram: the OOXML preset
/// geometry name + fill/line colors per <see cref="FlowShape"/>, plus the edge
/// color. Shared by the pptx and docx emitters so the two never drift (the
/// geometry strings — "rect", "diamond", "can", … — are valid DrawingML
/// <c>a:prstGeom@prst</c> values usable directly in docx and via
/// <c>TryParsePresetShape</c> in pptx).
/// </summary>
public static class DiagramStyles
{
    public static readonly IReadOnlyDictionary<FlowShape, (string Geometry, string Fill, string Line)> ByShape =
        new Dictionary<FlowShape, (string, string, string)>
        {
            [FlowShape.Process]       = ("rect",          "DAE8FC", "6C8EBF"),
            [FlowShape.Decision]      = ("diamond",       "FFF2CC", "D6B656"),
            [FlowShape.Terminator]    = ("roundRect",     "D5E8D4", "82B366"),
            [FlowShape.Stadium]       = ("roundRect",     "D5E8D4", "82B366"),
            [FlowShape.Circle]        = ("ellipse",       "F8CECC", "B85450"),
            [FlowShape.Hexagon]       = ("hexagon",       "FFF2CC", "D6B656"),
            [FlowShape.Parallelogram] = ("parallelogram", "DAE8FC", "6C8EBF"),
            [FlowShape.Database]      = ("can",           "E1D5E7", "9673A6"),
            [FlowShape.Subroutine]    = ("rect",          "DAE8FC", "6C8EBF"),
            [FlowShape.Flag]          = ("rect",          "DAE8FC", "6C8EBF"),
        };

    /// <summary>Connector / edge stroke color (dark grey).</summary>
    public const string EdgeColor = "4D4D4D";

    public static (string Geometry, string Fill, string Line) Resolve(FlowShape shape, DiagramTheme theme)
    {
        var geometry = ByShape[shape].Geometry;
        return shape switch
        {
            FlowShape.Decision or FlowShape.Hexagon => (geometry, theme.Warning, theme.MutedText),
            FlowShape.Terminator or FlowShape.Stadium => (geometry, theme.Positive, theme.MutedText),
            FlowShape.Circle => (geometry, theme.Danger, theme.MutedText),
            FlowShape.Database => (geometry, theme.Accent, theme.MutedText),
            _ => (geometry, theme.Surface, theme.Primary),
        };
    }

    /// <summary>
    /// Select a deterministic text colour with WCAG AA contrast against the
    /// node fill.  Native Office and SVG renderers share this decision so a
    /// dark branded node cannot silently inherit black host-default text.
    /// </summary>
    public static string TextColorFor(string fill, DiagramTheme theme)
    {
        var preferred = Normalize(theme.Text);
        var background = Normalize(fill);
        if (Contrast(preferred, background) >= 4.5) return preferred;
        return Contrast("FFFFFF", background) >= Contrast("000000", background) ? "FFFFFF" : "000000";
    }

    private static string Normalize(string value)
    {
        var normalized = value.Trim().TrimStart('#');
        return normalized.Length == 6 ? normalized.ToUpperInvariant() : "000000";
    }

    private static double Contrast(string foreground, string background)
    {
        var lighter = Math.Max(RelativeLuminance(foreground), RelativeLuminance(background));
        var darker = Math.Min(RelativeLuminance(foreground), RelativeLuminance(background));
        return (lighter + 0.05) / (darker + 0.05);
    }

    private static double RelativeLuminance(string hex)
    {
        static double Channel(string value)
        {
            var component = int.Parse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255d;
            return component <= 0.04045 ? component / 12.92 : Math.Pow((component + 0.055) / 1.055, 2.4);
        }
        return 0.2126 * Channel(hex[..2]) + 0.7152 * Channel(hex.Substring(2, 2)) + 0.0722 * Channel(hex.Substring(4, 2));
    }
}
