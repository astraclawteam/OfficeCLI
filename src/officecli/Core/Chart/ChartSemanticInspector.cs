// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Core;

internal sealed record ChartSemanticFinding(string Subtype, string Message, string Suggestion);

internal static class ChartSemanticInspector
{
    internal static IReadOnlyList<ChartSemanticFinding> Inspect(C.ChartSpace? chartSpace)
    {
        var findings = new List<ChartSemanticFinding>();
        var plotArea = chartSpace?.GetFirstChild<C.Chart>()?.GetFirstChild<C.PlotArea>();
        if (plotArea == null) return findings;
        var visibleText = string.Join(" ", chartSpace!.Descendants<OpenXmlElement>()
            .Where(element => element.LocalName is "t" or "v")
            .Select(element => element.InnerText));
        var series = plotArea.Descendants<OpenXmlCompositeElement>()
            .Where(element => element.LocalName == "ser" && element.Parent?.LocalName.Contains("Chart", StringComparison.OrdinalIgnoreCase) == true)
            .ToList();
        for (var index = 0; index < series.Count; index++)
        {
            var item = series[index];
            var category = item.Elements<OpenXmlCompositeElement>().FirstOrDefault(element => element.LocalName is "cat" or "xVal");
            var values = item.Elements<OpenXmlCompositeElement>().FirstOrDefault(element => element.LocalName is "val" or "yVal" or "bubbleSize");
            var categoryCount = PointCount(category);
            var valueCount = PointCount(values);
            if (categoryCount > 0 && valueCount > 0 && categoryCount != valueCount)
            {
                findings.Add(new(
                    IssueSubtypes.ChartCategorySeriesMismatch,
                    $"Chart series {index + 1} has {categoryCount} category point(s) but {valueCount} value point(s).",
                    "Bind the category axis and series values to ranges with the same number of rows."));
            }
            var text = item.Elements<OpenXmlCompositeElement>().FirstOrDefault(element => element.LocalName == "tx");
            if (series.Count > 1 && !HasTextOrFormula(text))
            {
                findings.Add(new(
                    IssueSubtypes.ChartSeriesNameMissing,
                    $"Chart series {index + 1} has no readable series name.",
                    "Bind the series name to a labelled cell or provide a meaningful literal name."));
            }
            var formatCode = values?.Descendants<OpenXmlElement>()
                .FirstOrDefault(element => element.LocalName == "formatCode")?.InnerText ?? "";
            var expected = ExpectedUnit(formatCode);
            if (expected != null && !ContainsVisibleUnit(visibleText, expected))
            {
                findings.Add(new(
                    IssueSubtypes.ChartUnitMissing,
                    $"Chart series {index + 1} uses {expected} number formatting but the chart title, axis title and series labels do not state that unit.",
                    $"Add {expected} to the value-axis title, chart title or series label so readers cannot misinterpret the scale."));
            }
        }
        return findings;
    }

    private static int PointCount(OpenXmlCompositeElement? data)
    {
        if (data == null) return 0;
        var declared = data.Descendants<OpenXmlElement>().FirstOrDefault(element => element.LocalName == "ptCount")
            ?.GetAttribute("val", "").Value;
        return int.TryParse(declared, out var count)
            ? count
            : data.Descendants<OpenXmlElement>().Count(element => element.LocalName == "pt");
    }

    private static bool HasTextOrFormula(OpenXmlCompositeElement? text) => text != null
        && text.Descendants<OpenXmlElement>().Any(element => element.LocalName is "v" or "f" && !string.IsNullOrWhiteSpace(element.InnerText));

    private static string? ExpectedUnit(string formatCode)
    {
        if (formatCode.Contains('%')) return "百分比（%）";
        if (formatCode.Contains('¥') || formatCode.Contains("CNY", StringComparison.OrdinalIgnoreCase)
            || formatCode.Contains("RMB", StringComparison.OrdinalIgnoreCase) || formatCode.Contains('￥')) return "人民币（CNY/¥）";
        if (formatCode.Contains('$')) return "货币（$）";
        return null;
    }

    private static bool ContainsVisibleUnit(string text, string expected) => expected switch
    {
        "百分比（%）" => text.Contains('%') || text.Contains("百分比", StringComparison.Ordinal) || text.Contains("百分率", StringComparison.Ordinal),
        "人民币（CNY/¥）" => text.Contains('¥') || text.Contains('￥') || text.Contains("CNY", StringComparison.OrdinalIgnoreCase)
            || text.Contains("RMB", StringComparison.OrdinalIgnoreCase) || text.Contains('元'),
        "货币（$）" => text.Contains('$') || text.Contains("USD", StringComparison.OrdinalIgnoreCase),
        _ => false,
    };
}
