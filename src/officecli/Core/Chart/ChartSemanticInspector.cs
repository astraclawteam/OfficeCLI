// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using System.Globalization;
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
            var seriesName = SeriesName(item);
            if (series.Count > 1 && (!HasTextOrFormula(text) || IsGenericSeriesName(seriesName)))
            {
                findings.Add(new(
                    IssueSubtypes.ChartSeriesNameMissing,
                    IsGenericSeriesName(seriesName)
                        ? $"Chart series {index + 1} uses the generic name {seriesName}."
                        : $"Chart series {index + 1} has no readable series name.",
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
        var valueAxis = plotArea.GetFirstChild<C.ValueAxis>();
        var axisFormat = valueAxis?.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value ?? "";
        var maximumMagnitude = series.Select(MaxAbsoluteValue).DefaultIfEmpty(0d).Max();
        if (maximumMagnitude >= 10_000_000d && !UsesCompactAxisLabels(valueAxis, axisFormat))
        {
            findings.Add(new(
                IssueSubtypes.ChartAxisLabelDensity,
                $"The value axis renders values up to {maximumMagnitude.ToString("0", CultureInfo.InvariantCulture)} without compact display units, which can clip or crowd multi-digit labels.",
                "Use a millions/billions display unit or a scaled axis number format such as 0.0,,\"M\", and state the unit in the chart or axis title."));
        }
        var seriesNames = series.Select(SeriesName).Where(value => value.Length > 0).ToList();
        var categories = series.Select(SeriesCategories).FirstOrDefault(values => values.Count > 0) ?? new List<string>();
        if (seriesNames.Count >= 2 && categories.Count >= 2
            && Majority(seriesNames, IsTimeCategory) && Majority(categories, IsBusinessMeasure))
        {
            findings.Add(new(
                IssueSubtypes.ChartAxisSeriesSemantics,
                "Chart series look like time periods while the category axis looks like business measures; the category and series roles are probably transposed.",
                "Use time periods on the category axis and business measures (such as revenue, cost, budget or profit) as named series, then verify the source ranges."));
        }
        if (series.Any(item => MixedTimeAndMeasure(SeriesCategories(item)) || MixedTimeAndMeasure(SeriesValues(item))))
        {
            findings.Add(new(
                IssueSubtypes.ChartAxisSeriesSemantics,
                "A chart category or value series mixes year-like values with large measure values; a header cell was probably included in the data range or the ranges are transposed.",
                "Bind the category axis only to labels such as years or months, bind each value series only to comparable measures, and verify the source range excludes header cells."));
        }
        return findings;
    }

    private static string SeriesName(OpenXmlCompositeElement series)
        => series.Elements<OpenXmlCompositeElement>().FirstOrDefault(element => element.LocalName == "tx")?
            .Descendants<OpenXmlElement>().FirstOrDefault(element => element.LocalName == "v")?.InnerText.Trim() ?? "";

    private static List<string> SeriesCategories(OpenXmlCompositeElement series)
    {
        var data = series.Elements<OpenXmlCompositeElement>().FirstOrDefault(element => element.LocalName is "cat" or "xVal");
        if (data == null) return new();
        return data.Descendants<OpenXmlElement>()
            .Where(element => element.LocalName == "pt")
            .Select(point => point.Descendants<OpenXmlElement>().FirstOrDefault(element => element.LocalName == "v")?.InnerText.Trim() ?? "")
            .Where(value => value.Length > 0).ToList();
    }

    private static List<string> SeriesValues(OpenXmlCompositeElement series)
    {
        var data = series.Elements<OpenXmlCompositeElement>()
            .FirstOrDefault(element => element.LocalName is "val" or "yVal" or "bubbleSize");
        if (data == null) return new();
        return data.Descendants<OpenXmlElement>()
            .Where(element => element.LocalName == "pt")
            .Select(point => point.Descendants<OpenXmlElement>().FirstOrDefault(element => element.LocalName == "v")?.InnerText.Trim() ?? "")
            .Where(value => value.Length > 0).ToList();
    }

    private static bool IsGenericSeriesName(string value)
        => System.Text.RegularExpressions.Regex.IsMatch(value.Trim(), @"^(?:Series|系列)\s*\d+$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant);

    private static bool MixedTimeAndMeasure(IReadOnlyList<string> values)
        => values.Any(IsNumericYear) && values.Any(IsLargeNumericMeasure);

    private static bool IsNumericYear(string value)
        => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            && number >= 1900 && number <= 2200 && Math.Abs(number - Math.Round(number)) < 1e-9;

    private static bool IsLargeNumericMeasure(string value)
        => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            && Math.Abs(number) >= 10_000;

    private static bool Majority(IReadOnlyList<string> values, Func<string, bool> predicate)
        => values.Count > 0 && values.Count(predicate) * 2 > values.Count;

    private static bool IsTimeCategory(string value)
    {
        var normalized = value.Trim().ToLowerInvariant();
        return System.Text.RegularExpressions.Regex.IsMatch(normalized,
            @"^(?:20\d{2}(?:[-/.年](?:0?[1-9]|1[0-2])(?:月)?)?|q[1-4]|[一二三四1-4]季度|(?:0?[1-9]|1[0-2])月|jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)$");
    }

    private static bool IsBusinessMeasure(string value)
    {
        var normalized = value.Trim().ToLowerInvariant();
        string[] terms = ["收入", "营收", "预算", "成本", "利润", "毛利", "费用", "金额", "销量", "订单", "实际", "目标", "完成率",
            "revenue", "sales", "budget", "cost", "profit", "margin", "expense", "actual", "target", "amount", "rate"];
        return terms.Any(normalized.Contains);
    }

    private static double MaxAbsoluteValue(OpenXmlCompositeElement series)
    {
        var values = series.Elements<OpenXmlCompositeElement>()
            .FirstOrDefault(element => element.LocalName is "val" or "yVal" or "bubbleSize");
        return values?.Descendants<OpenXmlElement>()
            .Where(element => element.LocalName is "v" && double.TryParse(element.InnerText, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
            .Select(element => double.TryParse(element.InnerText, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) ? Math.Abs(value) : 0d)
            .DefaultIfEmpty(0d)
            .Max() ?? 0d;
    }

    private static bool UsesCompactAxisLabels(C.ValueAxis? axis, string formatCode)
        => axis?.GetFirstChild<C.DisplayUnits>() != null
            || formatCode.Contains(",,", StringComparison.Ordinal)
            || formatCode.Contains('万')
            || formatCode.Contains('亿')
            || formatCode.Contains("M", StringComparison.OrdinalIgnoreCase)
            || formatCode.Contains("B", StringComparison.OrdinalIgnoreCase);

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
