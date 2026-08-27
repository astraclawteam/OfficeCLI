using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Tests;

public class ChartSemanticInspectorTests
{
    [Fact]
    public void PercentageChartBuildsPercentageAxisAndDeclaredThresholdLine()
    {
        using var temp = new TempDirectory();
        var path = temp.File("percentage-threshold.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Sheet1");
        using (var handler = new ExcelHandler(path, editable: true))
        {
            handler.Add("/Sheet1", "chart", null, new Dictionary<string, string>
            {
                ["chartType"] = "line",
                ["title"] = "错误回答率（扩大门槛 0.5%）",
                ["categories"] = "第1周,第2周",
                ["series1"] = "错误回答率:0.014,0.006",
                ["series1.valuesNumFmt"] = "0.0%",
            });
        }

        using var document = SpreadsheetDocument.Open(path, false);
        var chartSpace = Assert.IsType<C.ChartSpace>(
            Assert.Single(document.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts).ChartSpace);
        var valueAxis = chartSpace.Descendants<C.ValueAxis>().Single();
        Assert.Equal("0.0%", valueAxis.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
        Assert.DoesNotContain(ChartSemanticInspector.Inspect(chartSpace), finding =>
            finding.Subtype is IssueSubtypes.ChartPercentageAxisFormat or IssueSubtypes.ChartThresholdMissing);
    }

    [Fact]
    public void DetectsPercentageAxisAndDeclaredThresholdWithoutReferenceLine()
    {
        var series = new C.LineChartSeries(
            new C.Index { Val = 0U }, new C.Order { Val = 0U },
            new C.SeriesText(new C.NumericValue("错误回答率")),
            new C.CategoryAxisData(new C.StringLiteral(
                new C.PointCount { Val = 2U },
                new C.StringPoint(new C.NumericValue("第1周")) { Index = 0U },
                new C.StringPoint(new C.NumericValue("第2周")) { Index = 1U })),
            new C.Values(new C.NumberLiteral(
                new C.FormatCode("0.0%"), new C.PointCount { Val = 2U },
                new C.NumericPoint(new C.NumericValue("0.014")) { Index = 0U },
                new C.NumericPoint(new C.NumericValue("0.006")) { Index = 1U })));
        var titleText = new C.RichText(
            new DocumentFormat.OpenXml.Drawing.BodyProperties(),
            new DocumentFormat.OpenXml.Drawing.ListStyle(),
            new DocumentFormat.OpenXml.Drawing.Paragraph(
                new DocumentFormat.OpenXml.Drawing.Run(
                    new DocumentFormat.OpenXml.Drawing.Text("错误回答率（扩大门槛 0.5%）"))));
        var chart = new C.Chart(
            new C.Title(new C.ChartText(titleText)),
            new C.PlotArea(
                new C.LineChart(series),
                new C.ValueAxis(new C.NumberingFormat { FormatCode = "0", SourceLinked = false })));

        var findings = ChartSemanticInspector.Inspect(new C.ChartSpace(chart));

        Assert.Contains(findings, item => item.Subtype == IssueSubtypes.ChartPercentageAxisFormat);
        Assert.Contains(findings, item => item.Subtype == IssueSubtypes.ChartThresholdMissing);
    }

    [Fact]
    public void DetectsCategoryCountMissingSeriesNameAndPercentageUnit()
    {
        var series = new C.BarChartSeries(
            new C.Index { Val = 0U }, new C.Order { Val = 0U },
            new C.CategoryAxisData(new C.StringLiteral(
                new C.PointCount { Val = 2U },
                new C.StringPoint { Index = 0U, NumericValue = new C.NumericValue("A") },
                new C.StringPoint { Index = 1U, NumericValue = new C.NumericValue("B") })),
            new C.Values(new C.NumberLiteral(
                new C.FormatCode("0.0%"), new C.PointCount { Val = 1U },
                new C.NumericPoint { Index = 0U, NumericValue = new C.NumericValue("0.95") })));
        var space = new C.ChartSpace(new C.Chart(new C.PlotArea(new C.BarChart(series))));

        var findings = ChartSemanticInspector.Inspect(space);
        Assert.Contains(findings, item => item.Subtype == IssueSubtypes.ChartCategorySeriesMismatch);
        Assert.Contains(findings, item => item.Subtype == IssueSubtypes.ChartUnitMissing);
    }

    [Fact]
    public void DetectsTypedBulletInsideNativeBulletParagraph()
    {
        var paragraph = new Paragraph(
            new ParagraphProperties(new CharacterBullet { Char = "•" }),
            new Run(new Text("• 重复项目符号")));
        Assert.True(PresentationSemanticInspector.HasDuplicateBullet(paragraph, out _));
    }

    [Fact]
    public void DetectsTypedBulletInheritedFromListLevel()
    {
        var paragraph = new Paragraph(
            new ParagraphProperties { Level = 0 },
            new Run(new Text("• 重复项目符号")));
        var listStyle = new ListStyle(
            new Level1ParagraphProperties(new CharacterBullet { Char = "•" }));
        Assert.True(PresentationSemanticInspector.HasDuplicateBullet(paragraph, out _, listStyle));
    }

    [Fact]
    public void DetectsTypedBulletInheritedFromBodyPlaceholder()
    {
        var paragraph = new Paragraph(new Run(new Text("• 重复项目符号")));
        Assert.True(PresentationSemanticInspector.HasDuplicateBullet(paragraph, out _, inheritsPlaceholderBullet: true));
    }

    [Fact]
    public void DetectsTimePeriodsUsedAsSeriesAndMeasuresUsedAsCategories()
    {
        C.BarChartSeries MakeSeries(uint index, string name, params double[] values)
        {
            var literal = new C.NumberLiteral(
                new C.FormatCode("0.0"), new C.PointCount { Val = (uint)values.Length });
            foreach (var (value, point) in values.Select((value, point) => (value, point)))
            {
                literal.Append(new C.NumericPoint
                {
                    Index = (uint)point,
                    NumericValue = new C.NumericValue(value.ToString(System.Globalization.CultureInfo.InvariantCulture))
                });
            }

            return new C.BarChartSeries(
                new C.Index { Val = index }, new C.Order { Val = index },
                new C.SeriesText(new C.StringLiteral(
                    new C.PointCount { Val = 1U },
                    new C.StringPoint { Index = 0U, NumericValue = new C.NumericValue(name) })),
                new C.CategoryAxisData(new C.StringLiteral(
                    new C.PointCount { Val = 3U },
                    new C.StringPoint { Index = 0U, NumericValue = new C.NumericValue("收入") },
                    new C.StringPoint { Index = 1U, NumericValue = new C.NumericValue("成本") },
                    new C.StringPoint { Index = 2U, NumericValue = new C.NumericValue("利润") })),
                new C.Values(literal));
        }
        var space = new C.ChartSpace(new C.Chart(new C.PlotArea(new C.BarChart(
            MakeSeries(0U, "Q1", 10, 7, 3), MakeSeries(1U, "Q2", 12, 8, 4)))));

        Assert.Contains(ChartSemanticInspector.Inspect(space), item => item.Subtype == IssueSubtypes.ChartAxisSeriesSemantics);
    }

    [Fact]
    public void DetectsLargeUnscaledAxisLabelsAndAcceptsCompactMillionsFormat()
    {
        var series = new C.BarChartSeries(
            new C.Index { Val = 0U }, new C.Order { Val = 0U },
            new C.CategoryAxisData(new C.StringLiteral(
                new C.PointCount { Val = 1U },
                new C.StringPoint { Index = 0U, NumericValue = new C.NumericValue("7月") })),
            new C.Values(new C.NumberLiteral(
                new C.FormatCode("#,##0"), new C.PointCount { Val = 1U },
                new C.NumericPoint { Index = 0U, NumericValue = new C.NumericValue("12800000") })));
        var valueAxis = new C.ValueAxis(new C.NumberingFormat { FormatCode = "#,##0", SourceLinked = false });
        var space = new C.ChartSpace(new C.Chart(new C.PlotArea(new C.BarChart(series), valueAxis)));

        Assert.Contains(ChartSemanticInspector.Inspect(space), item => item.Subtype == IssueSubtypes.ChartAxisLabelDensity);

        valueAxis.NumberingFormat = new C.NumberingFormat { FormatCode = "0.0,,\"M\"", SourceLinked = false };
        Assert.DoesNotContain(ChartSemanticInspector.Inspect(space), item => item.Subtype == IssueSubtypes.ChartAxisLabelDensity);
    }

    [Fact]
    public void DetectsHeaderValuesMixedIntoChartRanges()
    {
        C.LineChartSeries MakeSeries(uint index, string name, string categoryYear, string categoryValue,
                                     string valueYear, string valueAmount)
        {
            return new C.LineChartSeries(
                new C.Index { Val = index }, new C.Order { Val = index },
                new C.SeriesText(new C.StringLiteral(
                    new C.PointCount { Val = 1U },
                    new C.StringPoint { Index = 0U, NumericValue = new C.NumericValue(name) })),
                new C.CategoryAxisData(new C.NumberLiteral(
                    new C.FormatCode("General"), new C.PointCount { Val = 2U },
                    new C.NumericPoint { Index = 0U, NumericValue = new C.NumericValue(categoryYear) },
                    new C.NumericPoint { Index = 1U, NumericValue = new C.NumericValue(categoryValue) })),
                new C.Values(new C.NumberLiteral(
                    new C.FormatCode("General"), new C.PointCount { Val = 2U },
                    new C.NumericPoint { Index = 0U, NumericValue = new C.NumericValue(valueYear) },
                    new C.NumericPoint { Index = 1U, NumericValue = new C.NumericValue(valueAmount) })));
        }
        var space = new C.ChartSpace(new C.Chart(new C.PlotArea(new C.LineChart(
            MakeSeries(0U, "Series 1", "2026", "2000000", "2027", "3200000"),
            MakeSeries(1U, "Series 2", "2026", "2000000", "2028", "5120000")))));

        var findings = ChartSemanticInspector.Inspect(space);
        Assert.Contains(findings, item => item.Subtype == IssueSubtypes.ChartSeriesNameMissing);
        Assert.Contains(findings, item => item.Subtype == IssueSubtypes.ChartAxisSeriesSemantics);
    }

    [Fact]
    public void FlagsOfficeOnlyLatinThemeFontsForCjkDecks()
    {
        var theme = new Theme(new ThemeElements(
            new ColorScheme() { Name = "test" },
            new FontScheme(
                new MajorFont(new LatinFont { Typeface = "Calibri Light" }),
                new MinorFont(new LatinFont { Typeface = "Calibri" })) { Name = "test" },
            new FormatScheme() { Name = "test" }));

        var risks = PresentationSemanticInspector.CrossSuiteFontRisks(theme, containsCjkText: true);
        Assert.Equal(new[] { "Calibri", "Calibri Light" }, risks);
        Assert.Empty(PresentationSemanticInspector.CrossSuiteFontRisks(theme, containsCjkText: false));
    }
}
