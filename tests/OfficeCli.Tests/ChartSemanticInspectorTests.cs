using DocumentFormat.OpenXml.Drawing;
using OfficeCli.Core;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeCli.Tests;

public class ChartSemanticInspectorTests
{
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
}
