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
}
