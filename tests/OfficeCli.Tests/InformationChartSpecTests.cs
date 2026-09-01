using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests;

public class InformationChartSpecTests
{
    [Fact]
    public void CatalogContainsEightInformationDesignCharts()
    {
        var charts = InformationChartEngine.List();
        Assert.Equal(8, charts.Count);
        Assert.All(charts, chart =>
        {
            Assert.NotEmpty(chart.RequiredValues);
            Assert.NotEmpty(chart.AdaptiveRules);
            Assert.False(string.IsNullOrWhiteSpace(chart.FallbackRepresentation));
        });
    }

    [Fact]
    public void NonzeroAxisRequiresAnExplicitReason()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli-chart-{Guid.NewGuid():N}.json");
        try
        {
            File.WriteAllText(path, """
            {"schemaVersion":1,"chartId":"trend","chartType":"annotated-trend","title":"Revenue accelerated","items":[{"label":"Jun","actual":10},{"label":"Jul","actual":12.8}],"factRefs":["revenue"],"claimRefs":["growth"],"axisPolicy":"nonzero","axisReason":""}
            """);
            var error = Assert.Throws<CliException>(() => InformationChartEngine.Parse(path));
            Assert.Equal("chart_axis_reason_missing", error.Code);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void NativeInformationChartRoundTripsAcrossWordExcelAndPowerPoint()
    {
        using var temp = new TempDirectory();
        var docx = temp.File("chart.docx");
        var xlsx = temp.File("chart.xlsx");
        var pptx = temp.File("chart.pptx");
        OpenXmlFixture.CreateDocument(docx);
        OpenXmlFixture.CreateWorkbook(xlsx, "Sheet1");
        global::OfficeCli.BlankDocCreator.Create(pptx, "zh-CN");
        using (var presentation = new PowerPointHandler(pptx, editable: true))
            presentation.Add("/", "slide", null, new Dictionary<string, string> { ["layout"] = "blank" });
        foreach (var path in new[] { docx, xlsx, pptx })
        {
            var spec = new InformationChartSpec
            {
                ChartId = "revenue-trend", ChartType = "annotated-trend", Title = "Revenue accelerated", Unit = "CNY million",
                FactRefs = ["revenue-jun", "revenue-jul"], ClaimRefs = ["growth-accelerated"], AxisPolicy = "zero",
                Items = [new InformationChartItem { Label = "Jun", Actual = 10.8 }, new InformationChartItem { Label = "Jul", Actual = 12.8 }],
            };
            using (var handler = DocumentHandlerFactory.Open(path, editable: true))
            {
                var receipt = InformationChartEngine.Apply(handler, path, spec);
                Assert.Equal("native-chart", receipt.Representation);
            }
            using (var handler = DocumentHandlerFactory.Open(path, editable: false))
            {
                var charts = InformationChartEngine.Read(handler);
                Assert.True(charts.Count == 1,
                    $"chart readback failed for {Path.GetExtension(path)}: {System.Text.Json.JsonSerializer.Serialize(handler.Query("chart").Select(item => item.Format))}");
                var chart = charts[0];
                Assert.Equal("revenue-trend", chart.ChartId);
                Assert.Equal("annotated-trend", chart.RequestedChartType);
                Assert.Equal(new[] { "revenue-jun", "revenue-jul" }, chart.FactRefs);
                Assert.Equal(new[] { "growth-accelerated" }, chart.ClaimRefs);
            }
        }
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        using (var package = WordprocessingDocument.Open(docx, false)) Assert.Empty(validator.Validate(package));
        using (var package = SpreadsheetDocument.Open(xlsx, false)) Assert.Empty(validator.Validate(package));
        using (var package = PresentationDocument.Open(pptx, false)) Assert.Empty(validator.Validate(package));
    }
}
