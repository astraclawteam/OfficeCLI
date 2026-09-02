using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests;

public class ProfessionalPageSpecTests
{
    [Fact]
    public void PageSpecRejectsAComponentBlockWithoutComponentSpec()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli-page-{Guid.NewGuid():N}.json");
        try
        {
            File.WriteAllText(path, """
            {"schemaVersion":1,"documentId":"review","format":"pptx","title":"Review","pages":[{"pageId":"decision","pageRole":"decision","primaryClaim":"Capacity limits growth","readerTakeaway":"The gap is material","readerAction":"Approve capacity","density":"balanced","blocks":[{"blockId":"missing","kind":"component","importance":"primary"}]}]}
            """);
            var error = Assert.Throws<CliException>(() => ProfessionalPageCompiler.Parse(path));
            Assert.Equal("page_spec_invalid", error.Code);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void PageSpecUsesDifferentiatedComposersAndKeepsNativeBindings()
    {
        using var temp = new TempDirectory();
        var files = new Dictionary<string, string>
        {
            ["docx"] = temp.File("composed.docx"), ["xlsx"] = temp.File("composed.xlsx"), ["pptx"] = temp.File("composed.pptx"),
        };
        OpenXmlFixture.CreateDocument(files["docx"]);
        OpenXmlFixture.CreateWorkbook(files["xlsx"], "Sheet1");
        global::OfficeCli.BlankDocCreator.Create(files["pptx"], "zh-CN");
        foreach (var (format, file) in files)
        {
            var spec = Spec(format);
            using (var handler = DocumentHandlerFactory.Open(file, editable: true))
            {
                var receipt = ProfessionalPageCompiler.Compile(handler, file, spec);
                Assert.Equal(format switch { "docx" => "WordComposer", "xlsx" => "ExcelComposer", _ => "PowerPointComposer" }, receipt.Composer);
                Assert.Single(receipt.Pages);
                Assert.Equal(4, receipt.Pages[0].Blocks.Count);
            }
            using (var handler = DocumentHandlerFactory.Open(file, editable: false))
            {
                Assert.Single(ProfessionalComponentCatalog.Read(handler, file, "revenue-kpi"));
                Assert.Single(InformationChartEngine.Read(handler));
            }
        }
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        using (var package = WordprocessingDocument.Open(files["docx"], false)) Assert.Empty(validator.Validate(package));
        using (var package = SpreadsheetDocument.Open(files["xlsx"], false)) Assert.Empty(validator.Validate(package));
        using (var package = PresentationDocument.Open(files["pptx"], false)) Assert.Empty(validator.Validate(package));
        using (var handler = DocumentHandlerFactory.Open(files["pptx"], editable: false))
        {
            var slide = handler.Get("/slide[1]", 2);
            var title = slide.Children.Single(item => item.Format.GetValueOrDefault("name")?.ToString() == "pagespec-title-decision");
            Assert.Equal("30pt", title.Format["size"]?.ToString());
            var narrative = slide.Children.Single(item => item.Format.GetValueOrDefault("name")?.ToString() == "pagespec-block-context");
            Assert.Equal("none", narrative.Format["fill"]?.ToString());
            Assert.Contains(slide.Children, item => item.Text == "Next action · Approve capacity");
            var chart = slide.Children.Single(item => item.Type == "chart");
            Assert.Equal("false", chart.Format["gridlines"]?.ToString());
            Assert.True(ParseLengthInPoints(chart.Format["height"]) > ParseLengthInPoints(narrative.Format["height"]));
        }
    }

    private static double ParseLengthInPoints(object? value)
    {
        var text = (value?.ToString() ?? "0").Trim();
        var factor = 1.0;
        if (text.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
        {
            text = text[..^2];
            factor = 72.0 / 2.54;
        }
        else if (text.EndsWith("in", StringComparison.OrdinalIgnoreCase))
        {
            text = text[..^2];
            factor = 72.0;
        }
        else if (text.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) text = text[..^2];
        return double.Parse(text, System.Globalization.CultureInfo.InvariantCulture) * factor;
    }

    private static ProfessionalPageSpec Spec(string format) => new()
    {
        DocumentId = "professional-review", Format = format, Title = "Business review",
        BrandTokens = new Dictionary<string, string> { ["primary"] = "1F4E78", ["background"] = "F7F9FC" },
        Pages = [new ProfessionalPage
        {
            PageId = "decision", PageRole = "decision", PrimaryClaim = "Revenue accelerated and requires a capacity decision",
            ReaderTakeaway = "Growth is above target", ReaderAction = "Approve capacity", Density = "balanced",
            Blocks = [
                new ProfessionalPageBlock { BlockId = "context", Kind = "narrative", Importance = "context", Title = "Context", Text = "Actual performance is above target.", FactRefs = ["revenue-jul"], ClaimRefs = ["growth-accelerated"] },
                new ProfessionalPageBlock { BlockId = "kpi", Kind = "component", Importance = "supporting", FactRefs = ["revenue-jul"], Component = new ProfessionalComponentSpec { ComponentId = "kpi-strip", InstanceId = "revenue-kpi", Title = "Revenue", Items = [new ProfessionalComponentItem { Label = "Revenue", Fields = new Dictionary<string, System.Text.Json.JsonElement> { ["value"] = System.Text.Json.JsonSerializer.SerializeToElement("CNY 12.8M") } }], FactRefs = ["revenue-jul"] } },
                new ProfessionalPageBlock { BlockId = "trend", Kind = "chart", Importance = "primary", FactRefs = ["revenue-jun", "revenue-jul"], ClaimRefs = ["growth-accelerated"], Chart = new InformationChartSpec { ChartId = "revenue-trend", ChartType = "annotated-trend", Title = "Revenue accelerated", Unit = "CNY million", AxisPolicy = "zero", FactRefs = ["revenue-jun", "revenue-jul"], ClaimRefs = ["growth-accelerated"], Items = [new InformationChartItem { Label = "Jun", Actual = 10.8 }, new InformationChartItem { Label = "Jul", Actual = 12.8 }] } },
            ],
        }],
    };
}
