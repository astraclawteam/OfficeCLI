using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests;

public class ProfessionalComponentTests
{
    [Fact]
    public void CatalogContainsTwelveImplementedSemanticComponents()
    {
        var components = ProfessionalComponentCatalog.List();
        Assert.Equal(12, components.Count);
        Assert.All(components, component =>
        {
            Assert.Equal(new[] { "docx", "xlsx", "pptx" }, component.TargetFormats);
            Assert.NotEmpty(component.RequiredSlots);
            Assert.NotEmpty(component.AdaptiveRules);
            Assert.Contains("editable", component.EditabilityContract, StringComparison.OrdinalIgnoreCase);
        });
    }

    [Fact]
    public void EveryComponentRejectsMissingRequiredSlotsAndBindings()
    {
        foreach (var definition in ProfessionalComponentCatalog.List())
        {
            var spec = new ProfessionalComponentSpec
            {
                ComponentId = definition.ComponentId,
                InstanceId = "component-test",
                Title = "Test",
                Items = [new ProfessionalComponentItem { Label = "Item" }],
            };
            using var handler = new RejectingHandler();
            Assert.Throws<CliException>(() => ProfessionalComponentCatalog.Apply(handler, "test.docx", spec, false));
        }
    }

    [Fact]
    public void ComponentSpecParserKeepsTypedScalarFields()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli-component-{Guid.NewGuid():N}.json");
        try
        {
            File.WriteAllText(path, """
            {"schemaVersion":1,"componentId":"kpi-strip","instanceId":"revenue-kpi","title":"Revenue exceeded budget","density":"balanced","items":[{"label":"Revenue","fields":{"value":"CNY 12.80M","delta":0.067}}],"factRefs":["revenue"]}
            """);
            var spec = ProfessionalComponentCatalog.Parse(path);
            Assert.Equal("CNY 12.80M", spec.Items[0].Fields["value"].GetString());
            Assert.Equal(0.067, spec.Items[0].Fields["delta"].GetDouble(), 3);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void NativeComponentRoundTripsAcrossWordExcelAndPowerPoint()
    {
        using var temp = new TempDirectory();
        var docx = temp.File("component.docx");
        var xlsx = temp.File("component.xlsx");
        var pptx = temp.File("component.pptx");
        OpenXmlFixture.CreateDocument(docx);
        OpenXmlFixture.CreateWorkbook(xlsx, "Sheet1");
        global::OfficeCli.BlankDocCreator.Create(pptx, "zh-CN");
        using (var presentation = new PowerPointHandler(pptx, editable: true))
            presentation.Add("/", "slide", null, new Dictionary<string, string> { ["layout"] = "blank" });

        foreach (var path in new[] { docx, xlsx, pptx })
        {
            var spec = KpiSpec(Path.GetExtension(path) == ".xlsx" ? "/Sheet1/B2" : null);
            using (var handler = DocumentHandlerFactory.Open(path, editable: true))
            {
                var inserted = ProfessionalComponentCatalog.Apply(handler, path, spec, update: false);
                Assert.Equal("inserted", inserted.Operation);
                spec.Items[0].Fields["value"] = Json("CNY 13.40M");
                var updated = ProfessionalComponentCatalog.Apply(handler, path, spec, update: true);
                Assert.Equal("updated", updated.Operation);
                Assert.Equal(inserted.InstanceId, updated.InstanceId);
            }
            using (var handler = DocumentHandlerFactory.Open(path, editable: false))
            {
                var components = ProfessionalComponentCatalog.Read(handler, path, "revenue-kpi");
                Assert.True(components.Count == 1,
                    $"component readback failed for {Path.GetExtension(path)}: {JsonSerializer.Serialize(handler.Query("table").Select(item => item.Format))}");
                var read = components[0];
                Assert.Equal("kpi-strip", read.ComponentId);
                Assert.Equal(1, read.ItemCount);
                Assert.Equal(new[] { "revenue" }, read.FactRefs);
            }
        }

        using (var package = WordprocessingDocument.Open(docx, false))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package).ToList();
            Assert.True(errors.Count == 0, string.Join("\n", errors.Select(error => $"{error.Description}\n{error.Node?.OuterXml}")));
        }
        using (var package = SpreadsheetDocument.Open(xlsx, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package));
        using (var package = PresentationDocument.Open(pptx, false))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(package));
    }

    [Fact]
    public void ChineseComponentsUseChineseBusinessHeaders()
    {
        using var temp = new TempDirectory();
        var pptx = temp.File("component-zh.pptx");
        global::OfficeCli.BlankDocCreator.Create(pptx, "zh-CN");
        using (var presentation = new PowerPointHandler(pptx, editable: true))
            presentation.Add("/", "slide", null, new Dictionary<string, string> { ["layout"] = "blank" });
        var spec = new ProfessionalComponentSpec
        {
            ComponentId = "owner-time-standard", InstanceId = "action-owner", Title = "行动责任",
            Items = [new ProfessionalComponentItem { Label = "增配工程师", Fields = new Dictionary<string, JsonElement>
            {
                ["owner"] = Json("交付负责人"), ["timeframe"] = Json("9月9日"), ["standard"] = Json("确认率不低于98%"),
            }}],
            ActionRefs = ["capacity-action"],
        };
        using (var handler = DocumentHandlerFactory.Open(pptx, editable: true))
            ProfessionalComponentCatalog.Apply(handler, pptx, spec, update: false);
        using var archive = System.IO.Compression.ZipFile.OpenRead(pptx);
        var xml = string.Concat(archive.Entries.Where(item => item.FullName.StartsWith("ppt/slides/", StringComparison.OrdinalIgnoreCase)
            && item.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .Select(entry => { using var reader = new StreamReader(entry.Open()); return reader.ReadToEnd(); }));
        Assert.Contains("事项", xml);
        Assert.Contains("负责人", xml);
        Assert.Contains("验收标准", xml);
    }

    private static ProfessionalComponentSpec KpiSpec(string? target) => new()
    {
        ComponentId = "kpi-strip",
        InstanceId = "revenue-kpi",
        Title = "Revenue",
        Target = target,
        Items = [new ProfessionalComponentItem
        {
            Label = "Revenue",
            Fields = new Dictionary<string, JsonElement> { ["value"] = Json("CNY 12.80M"), ["delta"] = Json(0.067) },
        }],
        FactRefs = ["revenue"],
    };

    private static JsonElement Json<T>(T value) => JsonSerializer.SerializeToElement(value);

    private sealed class RejectingHandler : IDocumentHandler
    {
        public void Dispose() { }
        public string ViewAsText(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null, string? range = null) => "";
        public string ViewAsAnnotated(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null) => "";
        public string ViewAsOutline() => "";
        public string ViewAsStats() => "";
        public System.Text.Json.Nodes.JsonNode ViewAsStatsJson() => new System.Text.Json.Nodes.JsonObject();
        public System.Text.Json.Nodes.JsonNode ViewAsOutlineJson() => new System.Text.Json.Nodes.JsonObject();
        public System.Text.Json.Nodes.JsonNode ViewAsTextJson(int? startLine = null, int? endLine = null, int? maxLines = null, HashSet<string>? cols = null, string? range = null) => new System.Text.Json.Nodes.JsonObject();
        public List<DocumentIssue> ViewAsIssues(string? issueType = null, int? limit = null) => [];
        public DocumentNode Get(string path, int depth = 1) => new();
        public List<DocumentNode> Query(string selector) => [];
        public List<string> Set(string path, Dictionary<string, string> properties) => [];
        public string Add(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties) => throw new InvalidOperationException();
        public string? Remove(string path, Dictionary<string, string>? properties = null) => null;
        public string Move(string sourcePath, string? targetParentPath, InsertPosition? position, Dictionary<string, string>? properties = null) => "";
        public string CopyFrom(string sourcePath, string targetParentPath, InsertPosition? position) => "";
        public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null) => "";
        public void RawSet(string partPath, string xpath, string action, string? xml) { }
        public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null) => ("", "");
        public List<ValidationError> Validate() => [];
        public bool TryExtractBinary(string path, string destPath, out string? contentType, out long byteCount) { contentType = null; byteCount = 0; return false; }
        public void Save() { }
    }
}
