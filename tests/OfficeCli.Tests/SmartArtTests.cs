using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeCli.Core;
using OfficeCli.Core.Diagram;
using OfficeCli.Handlers;
using Xunit;

namespace OfficeCli.Tests;

public sealed class SmartArtTests
{
    private const string SpecJson = """
    {
      "schemaVersion": 1,
      "diagramId": "release-flow",
      "type": "timeline",
      "title": "发布流程",
      "direction": "left-right",
      "facts": [{"factId":"f1","sourceId":"report","locator":"section-1","summary":"已确认流程","confidence":1}],
      "nodes": [
        {"id":"draft","label":"准备","factRefs":["f1"]},
        {"id":"review","label":"审核","factRefs":[]},
        {"id":"publish","label":"发布","factRefs":[]}
      ],
      "edges": [
        {"id":"e1","from":"draft","to":"review","label":"提交","dashed":false,"factRefs":["f1"]},
        {"id":"e2","from":"review","to":"publish","label":"通过","dashed":false,"factRefs":[]}
      ]
    }
    """;

    [Fact]
    public void NativeSmartArtGenerationIsValidEditableAndInspectableInWordAndPowerPoint()
    {
        using var temp = new TempDirectory();
        var specPath = temp.File("diagram.json");
        File.WriteAllText(specPath, SpecJson);
        var pptx = temp.File("diagram.pptx");
        var docx = temp.File("diagram.docx");

        BlankDocCreator.Create(pptx, "zh-CN");
        using (var handler = new PowerPointHandler(pptx, editable: true))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            Assert.Equal("/slide[1]/smartart[1]", handler.Add("/slide[1]", "diagram", null,
                new Dictionary<string, string> { ["spec"] = specPath, ["render"] = "smartart" }));
        }
        using (var package = PresentationDocument.Open(pptx, false))
            Assert.Empty(new OpenXmlValidator().Validate(package));
        var pptInspection = NativeSmartArtCodec.Inspect(pptx);
        var pptDiagram = Assert.Single(pptInspection.Diagrams);
        Assert.Equal("release-flow", pptDiagram.DiagramId);
        Assert.Equal(new[] { "准备", "审核", "发布" }, pptDiagram.Nodes.Select(n => n.Label));
        Assert.Equal(2, pptDiagram.Edges.Count);
        Assert.Contains("/ppt/slides/slide1.xml", pptDiagram.Hosts);

        OpenXmlFixture.CreateDocument(docx);
        using (var handler = new WordHandler(docx, editable: true))
            Assert.Equal("/body/smartart[1]", handler.Add("/body", "diagram", null,
                new Dictionary<string, string> { ["spec"] = specPath, ["render"] = "smartart" }));
        using (var package = WordprocessingDocument.Open(docx, false))
            Assert.Empty(new OpenXmlValidator().Validate(package));
        var wordDiagram = Assert.Single(NativeSmartArtCodec.Inspect(docx).Diagrams);
        Assert.Equal("release-flow", wordDiagram.DiagramId);
        Assert.Contains("/word/document.xml", wordDiagram.Hosts);
    }

    [Fact]
    public void DeepUpdateUsesTheSameSpecAndPreservesNativeLayoutStyleAndHost()
    {
        using var temp = new TempDirectory();
        var specPath = temp.File("diagram.json");
        File.WriteAllText(specPath, SpecJson);
        var pptx = temp.File("diagram.pptx");
        BlankDocCreator.Create(pptx, "zh-CN");
        using (var handler = new PowerPointHandler(pptx, editable: true))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "diagram", null,
                new Dictionary<string, string> { ["spec"] = specPath, ["render"] = "smartart" });
        }
        var before = OfficePackageEvidence.Snapshot(pptx);
        var part = Assert.Single(NativeSmartArtCodec.Inspect(pptx).Diagrams).DataPart;
        var updatedSpec = DiagramSpec.Parse(SpecJson.Replace("审核", "质量复核", StringComparison.Ordinal));

        NativeSmartArtCodec.Update(pptx, part, updatedSpec);

        var after = Assert.Single(NativeSmartArtCodec.Inspect(pptx).Diagrams);
        Assert.Contains(after.Nodes, node => node.NodeId == "review" && node.Label == "质量复核");
        Assert.Equal(part, after.DataPart);
        Assert.Contains("/ppt/slides/slide1.xml", after.Hosts);
        using (var package = PresentationDocument.Open(pptx, false))
            Assert.Empty(new OpenXmlValidator().Validate(package));
        var manifest = OfficePackageEvidence.Diff(before, pptx);
        Assert.True(manifest.Passed);
        var afterSnapshot = OfficePackageEvidence.Snapshot(pptx);
        Assert.Equal(1, afterSnapshot.Features["smartArtLayoutParts"]);
        Assert.Equal(1, afterSnapshot.Features["smartArtColorParts"]);
        Assert.Equal(1, afterSnapshot.Features["smartArtStyleParts"]);
    }

    [Fact]
    public void UnrelatedEditAndDumpReplayPreserveExistingSmartArt()
    {
        using var temp = new TempDirectory();
        var specPath = temp.File("diagram.json");
        File.WriteAllText(specPath, SpecJson);
        var source = temp.File("source.pptx");
        var replay = temp.File("replay.pptx");
        BlankDocCreator.Create(source, "zh-CN");
        using (var handler = new PowerPointHandler(source, editable: true))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "diagram", null,
                new Dictionary<string, string> { ["spec"] = specPath, ["render"] = "smartart" });
        }
        StripOfficeCliSemanticExtension(source);
        var externalInspection = Assert.Single(NativeSmartArtCodec.Inspect(source).Diagrams);
        Assert.Null(externalInspection.DiagramId);
        Assert.Equal(new[] { "准备", "审核", "发布" }, externalInspection.Nodes.Select(node => node.Label));
        Assert.Empty(externalInspection.Edges);
        var snapshot = OfficePackageEvidence.Snapshot(source);
        using (var handler = new PowerPointHandler(source, editable: true))
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string> { ["text"] = "旁注" });
        Assert.True(OfficePackageEvidence.Diff(snapshot, source).Passed);

        List<BatchItem> items;
        using (var sourceHandler = new PowerPointHandler(source, editable: false))
        {
            var emitted = PptxBatchEmitter.EmitPptx(sourceHandler);
            items = emitted.Items;
            Assert.DoesNotContain(emitted.Warnings, warning => warning.Element.Contains("SmartArt", StringComparison.OrdinalIgnoreCase));
        }
        BlankDocCreator.Create(replay, "zh-CN");
        using (var targetHandler = DocumentHandlerFactory.Open(replay, editable: true))
        {
            var results = CommandBuilder.ApplyBatchItems(targetHandler, items, stopOnError: true, json: true);
            Assert.All(results, result => Assert.True(result.Success, result.Error));
        }
        var replayed = Assert.Single(NativeSmartArtCodec.Inspect(replay).Diagrams);
        Assert.Null(replayed.DiagramId);
        Assert.Equal(3, replayed.Nodes.Count);
        Assert.Empty(replayed.Edges);
        using var reopened = PresentationDocument.Open(replay, false);
        Assert.Empty(new OpenXmlValidator().Validate(reopened));
    }

    [Fact]
    public void ExistingExcelSmartArtIsInspectableAndSurvivesWorkbookEdits()
    {
        using var temp = new TempDirectory();
        var path = temp.File("existing-smartart.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Sheet1");
        AddSmartArtToWorkbook(path, DiagramSpec.Parse(SpecJson));
        StripOfficeCliSemanticExtension(path);

        var beforeInspection = Assert.Single(NativeSmartArtCodec.Inspect(path).Diagrams);
        Assert.Contains("/xl/drawings/drawing1.xml", beforeInspection.Hosts);
        Assert.Equal(new[] { "准备", "审核", "发布" }, beforeInspection.Nodes.Select(node => node.Label));
        var snapshot = OfficePackageEvidence.Snapshot(path);

        using (var handler = new ExcelHandler(path, editable: true))
            handler.Import("/Sheet1", "Name,Value\nAlpha,42", ',', hasHeader: true, "A1");

        Assert.True(OfficePackageEvidence.Diff(snapshot, path).Passed);
        var afterInspection = Assert.Single(NativeSmartArtCodec.Inspect(path).Diagrams);
        Assert.Equal(3, afterInspection.Nodes.Count);
        using var reopened = SpreadsheetDocument.Open(path, false);
        Assert.Empty(new OpenXmlValidator().Validate(reopened));
    }

    private static void AddSmartArtToWorkbook(string path, DiagramSpec spec)
    {
        using var workbook = SpreadsheetDocument.Open(path, true);
        var worksheetPart = workbook.WorkbookPart!.WorksheetParts.Single();
        var worksheet = worksheetPart.Worksheet!;
        var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
        worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing
        {
            Id = worksheetPart.GetIdOfPart(drawingsPart),
        });
        worksheet.Save();

        var dataPart = drawingsPart.AddNewPart<DiagramDataPart>();
        var layoutPart = drawingsPart.AddNewPart<DiagramLayoutDefinitionPart>();
        var colorsPart = drawingsPart.AddNewPart<DiagramColorsPart>();
        var stylePart = drawingsPart.AddNewPart<DiagramStylePart>();
        WritePart(dataPart, NativeSmartArtCodec.BuildDataXml(spec));
        WritePart(layoutPart, NativeSmartArtCodec.BuildLayoutXml());
        WritePart(colorsPart, NativeSmartArtCodec.BuildColorsXml());
        WritePart(stylePart, NativeSmartArtCodec.BuildStyleXml());

        var anchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor(
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("1"),
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("1"),
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")),
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("9"),
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("18"),
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")));
        var frame = new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame(
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2U, Name = "Existing SmartArt" },
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()),
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.Transform(
                new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
                new DocumentFormat.OpenXml.Drawing.Extents { Cx = 0, Cy = 0 }));
        var relIds = new OpenXmlUnknownElement("dgm", "relIds", "http://schemas.openxmlformats.org/drawingml/2006/diagram");
        const string relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        relIds.SetAttribute(new OpenXmlAttribute("r", "dm", relationships, drawingsPart.GetIdOfPart(dataPart)));
        relIds.SetAttribute(new OpenXmlAttribute("r", "lo", relationships, drawingsPart.GetIdOfPart(layoutPart)));
        relIds.SetAttribute(new OpenXmlAttribute("r", "cs", relationships, drawingsPart.GetIdOfPart(colorsPart)));
        relIds.SetAttribute(new OpenXmlAttribute("r", "qs", relationships, drawingsPart.GetIdOfPart(stylePart)));
        frame.Append(new DocumentFormat.OpenXml.Drawing.Graphic(
            new DocumentFormat.OpenXml.Drawing.GraphicData(relIds)
            {
                Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            }));
        anchor.Append(frame, new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());
        drawingsPart.WorksheetDrawing.Append(anchor);
        drawingsPart.WorksheetDrawing.Save();
    }

    private static void WritePart(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
        writer.Write(xml);
    }

    private static void StripOfficeCliSemanticExtension(string path)
    {
        var dataPart = Assert.Single(NativeSmartArtCodec.Inspect(path).Diagrams).DataPart.TrimStart('/');
        using var archive = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update);
        var entry = archive.GetEntry(dataPart) ?? throw new InvalidOperationException($"Missing {dataPart}");
        string xml;
        using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
        var document = System.Xml.Linq.XDocument.Parse(xml);
        System.Xml.Linq.XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        System.Xml.Linq.XNamespace diagram = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
        foreach (var extension in document.Descendants(drawing + "ext")
                     .Where(element => (string?)element.Attribute("uri") == "urn:officecli:diagram:v1")
                     .ToList())
            extension.Remove();
        foreach (var properties in document.Descendants(diagram + "prSet"))
        {
            properties.Attribute("phldrT")?.Remove();
            if (((string?)properties.Attribute("presName"))?.StartsWith("officecli:", StringComparison.Ordinal) == true)
                properties.Attribute("presName")?.Remove();
        }
        using var stream = entry.Open();
        stream.SetLength(0);
        using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
        document.Save(writer, System.Xml.Linq.SaveOptions.DisableFormatting);
    }
}
