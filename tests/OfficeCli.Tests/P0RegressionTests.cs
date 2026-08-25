using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Tests;

public sealed class P0RegressionTests
{
    private const string OnePixelPngDataUri =
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

    [Fact]
    public void PresetWithoutHandlesRejectsAdjustment_Issue235()
    {
        var values = new A.AdjustValueList();
        var error = Assert.Throws<ArgumentException>(() =>
            PowerPointHandler.ApplyAdjustHandles(
                values, "adj:val 5000", A.ShapeTypeValues.Rectangle));

        Assert.Contains("has no adjust handles", error.Message);
        Assert.Empty(values.ChildElements);
    }

    [Fact]
    public void RemovingSheetDropsAndRenumbersScopedNames_Issue243()
    {
        using var temp = new TempDirectory();
        var path = temp.File("scoped-names.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Sheet1", "Sheet2", "Sheet3");
        using (var doc = SpreadsheetDocument.Open(path, true))
        {
            doc.WorkbookPart!.Workbook.Append(new DefinedNames(
                new DefinedName("Sheet2!$A$1") { Name = "RemovedScoped", LocalSheetId = 1U },
                new DefinedName("Sheet3!$A$1") { Name = "LaterScoped", LocalSheetId = 2U }));
            doc.WorkbookPart.Workbook.Save();
        }

        using (var handler = new ExcelHandler(path, editable: true))
            handler.Remove("/Sheet2");

        using var reopened = SpreadsheetDocument.Open(path, false);
        var names = reopened.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>().ToList();
        Assert.DoesNotContain(names, n => n.Name?.Value == "RemovedScoped");
        var later = Assert.Single(names, n => n.Name?.Value == "LaterScoped");
        Assert.Equal(1U, later.LocalSheetId!.Value);
    }

    [Fact]
    public void ValidateFindsOutOfRangeDefinedNameScope_Issue243()
    {
        using var temp = new TempDirectory();
        var path = temp.File("invalid-scope.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Sheet1");
        using (var doc = SpreadsheetDocument.Open(path, true))
        {
            doc.WorkbookPart!.Workbook.Append(new DefinedNames(
                new DefinedName("Sheet1!$A$1") { Name = "Broken", LocalSheetId = 8U }));
            doc.WorkbookPart.Workbook.Save();
        }

        using var handler = new ExcelHandler(path, editable: false);
        var errors = handler.Validate();
        Assert.Contains(errors, e =>
            e.ErrorType == "Semantic" && e.Description.Contains("localSheetId=8", StringComparison.Ordinal));
    }

    [Fact]
    public void RejectedSheetRemovalLeavesOriginalBytesUntouched_Issue243()
    {
        using var temp = new TempDirectory();
        var path = temp.File("referenced-name.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Sheet1", "Sheet2");
        using (var doc = SpreadsheetDocument.Open(path, true))
        {
            doc.WorkbookPart!.Workbook.Append(new DefinedNames(
                new DefinedName("Sheet2!$A$1") { Name = "InputValue", LocalSheetId = 1U }));
            var sheet1 = doc.WorkbookPart.WorksheetParts.First();
            var row = new Row { RowIndex = 1U };
            row.Append(new Cell
            {
                CellReference = "A1",
                CellFormula = new CellFormula("InputValue"),
                CellValue = new CellValue("0"),
            });
            sheet1.Worksheet.GetFirstChild<SheetData>()!.Append(row);
            sheet1.Worksheet.Save();
            doc.WorkbookPart.Workbook.Save();
        }
        var before = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));

        using (var handler = new ExcelHandler(path, editable: true))
            Assert.Throws<ArgumentException>(() => handler.Remove("/Sheet2"));

        var after = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
        Assert.Equal(before, after);
    }

    [Fact]
    public void InCellImageDumpReplayPreservesBytesAndAltText_Issue247()
    {
        using var temp = new TempDirectory();
        var source = temp.File("image-source.xlsx");
        var replayed = temp.File("image-replayed.xlsx");
        OpenXmlFixture.CreateWorkbook(source, "Sheet1");
        OpenXmlFixture.CreateWorkbook(replayed, "Sheet1");

        using (var handler = new ExcelHandler(source, editable: true))
        {
            handler.Set("/Sheet1/C3", new Dictionary<string, string>
            {
                ["image"] = OnePixelPngDataUri,
                ["alt"] = "Site photo",
            });
        }

        List<BatchItem> dump;
        using (var handler = new ExcelHandler(source, editable: false))
            dump = ExcelBatchEmitter.EmitExcel(handler).Items;

        var imageSet = Assert.Single(dump, item =>
            item.Command == "set" && item.Path == "/Sheet1/C3"
            && item.Props?.ContainsKey("image") == true);
        Assert.StartsWith("data:image/png;base64,", imageSet.Props!["image"]);
        Assert.Equal("Site photo", imageSet.Props["alt"]);
        Assert.DoesNotContain(dump, item => item.Text?.Contains("[image]", StringComparison.Ordinal) == true);

        using (var handler = new ExcelHandler(replayed, editable: true))
        {
            var result = BatchExecutor.ExecuteBatch(handler, JsonSerializer.Serialize(dump), json: true);
            using var envelope = JsonDocument.Parse(result);
            Assert.True(envelope.RootElement.GetProperty("success").GetBoolean(), result);
        }

        Assert.Equal(
            OpenXmlFixture.SingleEntrySha256(source, "xl/media/"),
            OpenXmlFixture.SingleEntrySha256(replayed, "xl/media/"));
        using var reopened = new ExcelHandler(replayed, editable: false);
        var node = reopened.Get("/Sheet1/C3");
        Assert.Equal("Image", node.Format["type"]);
        Assert.Equal("Site photo", node.Format["alt"]);
    }

    [Fact]
    public void TrackedParagraphFormatSnapshotUsesCanonicalPPrOrder_Issue277()
    {
        using var temp = new TempDirectory();
        var path = temp.File("tracked-format.docx");
        OpenXmlFixture.CreateDocument(path,
            new W.Paragraph(
                new W.ParagraphProperties(
                    new W.ParagraphStyleId { Val = "21" },
                    new W.ParagraphMarkRunProperties(new W.FontSize { Val = "36" })),
                new W.Run(new W.Text("Hello"))));

        using (var handler = new WordHandler(path, editable: true))
        {
            handler.Set("/body/p[1]", new Dictionary<string, string>
            {
                ["style"] = "31",
                ["revision.type"] = "format",
                ["revision.author"] = "T",
            });
        }

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var change = Assert.Single(doc.MainDocumentPart!.Document
                .Descendants<W.ParagraphPropertiesChange>());
            var snapshot = Assert.Single(change.ChildElements, child => child.LocalName == "pPr");
            var names = snapshot.ChildElements.Select(child => child.LocalName).ToList();
            Assert.True(names.IndexOf("pStyle") >= 0, string.Join(",", names));
            Assert.True(names.IndexOf("rPr") > names.IndexOf("pStyle"), string.Join(",", names));
            Assert.Equal("rPr", names[^1]);
        }

        using var validator = new WordHandler(path, editable: false);
        Assert.Empty(validator.Validate());
    }

    [Fact]
    public void ValidatorAcceptsSchemaOrderedPPrChangeRunProperties_Issue278()
    {
        using var temp = new TempDirectory();
        var path = temp.File("word-authored-ppr-change.docx");
        var previous = new W.ParagraphProperties(
            new W.ParagraphStyleId { Val = "21" },
            new W.ParagraphMarkRunProperties(new W.FontSize { Val = "36" }));
        var change = new W.ParagraphPropertiesChange
        {
            Author = "Word",
            Date = DateTime.UtcNow,
            Id = "99999999",
        };
        change.InnerXml = previous.OuterXml;
        OpenXmlFixture.CreateDocument(path,
            new W.Paragraph(
                new W.ParagraphProperties(change),
                new W.Run(new W.Text("Hello"))));

        using var handler = new WordHandler(path, editable: false);
        Assert.Empty(handler.Validate());
    }

    [Fact]
    public void ResidentPersistsAcknowledgedMutationBeforeResponse_Issue328()
    {
        using var temp = new TempDirectory();
        var path = temp.File("resident-durable.docx");
        OpenXmlFixture.CreateDocument(path,
            new W.Paragraph(new W.Run(new W.Text("before"))));
        using var server = new ResidentServer(path);
        var request = new ResidentRequest
        {
            Command = "set",
            Json = true,
            Args = { ["path"] = "/body/p[1]/r[1]" },
            Props = new Dictionary<string, string> { ["text"] = "after" },
        };

        var wire = server.ProcessRequest(JsonSerializer.Serialize(
            request, ResidentJsonContext.Default.ResidentRequest));
        var response = JsonSerializer.Deserialize(
            wire, ResidentJsonContext.Default.ResidentResponse)!;
        Assert.Equal(0, response.ExitCode);

        var documentXml = OpenXmlFixture.EntryTextShared(path, "word/document.xml");
        Assert.Contains(">after<", documentXml, StringComparison.Ordinal);
        Assert.DoesNotContain(">before<", documentXml, StringComparison.Ordinal);
    }

    [Fact]
    public void ResidentAtomicBatchFailureLeavesOriginalBytesUntouched_Issue244()
    {
        using var temp = new TempDirectory();
        var path = temp.File("resident-atomic-failure.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Sheet1");
        var before = OpenXmlFixture.FileSha256Shared(path);
        using var server = new ResidentServer(path);
        var items = new List<BatchItem>
        {
            new() { Command = "set", Path = "/Sheet1/A1", Props = new() { ["value"] = "must rollback" } },
            new() { Command = "set", Path = "/Missing/A1", Props = new() { ["value"] = "invalid" } },
        };
        var request = new ResidentRequest
        {
            Command = "batch",
            Json = true,
            Args =
            {
                ["batchJson"] = JsonSerializer.Serialize(items),
                ["stopOnError"] = "true",
                ["bestEffort"] = "false",
                ["force"] = "false",
            },
        };

        var wire = server.ProcessRequest(JsonSerializer.Serialize(
            request, ResidentJsonContext.Default.ResidentRequest));
        var response = JsonSerializer.Deserialize(
            wire, ResidentJsonContext.Default.ResidentResponse)!;
        Assert.NotEqual(0, response.ExitCode);
        Assert.Contains("atomicRolledBack", response.Stdout, StringComparison.Ordinal);

        var after = OpenXmlFixture.FileSha256Shared(path);
        Assert.Equal(before, after);
    }

    [Fact]
    public void ResidentAtomicBatchPersistsCompleteResultBeforeResponse_Issues244And328()
    {
        using var temp = new TempDirectory();
        var path = temp.File("resident-atomic-success.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Sheet1");
        using var server = new ResidentServer(path);
        var items = new List<BatchItem>
        {
            new() { Command = "set", Path = "/Sheet1/A1", Props = new() { ["value"] = "first" } },
            new() { Command = "set", Path = "/Sheet1/B1", Props = new() { ["value"] = "second" } },
        };
        var request = new ResidentRequest
        {
            Command = "batch",
            Json = true,
            Args =
            {
                ["batchJson"] = JsonSerializer.Serialize(items),
                ["stopOnError"] = "true",
                ["bestEffort"] = "false",
                ["force"] = "false",
            },
        };

        var wire = server.ProcessRequest(JsonSerializer.Serialize(
            request, ResidentJsonContext.Default.ResidentRequest));
        var response = JsonSerializer.Deserialize(
            wire, ResidentJsonContext.Default.ResidentResponse)!;
        Assert.Equal(0, response.ExitCode);

        using var reopened = new ExcelHandler(path, editable: false);
        Assert.Equal("first", reopened.Get("/Sheet1/A1").Text);
        Assert.Equal("second", reopened.Get("/Sheet1/B1").Text);
    }
}
