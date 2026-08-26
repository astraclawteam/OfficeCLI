using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Handlers;
using OfficeCli.Core;
using OfficeCli;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Tests;

public sealed class P1RegressionTests
{
    [Fact]
    public void NativeRenderReportsDocumentOpenFailureInsteadOfOfficeMissing_Issue326()
    {
        var inner = new System.Runtime.InteropServices.COMException(
            "The file or directory is corrupted and unreadable.", unchecked((int)0x80070570));
        var error = CommandBuilder.NativeRenderCliException(
            "PowerPoint", new NativeRenderException("PowerPoint", "open", inner));

        Assert.Equal("native_open_failed", error.Code);
        Assert.Contains("HRESULT 0x80070570", error.Message);
        Assert.DoesNotContain("requires Windows", error.Message);
    }

    [Fact]
    public void NativeRenderStillClassifiesActivationFailureAsUnavailable_Issue326()
    {
        var inner = new System.Runtime.InteropServices.COMException(
            "Class not registered", unchecked((int)0x80040154));
        var error = CommandBuilder.NativeRenderCliException(
            "Word", new NativeRenderException("Word", "activation", inner));

        Assert.Equal("native_app_unavailable", error.Code);
        Assert.Contains("installed and available", error.Message);
    }

    [Fact]
    public void ExcelPhoneticGuideDoesNotAlterVisibleCellText_Issue343()
    {
        using var temp = new TempDirectory();
        var path = temp.File("phonetic.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Sheet1");

        using (var doc = SpreadsheetDocument.Open(path, true))
        {
            var workbookPart = doc.WorkbookPart!;
            var shared = workbookPart.AddNewPart<SharedStringTablePart>();
            shared.SharedStringTable = new SharedStringTable(
                new SharedStringItem(
                    new S.Text("漢字"),
                    new PhoneticRun(new S.Text("カンジ")) { BaseTextStartIndex = 0, EndingBaseIndex = 2 },
                    new PhoneticProperties { FontId = 0 }));
            shared.SharedStringTable.Save();

            var sheetData = workbookPart.WorksheetParts.Single().Worksheet.GetFirstChild<SheetData>()!;
            sheetData.Append(new Row(new Cell
            {
                CellReference = "A1",
                DataType = CellValues.SharedString,
                CellValue = new CellValue("0"),
            }) { RowIndex = 1 });
            workbookPart.WorksheetParts.Single().Worksheet.Save();
        }

        using var handler = new ExcelHandler(path, editable: false);
        var html = handler.ViewAsHtml();
        Assert.Contains("漢字", html);
        Assert.DoesNotContain("カンジ", html);
    }

    [Fact]
    public void WordWhitespaceOnlyDirectUnderlineRendersFillLine_Issue306()
    {
        using var temp = new TempDirectory();
        var path = temp.File("fill-line.docx");
        OpenXmlFixture.CreateDocument(path,
            new W.Paragraph(
                new W.Run(new W.Text("Name: ")),
                new W.Run(
                    new W.RunProperties(new W.Underline { Val = W.UnderlineValues.Single }),
                    new W.Text("        ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve })));

        using var handler = new WordHandler(path, editable: false);
        var html = handler.ViewAsHtml();
        Assert.Contains("class=\"w-fill-line\"", html);
        Assert.Contains("border-bottom:1px solid", html);
        Assert.Contains("width:2em", html);
    }

    [Fact]
    public void WordFormattingCanBeUnsetToRestoreStyleInheritance_Issue274()
    {
        using var temp = new TempDirectory();
        var path = temp.File("inherit.docx");
        OpenXmlFixture.CreateDocument(path,
            new W.Paragraph(
                new W.ParagraphProperties(
                    new W.Indentation { Left = "480" },
                    new W.SpacingBetweenLines { After = "120" }),
                new W.Run(
                    new W.RunProperties(
                        new W.Bold(),
                        new W.Color { Val = "FF0000" },
                        new W.FontSize { Val = "22" }),
                    new W.Text("inherits"))));

        using (var handler = new WordHandler(path, editable: true))
        {
            var unsupported = handler.Set("/body/p[1]", new Dictionary<string, string>
            {
                ["indent"] = "inherit",
                ["spaceAfter"] = "unset",
                ["bold"] = "inherit",
                ["color"] = "inherit",
                ["size"] = "inherit",
            });
            Assert.Empty(unsupported);
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var paragraph = reopened.MainDocumentPart!.Document.Body!.Elements<W.Paragraph>().Single();
        Assert.Null(paragraph.ParagraphProperties?.Indentation);
        Assert.Null(paragraph.ParagraphProperties?.SpacingBetweenLines);
        var rPr = paragraph.Elements<W.Run>().Single().RunProperties;
        Assert.Null(rPr?.Bold);
        Assert.Null(rPr?.Color);
        Assert.Null(rPr?.FontSize);
    }

    [Fact]
    public void BentConnectorViewportExpandsForOutOfRangeAdjustments_Issue341()
    {
        var frame = PowerPointHandler.GetConnectorSvgFrame("0,0 -25,0 -25,140 125,140 100,100");

        Assert.Equal(-25, frame.MinX);
        Assert.Equal(0, frame.MinY);
        Assert.Equal(150, frame.Width);
        Assert.Equal(140, frame.Height);
        Assert.Equal("-25 0 150 140", frame.ViewBox);
    }

    [Fact]
    public void PptHtmlHonorsNoFillOutlineJapaneseFallbackAndMathMode_Issue341()
    {
        using var temp = new TempDirectory();
        var path = temp.File("ppt-html-fidelity.pptx");
        BlankDocCreator.Create(path, "ja-JP");

        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "日本語",
                ["font"] = "Yu Gothic",
                ["lang"] = "ja-JP",
                ["textOutline"] = "none",
            });
            handler.Add("/slide[1]", "equation", null, new Dictionary<string, string>
            {
                ["formula"] = "x^2+y^2",
            });
        }

        using var preview = new PowerPointHandler(path, editable: false);
        var html = preview.ViewAsHtml();
        Assert.Contains("'Yu Gothic','Hiragino Sans'", html);
        Assert.DoesNotContain("-webkit-text-stroke", html);
        Assert.Contains("class=\"katex-formula\" data-display=\"true\"", html);
        Assert.Contains("displayMode: el.dataset.display === 'true'", html);
    }

    [Fact]
    public void ExcelIssuesDetectNumericHashOverflow_Issue301()
    {
        using var temp = new TempDirectory();
        var path = temp.File("numeric-overflow.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Sheet1");
        using (var doc = SpreadsheetDocument.Open(path, true))
        {
            var worksheet = doc.WorkbookPart!.WorksheetParts.Single().Worksheet;
            worksheet.InsertAt(new S.Columns(new S.Column { Min = 1, Max = 1, Width = 2.0, CustomWidth = true }), 0);
            worksheet.GetFirstChild<SheetData>()!.Append(
                new Row(new Cell
                {
                    CellReference = "A1",
                    DataType = CellValues.Number,
                    CellValue = new CellValue("123456789.01"),
                }) { RowIndex = 1 });
            worksheet.Save();
        }

        using var handler = new ExcelHandler(path, editable: false);
        var issue = Assert.Single(handler.ViewAsIssues(), i => i.Path == "/Sheet1/A1");
        Assert.Contains("numeric overflow", issue.Message);
        Assert.Contains("suggest.width=", issue.Message);
    }

    [Fact]
    public void ExcelPreviewMaterializesTheRequestedScreenshotRange_Issue246()
    {
        using var temp = new TempDirectory();
        var path = temp.File("range-preview.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Dashboard");
        using (var doc = SpreadsheetDocument.Open(path, true))
        {
            doc.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetData>()!
                .Append(new Row(new Cell
                {
                    CellReference = "A1", DataType = CellValues.String,
                    CellValue = new CellValue("月份"),
                }) { RowIndex = 1 });
            doc.WorkbookPart.WorksheetParts.Single().Worksheet.Save();
        }

        using var handler = new ExcelHandler(path, editable: false);
        var html = handler.ViewAsHtml("Dashboard!A1:M20");
        Assert.Contains("data-path=\"/Dashboard/M20\"", html);
    }

    [Fact]
    public void ExcelPrintPreviewHonorsDeclaredPrintAreaInsteadOfDistantEvidenceCell()
    {
        using var temp = new TempDirectory();
        var path = temp.File("print-area-preview.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Dashboard");
        using (var doc = SpreadsheetDocument.Open(path, true))
        {
            var sheetData = doc.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<SheetData>()!;
            sheetData.Append(
                new Row(new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("看板") }) { RowIndex = 1 },
                new Row(new Cell { CellReference = "Z100", DataType = CellValues.String, CellValue = new CellValue("内部证据") }) { RowIndex = 100 });
            doc.WorkbookPart.Workbook.DefinedNames ??= new DefinedNames();
            doc.WorkbookPart.Workbook.DefinedNames.Append(new DefinedName("'Dashboard'!$A$1:$M$20")
            {
                Name = "_xlnm.Print_Area", LocalSheetId = 0U,
            });
            doc.WorkbookPart.WorksheetParts.Single().Worksheet.Save();
            doc.WorkbookPart.Workbook.Save();
        }

        using var handler = new ExcelHandler(path, editable: false);
        var html = handler.ViewAsHtml("print-area");
        Assert.Contains("看板", html);
        Assert.DoesNotContain("内部证据", html);
        Assert.DoesNotContain("data-path=\"/Dashboard/Z100\"", html);
    }

    [Fact]
    public void TrackedReplaceCanDeleteACompleteHyperlink_Issue279()
    {
        using var temp = new TempDirectory();
        var path = temp.File("tracked-hyperlink.docx");
        OpenXmlFixture.CreateDocument(path,
            new W.Paragraph(
                new W.Run(new W.Text("See ")),
                new W.Hyperlink(new W.Run(new W.Text("原文"))) { Id = "rIdLink" }));
        using (var doc = WordprocessingDocument.Open(path, true))
            doc.MainDocumentPart!.AddHyperlinkRelationship(new Uri("https://example.com"), true, "rIdLink");

        using (var handler = new WordHandler(path, editable: true))
        {
            var unsupported = handler.Set("/body", new Dictionary<string, string>
            {
                ["find"] = "原文",
                ["replace"] = "",
                ["revision.author"] = "Reviewer",
            });
            Assert.Empty(unsupported);
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var link = reopened.MainDocumentPart!.Document.Body!.Descendants<W.Hyperlink>().Single();
        var deletion = link.Descendants<W.DeletedRun>().Single();
        Assert.Equal("Reviewer", deletion.Author?.Value);
        Assert.Equal("原文", deletion.Descendants<W.DeletedText>().Single().Text);
        Assert.Empty(new OpenXmlValidator().Validate(reopened));
    }

    [Fact]
    public void TargetedPptEditPreservesUntouchedSlideBytes_Issue267()
    {
        using var temp = new TempDirectory();
        var path = temp.File("targeted.pptx");
        BlankDocCreator.Create(path, "en-US");
        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "textbox", null,
                new Dictionary<string, string> { ["text"] = "first" });
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[2]", "textbox", null,
                new Dictionary<string, string> { ["text"] = "second" });
        }
        var before = OpenXmlFixture.EntrySha256(path, "ppt/slides/slide2.xml");

        using (var handler = new PowerPointHandler(path, editable: true))
            handler.Set("/slide[1]/shape[1]",
                new Dictionary<string, string> { ["text"] = "changed" });

        Assert.Equal(before, OpenXmlFixture.EntrySha256(path, "ppt/slides/slide2.xml"));
    }

    [Fact]
    public void CharacterRelativeFirstLineIndentSatisfiesAudit_Issue283()
    {
        using var temp = new TempDirectory();
        var path = temp.File("first-line-chars.docx");
        OpenXmlFixture.CreateDocument(path,
            new W.Paragraph(
                new W.ParagraphProperties(new W.Indentation { FirstLineChars = 200 }),
                new W.Run(new W.Text("正文已经使用两个字符的首行缩进。"))));

        using var handler = new WordHandler(path, editable: false);
        Assert.DoesNotContain(handler.ViewAsIssues(), issue =>
            issue.Message.Contains("missing first-line indent", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void GroupResizeChangesOnlyRequestedDimensionUnlessKeepAspect_Issue237()
    {
        using var temp = new TempDirectory();
        var path = temp.File("group-resize.pptx");
        BlankDocCreator.Create(path, "en-US");
        using var handler = new PowerPointHandler(path, editable: true);
        handler.Add("/", "slide", null, new Dictionary<string, string>());
        handler.Add("/slide[1]", "group", null, new Dictionary<string, string>
        {
            ["x"] = "1", ["y"] = "1", ["width"] = "10", ["height"] = "4",
        });
        var originalHeight = handler.Get("/slide[1]/group[1]").Format["height"];
        handler.Set("/slide[1]/group[1]", new Dictionary<string, string> { ["width"] = "20" });
        var resized = handler.Get("/slide[1]/group[1]");
        Assert.Equal(originalHeight, resized.Format["height"]);

        handler.Set("/slide[1]/group[1]", new Dictionary<string, string>
        {
            ["width"] = "10", ["keepAspect"] = "true",
        });
        var proportional = handler.Get("/slide[1]/group[1]");
        Assert.Equal("2emu", proportional.Format["height"]);
    }

    [Fact]
    public void PptLayoutAuditIsOptInAndFindsDirectionalConnectorWithoutArrow_Issue301()
    {
        using var temp = new TempDirectory();
        var path = temp.File("layout-audit.pptx");
        BlankDocCreator.Create(path, "en-US");
        using (var handler = new PowerPointHandler(path, editable: true))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Source", ["x"] = "1", ["y"] = "5", ["width"] = "4", ["height"] = "2",
            });
            handler.Add("/slide[1]", "textbox", null, new Dictionary<string, string>
            {
                ["text"] = "Target", ["x"] = "10", ["y"] = "5", ["width"] = "4", ["height"] = "2",
            });
            handler.Add("/slide[1]", "connector", null, new Dictionary<string, string>
            {
                ["from"] = "/slide[1]/shape[1]", ["to"] = "/slide[1]/shape[2]",
            });
        }

        using var audit = new PowerPointHandler(path, editable: false);
        Assert.DoesNotContain(audit.ViewAsIssues(), i => i.Subtype == IssueSubtypes.PptLayout);
        Assert.Contains(audit.ViewAsIssues(IssueSubtypes.PptLayout), i =>
            i.Message.Contains("no arrowhead", StringComparison.OrdinalIgnoreCase));
    }

}
