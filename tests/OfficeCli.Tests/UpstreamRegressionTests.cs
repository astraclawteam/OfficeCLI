using System.Diagnostics;
using System.Collections.Concurrent;
using System.Text.Json;
using System.Threading.Channels;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using OfficeCli.Handlers;
using Xunit;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Tests;

public sealed class UpstreamRegressionTests
{
    [Fact]
    public void ExcelImportPersistsAfterDispose_Issue316()
    {
        using var temp = new TempDirectory();
        var path = temp.File("import.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Sheet1");

        using (var handler = new ExcelHandler(path, editable: true))
            handler.Import("/Sheet1", "Name,Value\nAlpha,42", ',', hasHeader: true, "A1");

        using var reopened = SpreadsheetDocument.Open(path, false);
        var cells = reopened.WorkbookPart!.WorksheetParts.Single().Worksheet
            .Descendants<S.Cell>().ToDictionary(c => c.CellReference!.Value!, c => c.InnerText);
        Assert.Equal("Name", cells["A1"]);
        Assert.Equal("42", cells["B2"]);
    }

    [Fact]
    public void WordTableIsoJustificationCanBeRead_Issue324()
    {
        using var temp = new TempDirectory();
        var path = temp.File("table.docx");
        var table = new Table(
            new TableProperties(new TableJustification { Val = TableRowAlignmentValues.Left }),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("value"))))));
        OpenXmlFixture.CreateDocument(path, table);
        OpenXmlFixture.ReplaceEntryText(path, "word/document.xml",
            xml => xml.Replace("w:val=\"left\"", "w:val=\"start\"", StringComparison.Ordinal));

        using var handler = new WordHandler(path, editable: false);
        var body = handler.Get("/body", depth: 5);
        var node = Assert.Single(body.Children, c => c.Type == "table");
        Assert.Equal("start", node.Format["align"]);
    }

    [Fact]
    public void ParagraphTextPreservesDrawingAndPictureRunRejectsText_Issues334And335()
    {
        using var temp = new TempDirectory();
        var path = temp.File("drawing.docx");
        OpenXmlFixture.CreateDocument(path,
            new Paragraph(
                new Run(new Text("before")),
                new Run(new Drawing())));

        using (var handler = new WordHandler(path, editable: true))
        {
            var paragraphUnsupported = handler.Set("/body/p[1]",
                new Dictionary<string, string> { ["text"] = "after" });
            Assert.Empty(paragraphUnsupported);
            var pictureUnsupported = handler.Set("/body/p[1]/r[2]",
                new Dictionary<string, string> { ["text"] = "must-not-apply" });
            Assert.Contains("text", pictureUnsupported);
        }

        using var reopened = WordprocessingDocument.Open(path, false);
        var paragraph = reopened.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Single();
        Assert.Equal("after", string.Concat(paragraph.Descendants<Text>().Select(t => t.Text)));
        Assert.Single(paragraph.Descendants<Drawing>());
    }

    [Fact]
    public void WatchPathsUseCanonicalCase_Issue330()
    {
        Assert.Equal(2, WatchMessage.ExtractSlideNum("/Slide[2]"));
        Assert.Equal("[data-path=\"/body/table[1]/tr[2]/tc[3]\"]",
            WatchMessage.ExtractWordScrollTarget("/BODY/Table[1]/TR[2]/TC[3]"));
    }

    [Fact]
    public void UntouchedWorksheetPartIsBytePreserved_Issue338()
    {
        using var temp = new TempDirectory();
        var path = temp.File("targeted.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Sheet1", "Sheet2");
        var before = OpenXmlFixture.EntrySha256(path, "xl/worksheets/sheet1.xml");

        using (var handler = new ExcelHandler(path, editable: true))
            handler.Set("/Sheet2/A1", new Dictionary<string, string> { ["value"] = "changed" });

        Assert.Equal(before, OpenXmlFixture.EntrySha256(path, "xl/worksheets/sheet1.xml"));
    }

    [Fact]
    public void TargetedWordEditPreservesAlternateContentMirrorParaIds_Issue336()
    {
        using var temp = new TempDirectory();
        var path = temp.File("mirror.docx");
        OpenXmlFixture.CreateDocument(path, new Paragraph(new Run(new Text("placeholder"))));
        OpenXmlFixture.ReplaceEntryText(path, "word/document.xml", _ =>
            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
            "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" " +
            "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
            "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 wps\">" +
            "<w:body>" +
            "<mc:AlternateContent><mc:Choice Requires=\"wps\">" +
            "<w:p w14:paraId=\"22222222\"><w:r><w:t>mirror-a</w:t></w:r></w:p>" +
            "</mc:Choice><mc:Fallback>" +
            "<w:p w14:paraId=\"22222222\"><w:r><w:t>mirror-b</w:t></w:r></w:p>" +
            "</mc:Fallback></mc:AlternateContent>" +
            "<w:p w14:paraId=\"11111111\"><w:r><w:t>target</w:t></w:r></w:p>" +
            "<w:sectPr/></w:body></w:document>");
        var before = OpenXmlFixture.EntryText(path, "word/document.xml");

        using (var handler = new WordHandler(path, editable: true))
            handler.Set("/body/p[@paraId=11111111]", new Dictionary<string, string> { ["text"] = "edited" });

        var after = OpenXmlFixture.EntryText(path, "word/document.xml");
        Assert.Equal(2, Count(after, "w14:paraId=\"22222222\""));
        Assert.Equal(2, Count(before, "w14:paraId=\"22222222\""));
    }

    [Fact]
    public void WordHtmlShowsCommentTextAndMarker_Issue342()
    {
        using var temp = new TempDirectory();
        var path = temp.File("comments.docx");
        OpenXmlFixture.CreateDocument(path,
            new Paragraph(
                new CommentRangeStart { Id = "0" },
                new Run(new Text("review me")),
                new CommentRangeEnd { Id = "0" },
                new Run(new CommentReference { Id = "0" })));
        using (var doc = WordprocessingDocument.Open(path, true))
        {
            var commentsPart = doc.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments(
                new Comment(new Paragraph(new Run(new Text("check wording"))))
                { Id = "0", Author = "Alice" });
            commentsPart.Comments.Save();
        }

        using var handler = new WordHandler(path, editable: false);
        var html = handler.ViewAsHtml();
        Assert.Contains("w-comment-range", html);
        Assert.Contains("w-comment-ref", html);
        Assert.Contains("Alice: check wording", html);
    }

    [Fact]
    public async Task ConsecutiveExplicitBatchCallsDoNotLoseMcpFrame_Issues339And340()
    {
        using var temp = new TempDirectory();
        var deck = temp.File("deck.pptx");
        var first = temp.File("first.json");
        var second = temp.File("second.json");
        await File.WriteAllTextAsync(first,
            "[{\"command\":\"add\",\"parent\":\"/\",\"type\":\"slide\"}]");
        await File.WriteAllTextAsync(second,
            "[{\"command\":\"add\",\"parent\":\"/slide[1]\",\"type\":\"textbox\",\"props\":{\"text\":\"Hello\"}}]");

        var executable = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory, "..", "..", "..", "..", "..",
            "src", "officecli", "bin", "Release", "net10.0", "win-x64", "officecli.exe"));
        Assert.True(File.Exists(executable), $"Missing test executable: {executable}");
        var startInfo = new ProcessStartInfo(executable, "mcp")
        {
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            WorkingDirectory = temp.Path,
        };
        startInfo.Environment["OFFICECLI_SKIP_UPDATE"] = "1";
        startInfo.Environment["OFFICECLI_NO_AUTO_RESIDENT"] = "1";
        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start MCP process.");
        var responses = Channel.CreateUnbounded<string>();
        var errors = new ConcurrentQueue<string>();
        process.OutputDataReceived += (_, e) =>
        {
            if (e.Data != null) responses.Writer.TryWrite(e.Data);
        };
        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data != null) errors.Enqueue(e.Data);
        };
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        try
        {
            await SendAsync(process, 1, "initialize", new
            {
                protocolVersion = "2024-11-05",
                capabilities = new { },
                clientInfo = new { name = "regression", version = "1" },
            });
            await ExpectIdAsync(responses.Reader, errors, 1);
            await SendToolAsync(process, 2, ["create", deck]);
            await ExpectIdAsync(responses.Reader, errors, 2);
            await SendToolAsync(process, 3, ["batch", deck, "--input", first]);
            await ExpectIdAsync(responses.Reader, errors, 3);
            await SendToolAsync(process, 4, ["batch", deck, "--input", second]);
            await ExpectIdAsync(responses.Reader, errors, 4);
        }
        finally
        {
            if (!process.HasExited) process.Kill(entireProcessTree: true);
        }
    }

    private static int Count(string value, string token) =>
        value.Split(token, StringSplitOptions.None).Length - 1;

    private static Task SendToolAsync(Process process, int id, string[] command) =>
        SendAsync(process, id, "tools/call", new
        {
            name = "officecli",
            arguments = new { command },
        });

    private static async Task SendAsync(Process process, int id, string method, object parameters)
    {
        var request = JsonSerializer.Serialize(new { jsonrpc = "2.0", id, method, @params = parameters });
        await process.StandardInput.WriteLineAsync(request);
        await process.StandardInput.FlushAsync();
    }

    private static async Task ExpectIdAsync(
        ChannelReader<string> responses,
        ConcurrentQueue<string> errors,
        int expected)
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
        while (true)
        {
            string line;
            try
            {
                line = await responses.ReadAsync(timeout.Token);
            }
            catch (OperationCanceledException)
            {
                Assert.Fail($"Timed out waiting for MCP response {expected}. stderr: {string.Join(" | ", errors)}");
                return;
            }
            using var json = JsonDocument.Parse(line);
            if (json.RootElement.TryGetProperty("id", out var id) && id.GetInt32() == expected) return;
        }
    }
}
