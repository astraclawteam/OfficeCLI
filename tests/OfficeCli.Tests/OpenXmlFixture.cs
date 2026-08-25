using System.IO.Compression;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Tests;

internal sealed class TempDirectory : IDisposable
{
    public string Path { get; } = System.IO.Path.Combine(
        System.IO.Path.GetTempPath(), $"officecli-tests-{Guid.NewGuid():N}");

    public TempDirectory() => Directory.CreateDirectory(Path);

    public string File(string name) => System.IO.Path.Combine(Path, name);

    public void Dispose()
    {
        try { Directory.Delete(Path, recursive: true); }
        catch { }
    }
}

internal static class OpenXmlFixture
{
    public static void CreateWorkbook(string path, params string[] sheetNames)
    {
        using var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook(new Sheets());
        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()!;
        uint id = 1;
        foreach (var name in sheetNames)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            sheets.Append(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = id++,
                Name = name,
            });
        }
        workbookPart.Workbook.Save();
    }

    public static void CreateDocument(string path, params OpenXmlElement[] bodyChildren)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body(bodyChildren));
        main.Document.Save();
    }

    public static string EntryText(string path, string entryName)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry(entryName)
            ?? throw new InvalidOperationException($"Missing ZIP entry: {entryName}");
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    public static string EntryTextShared(string path, string entryName)
    {
        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read, leaveOpen: false);
        var entry = archive.GetEntry(entryName)
            ?? throw new InvalidOperationException($"Missing ZIP entry: {entryName}");
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    public static string FileSha256Shared(string path)
    {
        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Convert.ToHexString(SHA256.HashData(file));
    }

    public static string EntrySha256(string path, string entryName)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.GetEntry(entryName)
            ?? throw new InvalidOperationException($"Missing ZIP entry: {entryName}");
        using var stream = entry.Open();
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    public static string SingleEntrySha256(string path, string entryPrefix)
    {
        using var archive = ZipFile.OpenRead(path);
        var entry = archive.Entries.Single(e =>
            e.FullName.StartsWith(entryPrefix, StringComparison.OrdinalIgnoreCase));
        using var stream = entry.Open();
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    public static void ReplaceEntryText(string path, string entryName, Func<string, string> transform)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        var entry = archive.GetEntry(entryName)
            ?? throw new InvalidOperationException($"Missing ZIP entry: {entryName}");
        string original;
        using (var reader = new StreamReader(entry.Open())) original = reader.ReadToEnd();
        entry.Delete();
        var replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var writer = new StreamWriter(replacement.Open());
        writer.Write(transform(original));
    }
}
