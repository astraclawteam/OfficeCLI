using System.IO.Compression;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using Xunit;

namespace OfficeCli.Tests;

public sealed class OfficeEvidenceTests
{
    [Fact]
    public void FidelityManifestSeparatesContentChangeFromPreservedPackageEvidence()
    {
        using var temp = new TempDirectory();
        var path = temp.File("report.docx");
        var snapshotPath = temp.File("snapshot.json");
        OpenXmlFixture.CreateDocument(path,
            new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text("before"))));
        var snapshot = OfficePackageEvidence.Snapshot(path);
        OfficePackageEvidence.WriteJson(snapshotPath, snapshot);

        OpenXmlFixture.ReplaceEntryText(path, "word/document.xml", xml => xml.Replace("before", "after", StringComparison.Ordinal));
        var manifest = OfficePackageEvidence.Diff(OfficePackageEvidence.ReadSnapshot(snapshotPath), path);

        Assert.Contains("word/document.xml", manifest.ModifiedParts);
        Assert.Contains("[Content_Types].xml", manifest.PreservedParts);
        Assert.Empty(manifest.RemovedParts);
        Assert.Equal(1d, manifest.FormatRetentionRate);
        Assert.True(manifest.Passed);
    }

    [Fact]
    public void FidelityManifestBlocksSilentLossOfExistingParts()
    {
        using var temp = new TempDirectory();
        var path = temp.File("report.docx");
        OpenXmlFixture.CreateDocument(path,
            new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text("body"))));
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var comments = archive.CreateEntry("word/comments.xml");
            using var writer = new StreamWriter(comments.Open());
            writer.Write("<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:comment w:id=\"0\"/></w:comments>");
        }
        var snapshot = OfficePackageEvidence.Snapshot(path);
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
            archive.GetEntry("word/comments.xml")!.Delete();

        var manifest = OfficePackageEvidence.Diff(snapshot, path);

        Assert.Contains("word/comments.xml", manifest.RemovedParts);
        Assert.Contains(manifest.FeatureChanges, change => change.Feature == "comments" && change.Status == "lost");
        Assert.False(manifest.Passed);
    }

    [Fact]
    public void BrandExtractionProducesLocalThemeAndSourceBoundProfile()
    {
        using var temp = new TempDirectory();
        var path = temp.File("template.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Dashboard");
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var media = archive.CreateEntry("xl/media/image1.png");
            using var stream = media.Open();
            stream.Write(new byte[] { 0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a });
        }

        var (profile, theme) = OfficePackageEvidence.ExtractBrand(
            path, "acme-brand", "ACME 品牌", temp.File("assets"));

        Assert.Equal("xlsx", profile.Source.Format);
        Assert.Equal(64, profile.Source.Sha256.Length);
        Assert.Equal(12, profile.Colors.Count);
        Assert.Equal(4, profile.Fonts.Count);
        Assert.True(profile.Formats.ContainsKey("docx"));
        Assert.True(profile.Formats.ContainsKey("pptx"));
        var asset = Assert.Single(profile.Assets);
        Assert.Equal("workbook-media-candidate", asset.Role);
        Assert.True(File.Exists(Path.Combine(temp.File("assets"), asset.FileName)));
        Assert.Equal("acme-brand", theme["themeId"]);
        var serialized = JsonSerializer.Serialize(profile, OfficeEvidenceJsonContext.Default.OfficeBrandProfile);
        Assert.DoesNotContain(temp.Path, serialized, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BrandExtractionMarksMasterAndHeaderImagesAsSourceBoundLogos()
    {
        using var temp = new TempDirectory();
        var path = temp.File("template.pptx");
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            var types = archive.CreateEntry("[Content_Types].xml");
            using (var writer = new StreamWriter(types.Open()))
                writer.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
            var logo = archive.CreateEntry("ppt/media/logo.png");
            using (var stream = logo.Open())
                stream.Write(new byte[] { 0x89, 0x50, 0x4e, 0x47, 1, 2, 3, 4 });
            var rels = archive.CreateEntry("ppt/slideMasters/_rels/slideMaster1.xml.rels");
            using (var writer = new StreamWriter(rels.Open()))
                writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/logo.png\"/></Relationships>");
        }

        var (profile, _) = OfficePackageEvidence.ExtractBrand(path, "brand-with-logo", "带 Logo 品牌", temp.File("assets"));

        var logoAsset = Assert.Single(profile.Assets);
        Assert.Equal("logo", logoAsset.Role);
        Assert.Equal("ppt/media/logo.png", logoAsset.PackagePath);
    }

    [Fact]
    public void FormulaLossIsAReleaseBlockingFidelityChange()
    {
        using var temp = new TempDirectory();
        var path = temp.File("book.xlsx");
        OpenXmlFixture.CreateWorkbook(path, "Sheet1");
        using (var document = SpreadsheetDocument.Open(path, true))
        {
            var worksheet = document.WorkbookPart?.WorksheetParts.Single().Worksheet
                ?? throw new InvalidOperationException("fixture worksheet is missing");
            var sheetData = worksheet.GetFirstChild<SheetData>()
                ?? throw new InvalidOperationException("fixture sheet data is missing");
            sheetData.Append(new Row(new Cell(new CellFormula("SUM(A1:A2)"), new CellValue("3"))));
        }
        var snapshot = OfficePackageEvidence.Snapshot(path);
        using (var document = SpreadsheetDocument.Open(path, true))
        {
            var worksheet = document.WorkbookPart?.WorksheetParts.Single().Worksheet
                ?? throw new InvalidOperationException("fixture worksheet is missing");
            worksheet.Descendants<CellFormula>().Single().Remove();
        }

        var manifest = OfficePackageEvidence.Diff(snapshot, path);

        Assert.False(manifest.Passed);
        Assert.Contains(manifest.FeatureChanges, item => item.Feature == "formulaCells" && item.Status == "lost");
    }

    [Fact]
    public void IntentionalStyleChangesRetainStructureButAreNotReportedAsByteIdentical()
    {
        using var temp = new TempDirectory();
        var path = temp.File("styled.docx");
        OpenXmlFixture.CreateDocument(path,
            new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text("body"))));
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var styles = archive.CreateEntry("word/styles.xml");
            using var writer = new StreamWriter(styles.Open());
            writer.Write("<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:style w:type=\"paragraph\" w:styleId=\"BrandTitle\"><w:name w:val=\"Brand title\"/></w:style></w:styles>");
        }
        var snapshot = OfficePackageEvidence.Snapshot(path);
        OpenXmlFixture.ReplaceEntryText(path, "word/styles.xml", xml => xml.Replace("Brand title", "Updated brand title", StringComparison.Ordinal));

        var manifest = OfficePackageEvidence.Diff(snapshot, path);

        Assert.True(manifest.Passed);
        Assert.Equal(1d, manifest.FormatRetentionRate);
        Assert.True(manifest.BytePreservationRate < 1d);
        Assert.Contains("word/styles.xml", manifest.ModifiedParts);
    }
}
