// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OfficeCli.Core;

internal sealed record OfficePackagePartEvidence(
    string Path,
    string Sha256,
    long SizeBytes,
    string Role);

internal sealed record OfficeFidelitySnapshot(
    int SchemaVersion,
    string Format,
    string SourceFileName,
    string SourceSha256,
    IReadOnlyList<OfficePackagePartEvidence> Parts,
    IReadOnlyDictionary<string, int> Features);

internal sealed record OfficeFidelityFeatureChange(
    string Feature,
    int Before,
    int After,
    string Status);

internal sealed record OfficeChangeManifest(
    int SchemaVersion,
    string Format,
    string BeforeSha256,
    string AfterSha256,
    IReadOnlyList<string> ModifiedParts,
    IReadOnlyList<string> PreservedParts,
    IReadOnlyList<string> AddedParts,
    IReadOnlyList<string> RemovedParts,
    IReadOnlyList<OfficeFidelityFeatureChange> FeatureChanges,
    double FormatRetentionRate,
    double BytePreservationRate,
    bool Passed);

internal sealed record OfficeBrandAsset(
    string AssetId,
    string FileName,
    string Sha256,
    long SizeBytes,
    string PackagePath,
    string Role);

internal sealed record OfficeBrandSource(
    string Format,
    string FileName,
    string Sha256);

internal sealed record OfficeBrandProfile(
    int SchemaVersion,
    string ProfileId,
    string Version,
    string DisplayName,
    OfficeBrandSource Source,
    IReadOnlyDictionary<string, string> Colors,
    IReadOnlyDictionary<string, string> Fonts,
    IReadOnlyDictionary<string, object> Formats,
    IReadOnlyList<string> TypeScale,
    IReadOnlyList<OfficeBrandAsset> Assets,
    IReadOnlyDictionary<string, int> Evidence);

internal sealed record OfficeEvidenceReceipt(
    string Output,
    string? Theme = null,
    string? AssetDirectory = null,
    int? AssetCount = null,
    string? SourceSha256 = null,
    int? PartCount = null,
    IReadOnlyDictionary<string, int>? Features = null,
    bool? Passed = null,
    double? FormatRetentionRate = null,
    double? BytePreservationRate = null,
    int? Modified = null,
    int? Preserved = null,
    int? Removed = null);

internal static class OfficePackageEvidence
{
    internal static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    private static readonly Regex HexColor = new(
        "(?i)(?:val|lastClr|rgb|color)=\"#?([0-9a-f]{6})\"",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex PointSize = new(
        "(?i)(?:sz|fontSize)=\"([0-9]{2,5})\"",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly IReadOnlyDictionary<string, string> DefaultColors =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["dk1"] = "000000", ["lt1"] = "FFFFFF", ["dk2"] = "1F2937", ["lt2"] = "F3F4F6",
            ["accent1"] = "2563EB", ["accent2"] = "0F766E", ["accent3"] = "D97706",
            ["accent4"] = "7C3AED", ["accent5"] = "DC2626", ["accent6"] = "0891B2",
            ["hlink"] = "0563C1", ["folHlink"] = "954F72",
        };

    internal static OfficeFidelitySnapshot Snapshot(string filePath)
    {
        var fullPath = Path.GetFullPath(filePath);
        var format = ValidateOfficePath(fullPath);
        using var file = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read, leaveOpen: true);
        var parts = archive.Entries
            .Where(entry => !string.IsNullOrEmpty(entry.Name))
            .OrderBy(entry => Normalize(entry.FullName), StringComparer.Ordinal)
            .Select(entry => new OfficePackagePartEvidence(
                Normalize(entry.FullName), HashEntry(entry), entry.Length, ClassifyPart(format, Normalize(entry.FullName))))
            .ToArray();
        var features = ExtractFeatures(format, archive);
        return new OfficeFidelitySnapshot(
            1, format, Path.GetFileName(fullPath), HashFileShared(fullPath), parts, features);
    }

    internal static OfficeChangeManifest Diff(OfficeFidelitySnapshot before, string afterPath)
    {
        var after = Snapshot(afterPath);
        if (!string.Equals(before.Format, after.Format, StringComparison.Ordinal))
            throw new ArgumentException("fidelity snapshot and current file formats do not match");

        var left = before.Parts.ToDictionary(part => part.Path, StringComparer.Ordinal);
        var right = after.Parts.ToDictionary(part => part.Path, StringComparer.Ordinal);
        var modified = left.Keys.Intersect(right.Keys, StringComparer.Ordinal)
            .Where(path => left[path].Sha256 != right[path].Sha256).Order(StringComparer.Ordinal).ToArray();
        var preserved = left.Keys.Intersect(right.Keys, StringComparer.Ordinal)
            .Where(path => left[path].Sha256 == right[path].Sha256).Order(StringComparer.Ordinal).ToArray();
        var added = right.Keys.Except(left.Keys, StringComparer.Ordinal).Order(StringComparer.Ordinal).ToArray();
        var removed = left.Keys.Except(right.Keys, StringComparer.Ordinal).Order(StringComparer.Ordinal).ToArray();

        var featureNames = before.Features.Keys.Union(after.Features.Keys, StringComparer.Ordinal).Order(StringComparer.Ordinal);
        var featureChanges = featureNames.Select(name =>
        {
            before.Features.TryGetValue(name, out var oldValue);
            after.Features.TryGetValue(name, out var newValue);
            var status = newValue == oldValue ? "preserved" : newValue > oldValue ? "added" : "lost";
            return new OfficeFidelityFeatureChange(name, oldValue, newValue, status);
        }).ToArray();

        var protectedBefore = before.Parts.Where(part => part.Role != "content").Select(part => part.Path).ToHashSet(StringComparer.Ordinal);
        var protectedRetained = protectedBefore.Count(right.ContainsKey);
        var protectedPreserved = protectedBefore.Count(path => right.TryGetValue(path, out var part) && part.Sha256 == left[path].Sha256);
        var retention = protectedBefore.Count == 0 ? 1d : Math.Round((double)protectedRetained / protectedBefore.Count, 6);
        var bytePreservation = protectedBefore.Count == 0 ? 1d : Math.Round((double)protectedPreserved / protectedBefore.Count, 6);
        var passed = removed.Length == 0 && featureChanges.All(change => change.Status != "lost");
        return new OfficeChangeManifest(
            1, before.Format, before.SourceSha256, after.SourceSha256,
            modified, preserved, added, removed, featureChanges, retention, bytePreservation, passed);
    }

    internal static (OfficeBrandProfile Profile, Dictionary<string, object> Theme) ExtractBrand(
        string filePath,
        string profileId,
        string displayName,
        string assetDirectory)
    {
        var fullPath = Path.GetFullPath(filePath);
        var snapshot = Snapshot(fullPath);
        using var file = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(file, ZipArchiveMode.Read, leaveOpen: false);

        var themeEntry = archive.Entries.FirstOrDefault(entry =>
            Normalize(entry.FullName).Contains("/theme/theme", StringComparison.OrdinalIgnoreCase));
        var themeXml = themeEntry is null ? "" : ReadEntryText(themeEntry);
        var colors = ExtractThemeColors(themeXml);
        var fonts = ExtractThemeFonts(themeXml);

        var colorCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var sizeCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var entry in archive.Entries.Where(entry => entry.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            var xml = ReadEntryText(entry);
            foreach (Match match in HexColor.Matches(xml))
            {
                var value = match.Groups[1].Value.ToUpperInvariant();
                colorCounts[value] = colorCounts.GetValueOrDefault(value) + 1;
            }
            foreach (var points in ExtractPointSizes(snapshot.Format, Normalize(entry.FullName), xml))
                sizeCounts[points] = sizeCounts.GetValueOrDefault(points) + 1;
        }

        var preferred = colorCounts.OrderByDescending(pair => pair.Value).ThenBy(pair => pair.Key, StringComparer.Ordinal).Select(pair => pair.Key).ToArray();
        var accentKeys = Enumerable.Range(1, 6).Select(index => $"accent{index}").ToArray();
        if (themeEntry is null)
            for (var index = 0; index < accentKeys.Length && index < preferred.Length; index++)
                colors[accentKeys[index]] = preferred[index];

        Directory.CreateDirectory(assetDirectory);
        var assets = new List<OfficeBrandAsset>();
        var logoMedia = FindLogoMediaCandidates(archive, snapshot.Format);
        foreach (var entry in archive.Entries.Where(IsBrandMediaCandidate).OrderBy(entry => Normalize(entry.FullName), StringComparer.Ordinal))
        {
            var bytes = ReadEntryBytes(entry);
            var digest = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
            var extension = Path.GetExtension(entry.Name).ToLowerInvariant();
            var fileName = $"asset-{assets.Count + 1:D2}-{digest[..12]}{extension}";
            var target = Path.Combine(assetDirectory, fileName);
            File.WriteAllBytes(target, bytes);
            assets.Add(new OfficeBrandAsset(
                $"brand-asset-{assets.Count + 1:D2}", fileName, digest, bytes.LongLength,
                Normalize(entry.FullName), logoMedia.Contains(Normalize(entry.FullName))
                    ? "logo"
                    : ClassifyBrandAssetRole(snapshot.Format)));
        }

        var formats = BuildFormatProfile(snapshot, archive);
        var typeScale = sizeCounts.OrderByDescending(pair => pair.Value).ThenByDescending(pair => double.Parse(pair.Key))
            .Take(12).Select(pair => pair.Key + "pt").ToArray();
        var profile = new OfficeBrandProfile(
            1, profileId, "1.0.0", displayName,
            new OfficeBrandSource(snapshot.Format, snapshot.SourceFileName, snapshot.SourceSha256),
            colors, fonts, formats, typeScale, assets, snapshot.Features);
        var theme = new Dictionary<string, object>(StringComparer.Ordinal)
        {
            ["schemaVersion"] = 1,
            ["themeId"] = profileId,
            ["version"] = "1.0.0",
            ["displayName"] = displayName,
            ["colors"] = colors,
            ["fonts"] = fonts,
            ["formats"] = formats,
        };
        return (profile, theme);
    }

    internal static OfficeFidelitySnapshot ReadSnapshot(string path)
    {
        var snapshot = JsonSerializer.Deserialize(File.ReadAllText(path), OfficeEvidenceJsonContext.Default.OfficeFidelitySnapshot);
        if (snapshot is null || snapshot.SchemaVersion != 1 || snapshot.Parts.Count == 0)
            throw new ArgumentException("invalid Office fidelity snapshot");
        return snapshot;
    }

    internal static void WriteJson<T>(string path, T value)
    {
        var fullPath = Path.GetFullPath(path);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var typeInfo = OfficeEvidenceJsonContext.Default.GetTypeInfo(typeof(T))
            ?? throw new InvalidOperationException($"Office evidence JSON type is not registered: {typeof(T).Name}");
        File.WriteAllText(fullPath, JsonSerializer.Serialize(value, typeInfo) + Environment.NewLine);
    }

    private static string ValidateOfficePath(string fullPath)
    {
        if (!File.Exists(fullPath)) throw new FileNotFoundException("Office file not found", fullPath);
        var format = Path.GetExtension(fullPath).TrimStart('.').ToLowerInvariant();
        if (format is not ("docx" or "xlsx" or "pptx"))
            throw new ArgumentException("brand and fidelity evidence support DOCX, XLSX and PPTX files");
        return format;
    }

    private static Dictionary<string, string> ExtractThemeColors(string xml)
    {
        var colors = new Dictionary<string, string>(DefaultColors, StringComparer.Ordinal);
        if (string.IsNullOrWhiteSpace(xml)) return colors;
        try
        {
            var root = XDocument.Parse(xml).Root;
            var scheme = root?.Descendants().FirstOrDefault(node => node.Name.LocalName == "clrScheme");
            if (scheme is null) return colors;
            foreach (var item in scheme.Elements())
            {
                var valueNode = item.Elements().FirstOrDefault();
                var value = valueNode?.Attribute("val")?.Value ?? valueNode?.Attribute("lastClr")?.Value;
                if (value is not null && Regex.IsMatch(value, "^[0-9A-Fa-f]{6}$"))
                    colors[item.Name.LocalName] = value.ToUpperInvariant();
            }
        }
        catch { }
        return colors;
    }

    private static Dictionary<string, string> ExtractThemeFonts(string xml)
    {
        var fonts = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["majorLatin"] = "Aptos Display", ["majorEastAsia"] = "Microsoft YaHei",
            ["minorLatin"] = "Aptos", ["minorEastAsia"] = "Microsoft YaHei",
        };
        if (string.IsNullOrWhiteSpace(xml)) return fonts;
        try
        {
            var root = XDocument.Parse(xml).Root;
            var scheme = root?.Descendants().FirstOrDefault(node => node.Name.LocalName == "fontScheme");
            foreach (var (containerName, prefix) in new[] { ("majorFont", "major"), ("minorFont", "minor") })
            {
                var container = scheme?.Elements().FirstOrDefault(node => node.Name.LocalName == containerName);
                var latin = container?.Elements().FirstOrDefault(node => node.Name.LocalName == "latin")?.Attribute("typeface")?.Value;
                var eastAsia = container?.Elements().FirstOrDefault(node => node.Name.LocalName == "ea")?.Attribute("typeface")?.Value;
                if (!string.IsNullOrWhiteSpace(latin)) fonts[prefix + "Latin"] = latin;
                if (!string.IsNullOrWhiteSpace(eastAsia)) fonts[prefix + "EastAsia"] = eastAsia;
            }
        }
        catch { }
        return fonts;
    }

    private static Dictionary<string, object> BuildFormatProfile(OfficeFidelitySnapshot snapshot, ZipArchive archive)
    {
        var paragraphStyles = ExtractAttributeValues(archive, path => path == "word/document.xml", "pStyle", "val");
        var tableStyles = ExtractAttributeValues(archive, path => path == "word/document.xml", "tblStyle", "val");
        var numberFormats = ExtractAttributeValues(archive, path => path == "xl/styles.xml", "numFmt", "formatCode");
        var workbookStyles = ExtractAttributeValues(archive, path => path == "xl/styles.xml", "cellStyle", "name");
        var chartTypes = ExtractElementNames(archive, path => path.StartsWith("xl/charts/chart", StringComparison.OrdinalIgnoreCase), name => name.EndsWith("Chart", StringComparison.Ordinal));
        var layoutNames = ExtractAttributeValues(archive, path => path.StartsWith("ppt/slideLayouts/", StringComparison.OrdinalIgnoreCase), "cSld", "name");
        var placeholderTypes = ExtractAttributeValues(archive, path => path.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase), "ph", "type");
        return new Dictionary<string, object>(StringComparer.Ordinal)
        {
            ["sourceFormat"] = snapshot.Format,
            ["docx"] = new Dictionary<string, object>
            {
                ["inheritSourceStyles"] = snapshot.Format == "docx",
                ["paragraphStyles"] = paragraphStyles,
                ["tableStyles"] = tableStyles,
                ["sectionCount"] = snapshot.Features.GetValueOrDefault("sections"),
                ["headerCount"] = snapshot.Features.GetValueOrDefault("headers"),
                ["footerCount"] = snapshot.Features.GetValueOrDefault("footers"),
                ["tocFieldCount"] = snapshot.Features.GetValueOrDefault("tocFields"),
            },
            ["xlsx"] = new Dictionary<string, object>
            {
                ["inheritNumberFormats"] = snapshot.Format == "xlsx",
                ["numberFormats"] = numberFormats,
                ["cellStyles"] = workbookStyles,
                ["chartTypes"] = chartTypes,
                ["definedNameCount"] = snapshot.Features.GetValueOrDefault("definedNames"),
                ["printAreaCount"] = snapshot.Features.GetValueOrDefault("printAreas"),
            },
            ["pptx"] = new Dictionary<string, object>
            {
                ["inheritMasterAndLayouts"] = snapshot.Format == "pptx",
                ["masterCount"] = snapshot.Features.GetValueOrDefault("slideMasters"),
                ["layoutCount"] = snapshot.Features.GetValueOrDefault("slideLayouts"),
                ["layoutNames"] = layoutNames,
                ["placeholderTypes"] = placeholderTypes,
                ["notesCount"] = snapshot.Features.GetValueOrDefault("notes"),
            },
            ["pdf"] = new Dictionary<string, object> { ["deriveFromEditableSource"] = true, ["preserveTheme"] = true },
        };
    }

    private static IReadOnlyList<string> ExtractAttributeValues(
        ZipArchive archive,
        Func<string, bool> pathPredicate,
        string localName,
        string attributeLocalName)
    {
        var values = new SortedSet<string>(StringComparer.Ordinal);
        foreach (var entry in archive.Entries.Where(entry => pathPredicate(Normalize(entry.FullName))))
        {
            try
            {
                var root = XDocument.Parse(ReadEntryText(entry));
                foreach (var node in root.Descendants().Where(node => node.Name.LocalName == localName))
                {
                    var value = node.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == attributeLocalName)?.Value;
                    if (!string.IsNullOrWhiteSpace(value)) values.Add(value.Trim());
                    if (values.Count >= 64) return values.ToArray();
                }
            }
            catch { }
        }
        return values.ToArray();
    }

    private static IReadOnlyList<string> ExtractElementNames(
        ZipArchive archive,
        Func<string, bool> pathPredicate,
        Func<string, bool> namePredicate)
    {
        var values = new SortedSet<string>(StringComparer.Ordinal);
        foreach (var entry in archive.Entries.Where(entry => pathPredicate(Normalize(entry.FullName))))
        {
            try
            {
                var root = XDocument.Parse(ReadEntryText(entry));
                foreach (var value in root.Descendants().Select(node => node.Name.LocalName).Where(namePredicate))
                    values.Add(value);
            }
            catch { }
        }
        return values.ToArray();
    }

    private static Dictionary<string, int> ExtractFeatures(string format, ZipArchive archive)
    {
        var paths = archive.Entries.Select(entry => Normalize(entry.FullName)).ToArray();
        int Count(string prefix, string suffix = ".xml") => paths.Count(path => path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase));
        var result = new Dictionary<string, int>(StringComparer.Ordinal)
        {
            ["themes"] = paths.Count(path => path.Contains("/theme/theme", StringComparison.OrdinalIgnoreCase)),
            ["media"] = paths.Count(path => path.Contains("/media/", StringComparison.OrdinalIgnoreCase)),
            ["charts"] = paths.Count(path => path.Contains("/charts/chart", StringComparison.OrdinalIgnoreCase) && path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)),
        };
        var smartArtDataPaths = paths.Where(path => path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            && (path.Contains("/diagrams/data", StringComparison.OrdinalIgnoreCase)
                || path.Contains("/graphics/data", StringComparison.OrdinalIgnoreCase)
                || path.StartsWith("graphics/data", StringComparison.OrdinalIgnoreCase))).ToArray();
        result["smartArtDataParts"] = smartArtDataPaths.Length;
        result["smartArtLayoutParts"] = paths.Count(path => path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            && (path.Contains("/diagrams/layout", StringComparison.OrdinalIgnoreCase)
                || path.Contains("/graphics/layout", StringComparison.OrdinalIgnoreCase)
                || path.StartsWith("graphics/layout", StringComparison.OrdinalIgnoreCase)));
        result["smartArtColorParts"] = paths.Count(path => path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            && (path.Contains("/diagrams/colors", StringComparison.OrdinalIgnoreCase)
                || path.Contains("/graphics/colors", StringComparison.OrdinalIgnoreCase)
                || path.StartsWith("graphics/colors", StringComparison.OrdinalIgnoreCase)));
        result["smartArtStyleParts"] = paths.Count(path => path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            && (path.Contains("/diagrams/quickStyle", StringComparison.OrdinalIgnoreCase)
                || path.Contains("/graphics/quickStyle", StringComparison.OrdinalIgnoreCase)
                || path.StartsWith("graphics/quickStyle", StringComparison.OrdinalIgnoreCase)));
        result["smartArtNodes"] = smartArtDataPaths.Sum(path => CountXmlNodes(archive, path, "pt"));
        result["smartArtConnections"] = smartArtDataPaths.Sum(path => CountXmlNodes(archive, path, "cxn"));
        if (format == "docx")
        {
            result["styles"] = CountXmlNodes(archive, "word/styles.xml", "style");
            result["sections"] = CountXmlNodes(archive, "word/document.xml", "sectPr");
            result["headers"] = Count("word/header"); result["footers"] = Count("word/footer");
            result["comments"] = CountXmlNodes(archive, "word/comments.xml", "comment");
            result["trackedChanges"] = CountXmlNodes(archive, "word/document.xml", "ins") + CountXmlNodes(archive, "word/document.xml", "del") + CountXmlNodes(archive, "word/document.xml", "pPrChange") + CountXmlNodes(archive, "word/document.xml", "rPrChange");
            result["tocFields"] = CountXmlTextMatches(archive, "word/document.xml", "instrText", @"\bTOC\b");
        }
        else if (format == "xlsx")
        {
            result["sheets"] = Count("xl/worksheets/sheet");
            result["formulaCells"] = paths.Where(path => path.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)).Sum(path => CountXmlNodes(archive, path, "f"));
            result["definedNames"] = CountXmlNodes(archive, "xl/workbook.xml", "definedName");
            result["conditionalFormats"] = paths.Where(path => path.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)).Sum(path => CountXmlNodes(archive, path, "conditionalFormatting"));
            result["dataValidations"] = paths.Where(path => path.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)).Sum(path => CountXmlNodes(archive, path, "dataValidation"));
            result["printAreas"] = CountXmlNodes(archive, "xl/workbook.xml", "definedName", "_xlnm.Print_Area");
        }
        else
        {
            result["slides"] = Count("ppt/slides/slide"); result["slideMasters"] = Count("ppt/slideMasters/slideMaster");
            result["slideLayouts"] = Count("ppt/slideLayouts/slideLayout"); result["notes"] = Count("ppt/notesSlides/notesSlide");
            result["animations"] = paths.Where(path => path.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase)).Sum(path => CountXmlNodes(archive, path, "timing"));
            result["placeholders"] = paths.Where(path => path.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase) && path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)).Sum(path => CountXmlNodes(archive, path, "ph"));
        }
        return result;
    }

    private static int CountXmlNodes(ZipArchive archive, string path, string localName, string? requiredAttributeValue = null)
    {
        var entry = archive.GetEntry(path);
        if (entry is null) return 0;
        try
        {
            var root = XDocument.Parse(ReadEntryText(entry));
            return root.Descendants().Count(node => node.Name.LocalName == localName &&
                (requiredAttributeValue is null || node.Attributes().Any(attribute => attribute.Value == requiredAttributeValue)));
        }
        catch { return 0; }
    }

    private static int CountXmlTextMatches(ZipArchive archive, string path, string localName, string pattern)
    {
        var entry = archive.GetEntry(path);
        if (entry is null) return 0;
        try
        {
            var root = XDocument.Parse(ReadEntryText(entry));
            return root.Descendants().Count(node => node.Name.LocalName == localName && Regex.IsMatch(node.Value, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
        }
        catch { return 0; }
    }

    private static IEnumerable<string> ExtractPointSizes(string format, string packagePath, string xml)
    {
        foreach (Match match in PointSize.Matches(xml))
        {
            if (!int.TryParse(match.Groups[1].Value, out var number) || number <= 0) continue;
            double points = format switch
            {
                "docx" when packagePath.StartsWith("word/", StringComparison.OrdinalIgnoreCase) => number / 2d,
                "pptx" when packagePath.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase) => number / 100d,
                "xlsx" when packagePath.StartsWith("xl/", StringComparison.OrdinalIgnoreCase) => number,
                _ => number >= 100 ? number / 100d : number,
            };
            if (points is < 4 or > 200) continue;
            yield return points.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    private static string ClassifyPart(string format, string path)
    {
        if (path is "[Content_Types].xml" || path.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) return "package";
        if (path.Contains("/theme/", StringComparison.OrdinalIgnoreCase) || path.Contains("/styles", StringComparison.OrdinalIgnoreCase) || path.Contains("/diagrams/", StringComparison.OrdinalIgnoreCase) || path.Contains("/graphics/", StringComparison.OrdinalIgnoreCase) || path.StartsWith("graphics/", StringComparison.OrdinalIgnoreCase) || path.Contains("/slideMasters/", StringComparison.OrdinalIgnoreCase) || path.Contains("/slideLayouts/", StringComparison.OrdinalIgnoreCase) || path.Contains("/header", StringComparison.OrdinalIgnoreCase) || path.Contains("/footer", StringComparison.OrdinalIgnoreCase) || path.Contains("/comments", StringComparison.OrdinalIgnoreCase)) return "design";
        if ((format == "docx" && path == "word/document.xml") || (format == "xlsx" && path.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)) || (format == "pptx" && path.StartsWith("ppt/slides/", StringComparison.OrdinalIgnoreCase))) return "content";
        return "structure";
    }

    private static bool IsBrandMediaCandidate(ZipArchiveEntry entry) =>
        !string.IsNullOrEmpty(entry.Name) && Normalize(entry.FullName).Contains("/media/", StringComparison.OrdinalIgnoreCase) && entry.Length <= 16 * 1024 * 1024;

    private static string ClassifyBrandAssetRole(string format)
    {
        if (format == "docx") return "word-media-candidate";
        if (format == "pptx") return "presentation-media-candidate";
        return "workbook-media-candidate";
    }

    private static HashSet<string> FindLogoMediaCandidates(ZipArchive archive, string format)
    {
        var candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in archive.Entries.Where(entry => entry.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            var path = Normalize(entry.FullName);
            var stableBrandHost = format switch
            {
                "pptx" => path.StartsWith("ppt/slideMasters/_rels/", StringComparison.OrdinalIgnoreCase)
                    || path.StartsWith("ppt/slideLayouts/_rels/", StringComparison.OrdinalIgnoreCase),
                "docx" => path.StartsWith("word/_rels/header", StringComparison.OrdinalIgnoreCase)
                    || path.StartsWith("word/_rels/footer", StringComparison.OrdinalIgnoreCase),
                _ => false,
            };
            if (!stableBrandHost) continue;
            XDocument relationships;
            try { relationships = XDocument.Parse(ReadEntryText(entry)); }
            catch { continue; }
            var marker = path.IndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
            if (marker < 0) continue;
            var ownerDirectory = path[..marker];
            foreach (var relationship in relationships.Descendants().Where(element => element.Name.LocalName == "Relationship"))
            {
                var type = relationship.Attribute("Type")?.Value ?? "";
                var target = relationship.Attribute("Target")?.Value ?? "";
                var targetMode = relationship.Attribute("TargetMode")?.Value ?? "";
                if (!type.EndsWith("/image", StringComparison.OrdinalIgnoreCase)
                    || targetMode.Equals("External", StringComparison.OrdinalIgnoreCase))
                    continue;
                var resolved = ResolvePackageTarget(ownerDirectory, target);
                if (resolved.Contains("/media/", StringComparison.OrdinalIgnoreCase)) candidates.Add(resolved);
            }
        }
        return candidates;
    }

    private static string ResolvePackageTarget(string ownerDirectory, string target)
    {
        var segments = new List<string>(ownerDirectory.Split('/', StringSplitOptions.RemoveEmptyEntries));
        foreach (var segment in Normalize(target).Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".") continue;
            if (segment == "..")
            {
                if (segments.Count > 0) segments.RemoveAt(segments.Count - 1);
                continue;
            }
            segments.Add(segment);
        }
        return string.Join('/', segments);
    }

    private static string HashEntry(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private static string HashFileShared(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private static string ReadEntryText(ZipArchiveEntry entry)
    {
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    private static byte[] ReadEntryBytes(ZipArchiveEntry entry)
    {
        using var input = entry.Open();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static string Normalize(string path) => path.Replace('\\', '/');
}

[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    WriteIndented = true,
    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull)]
[JsonSerializable(typeof(OfficeFidelitySnapshot))]
[JsonSerializable(typeof(OfficeChangeManifest))]
[JsonSerializable(typeof(OfficeBrandProfile))]
[JsonSerializable(typeof(OfficeEvidenceReceipt))]
[JsonSerializable(typeof(Dictionary<string, object>))]
internal partial class OfficeEvidenceJsonContext : JsonSerializerContext;
