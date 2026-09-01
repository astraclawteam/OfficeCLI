// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeCli.Core;

public sealed class ProfessionalPageSpec
{
    [JsonPropertyName("schemaVersion")] public int SchemaVersion { get; set; } = 1;
    [JsonPropertyName("documentId")] public string DocumentId { get; set; } = "";
    [JsonPropertyName("format")] public string Format { get; set; } = "";
    [JsonPropertyName("title")] public string Title { get; set; } = "";
    [JsonPropertyName("brandTokens")] public Dictionary<string, string> BrandTokens { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    [JsonPropertyName("pages")] public List<ProfessionalPage> Pages { get; set; } = [];
}

public sealed class ProfessionalPage
{
    [JsonPropertyName("pageId")] public string PageId { get; set; } = "";
    [JsonPropertyName("pageRole")] public string PageRole { get; set; } = "";
    [JsonPropertyName("primaryClaim")] public string PrimaryClaim { get; set; } = "";
    [JsonPropertyName("readerTakeaway")] public string ReaderTakeaway { get; set; } = "";
    [JsonPropertyName("readerAction")] public string ReaderAction { get; set; } = "";
    [JsonPropertyName("density")] public string Density { get; set; } = "balanced";
    [JsonPropertyName("blocks")] public List<ProfessionalPageBlock> Blocks { get; set; } = [];
}

public sealed class ProfessionalPageBlock
{
    [JsonPropertyName("blockId")] public string BlockId { get; set; } = "";
    [JsonPropertyName("kind")] public string Kind { get; set; } = "";
    [JsonPropertyName("importance")] public string Importance { get; set; } = "supporting";
    [JsonPropertyName("title")] public string? Title { get; set; }
    [JsonPropertyName("text")] public string? Text { get; set; }
    [JsonPropertyName("factRefs")] public List<string> FactRefs { get; set; } = [];
    [JsonPropertyName("claimRefs")] public List<string> ClaimRefs { get; set; } = [];
    [JsonPropertyName("decisionRefs")] public List<string> DecisionRefs { get; set; } = [];
    [JsonPropertyName("actionRefs")] public List<string> ActionRefs { get; set; } = [];
    [JsonPropertyName("component")] public ProfessionalComponentSpec? Component { get; set; }
    [JsonPropertyName("chart")] public InformationChartSpec? Chart { get; set; }
}

public sealed record ProfessionalPageBlockReceipt(string BlockId, string Kind, string NativeObjectPath,
    IReadOnlyList<string> FactRefs, IReadOnlyList<string> ClaimRefs, IReadOnlyList<string> DecisionRefs, IReadOnlyList<string> ActionRefs);
public sealed record ProfessionalPageReceipt(string PageId, string PageRole, string HostLocation, IReadOnlyList<ProfessionalPageBlockReceipt> Blocks);
public sealed record ProfessionalCompositionReceipt(bool Ok, string File, string Format, string DocumentId,
    string Composer, IReadOnlyList<ProfessionalPageReceipt> Pages);

public static class ProfessionalPageCompiler
{
    private static readonly HashSet<string> BlockKinds = new(StringComparer.OrdinalIgnoreCase)
    { "narrative", "component", "chart", "decision", "action", "source", "image" };

    public static ProfessionalPageSpec Parse(string path)
    {
        ProfessionalPageSpec? spec;
        try { spec = JsonSerializer.Deserialize(File.ReadAllText(path), ProfessionalPageJsonContext.Default.ProfessionalPageSpec); }
        catch (Exception ex) when (ex is IOException or JsonException)
        { throw new CliException($"PageSpec could not be read: {ex.Message}") { Code = "page_spec_invalid" }; }
        if (spec is null || spec.SchemaVersion != 1 || string.IsNullOrWhiteSpace(spec.DocumentId)
            || string.IsNullOrWhiteSpace(spec.Title) || spec.Pages.Count == 0 || spec.Format is not ("docx" or "xlsx" or "pptx"))
            throw new CliException("PageSpec requires schemaVersion=1, documentId, format, title and pages.") { Code = "page_spec_invalid" };
        var pageIds = new HashSet<string>(StringComparer.Ordinal);
        var blockIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var page in spec.Pages)
        {
            if (string.IsNullOrWhiteSpace(page.PageId) || !pageIds.Add(page.PageId) || string.IsNullOrWhiteSpace(page.PrimaryClaim)
                || string.IsNullOrWhiteSpace(page.ReaderTakeaway) || string.IsNullOrWhiteSpace(page.ReaderAction) || page.Blocks.Count == 0)
                throw new CliException("Every PageSpec page requires a unique pageId, claim, takeaway, action and blocks.") { Code = "page_spec_invalid" };
            if (page.Density is not ("compact" or "balanced" or "spacious"))
                throw new CliException("PageSpec page density must be compact, balanced or spacious.") { Code = "page_spec_invalid" };
            foreach (var block in page.Blocks)
            {
                if (string.IsNullOrWhiteSpace(block.BlockId) || !blockIds.Add(block.BlockId) || !BlockKinds.Contains(block.Kind))
                    throw new CliException("Every PageSpec block requires a unique blockId and supported kind.") { Code = "page_spec_invalid" };
                if (block.Kind == "component" && block.Component is null)
                    throw new CliException($"PageSpec block {block.BlockId} requires component.") { Code = "page_spec_invalid" };
                if (block.Kind == "chart" && block.Chart is null)
                    throw new CliException($"PageSpec block {block.BlockId} requires chart.") { Code = "page_spec_invalid" };
                if (block.Kind is not ("component" or "chart") && string.IsNullOrWhiteSpace(block.Text))
                    throw new CliException($"PageSpec block {block.BlockId} requires text.") { Code = "page_spec_invalid" };
            }
        }
        return spec;
    }

    public static ProfessionalCompositionReceipt Compile(IDocumentHandler handler, string filePath, ProfessionalPageSpec spec)
    {
        var actual = Path.GetExtension(filePath).TrimStart('.').ToLowerInvariant();
        if (actual != spec.Format)
            throw new CliException($"PageSpec format {spec.Format} does not match {actual}.") { Code = "page_spec_format_mismatch" };
        var pages = spec.Format switch
        {
            "docx" => CompileWord(handler, filePath, spec),
            "xlsx" => CompileExcel(handler, filePath, spec),
            _ => CompilePowerPoint(handler, filePath, spec),
        };
        handler.Save();
        return new ProfessionalCompositionReceipt(true, filePath, spec.Format, spec.DocumentId,
            spec.Format switch { "docx" => "WordComposer", "xlsx" => "ExcelComposer", _ => "PowerPointComposer" }, pages);
    }

    private static List<ProfessionalPageReceipt> CompileWord(IDocumentHandler handler, string filePath, ProfessionalPageSpec spec)
    {
        var result = new List<ProfessionalPageReceipt>();
        handler.Add("/body", "paragraph", null, new() { ["text"] = spec.Title, ["style"] = "Title" });
        foreach (var page in spec.Pages)
        {
            var blocks = new List<ProfessionalPageBlockReceipt>();
            var heading = handler.Add("/body", "paragraph", null, new() { ["text"] = page.PrimaryClaim, ["style"] = "Heading1" });
            blocks.Add(Receipt(new ProfessionalPageBlock { BlockId = page.PageId + "-claim", Kind = "narrative", ClaimRefs = page.Blocks.SelectMany(x => x.ClaimRefs).Distinct().ToList() }, heading));
            foreach (var block in page.Blocks)
                blocks.Add(CompileBlock(handler, filePath, spec, page, block, "/body", "docx"));
            handler.Add("/body", "paragraph", null, new() { ["text"] = $"Next action: {page.ReaderAction}", ["style"] = "Quote" });
            result.Add(new ProfessionalPageReceipt(page.PageId, page.PageRole, "/body", blocks));
        }
        return result;
    }

    private static List<ProfessionalPageReceipt> CompileExcel(IDocumentHandler handler, string filePath, ProfessionalPageSpec spec)
    {
        var result = new List<ProfessionalPageReceipt>();
        var sheet = handler.Query("sheet").FirstOrDefault()?.Path ?? "/Sheet1";
        var currentRow = 1;
        handler.Set($"{sheet}/A{currentRow}", new() { ["value"] = spec.Title, ["font.bold"] = "true", ["font.size"] = "20" });
        currentRow += 2;
        foreach (var page in spec.Pages)
        {
            var claimPath = $"{sheet}/A{currentRow}";
            handler.Set(claimPath, new() { ["value"] = page.PrimaryClaim, ["font.bold"] = "true", ["font.size"] = "15", ["fill"] = spec.BrandTokens.GetValueOrDefault("primary", "DCE6F1") });
            currentRow += 2;
            var blocks = new List<ProfessionalPageBlockReceipt>
            {
                Receipt(new ProfessionalPageBlock { BlockId = page.PageId + "-claim", Kind = "narrative", ClaimRefs = page.Blocks.SelectMany(x => x.ClaimRefs).Distinct().ToList() }, claimPath),
            };
            foreach (var block in page.Blocks)
            {
                var target = $"{sheet}/A{currentRow}";
                if (block.Chart is not null)
                    block.Chart.ThemeTokens["placement.anchor"] = $"A{currentRow}:H{currentRow + Math.Max(12, block.Chart.Items.Count + 4)}";
                blocks.Add(CompileBlock(handler, filePath, spec, page, block, target, "xlsx"));
                currentRow += block.Kind is "component" or "chart" ? Math.Max(5, block.Component?.Items.Count + 3 ?? block.Chart?.Items.Count + 3 ?? 5) : 2;
            }
            handler.Set($"{sheet}/A{currentRow}", new() { ["value"] = $"Decision/action: {page.ReaderAction}", ["font.bold"] = "true" });
            currentRow += 3;
            result.Add(new ProfessionalPageReceipt(page.PageId, page.PageRole, sheet, blocks));
        }
        return result;
    }

    private static List<ProfessionalPageReceipt> CompilePowerPoint(IDocumentHandler handler, string filePath, ProfessionalPageSpec spec)
    {
        var result = new List<ProfessionalPageReceipt>();
        var existingSlides = handler.Query("slide").Count;
        while (existingSlides < spec.Pages.Count)
        {
            handler.Add("/", "slide", null, new() { ["layout"] = "blank", ["background"] = spec.BrandTokens.GetValueOrDefault("background", "F7F9FC") });
            existingSlides++;
        }
        for (var pageIndex = 0; pageIndex < spec.Pages.Count; pageIndex++)
        {
            var page = spec.Pages[pageIndex];
            var slide = $"/slide[{pageIndex + 1}]";
            var blocks = new List<ProfessionalPageBlockReceipt>();
            var title = handler.Add(slide, "textbox", null, new()
            {
                ["name"] = "pagespec-title-" + page.PageId, ["text"] = page.PrimaryClaim,
                ["x"] = "1.2cm", ["y"] = "0.7cm", ["width"] = "31cm", ["height"] = "1.4cm",
                ["font.size"] = "24", ["font.bold"] = "true", ["font.color"] = spec.BrandTokens.GetValueOrDefault("text", "172033"),
            });
            blocks.Add(Receipt(new ProfessionalPageBlock { BlockId = page.PageId + "-claim", Kind = "narrative", ClaimRefs = page.Blocks.SelectMany(x => x.ClaimRefs).Distinct().ToList() }, title));
            var count = page.Blocks.Count;
            for (var blockIndex = 0; blockIndex < count; blockIndex++)
            {
                var block = page.Blocks[blockIndex];
                var target = slide;
                var columns = count <= 2 ? count : 2;
                var column = blockIndex % columns;
                var row = blockIndex / columns;
                var width = columns == 1 ? 31.0 : 15.1;
                var height = count <= 2 ? 10.2 : 5.0;
                var x = 1.2 + column * 15.9;
                var y = 2.5 + row * 5.45;
                if (block.Component is not null) block.Component.Target = target;
                if (block.Chart is not null) block.Chart.Target = target;
                if (block.Component is not null)
                {
                    block.Component.ThemeTokens["placement.x"] = $"{x:0.0}cm";
                    block.Component.ThemeTokens["placement.y"] = $"{y:0.0}cm";
                    block.Component.ThemeTokens["placement.width"] = $"{width:0.0}cm";
                    block.Component.ThemeTokens["placement.height"] = $"{height:0.0}cm";
                }
                if (block.Chart is not null)
                {
                    block.Chart.ThemeTokens["placement.x"] = $"{x:0.0}cm";
                    block.Chart.ThemeTokens["placement.y"] = $"{y:0.0}cm";
                    block.Chart.ThemeTokens["placement.width"] = $"{width:0.0}cm";
                    block.Chart.ThemeTokens["placement.height"] = $"{height:0.0}cm";
                }
                if (block.Kind is not ("component" or "chart"))
                {
                    target = handler.Add(slide, "textbox", null, new()
                    {
                        ["name"] = "pagespec-block-" + block.BlockId,
                        ["text"] = string.IsNullOrWhiteSpace(block.Title) ? block.Text! : $"{block.Title}\n{block.Text}",
                        ["x"] = $"{x:0.0}cm", ["y"] = $"{y:0.0}cm", ["width"] = $"{width:0.0}cm", ["height"] = $"{height:0.0}cm",
                        ["font.size"] = page.Density == "compact" ? "12" : "15", ["fill"] = "FFFFFF", ["line"] = "D9E2F3:1",
                    });
                    blocks.Add(Receipt(block, target));
                }
                else blocks.Add(CompileBlock(handler, filePath, spec, page, block, target, "pptx"));
            }
            result.Add(new ProfessionalPageReceipt(page.PageId, page.PageRole, slide, blocks));
        }
        return result;
    }

    private static ProfessionalPageBlockReceipt CompileBlock(IDocumentHandler handler, string filePath, ProfessionalPageSpec document,
        ProfessionalPage page, ProfessionalPageBlock block, string target, string format)
    {
        if (block.Component is not null)
        {
            block.Component.Target = target;
            block.Component.Density = page.Density;
            foreach (var fact in block.FactRefs) if (!block.Component.FactRefs.Contains(fact)) block.Component.FactRefs.Add(fact);
            foreach (var claim in block.ClaimRefs) if (!block.Component.ClaimRefs.Contains(claim)) block.Component.ClaimRefs.Add(claim);
            foreach (var decision in block.DecisionRefs) if (!block.Component.DecisionRefs.Contains(decision)) block.Component.DecisionRefs.Add(decision);
            foreach (var action in block.ActionRefs) if (!block.Component.ActionRefs.Contains(action)) block.Component.ActionRefs.Add(action);
            foreach (var token in document.BrandTokens) block.Component.ThemeTokens.TryAdd(token.Key, token.Value);
            var receipt = ProfessionalComponentCatalog.Apply(handler, filePath, block.Component, update: false);
            return Receipt(block, receipt.NativeObjectPath);
        }
        if (block.Chart is not null)
        {
            block.Chart.Target = format == "xlsx" ? target[..target.LastIndexOf('/')] : target;
            foreach (var fact in block.FactRefs) if (!block.Chart.FactRefs.Contains(fact)) block.Chart.FactRefs.Add(fact);
            foreach (var claim in block.ClaimRefs) if (!block.Chart.ClaimRefs.Contains(claim)) block.Chart.ClaimRefs.Add(claim);
            foreach (var token in document.BrandTokens) block.Chart.ThemeTokens.TryAdd(token.Key, token.Value);
            var receipt = InformationChartEngine.Apply(handler, filePath, block.Chart);
            return Receipt(block, receipt.NativeObjectPath);
        }
        string path;
        if (format == "docx")
            path = handler.Add("/body", "paragraph", null, new() { ["text"] = string.IsNullOrWhiteSpace(block.Title) ? block.Text! : $"{block.Title}: {block.Text}" });
        else if (format == "xlsx")
        {
            handler.Set(target, new() { ["value"] = string.IsNullOrWhiteSpace(block.Title) ? block.Text! : $"{block.Title}: {block.Text}", ["alignment.wrapText"] = "true" });
            path = target;
        }
        else path = target;
        return Receipt(block, path);
    }

    private static ProfessionalPageBlockReceipt Receipt(ProfessionalPageBlock block, string path) =>
        new(block.BlockId, block.Kind, path, block.FactRefs, block.ClaimRefs, block.DecisionRefs, block.ActionRefs);
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, WriteIndented = true)]
[JsonSerializable(typeof(ProfessionalPageSpec))]
[JsonSerializable(typeof(ProfessionalCompositionReceipt))]
internal partial class ProfessionalPageJsonContext : JsonSerializerContext;
