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
            handler.Add("/body", "paragraph", null, new() { ["text"] = $"{Label(spec, "下一步", "Next action")}: {page.ReaderAction}", ["style"] = "Quote" });
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
            handler.Set($"{sheet}/A{currentRow}", new() { ["value"] = $"{Label(spec, "决策与行动", "Decision and action")}: {page.ReaderAction}", ["font.bold"] = "true" });
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
            var cjkTitle = page.PrimaryClaim.Any(character => character is >= '\u3400' and <= '\u9FFF');
            var titleSize = page.PageRole.Equals("cover", StringComparison.OrdinalIgnoreCase)
                ? (cjkTitle && page.PrimaryClaim.Length > 24 ? 34 : 40)
                : (cjkTitle && page.PrimaryClaim.Length > 25 ? 26 : 30);
            var title = handler.Add(slide, "textbox", null, new()
            {
                ["name"] = "pagespec-title-" + page.PageId, ["text"] = page.PrimaryClaim,
                ["x"] = "1.2cm", ["y"] = "0.65cm", ["width"] = "31cm", ["height"] = "1.9cm",
                ["font.size"] = titleSize.ToString(System.Globalization.CultureInfo.InvariantCulture),
                ["font.bold"] = "true", ["font.color"] = spec.BrandTokens.GetValueOrDefault("text", "172033"),
            });
            blocks.Add(Receipt(new ProfessionalPageBlock { BlockId = page.PageId + "-claim", Kind = "narrative", ClaimRefs = page.Blocks.SelectMany(x => x.ClaimRefs).Distinct().ToList() }, title));
            var layouts = BuildPowerPointLayouts(page);
            for (var blockIndex = 0; blockIndex < page.Blocks.Count; blockIndex++)
            {
                var block = page.Blocks[blockIndex];
                var target = slide;
                var layout = layouts[blockIndex];
                var (x, y, width, height) = (layout.X, layout.Y, layout.Width, layout.Height);
                if (block.Component is not null) block.Component.Target = target;
                if (block.Chart is not null) block.Chart.Target = target;
                if (block.Component is not null)
                {
                    var tableHeight = Math.Min(height, Math.Max(2.9, (block.Component.Items.Count + 1) * 1.15));
                    block.Component.ThemeTokens["placement.x"] = $"{x:0.0}cm";
                    block.Component.ThemeTokens["placement.y"] = $"{y:0.0}cm";
                    block.Component.ThemeTokens["placement.width"] = $"{width:0.0}cm";
                    block.Component.ThemeTokens["placement.height"] = $"{tableHeight:0.0}cm";
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
                        ["font.size"] = block.Importance.Equals("primary", StringComparison.OrdinalIgnoreCase) ? "20" : page.Density == "compact" ? "12" : "15",
                        ["font.bold"] = block.Importance.Equals("primary", StringComparison.OrdinalIgnoreCase) ? "true" : "false",
                        ["font.color"] = spec.BrandTokens.GetValueOrDefault("text", "334155"), ["fill"] = "none", ["line"] = "none",
                    });
                    blocks.Add(Receipt(block, target));
                }
                else blocks.Add(CompileBlock(handler, filePath, spec, page, block, target, "pptx"));
            }
            handler.Add(slide, "textbox", null, new()
            {
                ["name"] = "pagespec-action-" + page.PageId,
                ["text"] = $"{Label(spec, "下一步", "Next action")} · {page.ReaderAction}",
                ["x"] = "1.2cm", ["y"] = "17.3cm", ["width"] = "31cm", ["height"] = "0.75cm",
                ["font.size"] = "12", ["font.bold"] = "true", ["font.color"] = spec.BrandTokens.GetValueOrDefault("primary", "1F4E78"),
                ["fill"] = "none", ["line"] = "none",
            });
            result.Add(new ProfessionalPageReceipt(page.PageId, page.PageRole, slide, blocks));
        }
        return result;
    }

    private sealed record PowerPointBlockLayout(double X, double Y, double Width, double Height);

    private static PowerPointBlockLayout[] BuildPowerPointLayouts(ProfessionalPage page)
    {
        const double x = 1.2, y = 3.05, width = 31.0, height = 13.55, gap = 0.65;
        var count = page.Blocks.Count;
        var result = new PowerPointBlockLayout[count];
        var primary = page.Blocks.FindIndex(block => block.Kind == "chart"
            && block.Importance.Equals("primary", StringComparison.OrdinalIgnoreCase));
        if (primary < 0) primary = page.Blocks.FindIndex(block => block.Kind == "chart");
        if (primary < 0) primary = page.Blocks.FindIndex(block => block.Kind == "component"
            && block.Importance.Equals("primary", StringComparison.OrdinalIgnoreCase));
        if (primary < 0) primary = page.Blocks.FindIndex(block => block.Kind == "component");
        if (primary < 0) primary = page.Blocks.FindIndex(block => block.Importance.Equals("primary", StringComparison.OrdinalIgnoreCase));
        if (primary < 0) primary = 0;
        if (count == 1)
        {
            result[0] = new(x, y, width, height);
            return result;
        }
        var supporting = Enumerable.Range(0, count).Where(index => index != primary).ToArray();
        if (count == 2)
        {
            var twoBlockPrimaryWidth = page.Blocks[primary].Kind == "chart" ? 18.4 : 17.0;
            result[primary] = new(x, y, twoBlockPrimaryWidth, height);
            result[supporting[0]] = new(x + twoBlockPrimaryWidth + gap, y, width - twoBlockPrimaryWidth - gap, height);
            return result;
        }
        if (count == 3
            && page.Blocks.Count(block => block.Kind == "narrative") == 1
            && page.Blocks.Any(block => block.Kind == "chart")
            && page.Blocks.Any(block => block.Kind == "component"))
        {
            var narrative = page.Blocks.FindIndex(block => block.Kind == "narrative");
            var chart = page.Blocks.FindIndex(block => block.Kind == "chart");
            var component = page.Blocks.FindIndex(block => block.Kind == "component");
            const double narrativeHeight = 2.15;
            const double visualWidth = 18.4;
            result[narrative] = new(x, y, width, narrativeHeight);
            result[chart] = new(x, y + narrativeHeight + gap, visualWidth, height - narrativeHeight - gap);
            result[component] = new(x + visualWidth + gap, y + narrativeHeight + gap, width - visualWidth - gap, height - narrativeHeight - gap);
            return result;
        }
        var decisionLed = page.PageRole.Equals("decision", StringComparison.OrdinalIgnoreCase)
            || page.PageRole.Equals("action", StringComparison.OrdinalIgnoreCase);
        if (count <= 4 && decisionLed)
        {
            const double primaryHeight = 8.25;
            result[primary] = new(x, y, width, primaryHeight);
            var supportWidth = (width - gap * (supporting.Length - 1)) / supporting.Length;
            for (var index = 0; index < supporting.Length; index++)
                result[supporting[index]] = new(x + index * (supportWidth + gap), y + primaryHeight + gap, supportWidth, height - primaryHeight - gap);
            return result;
        }
        const double primaryWidth = 20.1;
        result[primary] = new(x, y, primaryWidth, height);
        var supportHeight = (height - gap * (supporting.Length - 1)) / supporting.Length;
        for (var index = 0; index < supporting.Length; index++)
            result[supporting[index]] = new(x + primaryWidth + gap, y + index * (supportHeight + gap), width - primaryWidth - gap, supportHeight);
        return result;
    }

    private static string Label(ProfessionalPageSpec spec, string chinese, string english)
    {
        var text = spec.Title + string.Concat(spec.Pages.Select(page => page.PrimaryClaim + page.ReaderAction));
        return text.Any(character => character is >= '\u3400' and <= '\u9FFF') ? chinese : english;
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
