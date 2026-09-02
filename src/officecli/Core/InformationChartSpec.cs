// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeCli.Core;

public sealed class InformationChartSpec
{
    [JsonPropertyName("schemaVersion")] public int SchemaVersion { get; set; } = 1;
    [JsonPropertyName("chartId")] public string ChartId { get; set; } = "";
    [JsonPropertyName("chartType")] public string ChartType { get; set; } = "";
    [JsonPropertyName("title")] public string Title { get; set; } = "";
    [JsonPropertyName("unit")] public string? Unit { get; set; }
    [JsonPropertyName("target")] public string? Target { get; set; }
    [JsonPropertyName("items")] public List<InformationChartItem> Items { get; set; } = [];
    [JsonPropertyName("annotations")] public List<InformationChartAnnotation> Annotations { get; set; } = [];
    [JsonPropertyName("factRefs")] public List<string> FactRefs { get; set; } = [];
    [JsonPropertyName("claimRefs")] public List<string> ClaimRefs { get; set; } = [];
    [JsonPropertyName("axisPolicy")] public string AxisPolicy { get; set; } = "auto";
    [JsonPropertyName("axisReason")] public string AxisReason { get; set; } = "";
    [JsonPropertyName("themeTokens")] public Dictionary<string, string> ThemeTokens { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class InformationChartItem
{
    [JsonPropertyName("label")] public string Label { get; set; } = "";
    [JsonPropertyName("actual")] public double? Actual { get; set; }
    [JsonPropertyName("target")] public double? Target { get; set; }
    [JsonPropertyName("forecast")] public double? Forecast { get; set; }
    [JsonPropertyName("value")] public double? Value { get; set; }
    [JsonPropertyName("contribution")] public double? Contribution { get; set; }
    [JsonPropertyName("capacity")] public double? Capacity { get; set; }
    [JsonPropertyName("demand")] public double? Demand { get; set; }
    [JsonPropertyName("low")] public double? Low { get; set; }
    [JsonPropertyName("high")] public double? High { get; set; }
    [JsonPropertyName("probability")] public double? Probability { get; set; }
    [JsonPropertyName("impact")] public double? Impact { get; set; }
    [JsonPropertyName("fields")] public Dictionary<string, double> Fields { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class InformationChartAnnotation
{
    [JsonPropertyName("itemLabel")] public string ItemLabel { get; set; } = "";
    [JsonPropertyName("text")] public string Text { get; set; } = "";
}

public sealed record InformationChartDefinition(string ChartType, string SemanticIntent, string NativeChartType,
    string[] RequiredValues, string[] AdaptiveRules, string FallbackRepresentation);

public sealed record InformationChartReceipt(bool Ok, string File, string ChartId, string RequestedChartType,
    string Representation, string NativeObjectPath, string Title, int ItemCount, string? FallbackReason,
    IReadOnlyList<string> FactRefs, IReadOnlyList<string> ClaimRefs);

public sealed record InformationChartListResponse(bool Ok, int Count, IReadOnlyList<InformationChartDefinition> Charts);
public sealed record InformationChartReadItem(string ChartId, string RequestedChartType, string NativeObjectPath,
    string Title, string? Unit, IReadOnlyList<string> FactRefs, IReadOnlyList<string> ClaimRefs, DocumentNode NativeObject);
public sealed record InformationChartReadResponse(bool Ok, int Count, IReadOnlyList<InformationChartReadItem> Charts);

public static class InformationChartEngine
{
    private static readonly Dictionary<string, InformationChartDefinition> Definitions = new[]
    {
        new InformationChartDefinition("annotated-trend", "Show direction, inflection and material events", "line", ["actual|value"], ["use direct labels", "reduce tick density for long series", "annotate only decision-relevant points"], "table"),
        new InformationChartDefinition("actual-target-forecast", "Compare actual, target and forecast under one definition", "combo", ["actual", "target"], ["actual is primary", "target is a neutral baseline", "forecast uses a distinct dashed series"], "actual-target-forecast component"),
        new InformationChartDefinition("waterfall", "Explain a signed bridge from drivers to total", "waterfall", ["contribution"], ["preserve driver order", "aggregate immaterial drivers", "label totals"], "growth-contribution-waterfall component"),
        new InformationChartDefinition("supply-demand-gap", "Expose capacity shortfall or surplus over time", "combo", ["capacity", "demand"], ["align units", "highlight negative gaps", "show a zero gap reference"], "capacity-demand-gap component"),
        new InformationChartDefinition("dumbbell-comparison", "Compare two values while emphasizing the gap", "bar", ["low", "high"], ["sort by gap when order is not semantic", "label both endpoints"], "comparison table"),
        new InformationChartDefinition("probability-impact-scatter", "Prioritize risks by probability and impact", "scatter", ["probability", "impact"], ["label every risk", "do not imply precision absent from inputs", "keep threshold quadrants explicit"], "risk-probability-impact component"),
        new InformationChartDefinition("scenario-comparison", "Compare mutually exclusive scenarios on one outcome", "column", ["value"], ["use a common axis", "highlight the recommended scenario", "show assumptions in notes"], "option-comparison component"),
        new InformationChartDefinition("sensitivity-heatmap", "Show how two assumptions change an outcome", "table", ["fields"], ["use ordered semantic colors", "print values in every cell", "preserve row and column labels"], "native heatmap table"),
    }.ToDictionary(item => item.ChartType, StringComparer.OrdinalIgnoreCase);

    public static IReadOnlyList<InformationChartDefinition> List() => Definitions.Values.OrderBy(item => item.ChartType, StringComparer.Ordinal).ToList();

    public static InformationChartSpec Parse(string path)
    {
        InformationChartSpec? spec;
        try { spec = JsonSerializer.Deserialize(File.ReadAllText(path), InformationChartJsonContext.Default.InformationChartSpec); }
        catch (Exception ex) when (ex is IOException or JsonException)
        { throw new CliException($"ChartSpec could not be read: {ex.Message}") { Code = "chart_spec_invalid" }; }
        if (spec is null || spec.SchemaVersion != 1 || string.IsNullOrWhiteSpace(spec.ChartId)
            || string.IsNullOrWhiteSpace(spec.ChartType) || string.IsNullOrWhiteSpace(spec.Title) || spec.Items.Count == 0)
            throw new CliException("ChartSpec requires schemaVersion=1, chartId, chartType, title and items.") { Code = "chart_spec_invalid" };
        if (!Definitions.ContainsKey(spec.ChartType))
            throw new CliException($"Unknown information chart '{spec.ChartType}'. Use 'chart-spec list'.") { Code = "chart_type_unknown" };
        if (spec.FactRefs.Count == 0 || spec.ClaimRefs.Count == 0)
            throw new CliException("ChartSpec requires factRefs and claimRefs so the conclusion remains traceable.") { Code = "chart_binding_missing" };
        if (spec.AxisPolicy is not ("auto" or "zero" or "nonzero"))
            throw new CliException("ChartSpec axisPolicy must be auto, zero or nonzero.") { Code = "chart_spec_invalid" };
        if (spec.AxisPolicy == "nonzero" && string.IsNullOrWhiteSpace(spec.AxisReason))
            throw new CliException("A nonzero axis requires axisReason.") { Code = "chart_axis_reason_missing" };
        return spec;
    }

    public static InformationChartReceipt Apply(IDocumentHandler handler, string filePath, InformationChartSpec spec)
    {
        var definition = Definitions[spec.ChartType];
        var target = string.IsNullOrWhiteSpace(spec.Target) ? DefaultTarget(Path.GetExtension(filePath)) : spec.Target!;
        var missing = MissingRequiredValues(spec, definition).ToList();
        if (missing.Count > 0 || definition.NativeChartType == "table")
            return ApplyTableFallback(handler, filePath, spec, target,
                missing.Count > 0 ? $"missing values: {string.Join(", ", missing)}" : "heatmap is more truthful as an editable native table");

        var series = BuildSeries(spec);
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["chartType"] = definition.NativeChartType,
            ["title"] = spec.Title,
            ["categories"] = string.Join(',', spec.Items.Select(item => EscapeList(item.Label))),
            ["data"] = string.Join(';', series.Select(item => $"{EscapeList(item.Name)}:{string.Join(',', item.Values.Select(FormatNumber))}")),
            ["chartStyle"] = "10",
            ["legend"] = series.Count > 1 ? "bottom" : "none",
            ["legend.font"] = "10:#526074",
            ["gridlines"] = "false",
            ["dataLabels"] = spec.Items.Count <= 12 ? "value" : "none",
            ["datalabels.numfmt"] = string.IsNullOrWhiteSpace(spec.Unit) ? "0.0" : UnitFormat(spec.Unit),
            ["colors"] = string.Join(',', ChartColors(spec, series.Count)),
            ["width"] = Path.GetExtension(filePath).Equals(".pptx", StringComparison.OrdinalIgnoreCase) ? "30cm" : "16cm",
            ["height"] = Path.GetExtension(filePath).Equals(".pptx", StringComparison.OrdinalIgnoreCase) ? "12cm" : "9cm",
            ["axisNumFmt"] = string.IsNullOrWhiteSpace(spec.Unit) ? "0.0" : UnitFormat(spec.Unit),
            ["name"] = Marker(spec),
        };
        foreach (var key in new[] { "x", "y", "width", "height", "anchor" })
            if (spec.ThemeTokens.TryGetValue("placement." + key, out var placement)) props[key] = placement;
        if (spec.AxisPolicy == "zero") props["axismin"] = "0";
        if (spec.ChartType == "actual-target-forecast") props["combotypes"] = "column,line,line";
        if (spec.ChartType == "supply-demand-gap") props["combotypes"] = "column,line";
        if (spec.ChartType == "probability-impact-scatter") props["scatterStyle"] = "marker";
        var path = handler.Add(target, "chart", null, props);
        if (Path.GetExtension(filePath).Equals(".pptx", StringComparison.OrdinalIgnoreCase) && spec.Annotations.Count > 0)
            AddPowerPointAnnotations(handler, target, spec);
        handler.Save();
        return new InformationChartReceipt(true, filePath, spec.ChartId, spec.ChartType, "native-chart", path,
            spec.Title, spec.Items.Count, null, spec.FactRefs, spec.ClaimRefs);
    }

    private static void AddPowerPointAnnotations(IDocumentHandler handler, string target, InformationChartSpec spec)
    {
        var x = Centimeters(spec.ThemeTokens.GetValueOrDefault("placement.x"), 1.2);
        var y = Centimeters(spec.ThemeTokens.GetValueOrDefault("placement.y"), 2.5);
        var width = Centimeters(spec.ThemeTokens.GetValueOrDefault("placement.width"), 30.0);
        var accent = spec.ThemeTokens.GetValueOrDefault("accent", "0E7490").TrimStart('#');
        for (var index = 0; index < Math.Min(3, spec.Annotations.Count); index++)
        {
            var annotation = spec.Annotations[index];
            handler.Add(target, "textbox", null, new Dictionary<string, string>
            {
                ["name"] = $"officecli-chart-annotation-{spec.ChartId}-{index + 1}",
                ["text"] = string.IsNullOrWhiteSpace(annotation.ItemLabel) ? annotation.Text : $"{annotation.ItemLabel} · {annotation.Text}",
                ["x"] = $"{Math.Max(x + 0.5, x + width - 7.0):0.0}cm", ["y"] = $"{y + 0.8 + index * 1.25:0.0}cm",
                ["width"] = "6.5cm", ["height"] = "0.95cm", ["font.size"] = "10", ["font.bold"] = "true",
                ["font.color"] = accent, ["fill"] = "FFFFFF", ["fill.transparency"] = "8", ["line"] = $"{accent}:1",
            });
        }
    }

    private static double Centimeters(string? value, double fallback)
    {
        if (string.IsNullOrWhiteSpace(value)) return fallback;
        var normalized = value.Trim();
        if (normalized.EndsWith("cm", StringComparison.OrdinalIgnoreCase)) normalized = normalized[..^2];
        return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) ? parsed : fallback;
    }

    public static IReadOnlyList<InformationChartReadItem> Read(IDocumentHandler handler)
    {
        var result = new List<InformationChartReadItem>();
        foreach (var node in handler.Query("chart"))
        {
            if (!node.Format.TryGetValue("name", out var raw) || raw is null || !TryParseMarker(raw.ToString() ?? "", out var values))
                continue;
            result.Add(new InformationChartReadItem(
                values["id"], values["type"], node.Path,
                node.Format.GetValueOrDefault("title")?.ToString() ?? "",
                values.GetValueOrDefault("unit"), SplitRefs(values.GetValueOrDefault("facts", "")),
                SplitRefs(values.GetValueOrDefault("claims", "")), node));
        }
        return result;
    }

    private static InformationChartReceipt ApplyTableFallback(IDocumentHandler handler, string filePath, InformationChartSpec spec, string target, string reason)
    {
        var headers = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in spec.Items)
        {
            foreach (var key in new[] { "actual", "target", "forecast", "value", "contribution", "capacity", "demand", "low", "high", "probability", "impact" })
                if (Read(item, key).HasValue) headers.Add(key);
            foreach (var key in item.Fields.Keys) headers.Add(key);
        }
        var rows = new List<List<string>> { new() { "Item" } };
        rows[0].AddRange(headers);
        foreach (var item in spec.Items)
        {
            var row = new List<string> { item.Label };
            row.AddRange(headers.Select(key => item.Fields.TryGetValue(key, out var field) ? FormatNumber(field) : Read(item, key) is double value ? FormatNumber(value) : ""));
            rows.Add(row);
        }
        var props = new Dictionary<string, string>
        {
            ["data"] = string.Join(';', rows.Select(row => string.Join(',', row.Select(QuoteCell)))),
            ["style"] = "medium2",
        };
        if (Path.GetExtension(filePath).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
        {
            props["name"] = "officecli-chart-fallback-" + spec.ChartId;
            props["x"] = "1.2cm"; props["y"] = "2.5cm"; props["width"] = "31cm"; props["height"] = "11cm";
        }
        if (Path.GetExtension(filePath).Equals(".docx", StringComparison.OrdinalIgnoreCase)) props["caption"] = "officecli-chart-fallback-" + spec.ChartId;
        if (Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new CliException("XLSX chart fallback requires PageSpec or ComponentSpec so source cells and table range remain explicit.") { Code = "chart_fallback_requires_component" };
        var path = handler.Add(target, "table", null, props);
        handler.Save();
        return new InformationChartReceipt(true, filePath, spec.ChartId, spec.ChartType, "native-table", path,
            spec.Title, spec.Items.Count, reason, spec.FactRefs, spec.ClaimRefs);
    }

    private static IEnumerable<string> MissingRequiredValues(InformationChartSpec spec, InformationChartDefinition definition)
    {
        if (spec.ChartType == "annotated-trend" && spec.Items.Count < 2) yield return "at least two periods";
        foreach (var required in definition.RequiredValues)
        foreach (var item in spec.Items)
        {
            if (required == "fields") { if (item.Fields.Count == 0) yield return $"{item.Label}.fields"; continue; }
            if (required.Contains('|'))
            {
                if (!required.Split('|').Any(key => Read(item, key).HasValue)) yield return $"{item.Label}.{required}";
            }
            else if (!Read(item, required).HasValue) yield return $"{item.Label}.{required}";
        }
    }

    private static List<(string Name, List<double> Values)> BuildSeries(InformationChartSpec spec)
    {
        string[] keys = spec.ChartType switch
        {
            "annotated-trend" => spec.Items.Any(item => item.Actual.HasValue) ? ["actual"] : ["value"],
            "actual-target-forecast" => spec.Items.Any(item => item.Forecast.HasValue) ? ["actual", "target", "forecast"] : ["actual", "target"],
            "waterfall" => ["contribution"],
            "supply-demand-gap" => ["demand", "capacity"],
            "dumbbell-comparison" => ["low", "high"],
            "probability-impact-scatter" => ["impact", "probability"],
            _ => ["value"],
        };
        return keys.Select(key => (ToHeader(key), spec.Items.Select(item => Read(item, key) ?? 0).ToList())).ToList();
    }

    private static double? Read(InformationChartItem item, string key) => key switch
    {
        "actual" => item.Actual, "target" => item.Target, "forecast" => item.Forecast, "value" => item.Value,
        "contribution" => item.Contribution, "capacity" => item.Capacity, "demand" => item.Demand,
        "low" => item.Low, "high" => item.High, "probability" => item.Probability, "impact" => item.Impact,
        _ => item.Fields.TryGetValue(key, out var value) ? value : null,
    };
    private static string[] ChartColors(InformationChartSpec spec, int count)
    {
        var primary = spec.ThemeTokens.GetValueOrDefault("primary", "1F4E78").TrimStart('#');
        var neutral = spec.ThemeTokens.GetValueOrDefault("neutral", "94A3B8").TrimStart('#');
        var accent = spec.ThemeTokens.GetValueOrDefault("accent", "0E7490").TrimStart('#');
        return Enumerable.Range(0, count).Select(index => index == 0 ? primary : index == 1 ? neutral : accent).ToArray();
    }
    private static string UnitFormat(string unit) => unit.Contains('%') || unit.Contains("percent", StringComparison.OrdinalIgnoreCase) ? "0.0%" : "0.0";
    private static string Marker(InformationChartSpec spec)
    {
        var marker = "officecli-chart|" + string.Join('|', new Dictionary<string, string>
        {
            ["id"] = spec.ChartId, ["type"] = spec.ChartType, ["unit"] = spec.Unit ?? "",
            ["facts"] = string.Join(',', spec.FactRefs), ["claims"] = string.Join(',', spec.ClaimRefs),
        }.Select(pair => $"{pair.Key}={Uri.EscapeDataString(pair.Value)}"));
        if (marker.Length > 240)
            throw new CliException("ChartSpec trace bindings exceed the native Office object-name limit; bind a smaller set of authoritative facts and claims.")
            { Code = "chart_binding_too_large" };
        return marker;
    }
    private static bool TryParseMarker(string marker, out Dictionary<string, string> values)
    {
        values = new(StringComparer.OrdinalIgnoreCase);
        if (!marker.StartsWith("officecli-chart|", StringComparison.Ordinal)) return false;
        foreach (var segment in marker.Split('|').Skip(1))
        {
            var split = segment.IndexOf('=');
            if (split > 0) values[segment[..split]] = Uri.UnescapeDataString(segment[(split + 1)..]);
        }
        return values.ContainsKey("id") && values.ContainsKey("type");
    }
    private static IReadOnlyList<string> SplitRefs(string value) => value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
    private static string FormatNumber(double value) => value.ToString("0.###############", CultureInfo.InvariantCulture);
    private static string EscapeList(string value) => value.Replace(",", " ").Replace(";", " ");
    private static string QuoteCell(string value) => value.IndexOfAny([',', ';', '"', '\n', '\r']) >= 0 ? $"\"{value.Replace("\"", "\"\"")}\"" : value;
    private static string ToHeader(string value) => char.ToUpperInvariant(value[0]) + value[1..];
    private static string DefaultTarget(string extension) => extension.ToLowerInvariant() switch { ".docx" => "/body", ".xlsx" => "/Sheet1", _ => "/slide[1]" };
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, WriteIndented = true)]
[JsonSerializable(typeof(InformationChartSpec))]
[JsonSerializable(typeof(InformationChartReceipt))]
[JsonSerializable(typeof(InformationChartListResponse))]
[JsonSerializable(typeof(InformationChartReadResponse))]
internal partial class InformationChartJsonContext : JsonSerializerContext;
