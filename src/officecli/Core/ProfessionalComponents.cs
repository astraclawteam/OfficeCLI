// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeCli.Core;

public sealed class ProfessionalComponentSpec
{
    [JsonPropertyName("schemaVersion")] public int SchemaVersion { get; set; } = 1;
    [JsonPropertyName("componentId")] public string ComponentId { get; set; } = "";
    [JsonPropertyName("instanceId")] public string InstanceId { get; set; } = "";
    [JsonPropertyName("title")] public string Title { get; set; } = "";
    [JsonPropertyName("density")] public string Density { get; set; } = "balanced";
    [JsonPropertyName("target")] public string? Target { get; set; }
    [JsonPropertyName("items")] public List<ProfessionalComponentItem> Items { get; set; } = [];
    [JsonPropertyName("factRefs")] public List<string> FactRefs { get; set; } = [];
    [JsonPropertyName("claimRefs")] public List<string> ClaimRefs { get; set; } = [];
    [JsonPropertyName("decisionRefs")] public List<string> DecisionRefs { get; set; } = [];
    [JsonPropertyName("actionRefs")] public List<string> ActionRefs { get; set; } = [];
    [JsonPropertyName("themeTokens")] public Dictionary<string, string> ThemeTokens { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class ProfessionalComponentItem
{
    [JsonPropertyName("label")] public string Label { get; set; } = "";
    [JsonPropertyName("fields")] public Dictionary<string, JsonElement> Fields { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed record ProfessionalComponentDefinition(
    string ComponentId,
    string Category,
    string SemanticIntent,
    string[] ApplicablePageRoles,
    string[] RequiredSlots,
    string[] OptionalSlots,
    string AcceptedDataShape,
    string[] UnsupportedConditions,
    string[] CompositionVariants,
    string[] AdaptiveRules,
    string[] DensityModes,
    string[] TargetFormats,
    string EditabilityContract,
    string FactBindingContract);

public sealed record ProfessionalComponentReceipt(
    bool Ok,
    string Operation,
    string File,
    string Format,
    string ComponentId,
    string InstanceId,
    string NativeObjectPath,
    string Variant,
    string Density,
    int ItemCount,
    IReadOnlyList<string> FactRefs,
    IReadOnlyList<string> ClaimRefs,
    IReadOnlyList<string> DecisionRefs,
    IReadOnlyList<string> ActionRefs);

public sealed record ProfessionalComponentListResponse(bool Ok, int Count, IReadOnlyList<ProfessionalComponentDefinition> Components);
public sealed record ProfessionalComponentDescribeResponse(bool Ok, ProfessionalComponentDefinition Component);
public sealed record ProfessionalComponentReadResponse(bool Ok, int Count, IReadOnlyList<ProfessionalComponentReceipt> Components);

public static class ProfessionalComponentCatalog
{
    private static readonly string[] Formats = ["docx", "xlsx", "pptx"];
    private static readonly string[] Density = ["compact", "balanced", "spacious"];
    private static readonly Dictionary<string, ProfessionalComponentDefinition> Definitions = BuildDefinitions();

    public static IReadOnlyList<ProfessionalComponentDefinition> List() =>
        Definitions.Values.OrderBy(item => item.Category, StringComparer.Ordinal)
            .ThenBy(item => item.ComponentId, StringComparer.Ordinal).ToList();

    public static ProfessionalComponentDefinition Get(string componentId)
    {
        if (!Definitions.TryGetValue(componentId, out var definition))
            throw new CliException($"Unknown professional component '{componentId}'. Use 'component list'.")
            { Code = "component_unknown" };
        return definition;
    }

    public static ProfessionalComponentSpec Parse(string path)
    {
        ProfessionalComponentSpec? spec;
        try
        {
            spec = JsonSerializer.Deserialize(File.ReadAllText(path), ProfessionalComponentJsonContext.Default.ProfessionalComponentSpec);
        }
        catch (Exception ex) when (ex is IOException or JsonException)
        {
            throw new CliException($"ComponentSpec could not be read: {ex.Message}") { Code = "component_spec_invalid" };
        }
        if (spec is null || spec.SchemaVersion != 1 || string.IsNullOrWhiteSpace(spec.ComponentId)
            || string.IsNullOrWhiteSpace(spec.InstanceId) || string.IsNullOrWhiteSpace(spec.Title))
            throw new CliException("ComponentSpec requires schemaVersion=1, componentId, instanceId and title.")
            { Code = "component_spec_invalid" };
        _ = Get(spec.ComponentId);
        if (spec.Items.Count == 0)
            throw new CliException("ComponentSpec items must contain at least one business item.")
            { Code = "component_spec_invalid" };
        if (spec.Items.Any(item => string.IsNullOrWhiteSpace(item.Label)))
            throw new CliException("Every ComponentSpec item requires a non-empty label.")
            { Code = "component_spec_invalid" };
        if (spec.Density is not ("compact" or "balanced" or "spacious"))
            throw new CliException("ComponentSpec density must be compact, balanced or spacious.")
            { Code = "component_spec_invalid" };
        EnsureUnique(spec.FactRefs, "factRefs");
        EnsureUnique(spec.ClaimRefs, "claimRefs");
        EnsureUnique(spec.DecisionRefs, "decisionRefs");
        EnsureUnique(spec.ActionRefs, "actionRefs");
        return spec;
    }

    public static ProfessionalComponentReceipt Apply(IDocumentHandler handler, string filePath, ProfessionalComponentSpec spec, bool update)
    {
        var extension = Path.GetExtension(filePath).TrimStart('.').ToLowerInvariant();
        if (!Formats.Contains(extension, StringComparer.Ordinal))
            throw new CliException("Professional components support DOCX, XLSX and PPTX.") { Code = "component_format_unsupported" };
        var definition = Get(spec.ComponentId);
        ValidateRequiredSlots(definition, spec);
        var marker = Marker(spec);
        if (update)
        {
            var existing = FindNativeObject(handler, extension, spec);
            if (existing is null)
                throw new CliException($"Component instance '{spec.InstanceId}' was not found for update.") { Code = "component_not_found" };
            handler.Remove(existing.Path);
        }
        else if (FindNativeObject(handler, extension, spec) is not null)
        {
            throw new CliException($"Component instance '{spec.InstanceId}' already exists; use 'component update'.")
            { Code = "component_already_exists" };
        }

        var rows = BuildRows(definition, spec);
        var target = string.IsNullOrWhiteSpace(spec.Target) ? DefaultTarget(extension) : spec.Target!;
        if (extension == "xlsx")
        {
            var slash = target.LastIndexOf('/');
            var sheet = slash > 0 ? target[..slash] : target;
            if (string.IsNullOrWhiteSpace(sheet)) sheet = "/Sheet1";
            var start = ParseExcelStart(target);
            var startColumn = ExcelColumnNumber(start);
            var startRow = ExcelRow(start);
            for (var row = 0; row < rows.Count; row++)
            for (var column = 0; column < rows[row].Count; column++)
                handler.Set($"{sheet}/{ExcelColumn(startColumn + column)}{startRow + row}", new Dictionary<string, string>
                {
                    ["value"] = rows[row][column],
                    ["bold"] = row == 0 ? "true" : "false",
                    ["fill"] = row == 0 ? spec.ThemeTokens.GetValueOrDefault("primary", "1F4E78") : "FFFFFF",
                    ["font.color"] = row == 0 ? "FFFFFF" : "1F2937",
                });
            target = sheet;
        }
        var props = BuildTableProperties(extension, definition, spec, rows, marker);
        var path = handler.Add(target, "table", null, props);
        handler.Save();
        var variant = SelectVariant(spec, rows[0].Count);
        return new ProfessionalComponentReceipt(true, update ? "updated" : "inserted", filePath, extension,
            spec.ComponentId, spec.InstanceId, path, variant, spec.Density, spec.Items.Count,
            spec.FactRefs, spec.ClaimRefs, spec.DecisionRefs, spec.ActionRefs);
    }

    public static IReadOnlyList<ProfessionalComponentReceipt> Read(IDocumentHandler handler, string filePath, string? instanceId)
    {
        var extension = Path.GetExtension(filePath).TrimStart('.').ToLowerInvariant();
        var candidates = handler.Query("table");
        var receipts = new List<ProfessionalComponentReceipt>();
        foreach (var node in candidates)
        {
            var marker = ComponentMarker(node, extension);
            if (marker is null || !TryParseAnyMarker(marker, out var data)) continue;
            if (!string.IsNullOrWhiteSpace(instanceId) && !string.Equals(instanceId, data["instance"], StringComparison.Ordinal)) continue;
            receipts.Add(new ProfessionalComponentReceipt(true, "read", filePath, extension,
                data["component"], data["instance"], node.Path, data.GetValueOrDefault("variant", "adaptive"),
                data.GetValueOrDefault("density", "balanced"), node.Children.Count,
                SplitRefs(data.GetValueOrDefault("facts", "")), SplitRefs(data.GetValueOrDefault("claims", "")),
                SplitRefs(data.GetValueOrDefault("decisions", "")), SplitRefs(data.GetValueOrDefault("actions", ""))));
        }
        return receipts;
    }

    private static Dictionary<string, ProfessionalComponentDefinition> BuildDefinitions()
    {
        ProfessionalComponentDefinition D(string id, string category, string intent, string[] required, string[] optional,
            string shape, string[] unsupported, string[] variants, string[] rules, string editability, string bindings) =>
            new(id, category, intent,
                category == "operating-analysis" ? ["evidence", "comparison"] : category == "decision" ? ["decision", "risk"] : ["action"],
                required, optional, shape, unsupported, variants, rules, Density, Formats, editability, bindings);
        return new[]
        {
            D("kpi-strip", "operating-analysis", "Summarize a small set of decision-relevant metrics", ["value"], ["delta", "status", "note"],
                "one item per KPI", ["more than 12 KPIs without grouping"], ["horizontal-strip", "two-row-grid", "compact-table"],
                ["switch to two rows above four KPIs", "move secondary notes below the metric in compact hosts"], "native table cells and text remain editable", "each KPI must bind at least one factId"),
            D("actual-target-forecast", "operating-analysis", "Compare actual, target and forecast under one metric definition", ["actual", "target"], ["forecast", "gap", "unit"],
                "one item per metric or period", ["mixed units without normalization"], ["comparison-table", "grouped-comparison"],
                ["omit forecast only when unavailable", "use horizontal labels when labels are long"], "native values remain editable", "metric rows bind authoritative facts"),
            D("growth-contribution-waterfall", "operating-analysis", "Explain how drivers bridge a start value to an end value", ["contribution"], ["start", "end", "unit"],
                "ordered signed contributions", ["unordered or non-additive drivers"], ["bridge-table", "waterfall-chart"],
                ["aggregate minor drivers above eight points", "retain start and total rows"], "native bridge data remain editable and can feed a chart", "every contribution binds a factId"),
            D("capacity-demand-gap", "operating-analysis", "Show demand, available capacity and the resulting service gap", ["capacity", "demand"], ["gap", "period", "threshold"],
                "aligned capacity and demand series", ["incomparable time or capacity units"], ["gap-table", "paired-series"],
                ["compute gap only from compatible units", "switch to compact period rows above six periods"], "native series remain editable", "capacity and demand values bind facts"),
            D("option-comparison", "decision", "Compare mutually exclusive options against explicit decision criteria", ["benefit", "cost", "risk"], ["constraint", "score", "recommendation"],
                "one item per option", ["options are not mutually comparable"], ["criteria-matrix", "recommended-option-emphasis"],
                ["transpose when criteria exceed options", "highlight exactly one recommendation when evidence supports it"], "native comparison matrix remains editable", "options bind decisionId and supporting factIds"),
            D("risk-probability-impact", "decision", "Position risks by probability and impact without inventing precision", ["probability", "impact"], ["owner", "mitigation", "status"],
                "one item per risk", ["qualitative inputs presented as false numeric precision"], ["risk-register", "probability-impact-plot"],
                ["use bands for qualitative probability", "label every plotted risk"], "native risk register remains editable", "risk assessments bind claims and facts"),
            D("decision-request", "decision", "Make the requested decision, recommendation and deadline explicit", ["request", "recommendation"], ["deadline", "owner", "conditions"],
                "one item per decision request", ["no accountable decision owner"], ["decision-contract", "approval-strip"],
                ["separate conditions from recommendation", "keep the requested action visible in compact density"], "native decision contract remains editable", "each row binds decisionId"),
            D("cost-of-inaction", "decision", "Quantify consequences of delaying or rejecting action", ["consequence"], ["amount", "timing", "probability", "affectedParty"],
                "one item per consequence", ["unsupported amounts presented as facts"], ["consequence-table", "time-escalation"],
                ["separate evidenced amount from qualitative consequence", "sort by decision relevance"], "native consequence rows remain editable", "quantitative consequences bind facts"),
            D("milestone-timeline", "execution", "Show dated milestones, owners and acceptance conditions", ["date", "owner"], ["acceptance", "status", "dependency"],
                "ordered milestones", ["undated milestones"], ["timeline-table", "phase-bands"],
                ["group dense milestones by phase", "preserve chronological order"], "native timeline rows remain editable", "milestones bind actionId"),
            D("owner-time-standard", "execution", "Turn recommendations into accountable action contracts", ["owner", "timeframe", "standard"], ["contributors", "threshold", "escalation"],
                "one item per action", ["multiple final owners for one action"], ["action-contract-table", "owner-led-groups"],
                ["keep owner, time and standard visible in every density", "move contributors to notes when compact"], "native action rows remain editable", "each row binds actionId and facts"),
            D("dependency-graph", "execution", "Expose predecessor-successor dependencies and blocking conditions", ["from", "to"], ["condition", "owner", "status"],
                "directed dependency edges", ["cyclic plan without an explicit iteration rule"], ["dependency-edge-table", "native-diagram"],
                ["topologically order acyclic dependencies", "surface blocking conditions"], "native edge data remain editable and can feed DiagramSpec", "edges bind actionId"),
            D("rag-status", "execution", "Report red-amber-green status with evidence and required action", ["status"], ["evidence", "owner", "action", "threshold"],
                "one item per monitored outcome", ["status without a stated threshold or evidence"], ["status-table", "exception-first"],
                ["sort red then amber then green", "show the triggering evidence beside status"], "native cells and conditional semantics remain editable", "status binds facts and actions"),
        }.ToDictionary(item => item.ComponentId, StringComparer.OrdinalIgnoreCase);
    }

    private static List<List<string>> BuildRows(ProfessionalComponentDefinition definition, ProfessionalComponentSpec spec)
    {
        var fields = definition.RequiredSlots.Concat(definition.OptionalSlots)
            .Where(field => definition.RequiredSlots.Contains(field, StringComparer.OrdinalIgnoreCase)
                || spec.Items.Any(item => item.Fields.ContainsKey(field)))
            .ToList();
        var rows = new List<List<string>> { new() { "Item" } };
        rows[0].AddRange(fields.Select(ToHeader));
        foreach (var item in OrderItems(spec.ComponentId, spec.Items))
        {
            var row = new List<string> { item.Label };
            row.AddRange(fields.Select(field => item.Fields.TryGetValue(field, out var value) ? Display(value) : ""));
            rows.Add(row);
        }
        return rows;
    }

    private static IEnumerable<ProfessionalComponentItem> OrderItems(string componentId, IEnumerable<ProfessionalComponentItem> items)
    {
        if (!componentId.Equals("rag-status", StringComparison.OrdinalIgnoreCase)) return items;
        static int Rank(ProfessionalComponentItem item)
        {
            if (!item.Fields.TryGetValue("status", out var value)) return 4;
            return Display(value).Trim().ToLowerInvariant() switch
            {
                "red" or "r" or "红" or "红色" => 0,
                "amber" or "yellow" or "a" or "黄" or "黄色" => 1,
                "green" or "g" or "绿" or "绿色" => 2,
                _ => 3,
            };
        }
        return items.OrderBy(Rank);
    }

    private static Dictionary<string, string> BuildTableProperties(string extension, ProfessionalComponentDefinition definition,
        ProfessionalComponentSpec spec, List<List<string>> rows, string marker)
    {
        var variant = SelectVariant(spec, rows[0].Count);
        var primary = spec.ThemeTokens.GetValueOrDefault("primary", "1F4E78").TrimStart('#');
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (extension is "docx" or "pptx")
        {
            props["data"] = string.Join(';', rows.Select(row => string.Join(',', row.Select(QuoteCell))));
            props["style"] = spec.Density == "compact" ? "light1" : "medium2";
            if (extension == "docx")
            {
                props["caption"] = marker;
                props["description"] = $"{spec.Title}; {definition.SemanticIntent}";
                props["layout"] = spec.Density == "compact" ? "autofit" : "fixed";
                props["width"] = "100%";
            }
            else
            {
                props["name"] = marker;
                props["x"] = spec.ThemeTokens.GetValueOrDefault("placement.x", "1.2cm");
                props["y"] = spec.ThemeTokens.GetValueOrDefault("placement.y", "2.5cm");
                props["width"] = spec.ThemeTokens.GetValueOrDefault("placement.width", "31.4cm");
                props["height"] = spec.ThemeTokens.GetValueOrDefault("placement.height",
                    spec.Density switch { "compact" => "9cm", "spacious" => "13.5cm", _ => "11.2cm" });
                props["headerFill"] = primary;
                props["bodyFill"] = "F7F9FC";
                props["firstRow"] = "true";
                props["bandedRows"] = "true";
            }
        }
        else
        {
            var start = ParseExcelStart(spec.Target);
            props["ref"] = $"{start}:{ExcelColumn(ExcelColumnNumber(start) + rows[0].Count - 1)}{ExcelRow(start) + rows.Count - 1}";
            props["name"] = ExcelMarker(spec);
            props["displayName"] = ExcelMarker(spec);
            props["comment"] = marker;
            props["style"] = spec.Density == "compact" ? "light1" : "medium2";
            props["showRowStripes"] = "true";
        }
        return props;
    }

    private static DocumentNode? FindNativeObject(IDocumentHandler handler, string extension, ProfessionalComponentSpec spec) =>
        handler.Query("table").FirstOrDefault(node =>
        {
            var marker = ComponentMarker(node, extension);
            return marker is not null && TryParseAnyMarker(marker, out var values)
                && string.Equals(values.GetValueOrDefault("component"), spec.ComponentId, StringComparison.Ordinal)
                && string.Equals(values.GetValueOrDefault("instance"), spec.InstanceId, StringComparison.Ordinal);
        });

    private static string? ComponentMarker(DocumentNode node, string extension)
    {
        var keys = extension switch
        {
            "docx" => new[] { "caption" },
            "xlsx" => new[] { "comment", "displayName", "name" },
            _ => new[] { "name" },
        };
        foreach (var key in keys)
            if (node.Format.TryGetValue(key, out var value) && value is not null)
                return value.ToString();
        return null;
    }

    private static string Marker(ProfessionalComponentSpec spec)
    {
        var values = new Dictionary<string, string>
        {
            ["instance"] = spec.InstanceId,
            ["component"] = spec.ComponentId,
            ["density"] = spec.Density,
            ["variant"] = SelectVariant(spec, 1),
            ["facts"] = string.Join(',', spec.FactRefs),
            ["claims"] = string.Join(',', spec.ClaimRefs),
            ["decisions"] = string.Join(',', spec.DecisionRefs),
            ["actions"] = string.Join(',', spec.ActionRefs),
        };
        return "officecli-component|" + string.Join('|', values.Select(pair => $"{pair.Key}={Uri.EscapeDataString(pair.Value)}"));
    }

    private static bool TryParseMarker(string marker, out Dictionary<string, string> values)
    {
        values = new(StringComparer.OrdinalIgnoreCase);
        if (!marker.StartsWith("officecli-component|", StringComparison.Ordinal)) return false;
        foreach (var segment in marker.Split('|').Skip(1))
        {
            var split = segment.IndexOf('=');
            if (split <= 0) continue;
            values[segment[..split]] = Uri.UnescapeDataString(segment[(split + 1)..]);
        }
        return values.ContainsKey("instance") && values.ContainsKey("component");
    }

    private static bool TryParseAnyMarker(string marker, out Dictionary<string, string> values) =>
        TryParseMarker(marker, out values) || TryParseExcelMarker(marker, out values);

    private static string ExcelMarker(ProfessionalComponentSpec spec) =>
        ExcelObjectName($"oc__{EncodeExcelMarkerPart(spec.ComponentId)}__{EncodeExcelMarkerPart(spec.InstanceId)}__{EncodeExcelMarkerPart(spec.Density)}");

    private static bool TryParseExcelMarker(string marker, out Dictionary<string, string> values)
    {
        values = new(StringComparer.OrdinalIgnoreCase);
        if (!marker.StartsWith("oc__", StringComparison.OrdinalIgnoreCase)) return false;
        var parts = marker.Split("__", StringSplitOptions.None);
        if (parts.Length < 4) return false;
        if (!TryDecodeExcelMarkerPart(parts[1], out var component)
            || !TryDecodeExcelMarkerPart(parts[2], out var instance)
            || !TryDecodeExcelMarkerPart(parts[3], out var density)) return false;
        values["component"] = component;
        values["instance"] = instance;
        values["density"] = density;
        values["variant"] = "adaptive-table";
        return true;
    }

    private static string SelectVariant(ProfessionalComponentSpec spec, int columns) =>
        spec.Items.Count > 8 || columns > 6 ? "compact-table" : spec.Items.Count <= 4 && spec.Density == "spacious" ? "spacious-focus" : "balanced-table";

    private static void ValidateRequiredSlots(ProfessionalComponentDefinition definition, ProfessionalComponentSpec spec)
    {
        foreach (var item in spec.Items)
        foreach (var required in definition.RequiredSlots)
            if (!item.Fields.TryGetValue(required, out var value) || string.IsNullOrWhiteSpace(Display(value)))
                throw new CliException($"Component '{definition.ComponentId}' item '{item.Label}' requires field '{required}'.")
                { Code = "component_slot_missing" };
        if (spec.FactRefs.Count == 0 && definition.Category == "operating-analysis")
            throw new CliException($"Component '{definition.ComponentId}' requires factRefs for traceability.")
            { Code = "component_binding_missing" };
        if (spec.DecisionRefs.Count == 0 && definition.Category == "decision")
            throw new CliException($"Component '{definition.ComponentId}' requires decisionRefs for traceability.")
            { Code = "component_binding_missing" };
        if (spec.ActionRefs.Count == 0 && definition.Category == "execution")
            throw new CliException($"Component '{definition.ComponentId}' requires actionRefs for traceability.")
            { Code = "component_binding_missing" };
    }

    private static string DefaultTarget(string extension) => extension switch { "docx" => "/body", "xlsx" => "/Sheet1", _ => "/slide[1]" };
    private static string ParseExcelStart(string? target)
    {
        if (string.IsNullOrWhiteSpace(target)) return "A1";
        var last = target.Split('/').Last();
        return System.Text.RegularExpressions.Regex.IsMatch(last, "^[A-Za-z]{1,3}[1-9][0-9]*$") ? last.ToUpperInvariant() : "A1";
    }
    private static int ExcelRow(string cell) => int.Parse(new string(cell.Where(char.IsDigit).ToArray()), CultureInfo.InvariantCulture);
    private static int ExcelColumnNumber(string cell)
    {
        var result = 0;
        foreach (var c in cell.TakeWhile(char.IsLetter)) result = result * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        return result;
    }
    private static string ExcelColumn(int number)
    {
        var result = "";
        while (number > 0) { number--; result = (char)('A' + number % 26) + result; number /= 26; }
        return result;
    }
    private static string ExcelObjectName(string value)
    {
        var sanitized = new string(value.Select(c => char.IsLetterOrDigit(c) || c == '_' ? c : '_').ToArray());
        if (string.IsNullOrEmpty(sanitized) || char.IsDigit(sanitized[0])) sanitized = "component_" + sanitized;
        return sanitized.Length > 240 ? sanitized[..240] : sanitized;
    }
    private static string EncodeExcelMarkerPart(string value) => Convert.ToHexString(System.Text.Encoding.UTF8.GetBytes(value));
    private static bool TryDecodeExcelMarkerPart(string value, out string decoded)
    {
        decoded = "";
        try
        {
            decoded = System.Text.Encoding.UTF8.GetString(Convert.FromHexString(value));
            return decoded.Length > 0;
        }
        catch (FormatException)
        {
            return false;
        }
    }
    private static string Display(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.String => value.GetString() ?? "",
        JsonValueKind.Number => value.GetRawText(),
        JsonValueKind.True => "true",
        JsonValueKind.False => "false",
        JsonValueKind.Null => "",
        _ => value.GetRawText(),
    };
    private static string QuoteCell(string value) => value.IndexOfAny([',', ';', '"', '\n', '\r']) >= 0
        ? $"\"{value.Replace("\"", "\"\"")}\"" : value;
    private static string ToHeader(string value) => string.Concat(value.Select((c, i) => char.IsUpper(c) && i > 0 ? " " + c : c.ToString()));
    private static IReadOnlyList<string> SplitRefs(string value) => value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
    private static void EnsureUnique(List<string> values, string field)
    {
        if (values.Count != values.Distinct(StringComparer.Ordinal).Count() || values.Any(string.IsNullOrWhiteSpace))
            throw new CliException($"ComponentSpec {field} must contain unique, non-empty identifiers.") { Code = "component_spec_invalid" };
    }
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, WriteIndented = true)]
[JsonSerializable(typeof(ProfessionalComponentSpec))]
[JsonSerializable(typeof(ProfessionalComponentReceipt))]
[JsonSerializable(typeof(ProfessionalComponentListResponse))]
[JsonSerializable(typeof(ProfessionalComponentDescribeResponse))]
[JsonSerializable(typeof(ProfessionalComponentReadResponse))]
internal partial class ProfessionalComponentJsonContext : JsonSerializerContext;
