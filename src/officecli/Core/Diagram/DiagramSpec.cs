// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace OfficeCli.Core.Diagram;

/// <summary>
/// Versioned, format-neutral contract between semantic planning and deterministic
/// Office rendering.  It deliberately contains no OOXML and no coordinates:
/// callers describe meaning, evidence and relationships; OfficeCLI owns layout.
/// </summary>
public sealed class DiagramSpec
{
    public int SchemaVersion { get; set; }
    public string DiagramId { get; set; } = "";
    public string Type { get; set; } = "flowchart";
    public string? Title { get; set; }
    public string Direction { get; set; } = "top-down";
    public DiagramCommunication? Communication { get; set; }
    public List<DiagramSpecFact> Facts { get; set; } = new();
    public List<DiagramSpecNode> Nodes { get; set; } = new();
    public List<DiagramSpecEdge> Edges { get; set; } = new();

    public static DiagramSpec Load(string path)
    {
        if (!File.Exists(path))
            throw new ArgumentException($"DiagramSpec file not found: '{path}'.");
        return Parse(File.ReadAllText(path));
    }

    public static DiagramSpec Parse(string json)
    {
        DiagramSpec? spec;
        try { spec = JsonSerializer.Deserialize(json, DiagramJsonContext.Default.DiagramSpec); }
        catch (JsonException ex) { throw new ArgumentException($"invalid DiagramSpec JSON: {ex.Message}", ex); }
        if (spec is null) throw new ArgumentException("DiagramSpec JSON is empty.");
        spec.Validate();
        return spec;
    }

    public void Validate()
    {
        if (SchemaVersion != 1)
            throw new ArgumentException("DiagramSpec schemaVersion must be 1.");
        RequireId(DiagramId, "diagramId");
        if (!SupportedTypes.Contains(Type))
            throw new ArgumentException($"unsupported DiagramSpec type '{Type}' (supported: {string.Join(", ", SupportedTypes)}).");
        if (Direction is not ("top-down" or "left-right"))
            throw new ArgumentException("DiagramSpec direction must be 'top-down' or 'left-right'.");
        if (Title is { Length: > 240 }) throw new ArgumentException("DiagramSpec title exceeds 240 characters.");
        if (Nodes.Count is < 1 or > 24)
            throw new ArgumentException("DiagramSpec requires 1-24 nodes; split larger diagrams at a semantic subsystem or phase boundary.");
        if (Communication is not null)
        {
            RequireId(Communication.IntentId, "communication intentId");
            RequireId(Communication.RepresentationChoiceId, "communication representationChoiceId");
            RequireText(Communication.Purpose, "communication purpose", 500);
            RequireText(Communication.Audience, "communication audience", 240);
            RequireText(Communication.DesiredResponse, "communication desiredResponse", 500);
            RequireText(Communication.CoreMessage, "communication coreMessage", 500);
            if (Communication.MaxNodesPerDiagram is < 3 or > 24)
                throw new ArgumentException("communication maxNodesPerDiagram must be between 3 and 24.");
            if (Nodes.Count > Communication.MaxNodesPerDiagram)
                throw new ArgumentException("DiagramSpec exceeds the CommunicationIntent node-density budget; split it at a semantic boundary.");
        }
        if (Edges.Count > 400) throw new ArgumentException("DiagramSpec supports at most 400 edges.");
        if (Facts.Count > 1000) throw new ArgumentException("DiagramSpec supports at most 1000 fact references.");

        var factIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var fact in Facts)
        {
            RequireId(fact.FactId, "factId");
            if (!factIds.Add(fact.FactId)) throw new ArgumentException($"duplicate DiagramSpec factId '{fact.FactId}'.");
            RequireText(fact.SourceId, "fact sourceId", 200);
            RequireText(fact.Locator, "fact locator", 500);
            RequireText(fact.Summary, "fact summary", 1000);
            if (fact.Confidence is < 0 or > 1 || double.IsNaN(fact.Confidence) || double.IsInfinity(fact.Confidence))
                throw new ArgumentException($"fact '{fact.FactId}' confidence must be between 0 and 1.");
        }

        var nodeIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var node in Nodes)
        {
            RequireId(node.Id, "node id");
            if (!nodeIds.Add(node.Id)) throw new ArgumentException($"duplicate DiagramSpec node id '{node.Id}'.");
            RequireText(node.Label, $"node '{node.Id}' label", 500);
            if (node.Shape is not null && !SupportedShapes.Contains(node.Shape))
                throw new ArgumentException($"node '{node.Id}' uses unsupported shape '{node.Shape}'.");
            ValidateFactRefs(node.FactRefs, factIds, $"node '{node.Id}'");
        }

        var edgeIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var edge in Edges)
        {
            RequireId(edge.Id, "edge id");
            if (!edgeIds.Add(edge.Id)) throw new ArgumentException($"duplicate DiagramSpec edge id '{edge.Id}'.");
            if (!nodeIds.Contains(edge.From) || !nodeIds.Contains(edge.To))
                throw new ArgumentException($"edge '{edge.Id}' references an unknown node.");
            if (edge.Label is { Length: > 500 }) throw new ArgumentException($"edge '{edge.Id}' label exceeds 500 characters.");
            ValidateFactRefs(edge.FactRefs, factIds, $"edge '{edge.Id}'");
        }

        if (Type == "sequence" && Edges.Count == 0)
            throw new ArgumentException("sequence DiagramSpec requires at least one message edge.");
    }

    internal static readonly string[] SupportedTypes =
        ["flowchart", "mindmap", "relationship", "architecture", "sequence", "timeline"];
    internal static readonly HashSet<string> SupportedShapes = new(StringComparer.Ordinal)
        { "process", "decision", "terminator", "circle", "hexagon", "parallelogram", "database", "subroutine" };
    private static readonly Regex IdPattern = new("^[a-z][a-z0-9._-]{0,63}$", RegexOptions.CultureInvariant);

    private static void RequireId(string value, string label)
    {
        if (!IdPattern.IsMatch(value)) throw new ArgumentException($"{label} must match ^[a-z][a-z0-9._-]{{0,63}}$.");
    }

    private static void RequireText(string value, string label, int max)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > max)
            throw new ArgumentException($"{label} must contain 1-{max} characters.");
    }

    private static void ValidateFactRefs(IEnumerable<string> refs, HashSet<string> factIds, string owner)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var factId in refs)
        {
            if (!seen.Add(factId)) throw new ArgumentException($"{owner} repeats factRef '{factId}'.");
            if (!factIds.Contains(factId)) throw new ArgumentException($"{owner} references unknown fact '{factId}'.");
        }
    }
}

public sealed record DiagramEvidence(
    int SchemaVersion,
    string DiagramId,
    string Type,
    int NodeCount,
    int EdgeCount,
    int GroupCount,
    int FactCount,
    double WidthCm,
    double HeightCm,
    List<string> NodeOverlaps,
    List<string> TextOverflows,
    bool FactBindingsComplete,
    bool StructureComplete,
    bool ConnectorAttachmentsComplete,
    bool TypeSelectionPassed,
    string TypeSelectionReason,
    bool CommunicationBindingComplete,
    string? CommunicationIntentId,
    string? RepresentationChoiceId,
    string ThemeFingerprint,
    string MajorFont,
    string BodyFont,
    bool RequiresSemanticSplit,
    int SuggestedPageCount,
    int EstimatedManualRepairActions);

public sealed record DiagramCommandReceipt(
    bool Ok,
    int SchemaVersion,
    string DiagramId,
    string Type,
    string Output,
    string? Evidence,
    int NodeCount,
    int EdgeCount);

public sealed record DiagramRefreshReceipt(
    bool Ok,
    int SchemaVersion,
    string DiagramId,
    string File,
    string Host,
    int NodeCount,
    int EdgeCount,
    string Sha256);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, WriteIndented = true,
    UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow)]
[JsonSerializable(typeof(DiagramSpec))]
[JsonSerializable(typeof(DiagramEvidence))]
[JsonSerializable(typeof(DiagramCommandReceipt))]
[JsonSerializable(typeof(DiagramRefreshReceipt))]
[JsonSerializable(typeof(SmartArtPackageInspection))]
[JsonSerializable(typeof(SmartArtUpdateReceipt))]
internal partial class DiagramJsonContext : JsonSerializerContext;

public sealed class DiagramCommunication
{
    public string IntentId { get; set; } = "";
    public string RepresentationChoiceId { get; set; } = "";
    public string Purpose { get; set; } = "";
    public string Audience { get; set; } = "";
    public string DesiredResponse { get; set; } = "";
    public string CoreMessage { get; set; } = "";
    public int MaxNodesPerDiagram { get; set; }
}

public sealed class DiagramSpecFact
{
    public string FactId { get; set; } = "";
    public string SourceId { get; set; } = "";
    public string Locator { get; set; } = "";
    public string Summary { get; set; } = "";
    public double Confidence { get; set; } = 1;
}

public sealed class DiagramSpecNode
{
    public string Id { get; set; } = "";
    public string Label { get; set; } = "";
    public string? Shape { get; set; }
    public List<string> FactRefs { get; set; } = new();
}

public sealed class DiagramSpecEdge
{
    public string Id { get; set; } = "";
    public string From { get; set; } = "";
    public string To { get; set; } = "";
    public string? Label { get; set; }
    public bool Dashed { get; set; }
    public List<string> FactRefs { get; set; } = new();
}
