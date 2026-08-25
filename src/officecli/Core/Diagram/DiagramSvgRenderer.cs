// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Security;
using System.Text;
using System.Text.Json;

namespace OfficeCli.Core.Diagram;

public static class DiagramSvgRenderer
{
    public static string Render(DiagramSpec spec, LaidOutGraph graph, DiagramTheme theme)
    {
        string N(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);
        string E(string? value) => SecurityElement.Escape(value ?? "") ?? "";
        var nodeSpecs = spec.Nodes.ToDictionary(item => item.Id, StringComparer.Ordinal);
        var builder = new StringBuilder();
        builder.Append($"<svg xmlns=\"http://www.w3.org/2000/svg\" data-diagram-id=\"{E(spec.DiagramId)}\"");
        if (spec.Communication is not null)
            builder.Append($" data-communication-intent-id=\"{E(spec.Communication.IntentId)}\" data-representation-choice-id=\"{E(spec.Communication.RepresentationChoiceId)}\"");
        builder.Append($" width=\"{N(graph.SlideWidthCm)}cm\" height=\"{N(graph.SlideHeightCm)}cm\" viewBox=\"0 0 {N(graph.SlideWidthCm)} {N(graph.SlideHeightCm)}\" role=\"img\" aria-labelledby=\"diagram-title\">");
        builder.Append($"<title id=\"diagram-title\">{E(spec.Title ?? spec.DiagramId)}</title>");
        builder.Append($"<rect width=\"100%\" height=\"100%\" fill=\"#{theme.Background}\"/>");
        builder.Append("<g class=\"diagram-edges\" fill=\"none\">");
        for (var edgeIndex = 0; edgeIndex < graph.Edges.Count; edgeIndex++)
        {
            var edge = graph.Edges[edgeIndex];
            if (edge.Points.Count < 2) continue;
            var points = string.Join(" ", edge.Points.Select(point => $"{N(point.X)},{N(point.Y)}"));
            builder.Append($"<polyline points=\"{points}\" stroke=\"#{theme.MutedText}\" stroke-width=\"0.05\" stroke-linejoin=\"round\" stroke-linecap=\"round\"");
            if (edge.Dashed) builder.Append(" stroke-dasharray=\"0.18 0.12\"");
            builder.Append("/>");
            if (edge.ArrowAtEnd) AppendArrow(builder, edge, theme, N);
        }
        builder.Append("</g><g class=\"diagram-nodes\">");
        foreach (var node in graph.Nodes)
        {
            var style = DiagramStyles.Resolve(node.Shape, theme);
            var factRefs = nodeSpecs.TryGetValue(node.Id, out var nodeSpec) ? string.Join(" ", nodeSpec.FactRefs) : "";
            builder.Append($"<g id=\"node-{E(node.Id)}\" data-node-id=\"{E(node.Id)}\" data-fact-refs=\"{E(factRefs)}\">");
            AppendShape(builder, node, style.Fill, style.Line, N);
            AppendText(builder, node.Label, node.X + node.W / 2, node.Y + node.H / 2, node.W, theme.Text,
                theme.MinorLatinFont, theme.MinorEastAsiaFont, 0.42, N, E);
            builder.Append("</g>");
        }
        builder.Append("</g><g class=\"diagram-labels\">");
        foreach (var label in graph.Labels)
        {
            if (label.Opaque) builder.Append($"<rect x=\"{N(label.Cx - label.W / 2)}\" y=\"{N(label.Cy - label.H / 2)}\" width=\"{N(label.W)}\" height=\"{N(label.H)}\" fill=\"#{theme.Background}\"/>");
            AppendText(builder, label.Text, label.Cx, label.Cy, label.W, theme.Text,
                theme.MinorLatinFont, theme.MinorEastAsiaFont, 0.34, N, E);
        }
        builder.Append("</g></svg>");
        return builder.ToString();
    }

    public static DiagramEvidence Evidence(DiagramSpec spec, LaidOutGraph graph, DiagramTheme? theme = null)
    {
        theme ??= DiagramTheme.Default;
        var overlaps = FindOverlaps(graph.Nodes);
        var textOverflows = FindTextOverflows(graph);
        var factsComplete = spec.Nodes.All(node => node.FactRefs.All(id => spec.Facts.Any(fact => fact.FactId == id)))
            && spec.Edges.All(edge => edge.FactRefs.All(id => spec.Facts.Any(fact => fact.FactId == id)));
        var nodeIds = graph.Nodes.Select(node => node.Id).ToHashSet(StringComparer.Ordinal);
        var semanticEdges = graph.Edges.Where(edge => spec.Edges.Any(item => item.Id == edge.Id)).ToList();
        var structureComplete = spec.Nodes.All(node => nodeIds.Contains(node.Id))
            && spec.Edges.All(edge => semanticEdges.Count(item => item.Id == edge.Id) == 1);
        var attachmentsComplete = semanticEdges.All(edge => !string.IsNullOrWhiteSpace(edge.SourceNodeId)
            && !string.IsNullOrWhiteSpace(edge.TargetNodeId));
        var (typePassed, typeReason) = EvaluateType(spec);
        var communicationComplete = spec.Communication is null ||
            (!string.IsNullOrWhiteSpace(spec.Communication.IntentId)
             && !string.IsNullOrWhiteSpace(spec.Communication.RepresentationChoiceId)
             && !string.IsNullOrWhiteSpace(spec.Communication.Purpose)
             && !string.IsNullOrWhiteSpace(spec.Communication.Audience)
             && !string.IsNullOrWhiteSpace(spec.Communication.DesiredResponse)
             && !string.IsNullOrWhiteSpace(spec.Communication.CoreMessage));
        // Split for semantic density as well as physical extent.  A short vertical
        // architecture chain can safely be fit to a slide/page even when its
        // conservative CJK boxes are just over 24 cm high; treating that as a
        // second page creates needless manual repair.  Conversely, 19+ semantic
        // nodes are too dense for one editable Office canvas even if the layout
        // engine happens to pack them into the physical bounds.
        var nodeBudget = Math.Min(18, spec.Communication?.MaxNodesPerDiagram ?? 18);
        var requiresSplit = spec.Nodes.Count > nodeBudget || graph.SlideWidthCm > 32 || graph.SlideHeightCm > 26;
        var suggestedPages = requiresSplit ? Math.Max(2, (int)Math.Ceiling(spec.Nodes.Count / (double)nodeBudget)) : 1;
        var repairActions = overlaps.Count + textOverflows.Count
            + (factsComplete ? 0 : 1) + (structureComplete ? 0 : 1)
            + (attachmentsComplete ? 0 : 1) + (typePassed ? 0 : 1)
            + (communicationComplete ? 0 : 1) + (requiresSplit ? 1 : 0);
        var themeFingerprint = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(
            Encoding.UTF8.GetBytes(string.Join("|", theme.Primary, theme.Secondary, theme.Text,
                theme.MajorLatinFont, theme.MajorEastAsiaFont, theme.MinorLatinFont, theme.MinorEastAsiaFont))))
            .ToLowerInvariant();
        return new DiagramEvidence(1, spec.DiagramId, spec.Type, graph.Nodes.Count, spec.Edges.Count, 1,
            spec.Facts.Count, graph.SlideWidthCm, graph.SlideHeightCm, overlaps, textOverflows,
            factsComplete, structureComplete, attachmentsComplete, typePassed, typeReason, communicationComplete,
            spec.Communication?.IntentId, spec.Communication?.RepresentationChoiceId, themeFingerprint,
            theme.MajorEastAsiaFont, theme.MinorEastAsiaFont, requiresSplit, suggestedPages, repairActions);
    }

    private static List<string> FindOverlaps(IReadOnlyList<PlacedNode> nodes)
    {
        var overlaps = new List<string>();
        for (var i = 0; i < nodes.Count; i++)
            for (var j = i + 1; j < nodes.Count; j++)
                if (nodes[i].X < nodes[j].X + nodes[j].W && nodes[j].X < nodes[i].X + nodes[i].W &&
                    nodes[i].Y < nodes[j].Y + nodes[j].H && nodes[j].Y < nodes[i].Y + nodes[i].H)
                    overlaps.Add($"{nodes[i].Id}:{nodes[j].Id}");
        return overlaps;
    }

    private static List<string> FindTextOverflows(LaidOutGraph graph)
    {
        var issues = new List<string>();
        foreach (var node in graph.Nodes)
        {
            var lines = DiagramTextMetrics.Wrap(node.Label, Math.Max(0.8, node.W - 0.6)).Count;
            if (lines * DiagramTextMetrics.NodeLineHeightCm + 0.35 > node.H)
                issues.Add("node:" + node.Id);
        }
        for (var index = 0; index < graph.Labels.Count; index++)
        {
            var label = graph.Labels[index];
            var lines = DiagramTextMetrics.Wrap(label.Text, Math.Max(0.6, label.W - 0.3)).Count;
            if (lines * DiagramTextMetrics.EdgeLineHeightCm + 0.12 > label.H)
                issues.Add("edge-label:" + index);
        }
        return issues;
    }

    private static (bool Passed, string Reason) EvaluateType(DiagramSpec spec)
    {
        if (spec.Type == "sequence")
            return (spec.Edges.Count > 0, "sequence requires participant messages");
        if (spec.Type == "timeline")
        {
            var degree = spec.Nodes.ToDictionary(node => node.Id, _ => (In: 0, Out: 0), StringComparer.Ordinal);
            foreach (var edge in spec.Edges)
            {
                degree[edge.From] = (degree[edge.From].In, degree[edge.From].Out + 1);
                degree[edge.To] = (degree[edge.To].In + 1, degree[edge.To].Out);
            }
            var linear = degree.Values.All(value => value.In <= 1 && value.Out <= 1)
                && (spec.Edges.Count == 0 || spec.Edges.Count == spec.Nodes.Count - 1);
            return (linear, linear ? "timeline is a single ordered chain" : "timeline contains branching or disconnected events");
        }
        if (spec.Type == "mindmap")
        {
            var rootCount = spec.Nodes.Count(node => spec.Edges.All(edge => edge.To != node.Id));
            var tree = spec.Edges.Count == spec.Nodes.Count - 1 && rootCount == 1 && IsWeaklyConnected(spec);
            return (tree, tree ? "mindmap is one rooted hierarchy" : "mindmap is not a connected rooted hierarchy");
        }
        return (true, spec.Type switch
        {
            "flowchart" => "process or decision graph",
            "relationship" => "general many-to-many relationship graph",
            "architecture" => "component and boundary graph",
            _ => "supported semantic graph",
        });
    }

    private static bool IsWeaklyConnected(DiagramSpec spec)
    {
        var adjacent = spec.Nodes.ToDictionary(node => node.Id, _ => new List<string>(), StringComparer.Ordinal);
        foreach (var edge in spec.Edges) { adjacent[edge.From].Add(edge.To); adjacent[edge.To].Add(edge.From); }
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var pending = new Queue<string>(); pending.Enqueue(spec.Nodes[0].Id);
        while (pending.Count > 0)
        {
            var node = pending.Dequeue(); if (!seen.Add(node)) continue;
            foreach (var next in adjacent[node]) pending.Enqueue(next);
        }
        return seen.Count == spec.Nodes.Count;
    }

    private static void AppendArrow(StringBuilder builder, RoutedEdge edge, DiagramTheme theme, Func<double, string> n)
    {
        var end = edge.Points[^1]; var previous = edge.Points[^2];
        var angle = Math.Atan2(end.Y - previous.Y, end.X - previous.X);
        const double size = 0.18;
        var left = new Pt(end.X - size * Math.Cos(angle - 0.55), end.Y - size * Math.Sin(angle - 0.55));
        var right = new Pt(end.X - size * Math.Cos(angle + 0.55), end.Y - size * Math.Sin(angle + 0.55));
        builder.Append($"<path d=\"M {n(end.X)} {n(end.Y)} L {n(left.X)} {n(left.Y)} L {n(right.X)} {n(right.Y)} Z\" fill=\"#{theme.MutedText}\"/>");
    }

    private static void AppendShape(StringBuilder builder, PlacedNode node, string fill, string line, Func<double, string> n)
    {
        var common = $"fill=\"#{fill}\" stroke=\"#{line}\" stroke-width=\"0.05\"";
        if (node.Shape == FlowShape.Circle)
            builder.Append($"<ellipse cx=\"{n(node.X + node.W / 2)}\" cy=\"{n(node.Y + node.H / 2)}\" rx=\"{n(node.W / 2)}\" ry=\"{n(node.H / 2)}\" {common}/>");
        else if (node.Shape == FlowShape.Decision)
            builder.Append($"<polygon points=\"{n(node.X + node.W / 2)},{n(node.Y)} {n(node.X + node.W)},{n(node.Y + node.H / 2)} {n(node.X + node.W / 2)},{n(node.Y + node.H)} {n(node.X)},{n(node.Y + node.H / 2)}\" {common}/>");
        else
            builder.Append($"<rect x=\"{n(node.X)}\" y=\"{n(node.Y)}\" width=\"{n(node.W)}\" height=\"{n(node.H)}\" rx=\"{(node.Shape is FlowShape.Terminator or FlowShape.Stadium ? "0.35" : "0.08")}\" {common}/>");
    }

    private static void AppendText(StringBuilder builder, string text, double cx, double cy, double width, string color,
                                   string latinFont, string eastAsiaFont, double fontSize,
                                   Func<double, string> n, Func<string?, string> e)
    {
        var lines = DiagramTextMetrics.Wrap(text, Math.Max(0.6, width - 0.3));
        var start = cy - (lines.Count - 1) * 0.22;
        builder.Append($"<text x=\"{n(cx)}\" y=\"{n(start)}\" text-anchor=\"middle\" dominant-baseline=\"middle\" font-family=\"{e(eastAsiaFont)},{e(latinFont)},sans-serif\" font-size=\"{n(fontSize)}\" fill=\"#{color}\">");
        for (var index = 0; index < lines.Count; index++)
            builder.Append($"<tspan x=\"{n(cx)}\" dy=\"{(index == 0 ? "0" : "0.44")}\">{e(lines[index])}</tspan>");
        builder.Append("</text>");
    }

}
