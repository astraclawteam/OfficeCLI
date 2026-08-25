// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Text.RegularExpressions;

namespace OfficeCli.Core.Diagram;

/// <summary>
/// Single semantic-to-geometric compiler for native Office shapes and SVG.
/// DiagramSpec is the primary contract; Mermaid remains a compatibility input
/// adapter that maps into the same layout engine. Unsupported compatibility
/// syntax is rejected rather than creating a second, lower-fidelity pipeline.
/// </summary>
public static class DiagramCompiler
{
    public static LaidOutGraph Compile(DiagramSpec spec)
    {
        spec.Validate();
        if (spec.Type == "sequence")
        {
            var sequence = new SequenceDiagram();
            foreach (var node in spec.Nodes)
            {
                var participant = sequence.See(node.Id);
                participant.Label = node.Label;
                participant.FactRefs = new List<string>(node.FactRefs);
            }
            foreach (var edge in spec.Edges)
                sequence.Messages.Add(new SeqMessage
                {
                    Id = edge.Id, From = edge.From, To = edge.To, Label = edge.Label ?? "", Dashed = edge.Dashed,
                    Arrow = true, FactRefs = new List<string>(edge.FactRefs),
                });
            var layout = SequenceLayout.Layout(sequence);
            layout.DiagramId = spec.DiagramId;
            return layout;
        }

        var graph = new DiagramGraph
        {
            Direction = spec.Direction == "left-right" || spec.Type is "timeline" or "mindmap"
                ? FlowDirection.LeftRight : FlowDirection.TopDown,
        };
        foreach (var source in spec.Nodes)
        {
            var node = graph.GetOrAdd(source.Id);
            node.Label = source.Label;
            node.Shape = ParseShape(source.Shape, spec.Type);
            node.FactRefs = new List<string>(source.FactRefs);
        }
        var edges = spec.Edges;
        if (spec.Type == "timeline" && edges.Count == 0)
            edges = spec.Nodes.Zip(spec.Nodes.Skip(1), (from, to) => new DiagramSpecEdge
                { Id = $"{from.Id}-{to.Id}", From = from.Id, To = to.Id }).ToList();
        foreach (var edge in edges)
            graph.Edges.Add(new DiagramEdge { Id = edge.Id, From = edge.From, To = edge.To,
                Label = edge.Label ?? "", FactRefs = new List<string>(edge.FactRefs) });
        var result = FlowchartLayout.Layout(graph);
        result.DiagramId = spec.DiagramId;
        return result;
    }

    public static LaidOutGraph Compile(string mermaid)
    {
        var header = FirstMeaningfulLine(mermaid);

        if (Regex.IsMatch(header, @"^(flowchart|graph)\b", RegexOptions.IgnoreCase))
            return FlowchartLayout.Layout(MermaidParser.Parse(mermaid));

        if (Regex.IsMatch(header, @"^sequenceDiagram\b", RegexOptions.IgnoreCase))
            return SequenceLayout.Layout(SequenceLayout.Parse(mermaid));

        // No explicit header → assume flowchart (mermaid's own lenient default).
        if (header.Length == 0 || !Regex.IsMatch(header, @"^[A-Za-z]"))
            return FlowchartLayout.Layout(MermaidParser.Parse(mermaid));

        var kind = Regex.Match(header, @"^[A-Za-z]+").Value;
        throw new ArgumentException(
            $"diagram type '{kind}' is not supported yet (currently: flowchart, sequenceDiagram).");
    }

    private static string FirstMeaningfulLine(string text)
    {
        foreach (var raw in text.Split('\n'))
        {
            var s = raw.Trim();
            if (s.Length > 0 && !s.StartsWith("%%"))
                return s;
        }
        return "";
    }

    private static FlowShape ParseShape(string? value, string diagramType) => value switch
    {
        "decision" => FlowShape.Decision,
        "terminator" => FlowShape.Terminator,
        "circle" => FlowShape.Circle,
        "hexagon" => FlowShape.Hexagon,
        "parallelogram" => FlowShape.Parallelogram,
        "database" => FlowShape.Database,
        "subroutine" => FlowShape.Subroutine,
        "process" => FlowShape.Process,
        null when diagramType == "mindmap" => FlowShape.Stadium,
        null when diagramType == "architecture" => FlowShape.Subroutine,
        _ => FlowShape.Process,
    };
}
