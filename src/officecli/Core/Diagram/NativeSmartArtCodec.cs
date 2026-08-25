// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace OfficeCli.Core.Diagram;

/// <summary>
/// Package-level SmartArt reader and deterministic DiagramSpec adapter.
/// SmartArt is deliberately a renderer of the same DiagramSpec contract used
/// by native shapes; it is not a second semantic model and never accepts raw
/// OOXML from an Agent.
/// </summary>
internal static class NativeSmartArtCodec
{
    private static readonly XNamespace Dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Ocli = "urn:officecli:diagram:v1";

    internal static string BuildDataXml(DiagramSpec spec, IReadOnlyDictionary<string, string>? sourceProfile = null)
    {
        spec.Validate();
        var docId = StableGuid(spec.DiagramId, "document");
        var modelIds = spec.Nodes.ToDictionary(n => n.Id, n => StableGuid(spec.DiagramId, $"node:{n.Id}"), StringComparer.Ordinal);

        var docProps = new XElement(Dgm + "prSet",
            new XAttribute("loTypeId", Profile(sourceProfile, "loTypeId", "urn:microsoft.com/office/officeart/2005/8/layout/default")),
            new XAttribute("loCatId", Profile(sourceProfile, "loCatId", "list")),
            new XAttribute("qsTypeId", Profile(sourceProfile, "qsTypeId", "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple5")),
            new XAttribute("qsCatId", Profile(sourceProfile, "qsCatId", "simple")),
            new XAttribute("csTypeId", Profile(sourceProfile, "csTypeId", "urn:microsoft.com/office/officeart/2005/8/colors/colorful4")),
            new XAttribute("csCatId", Profile(sourceProfile, "csCatId", "colorful")),
            new XAttribute("presName", $"officecli:{spec.DiagramId}"));

        var points = new XElement(Dgm + "ptLst",
            new XElement(Dgm + "pt",
                new XAttribute("modelId", docId),
                new XAttribute("type", "doc"),
                docProps,
                new XElement(Dgm + "spPr"),
                EmptyText()));

        foreach (var node in spec.Nodes)
        {
            points.Add(new XElement(Dgm + "pt",
                new XAttribute("modelId", modelIds[node.Id]),
                new XElement(Dgm + "prSet", new XAttribute("phldrT", $"[officecli:{node.Id}]")),
                new XElement(Dgm + "spPr"),
                TextBody(node.Label)));
        }

        // The public default layout is an ordered SmartArt list. Every semantic
        // node is therefore a direct document child so it is visibly editable in
        // Office. The complete graph remains in the standard extension point
        // below and round-trips through inspect/update without inventing a second
        // model. Graphs needing visible arbitrary edges use the default
        // Shape+Connector renderer instead.
        var connections = new XElement(Dgm + "cxnLst");
        for (var i = 0; i < spec.Nodes.Count; i++)
        {
            var node = spec.Nodes[i];
            connections.Add(new XElement(Dgm + "cxn",
                new XAttribute("modelId", StableGuid(spec.DiagramId, $"list:{node.Id}")),
                new XAttribute("srcId", docId),
                new XAttribute("destId", modelIds[node.Id]),
                new XAttribute("srcOrd", i),
                new XAttribute("destOrd", 0)));
        }

        var semantic = new XElement(Ocli + "diagram",
            new XAttribute("schemaVersion", 1),
            new XAttribute("diagramId", spec.DiagramId),
            new XAttribute("type", spec.Type),
            new XAttribute("direction", spec.Direction));
        if (!string.IsNullOrWhiteSpace(spec.Title)) semantic.Add(new XAttribute("title", spec.Title));
        foreach (var node in spec.Nodes)
        {
            semantic.Add(new XElement(Ocli + "node",
                new XAttribute("id", node.Id),
                new XAttribute("modelId", modelIds[node.Id]),
                new XAttribute("shape", node.Shape ?? "process"),
                new XAttribute("factRefs", string.Join(";", node.FactRefs))));
        }
        foreach (var edge in spec.Edges)
        {
            semantic.Add(new XElement(Ocli + "edge",
                new XAttribute("id", edge.Id),
                new XAttribute("from", edge.From),
                new XAttribute("to", edge.To),
                new XAttribute("label", edge.Label ?? ""),
                new XAttribute("dashed", edge.Dashed ? "true" : "false"),
                new XAttribute("factRefs", string.Join(";", edge.FactRefs))));
        }
        foreach (var fact in spec.Facts)
        {
            semantic.Add(new XElement(Ocli + "fact",
                new XAttribute("id", fact.FactId),
                new XAttribute("sourceId", fact.SourceId),
                new XAttribute("locator", fact.Locator),
                new XAttribute("confidence", fact.Confidence),
                new XAttribute("summary", fact.Summary)));
        }

        var root = new XElement(Dgm + "dataModel",
            new XAttribute(XNamespace.Xmlns + "dgm", Dgm),
            new XAttribute(XNamespace.Xmlns + "a", A),
            points,
            connections,
            new XElement(Dgm + "bg"),
            new XElement(Dgm + "whole"),
            new XElement(Dgm + "extLst",
                new XElement(A + "ext",
                    new XAttribute("uri", "urn:officecli:diagram:v1"),
                    semantic)));
        return root.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>Official Office default-list layout, reduced only by removing gallery sample decoration.</summary>
    internal static string BuildLayoutXml() => """
    <dgm:layoutDef uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/default" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="list" pri="400"/></dgm:catLst>
      <dgm:layoutNode name="diagram">
        <dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst>
        <dgm:choose name="direction"><dgm:if name="normal" func="var" arg="dir" op="equ" val="norm"><dgm:alg type="snake"><dgm:param type="grDir" val="tL"/><dgm:param type="flowDir" val="row"/><dgm:param type="contDir" val="sameDir"/><dgm:param type="off" val="ctr"/></dgm:alg></dgm:if><dgm:else name="reverse"><dgm:alg type="snake"><dgm:param type="grDir" val="tR"/><dgm:param type="flowDir" val="row"/><dgm:param type="contDir" val="sameDir"/><dgm:param type="off" val="ctr"/></dgm:alg></dgm:else></dgm:choose>
        <dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/>
        <dgm:constrLst><dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="0.6"/><dgm:constr type="w" for="ch" forName="sibTrans" refType="w" refFor="ch" refForName="node" fact="0.1"/><dgm:constr type="sp" refType="w" refFor="ch" refForName="sibTrans"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/></dgm:constrLst><dgm:ruleLst/>
        <dgm:forEach name="nodes" axis="ch" ptType="node"><dgm:layoutNode name="node"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"/><dgm:shape type="rect"><dgm:adjLst/></dgm:shape><dgm:presOf axis="desOrSelf" ptType="node"/><dgm:constrLst><dgm:constr type="lMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="rMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="tMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="bMarg" refType="primFontSz" fact="0.3"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst></dgm:layoutNode><dgm:forEach name="transitions" axis="followSib" ptType="sibTrans" cnt="1"><dgm:layoutNode name="sibTrans"><dgm:alg type="sp"/><dgm:shape><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:forEach></dgm:forEach>
      </dgm:layoutNode>
    </dgm:layoutDef>
    """;

    internal static string BuildColorsXml(DiagramTheme? theme = null)
    {
        theme ??= DiagramTheme.Default;
        return $"""
        <dgm:colorsDef uniqueId="urn:officecli:diagram:colors:v1" minVer="12.0" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dgm:title lang="" val="OfficeCLI"/><dgm:desc lang="" val="OfficeCLI editable diagram colors"/><dgm:catLst><dgm:cat type="mainScheme" pri="10300"/></dgm:catLst>
          <dgm:styleLbl name="node0"><dgm:fillClrLst><a:srgbClr val="{theme.Primary}"/></dgm:fillClrLst><dgm:linClrLst meth="repeat"><a:srgbClr val="{theme.Text}"/></dgm:linClrLst><dgm:effectClrLst><a:srgbClr val="{theme.Primary}"/></dgm:effectClrLst><dgm:txLinClrLst/><dgm:txFillClrLst><a:srgbClr val="{theme.Text}"/></dgm:txFillClrLst><dgm:txEffectClrLst/></dgm:styleLbl>
        </dgm:colorsDef>
        """;
    }

    internal static string BuildStyleXml(DiagramTheme? theme = null)
    {
        theme ??= DiagramTheme.Default;
        return $"""
        <dgm:styleDef uniqueId="urn:officecli:diagram:style:v1" minVer="12.0" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dgm:title lang="" val="OfficeCLI"/><dgm:desc lang="" val="OfficeCLI editable diagram style"/><dgm:catLst><dgm:cat type="simple" pri="100"/></dgm:catLst>
          <dgm:styleLbl name="node0"><dgm:txPr/><dgm:style><a:lnRef idx="0"><a:srgbClr val="{theme.Text}"/></a:lnRef><a:fillRef idx="1"><a:srgbClr val="{theme.Primary}"/></a:fillRef><a:effectRef idx="0"><a:srgbClr val="{theme.Primary}"/></a:effectRef><a:fontRef idx="minor"><a:srgbClr val="{theme.Text}"/></a:fontRef></dgm:style></dgm:styleLbl>
        </dgm:styleDef>
        """;
    }

    internal static SmartArtPackageInspection Inspect(string filePath)
    {
        if (!File.Exists(filePath)) throw new ArgumentException($"Office file not found: '{filePath}'.");
        using var archive = ZipFile.OpenRead(filePath);
        var diagrams = new List<SmartArtDiagramInfo>();
        foreach (var entry in archive.Entries.Where(IsDiagramDataPart).OrderBy(e => e.FullName, StringComparer.OrdinalIgnoreCase))
        {
            using var reader = new StreamReader(entry.Open());
            var xml = reader.ReadToEnd();
            diagrams.Add(InspectDataXml('/' + entry.FullName.Replace('\\', '/'), xml, FindHosts(archive, entry.FullName)));
        }
        return new SmartArtPackageInspection(filePath, diagrams.Count, diagrams);
    }

    internal static SmartArtUpdateReceipt Update(string filePath, string dataPartPath, DiagramSpec spec)
    {
        if (!File.Exists(filePath)) throw new ArgumentException($"Office file not found: '{filePath}'.");
        var normalized = dataPartPath.Replace('\\', '/').TrimStart('/');
        if (normalized.Contains("..", StringComparison.Ordinal)) throw new ArgumentException("SmartArt data part path cannot contain '..'.");

        using var archive = ZipFile.Open(filePath, ZipArchiveMode.Update);
        var entry = archive.GetEntry(normalized)
            ?? throw new ArgumentException($"SmartArt data part not found: '/{normalized}'. Run 'smartart inspect' to list valid data parts.");
        string before;
        using (var reader = new StreamReader(entry.Open())) before = reader.ReadToEnd();
        var profile = ReadProfile(before);
        var updatedXml = BuildDataXml(spec, profile);
        entry.Delete();
        var replacement = archive.CreateEntry(normalized, CompressionLevel.Optimal);
        using (var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(false))) writer.Write(updatedXml);
        return new SmartArtUpdateReceipt(true, filePath, '/' + normalized, spec.DiagramId, spec.Nodes.Count, spec.Edges.Count,
            Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(updatedXml))).ToLowerInvariant());
    }

    private static SmartArtDiagramInfo InspectDataXml(string partPath, string xml, IReadOnlyList<string> hosts)
    {
        var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var semantic = doc.Descendants(Ocli + "diagram").FirstOrDefault();
        var semanticNodes = semantic?.Elements(Ocli + "node").ToDictionary(e => (string?)e.Attribute("modelId") ?? "", StringComparer.Ordinal)
            ?? new Dictionary<string, XElement>(StringComparer.Ordinal);
        var nodes = new List<SmartArtNodeInfo>();
        var regularIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var point in doc.Descendants(Dgm + "pt"))
        {
            var type = (string?)point.Attribute("type") ?? "node";
            if (type is "doc" or "parTrans" or "sibTrans" or "pres") continue;
            var modelId = (string?)point.Attribute("modelId") ?? "";
            regularIds.Add(modelId);
            semanticNodes.TryGetValue(modelId, out var metadata);
            var placeholder = (string?)point.Element(Dgm + "prSet")?.Attribute("phldrT");
            var logicalId = (string?)metadata?.Attribute("id")
                ?? ParseLogicalId(placeholder)
                ?? $"node-{nodes.Count + 1}";
            var label = string.Join("\n", point.Descendants(A + "p")
                .Select(p => string.Concat(p.Descendants(A + "t").Select(t => t.Value)))
                .Where(s => s.Length > 0));
            nodes.Add(new SmartArtNodeInfo(logicalId, modelId, label,
                Split((string?)metadata?.Attribute("factRefs"))));
        }

        var edges = new List<SmartArtEdgeInfo>();
        if (semantic is not null)
        {
            foreach (var edge in semantic.Elements(Ocli + "edge"))
            {
                edges.Add(new SmartArtEdgeInfo(
                    (string?)edge.Attribute("id") ?? $"edge-{edges.Count + 1}",
                    (string?)edge.Attribute("from") ?? "",
                    (string?)edge.Attribute("to") ?? "",
                    (string?)edge.Attribute("label"),
                    string.Equals((string?)edge.Attribute("dashed"), "true", StringComparison.OrdinalIgnoreCase),
                    Split((string?)edge.Attribute("factRefs"))));
            }
        }
        else
        {
            var logicalByModel = nodes.ToDictionary(n => n.ModelId, n => n.NodeId, StringComparer.Ordinal);
            foreach (var cxn in doc.Descendants(Dgm + "cxn"))
            {
                var src = (string?)cxn.Attribute("srcId") ?? "";
                var dest = (string?)cxn.Attribute("destId") ?? "";
                if (!regularIds.Contains(src) || !regularIds.Contains(dest)) continue;
                edges.Add(new SmartArtEdgeInfo(
                    $"edge-{edges.Count + 1}", logicalByModel[src], logicalByModel[dest], null, false, new List<string>()));
            }
        }

        var docProps = doc.Descendants(Dgm + "pt").FirstOrDefault(p => (string?)p.Attribute("type") == "doc")?.Element(Dgm + "prSet");
        var diagramId = (string?)semantic?.Attribute("diagramId") ?? ParseDiagramId((string?)docProps?.Attribute("presName"));
        return new SmartArtDiagramInfo(partPath, hosts, diagramId,
            (string?)semantic?.Attribute("type") ?? "smartart", nodes, edges);
    }

    private static IReadOnlyDictionary<string, string> ReadProfile(string xml)
    {
        var doc = XDocument.Parse(xml);
        var props = doc.Descendants(Dgm + "pt").FirstOrDefault(p => (string?)p.Attribute("type") == "doc")?.Element(Dgm + "prSet");
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        if (props is null) return result;
        foreach (var name in new[] { "loTypeId", "loCatId", "qsTypeId", "qsCatId", "csTypeId", "csCatId" })
            if ((string?)props.Attribute(name) is { Length: > 0 } value) result[name] = value;
        return result;
    }

    private static string Profile(IReadOnlyDictionary<string, string>? profile, string key, string fallback)
        => profile is not null && profile.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value) ? value : fallback;

    private static bool IsDiagramDataPart(ZipArchiveEntry entry)
    {
        var p = entry.FullName.Replace('\\', '/');
        return p.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            && (p.Contains("/diagrams/data", StringComparison.OrdinalIgnoreCase)
                || p.Contains("/graphics/data", StringComparison.OrdinalIgnoreCase)
                || p.StartsWith("graphics/data", StringComparison.OrdinalIgnoreCase));
    }

    private static IReadOnlyList<string> FindHosts(ZipArchive archive, string dataPart)
    {
        var hosts = new List<string>();
        foreach (var relEntry in archive.Entries.Where(e => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
        {
            XDocument rels;
            try { using var r = new StreamReader(relEntry.Open()); rels = XDocument.Parse(r.ReadToEnd()); }
            catch { continue; }
            var owner = OwnerPart(relEntry.FullName);
            foreach (var rel in rels.Root?.Elements() ?? Enumerable.Empty<XElement>())
            {
                if (!((string?)rel.Attribute("Type") ?? "").EndsWith("/diagramData", StringComparison.OrdinalIgnoreCase)) continue;
                var target = (string?)rel.Attribute("Target");
                if (string.IsNullOrWhiteSpace(target)) continue;
                if (ResolveTarget(owner, target).Equals(dataPart.Replace('\\', '/'), StringComparison.OrdinalIgnoreCase))
                    hosts.Add('/' + owner);
            }
        }
        return hosts.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
    }

    private static string OwnerPart(string relPath)
    {
        var p = relPath.Replace('\\', '/');
        if (p.Equals("_rels/.rels", StringComparison.OrdinalIgnoreCase)) return "";
        var marker = "/_rels/";
        var i = p.LastIndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (i < 0) return "";
        var prefix = p[..i];
        var name = p[(i + marker.Length)..];
        return (prefix.Length == 0 ? "" : prefix + "/") + name[..^5];
    }

    private static string ResolveTarget(string owner, string target)
    {
        if (target.StartsWith('/')) return target.TrimStart('/');
        var segments = new List<string>();
        var ownerDir = owner.Contains('/') ? owner[..owner.LastIndexOf('/')] : "";
        foreach (var segment in (ownerDir + "/" + target).Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".") continue;
            if (segment == "..") { if (segments.Count > 0) segments.RemoveAt(segments.Count - 1); continue; }
            segments.Add(segment);
        }
        return string.Join('/', segments);
    }

    private static XElement EmptyText() => new(Dgm + "t", new XElement(A + "bodyPr"), new XElement(A + "lstStyle"),
        new XElement(A + "p", new XElement(A + "endParaRPr", new XAttribute("lang", "zh-CN"))));

    private static XElement TextBody(string text) => new(Dgm + "t", new XElement(A + "bodyPr"), new XElement(A + "lstStyle"),
        new XElement(A + "p", new XElement(A + "r", new XElement(A + "rPr", new XAttribute("lang", "zh-CN")), new XElement(A + "t", text))));

    private static string StableGuid(string diagramId, string token)
    {
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(diagramId + "\n" + token));
        Span<byte> guid = stackalloc byte[16];
        hash.AsSpan(0, 16).CopyTo(guid);
        guid[6] = (byte)((guid[6] & 0x0f) | 0x50);
        guid[8] = (byte)((guid[8] & 0x3f) | 0x80);
        return new Guid(guid).ToString("B").ToUpperInvariant();
    }

    private static string? ParseLogicalId(string? value)
        => value is { Length: > 14 } && value.StartsWith("[officecli:", StringComparison.Ordinal) && value.EndsWith(']') ? value[11..^1] : null;

    private static string? ParseDiagramId(string? value)
        => value is { Length: > 10 } && value.StartsWith("officecli:", StringComparison.Ordinal) ? value[10..] : null;

    private static List<string> Split(string? value) => string.IsNullOrWhiteSpace(value)
        ? new List<string>()
        : value.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
}

internal sealed record SmartArtPackageInspection(string File, int DiagramCount, List<SmartArtDiagramInfo> Diagrams);
internal sealed record SmartArtDiagramInfo(string DataPart, IReadOnlyList<string> Hosts, string? DiagramId, string Type,
    List<SmartArtNodeInfo> Nodes, List<SmartArtEdgeInfo> Edges);
internal sealed record SmartArtNodeInfo(string NodeId, string ModelId, string Label, List<string> FactRefs);
internal sealed record SmartArtEdgeInfo(string EdgeId, string From, string To, string? Label, bool Dashed, List<string> FactRefs);
internal sealed record SmartArtUpdateReceipt(bool Ok, string File, string DataPart, string DiagramId, int NodeCount, int EdgeCount, string DataSha256);
