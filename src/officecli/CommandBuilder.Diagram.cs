// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using OfficeCli.Core.Diagram;
using OfficeCli.Handlers;
using GroupShape = DocumentFormat.OpenXml.Presentation.GroupShape;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildDiagramCommand(Option<bool> jsonOption)
    {
        var specArg = new Argument<FileInfo>("spec") { Description = "DiagramSpec v1 JSON file" };
        var outOption = new Option<FileInfo>("--out") { Description = "Output SVG file", Required = true };
        var themeOption = new Option<FileInfo?>("--theme") { Description = "Optional Office theme v1 JSON" };
        var evidenceOption = new Option<FileInfo?>("--evidence-out") { Description = "Optional deterministic layout/fact-binding evidence JSON" };
        var command = new Command("diagram", "Validate and render a DiagramSpec without exposing OOXML")
        {
            specArg, outOption, themeOption, evidenceOption, jsonOption,
        };
        command.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var specPath = result.GetValue(specArg)!.FullName;
            var output = result.GetValue(outOption)!.FullName;
            if (!Path.GetExtension(output).Equals(".svg", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("diagram --out must use the .svg extension; use OfficePDF to produce PDF from the SVG.");
            var spec = DiagramSpec.Load(specPath);
            var graph = DiagramCompiler.Compile(spec);
            var theme = DiagramTheme.Load(result.GetValue(themeOption)?.FullName);
            var svg = DiagramSvgRenderer.Render(spec, graph, theme);
            Directory.CreateDirectory(Path.GetDirectoryName(output) ?? ".");
            File.WriteAllText(output, svg, new System.Text.UTF8Encoding(false));
            var evidence = DiagramSvgRenderer.Evidence(spec, graph, theme);
            var evidencePath = result.GetValue(evidenceOption)?.FullName;
            if (!string.IsNullOrWhiteSpace(evidencePath))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(evidencePath) ?? ".");
                File.WriteAllText(evidencePath, JsonSerializer.Serialize(evidence, DiagramJsonContext.Default.DiagramEvidence), new System.Text.UTF8Encoding(false));
            }
            var receipt = new DiagramCommandReceipt(true, 1, spec.DiagramId, spec.Type,
                output, evidencePath, graph.Nodes.Count, spec.Edges.Count);
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelope(JsonSerializer.Serialize(receipt, DiagramJsonContext.Default.DiagramCommandReceipt)));
            else Console.WriteLine($"Diagram SVG written: {output}");
            return 0;
        }, json); });
        return command;
    }

    private static Command BuildDiagramRefreshCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "DOCX or PPTX containing an OfficeCLI native diagram" };
        var specOption = new Option<FileInfo>("--spec") { Description = "Updated DiagramSpec v1 JSON", Required = true };
        var themeOption = new Option<FileInfo?>("--theme") { Description = "Optional Office theme v1 JSON" };
        var command = new Command("diagram-refresh", "Replace one generated Shape+Connector diagram by stable diagramId without rebuilding the document")
        {
            fileArg, specOption, themeOption, jsonOption,
        };
        command.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var receipt = RefreshNativeDiagram(result.GetValue(fileArg)!.FullName,
                DiagramSpec.Load(result.GetValue(specOption)!.FullName),
                result.GetValue(themeOption)?.FullName);
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelope(JsonSerializer.Serialize(receipt, DiagramJsonContext.Default.DiagramRefreshReceipt)));
            else Console.WriteLine($"Diagram refreshed: {receipt.DiagramId} in {receipt.Host}");
            return 0;
        }, json); });
        return command;
    }

    internal static DiagramRefreshReceipt RefreshNativeDiagram(string file, DiagramSpec spec, string? themePath)
    {
        var fullPath = Path.GetFullPath(file);
        if (!File.Exists(fullPath)) throw new ArgumentException($"Office file not found: '{fullPath}'.");
        var extension = Path.GetExtension(fullPath).ToLowerInvariant();
        if (extension is not (".docx" or ".pptx"))
            throw new ArgumentException("diagram-refresh supports DOCX and PPTX native Shape+Connector diagrams.");
        var temp = Path.Combine(Path.GetDirectoryName(fullPath)!, $".{Path.GetFileName(fullPath)}.{Guid.NewGuid():N}.diagram-refresh.tmp{extension}");
        File.Copy(fullPath, temp, overwrite: false);
        try
        {
            var host = extension == ".pptx"
                ? RefreshPowerPointDiagram(temp, spec, themePath)
                : RefreshWordDiagram(temp, spec, themePath);
            ValidateOfficePackage(temp, extension);
            var digest = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(temp))).ToLowerInvariant();
            File.Move(temp, fullPath, overwrite: true);
            return new DiagramRefreshReceipt(true, 1, spec.DiagramId, fullPath, host,
                spec.Nodes.Count, spec.Edges.Count, digest);
        }
        finally
        {
            try { if (File.Exists(temp)) File.Delete(temp); } catch { /* best effort */ }
        }
    }

    private static string RefreshPowerPointDiagram(string path, DiagramSpec spec, string? themePath)
    {
        int slideIndex = 0; uint oldGroupId = 0; long x = 0, y = 0, cx = 0, cy = 0;
        using (var package = PresentationDocument.Open(path, false))
        {
            var slides = package.PresentationPart!.SlideParts.ToList();
            for (var index = 0; index < slides.Count; index++)
            {
                var slide = slides[index].Slide ?? throw new InvalidOperationException("PowerPoint slide is missing.");
                var group = slide.CommonSlideData?.ShapeTree?.Elements<GroupShape>()
                    .SingleOrDefault(item => HasDiagramId(
                        item.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Description?.Value,
                        spec.DiagramId));
                if (group is null) continue;
                var nv = group.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!;
                var transform = group.GroupShapeProperties!.TransformGroup!;
                slideIndex = index + 1; oldGroupId = nv.Id!.Value;
                x = transform.Offset!.X!.Value; y = transform.Offset.Y!.Value;
                cx = transform.Extents!.Cx!.Value; cy = transform.Extents.Cy!.Value;
                break;
            }
        }
        if (slideIndex == 0) throw new ArgumentException($"native diagramId '{spec.DiagramId}' was not found in the PPTX.");
        var specPath = WriteRefreshSpec(spec);
        var properties = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["spec"] = specPath, ["render"] = "shapes",
            ["x"] = EmuCm(x), ["y"] = EmuCm(y), ["width"] = EmuCm(cx), ["height"] = EmuCm(cy),
        };
        if (!string.IsNullOrWhiteSpace(themePath)) properties["themeFile"] = themePath;
        try
        {
            using var handler = new PowerPointHandler(path, editable: true);
            handler.Add($"/slide[{slideIndex}]", "diagram", null, properties);
        }
        finally { DeleteRefreshSpec(specPath); }
        using (var package = PresentationDocument.Open(path, true))
        {
            var slide = package.PresentationPart!.SlideParts.ElementAt(slideIndex - 1).Slide
                ?? throw new InvalidOperationException("PowerPoint slide is missing.");
            var shapeTree = slide.CommonSlideData?.ShapeTree
                ?? throw new InvalidOperationException("PowerPoint shape tree is missing.");
            var oldGroup = shapeTree.Elements<GroupShape>().Single(item =>
                item.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value == oldGroupId);
            var replacement = shapeTree.Elements<GroupShape>().Last(item =>
                item.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value != oldGroupId &&
                HasDiagramId(item.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Description?.Value, spec.DiagramId));
            replacement.Remove(); shapeTree.InsertBefore(replacement, oldGroup); oldGroup.Remove();
            slide.Save();
        }
        return $"/slide[{slideIndex}]";
    }

    private static string RefreshWordDiagram(string path, DiagramSpec spec, string? themePath)
    {
        int oldParagraphIndex; long width; long height;
        using (var package = WordprocessingDocument.Open(path, false))
        {
            var document = package.MainDocumentPart?.Document
                ?? throw new InvalidOperationException("Word document part is missing.");
            var paragraphs = (document.Body ?? throw new InvalidOperationException("Word document body is missing."))
                .Elements<Paragraph>().ToList();
            oldParagraphIndex = paragraphs.FindIndex(paragraph => paragraph.OuterXml.Contains("DiagramId:" + spec.DiagramId, StringComparison.Ordinal));
            if (oldParagraphIndex < 0) throw new ArgumentException($"native diagramId '{spec.DiagramId}' was not found in the DOCX.");
            var extent = paragraphs[oldParagraphIndex].Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>().Single();
            width = extent.Cx!.Value; height = extent.Cy!.Value;
        }
        var specPath = WriteRefreshSpec(spec);
        var properties = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["spec"] = specPath, ["render"] = "shapes", ["width"] = EmuCm(width), ["height"] = EmuCm(height),
        };
        if (!string.IsNullOrWhiteSpace(themePath)) properties["themeFile"] = themePath;
        try
        {
            using var handler = new WordHandler(path, editable: true);
            handler.Add("/body", "diagram", null, properties);
        }
        finally { DeleteRefreshSpec(specPath); }
        using (var package = WordprocessingDocument.Open(path, true))
        {
            var document = package.MainDocumentPart?.Document
                ?? throw new InvalidOperationException("Word document part is missing.");
            var body = document.Body ?? throw new InvalidOperationException("Word document body is missing.");
            var matches = body.Elements<Paragraph>().Where(paragraph =>
                paragraph.OuterXml.Contains("DiagramId:" + spec.DiagramId, StringComparison.Ordinal)).ToList();
            if (matches.Count != 2) throw new InvalidOperationException("diagram refresh did not produce exactly one replacement Word group.");
            var oldParagraph = body.Elements<Paragraph>().ElementAt(oldParagraphIndex);
            var replacement = matches.Single(item => !ReferenceEquals(item, oldParagraph));
            replacement.Remove(); body.InsertBefore(replacement, oldParagraph); oldParagraph.Remove();
            document.Save();
        }
        return "/body";
    }

    private static string WriteRefreshSpec(DiagramSpec spec)
    {
        var path = Path.Combine(Path.GetTempPath(), $"officecli-diagram-{Guid.NewGuid():N}.json");
        File.WriteAllText(path, JsonSerializer.Serialize(spec, DiagramJsonContext.Default.DiagramSpec), new System.Text.UTF8Encoding(false));
        return path;
    }

    private static bool HasDiagramId(string? metadata, string diagramId) => metadata?.Split(';')
        .Contains("DiagramId:" + diagramId, StringComparer.Ordinal) == true;
    private static void DeleteRefreshSpec(string path) { try { File.Delete(path); } catch { /* best effort */ } }
    private static string EmuCm(long value) => (value / 360000d).ToString("0.########", CultureInfo.InvariantCulture) + "cm";

    private static void ValidateOfficePackage(string path, string extension)
    {
        ValidationErrorInfo[] errors;
        if (extension == ".pptx")
        {
            using var package = PresentationDocument.Open(path, false);
            errors = new OpenXmlValidator().Validate(package).ToArray();
        }
        else
        {
            using var package = WordprocessingDocument.Open(path, false);
            errors = new OpenXmlValidator().Validate(package).ToArray();
        }
        if (errors.Length > 0)
            throw new InvalidOperationException("refreshed Office diagram failed OOXML validation: " + errors[0].Description);
    }
}
