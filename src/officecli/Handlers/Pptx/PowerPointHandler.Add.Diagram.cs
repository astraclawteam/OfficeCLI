// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using OfficeCli.Core.Diagram;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // DiagramSpec has one semantic model and two editable renderers. The default
    // expands into native Shape+Connector elements; explicit render=smartart
    // writes a persistent native SmartArt data model that smartart inspect/update
    // can round-trip. Legacy Mermaid input remains an adapter to the shape or
    // image paths. Layout and semantics stay format-agnostic in Core/Diagram.
    private const double CmToEmu = 360000.0;

    private string AddDiagram(string parentPath, int? index, Dictionary<string, string> properties)
    {
        var specPath = properties.GetValueOrDefault("spec") ?? properties.GetValueOrDefault("diagramSpec");
        if (!string.IsNullOrWhiteSpace(specPath))
        {
            var spec = DiagramSpec.Load(specPath);
            var specRenderer = (properties.GetValueOrDefault("render") ?? "native").Trim().ToLowerInvariant();
            if (specRenderer is "smartart" or "native-smartart")
                return AddDiagramSmartArt(parentPath, properties, spec);
            var theme = DiagramTheme.Load(properties.GetValueOrDefault("themeFile"));
            return AddDiagramNative(parentPath, index, properties, DiagramCompiler.Compile(spec), theme);
        }

        // Input mirrors `equation` (canonical domain word `formula` + alias `text`):
        //   mermaid / text / dsl   → inline flowchart text
        //   src / path             → load the text from a .mmd file (consistent with
        //                            picture/media `src`, which is also a file path)
        var mermaidText = properties.GetValueOrDefault("mermaid")
                          ?? properties.GetValueOrDefault("text")
                          ?? properties.GetValueOrDefault("dsl");
        if (string.IsNullOrWhiteSpace(mermaidText)
            && (properties.TryGetValue("src", out var srcFile) || properties.TryGetValue("path", out srcFile))
            && !string.IsNullOrWhiteSpace(srcFile))
        {
            if (!System.IO.File.Exists(srcFile))
                throw new ArgumentException($"diagram source file not found: '{srcFile}'.");
            mermaidText = System.IO.File.ReadAllText(srcFile);
        }
        if (string.IsNullOrWhiteSpace(mermaidText))
            throw new ArgumentException("diagram requires inline 'mermaid' text (aliases: text, dsl) or a 'src' .mmd file path.");

        // render mode: native (built-in editable shapes) | image (real mermaid.js in
        // a headless browser → embedded SVG, covers EVERY mermaid type at full
        // fidelity) | auto (default: image when a browser is available, else native).
        var renderMode = (properties.GetValueOrDefault("render") ?? "auto").Trim().ToLowerInvariant();
        bool forceImage = renderMode is "image" or "svg" or "browser";
        if (forceImage && !MermaidImageRenderer.IsAvailable())
            throw new ArgumentException(
                "render=image needs mermaid-cli (mmdc) or a headless browser (Chrome/Chromium/Edge). "
                + "Install one, or use render=native for the built-in synthesizer.");
        bool wantImage = forceImage
            || (renderMode is not ("native" or "shapes") && MermaidImageRenderer.IsAvailable());
        if (wantImage)
            return AddDiagramAsImage(parentPath, index, properties, mermaidText, allowNativeFallback: !forceImage);

        return AddDiagramNative(parentPath, index, properties, mermaidText);
    }

    private string AddDiagramSmartArt(string parentPath, Dictionary<string, string> properties, DiagramSpec spec)
    {
        var theme = DiagramTheme.Load(properties.GetValueOrDefault("themeFile"));
        var m = Regex.Match(parentPath, @"/slide\[(\d+)\]", RegexOptions.IgnoreCase);
        if (!m.Success)
            throw new ArgumentException($"SmartArt diagram parent must be a slide (e.g. /slide[1]); got '{parentPath}'.");
        int slideIdx = int.Parse(m.Groups[1].Value);
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"slide {slideIdx} not found (total: {slideParts.Count}).");

        AddPart(parentPath, "smartart", new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["dataXml"] = NativeSmartArtCodec.BuildDataXml(spec),
            ["layoutXml"] = NativeSmartArtCodec.BuildLayoutXml(),
            ["colorsXml"] = NativeSmartArtCodec.BuildColorsXml(theme),
            ["quickStyleXml"] = NativeSmartArtCodec.BuildStyleXml(theme),
        });

        var slide = GetSlide(slideParts[slideIdx - 1]);
        var shapeTree = slide.CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("slide shape tree is missing.");
        var frame = shapeTree.Elements<GraphicFrame>().Last();
        var nv = frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
        if (nv != null)
        {
            nv.Name = string.IsNullOrWhiteSpace(spec.Title) ? $"SmartArt {spec.DiagramId}" : spec.Title;
            nv.Description = "DiagramId:" + spec.DiagramId
                + ";DiagramRenderer:SmartArt;DiagramFactRefs:"
                + string.Join(',', spec.Facts.Select(f => f.FactId));
        }
        var transform = frame.Transform;
        if (transform != null)
        {
            if (properties.TryGetValue("x", out var x)) transform.Offset!.X = ParseEmu(x);
            if (properties.TryGetValue("y", out var y)) transform.Offset!.Y = ParseEmu(y);
            if (properties.TryGetValue("width", out var width)) transform.Extents!.Cx = ParseEmu(width);
            if (properties.TryGetValue("height", out var height)) transform.Extents!.Cy = ParseEmu(height);
        }
        slide.Save();
        var smartArtCount = shapeTree.Elements<GraphicFrame>().Count(g =>
            g.Descendants().Any(e => e.LocalName == "relIds"
                && e.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/diagram"));
        return $"/slide[{slideIdx}]/smartart[{smartArtCount}]";
    }

    // Built-in synthesizer: mermaid → laid-out graph → native editable shapes.
    private string AddDiagramNative(string parentPath, int? index, Dictionary<string, string> properties, string mermaidText)
    {
        return AddDiagramNative(parentPath, index, properties, DiagramCompiler.Compile(mermaidText), DiagramTheme.Default);
    }

    private string AddDiagramNative(string parentPath, int? index, Dictionary<string, string> properties,
                                    LaidOutGraph lo, DiagramTheme theme)
    {
        if (lo.Nodes.Count == 0)
            throw new ArgumentException("diagram parsed to zero nodes — check the mermaid syntax.");

        var m = Regex.Match(parentPath, @"/slide\[(\d+)\]", RegexOptions.IgnoreCase);
        if (!m.Success)
            throw new ArgumentException($"diagram parent must be a slide (e.g. /slide[1]); got '{parentPath}'.");
        int slideIdx = int.Parse(m.Groups[1].Value);
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count)
            throw new ArgumentException($"slide {slideIdx} not found (total: {slideParts.Count}).");
        var shapeTree = GetSlide(slideParts[PathIndex.ToArrayIndex(slideIdx)]).CommonSlideData!.ShapeTree!;

        // Placement: the slide size is the user's, not ours. By default we FIT the
        // diagram into a box on the UNCHANGED slide (a lone flowchart must not
        // silently resize someone's deck). `poster=true` is the explicit opt-in to
        // grow the slide to the whole diagram instead (export-a-diagram use case).
        // x/y/width/height define the box, mirroring picture/chart. Without an
        // explicit box, reserve any wide title band already present near the top
        // and use the remaining content area. Uniform scale keeps the aspect ratio
        // while allowing a standalone diagram to use the available canvas.
        double natW = lo.SlideWidthCm, natH = lo.SlideHeightCm;
        bool hasX = properties.TryGetValue("x", out var xs);
        bool hasY = properties.TryGetValue("y", out var ys);
        bool hasW = properties.TryGetValue("width", out var ws);
        bool hasH = properties.TryGetValue("height", out var hs);
        double sc, ox, oy;
        if (OfficeCli.Core.ParseHelpers.IsTruthy(properties.GetValueOrDefault("poster")))
        {
            SetSlideSizeCm(natW, natH);
            sc = 1; ox = 0; oy = 0;
        }
        else
        {
            var (slideWEmu, slideHEmu) = GetSlideSize();
            double slideW = slideWEmu / CmToEmu, slideH = slideHEmu / CmToEmu;
            const double margin = 0.6;
            double boxX = hasX ? ParseEmu(xs!) / CmToEmu : margin;
            double boxY = hasY ? ParseEmu(ys!) / CmToEmu : margin;
            if (!hasY && !hasH)
            {
                double occupiedTop = margin;
                foreach (var existingShape in shapeTree.Elements<Shape>())
                {
                    var transform = existingShape.ShapeProperties?.Transform2D;
                    if (transform?.Offset?.Y?.Value is not long titleYEmu
                        || transform.Extents?.Cx?.Value is not long titleWidthEmu
                        || transform.Extents.Cy?.Value is not long titleHeightEmu)
                        continue;
                    double titleY = titleYEmu / CmToEmu;
                    double titleWidth = titleWidthEmu / CmToEmu;
                    double titleHeight = titleHeightEmu / CmToEmu;
                    if (titleY <= slideH * 0.3 && titleWidth >= slideW * 0.35 && titleHeight > 0)
                        occupiedTop = Math.Max(occupiedTop, titleY + titleHeight + 0.35);
                }
                boxY = Math.Min(occupiedTop, Math.Max(margin, slideH - margin - 1.0));
            }
            double boxW = hasW ? ParseEmu(ws!) / CmToEmu : Math.Max(0.1, slideW - boxX - margin);
            double boxH = hasH ? ParseEmu(hs!) / CmToEmu : Math.Max(0.1, slideH - boxY - margin);
            double fit = Math.Min(boxW / natW, boxH / natH);
            sc = fit;
            // Uniform scale leaves slack on one axis; CENTRE the fitted diagram in
            // its box (slack split evenly) rather than pinning it to the top-left.
            // Mirrors the image path (AddDiagramAsImage) so native and PNG place
            // identically, and honours the "centred" contract for the default
            // (no-position) box: boxX=margin, boxW=slideW-2margin →
            // margin+(boxW-natW*sc)/2 == (slideW-natW*sc)/2, unchanged.
            ox = boxX + (boxW - natW * sc) / 2;
            oy = boxY + (boxH - natH * sc) / 2;
        }

        uint nextId = AcquireShapeId(shapeTree, new Dictionary<string, string>());
        long Emu(double cm) => (long)Math.Round(cm * CmToEmu);
        double TX(double cm) => cm * sc + ox;   // natural cm → placed cm (x-axis)
        double TY(double cm) => cm * sc + oy;    // natural cm → placed cm (y-axis)
        // Use a conservative, explicit presentation font size.  Microsoft Office
        // and WPS do not apply DrawingML normAutofit identically, especially after
        // CJK font substitution; 18pt can therefore cross a shape boundary even
        // when the deterministic SVG fits.  Fourteen points remains readable on a
        // slide and preserves enough host-independent wrap margin.
        int fontPt = Math.Max(1, (int)Math.Floor(14 * lo.FontScale * sc));

        // Wrap the whole diagram in ONE group so it stays adjustable as a unit
        // AFTER Add: a human drags a single object; an agent addresses one stable
        // path `/slide[N]/group[K]` and `set width/height` scales every child via
        // the chOff/chExt baseline (see Set.Shape.cs group-scale-baseline). Child
        // coordinates ARE slide EMU here (chOff==off, chExt==ext → identity map),
        // so each child keeps the absolute placement it already computed — the
        // group is a transparent wrapper until someone resizes it.
        long gx = Emu(TX(0)), gy = Emu(TY(0));
        long gcx = Emu(natW * sc), gcy = Emu(natH * sc);
        uint groupId = nextId++;
        var group = new GroupShape(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = groupId, Name = $"Diagram {groupId}",
                    Description = DiagramMetadata(lo.DiagramId, null) },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties(
                new Drawing.TransformGroup(
                    new Drawing.Offset { X = gx, Y = gy },
                    new Drawing.Extents { Cx = gcx, Cy = gcy },
                    new Drawing.ChildOffset { X = gx, Y = gy },
                    new Drawing.ChildExtents { Cx = gcx, Cy = gcy })));

        // nodes (appended first → behind connectors/labels in z-order)
        var nodeShapeIds = new Dictionary<string, uint>(StringComparer.Ordinal);
        foreach (var n in lo.Nodes)
        {
            var (geom, fill, line) = DiagramStyles.Resolve(n.Shape, theme);
            var textColor = DiagramStyles.TextColorFor(fill, theme);
            var shapeId = nextId++;
            nodeShapeIds[n.Id] = shapeId;
            group.AppendChild(BuildDiagramShape(shapeId, geom, fill, line,
                DiagramTextMetrics.WrappedText(n.Label, Math.Max(0.8, n.W - 0.6)), fontPt,
                Emu(TX(n.X)), Emu(TY(n.Y)), Emu(n.W * sc), Emu(n.H * sc), n.FactRefs, lo.DiagramId,
                theme.MinorLatinFont, theme.MinorEastAsiaFont, textColor));
        }

        // One native connector per semantic edge.  Start/end connection references
        // make the line follow a node when the user moves it in Office or WPS.
        // SVG/PDF retain the exact deterministic orthogonal route; the editable
        // Office view delegates elbow rerouting to the host application.
        foreach (var e in lo.Edges)
        {
            if (e.Points.Count < 2) continue;
            // Some Office-compatible renderers shorten an attached connector by
            // one or two pixels at the shape boundary (most visible when two
            // metrics enter the top point of a decision diamond).  Keep the real
            // start/end connection references, but extend the deterministic
            // geometry a small distance into each attached node so PPT/PDF/SVG
            // previews all read as physically connected.
            var (x1, y1, x2, y2) = NativeConnectorAnchorPoints(e);
            uint? sourceShapeId = e.SourceNodeId is not null && nodeShapeIds.TryGetValue(e.SourceNodeId, out var source) ? source : null;
            uint? targetShapeId = e.TargetNodeId is not null && nodeShapeIds.TryGetValue(e.TargetNodeId, out var target) ? target : null;
            group.AppendChild(BuildDiagramConnector(nextId++, TX(x1), TY(y1), TX(x2), TY(y2),
                theme.MutedText, e.ArrowAtEnd, e.Dashed, e.FactRefs, lo.DiagramId,
                sourceShapeId, targetShapeId, e.StartConnectionIndex, e.EndConnectionIndex,
                e.Points.Count > 2));
        }

        // edge labels (appended last → white masks sit on top of the lines)
        foreach (var lbl in lo.Labels)
        {
            double w = lbl.W;
            // Opaque (flowchart) labels mask the edge line they sit on; sequence
            // labels sit in empty space above the arrow → no fill, so they don't
            // punch a white hole in whatever lifeline they overlap.
            group.AppendChild(BuildDiagramShape(nextId++, "rect", lbl.Opaque ? "FFFFFF" : null, null,
                DiagramTextMetrics.WrappedText(lbl.Text, Math.Max(0.6, lbl.W - 0.3)),
                Math.Max(1, (int)Math.Round(10 * sc)),
                Emu(TX(lbl.Cx - w / 2)), Emu(TY(lbl.Cy - lbl.H / 2)), Emu(w * sc), Emu(lbl.H * sc), null,
                lo.DiagramId, theme.MinorLatinFont, theme.MinorEastAsiaFont, theme.Text));
        }

        shapeTree.AppendChild(group);
        return $"/slide[{slideIdx}]/group[{shapeTree.Elements<GroupShape>().Count()}]";
    }

    // High-fidelity path: render the mermaid with the real mermaid.js (headless
    // browser) to SVG and embed it as a picture, stamping the source into alt-text
    // so the diagram travels in the file and is regenerable. In auto mode any
    // render failure falls back to the native synthesizer.
    private string AddDiagramAsImage(string parentPath, int? index, Dictionary<string, string> properties,
                                     string mermaidText, bool allowNativeFallback)
    {
        // Bake theme/layout/look into the source as frontmatter so they render and
        // round-trip via alt-text (the composed source is what gets stamped). The
        // image backend renders elk/handDrawn/themes at full fidelity. The native
        // fallback keeps the ORIGINAL source — its parser has no frontmatter/elk
        // support and would emit garbage nodes from the `---` lines.
        var composedText = MermaidImageRenderer.ComposeSource(mermaidText,
            properties.GetValueOrDefault("theme"),
            properties.GetValueOrDefault("layout"),
            properties.GetValueOrDefault("look"));
        var background = properties.GetValueOrDefault("background");

        string imgPath;
        try { imgPath = MermaidImageRenderer.RenderToPngFile(composedText, background); }
        // A syntax error is bad input — surface it (with mermaid's line-numbered
        // message) so the caller can fix the source. Never fall back to native: the
        // synthesizer would reject the same broken text or, worse, draw garbage.
        catch (MermaidSyntaxException) { throw; }
        catch when (allowNativeFallback) { return AddDiagramNative(parentPath, index, properties, mermaidText); }
        try
        {
            var pic = new Dictionary<string, string>(properties);
            foreach (var k in new[] { "mermaid", "text", "dsl", "src", "path", "render", "poster",
                                      "theme", "layout", "look", "background" })
                pic.Remove(k);
            pic["src"] = imgPath;
            // Stamp the COMPOSED source (frontmatter included) so theme/layout/look
            // travel in the document and a regenerate reproduces the same styling.
            if (!(pic.TryGetValue("alt", out var a) && !string.IsNullOrEmpty(a)))
                pic["alt"] = MermaidImageRenderer.SourceTag + composedText;

            // Sizing parity with the native path: the diagram is ALWAYS scaled to FIT
            // its box with aspect preserved (a mermaid diagram is never stretched).
            // width/height define the box (else the slide content area); passing them
            // straight to AddPicture would stretch — e.g. a tall flowchart forced into
            // a wide 30x14cm box comes out squashed. Fit-into-box, then centre in the
            // box (explicit x/y = box origin) or in the slide when position is implicit.
            {
                using var s = System.IO.File.OpenRead(imgPath);
                var dims = OfficeCli.Core.ImageSource.TryGetDimensions(s);

                // poster resolution. Explicit poster=true always grows the slide;
                // poster=false always fits one slide. When poster is UNSET, the
                // ADAPTIVE DEFAULT grows the slide only when fitting the diagram to
                // the slide would shrink it below the readability floor (a long
                // flowchart otherwise becomes a 1cm sliver) — a normal diagram still
                // fits the slide unchanged. Auto-poster stands down when the caller
                // pinned an explicit box (x/y/width/height): that is an explicit
                // placement request, honor it.
                bool posterSet = properties.ContainsKey("poster");
                bool posterOn = OfficeCli.Core.ParseHelpers.IsTruthy(properties.GetValueOrDefault("poster"));
                bool hasExplicitBox = pic.ContainsKey("width") || pic.ContainsKey("height")
                                      || pic.ContainsKey("x") || pic.ContainsKey("y");
                bool grow = posterOn;
                if (!posterSet && !hasExplicitBox && dims is { Width: > 0, Height: > 0 } ad)
                {
                    var (sw0, sh0) = GetSlideSize();
                    double m0 = 0.6 * CmToEmu;
                    grow = OfficeCli.Core.Diagram.MermaidImageRenderer.ExceedsOnePageReadably(
                        ad.Width, ad.Height, (sw0 - 2 * m0) / CmToEmu, (sh0 - 2 * m0) / CmToEmu);
                }

                // Grow the SLIDE to the whole diagram. The raster px are read as
                // 96-DPI CSS pixels; both axes are clamped, aspect preserved, to
                // PowerPoint's maximum slide edge (56in / 142.24cm) so an extremely
                // long chart yields a valid — if tall — single slide rather than a
                // file PowerPoint refuses to open.
                if (grow && dims is { Width: > 0, Height: > 0 } pd)
                {
                    double wCm = pd.Width / 96.0 * 2.54, hCm = pd.Height / 96.0 * 2.54;
                    double clamp = Math.Min(1.0, MaxSlideEdgeCm / Math.Max(wCm, hCm));
                    wCm *= clamp; hCm *= clamp;
                    SetSlideSizeCm(wCm, hCm);
                    long cxp = (long)(wCm * CmToEmu), cyp = (long)(hCm * CmToEmu);
                    pic["x"] = "0"; pic["y"] = "0";
                    pic["width"] = cxp.ToString();
                    pic["height"] = cyp.ToString();
                    return AddPicture(parentPath, index, pic);
                }
                if (dims is { Width: > 0, Height: > 0 } d)
                {
                    var (sw, sh) = GetSlideSize();
                    double margin = 0.6 * CmToEmu;
                    bool hasX = pic.TryGetValue("x", out var xs);
                    bool hasY = pic.TryGetValue("y", out var ys);
                    double boxX = hasX ? ParseEmu(xs!) : margin;
                    double boxY = hasY ? ParseEmu(ys!) : margin;
                    double boxW = pic.TryGetValue("width", out var ws) ? ParseEmu(ws) : sw - 2 * margin;
                    double boxH = pic.TryGetValue("height", out var hs) ? ParseEmu(hs) : sh - 2 * margin;
                    double fit = Math.Min(boxW / d.Width, boxH / d.Height);
                    long cx = (long)(d.Width * fit), cy = (long)(d.Height * fit);
                    pic["width"] = cx.ToString();
                    pic["height"] = cy.ToString();
                    // Centre the fitted image inside its box (letterbox slack split
                    // evenly); with no explicit position that box is the whole slide.
                    pic["x"] = ((long)(boxX + (boxW - cx) / 2)).ToString();
                    pic["y"] = ((long)(boxY + (boxH - cy) / 2)).ToString();
                }
            }
            return AddPicture(parentPath, index, pic);
        }
        finally { try { System.IO.File.Delete(imgPath); } catch { /* best effort */ } }
    }

    private Shape BuildDiagramShape(uint id, string geometry, string? fill, string? line, string text,
                                    int fontPt, long x, long y, long cx, long cy, IReadOnlyList<string>? factRefs,
                                    string? diagramId = null, string? latinFont = null, string? eastAsiaFont = null,
                                    string? textColor = null)
    {
        var shape = new Shape
        {
            NonVisualShapeProperties = new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"DiagramShape {id}",
                    Description = DiagramMetadata(diagramId, factRefs) },
                new NonVisualShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            ShapeProperties = new ShapeProperties(),
        };
        var sp = shape.ShapeProperties!;
        sp.Transform2D = new Drawing.Transform2D(
            new Drawing.Offset { X = x, Y = y },
            new Drawing.Extents { Cx = cx, Cy = cy });
        var preset = TryParsePresetShape(geometry, out var geomEnum) ? geomEnum : Drawing.ShapeTypeValues.Rectangle;
        sp.AppendChild(new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = preset });
        if (!string.IsNullOrEmpty(fill))
            sp.AppendChild(BuildSolidFill(fill));
        if (!string.IsNullOrEmpty(line))
            sp.AppendChild(new Drawing.Outline(BuildSolidFill(line)) { Width = 9525 }); // ~0.75pt

        shape.TextBody = new TextBody(
            // Zero insets: default text insets (~0.25cm L/R) are fixed and would
            // eat a fit-shrunk box's width, wrapping/clipping the label. Padding is
            // already in the box geometry. normAutofit shrinks any residual overflow.
            new Drawing.BodyProperties(new Drawing.NormalAutoFit())
            {
                Anchor = Drawing.TextAnchoringTypeValues.Center, Wrap = Drawing.TextWrappingValues.Square,
                LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0,
            },
            new Drawing.ListStyle(),
            new Drawing.Paragraph(
                new Drawing.ParagraphProperties { Alignment = Drawing.TextAlignmentTypeValues.Center },
                new Drawing.Run(
                    new Drawing.RunProperties(
                        new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = textColor ?? "000000" }),
                        new Drawing.LatinFont { Typeface = latinFont ?? "Arial" },
                        new Drawing.EastAsianFont { Typeface = eastAsiaFont ?? "Microsoft YaHei" })
                    { FontSize = fontPt * 100, Language = "zh-CN" },
                    new Drawing.Text(text))));
        return shape;
    }

    internal static (double X1, double Y1, double X2, double Y2) NativeConnectorAnchorPoints(
        RoutedEdge edge, double overlapCm = 0.08)
    {
        if (edge.Points.Count < 2)
            throw new ArgumentException("A connector requires at least two points.", nameof(edge));

        var start = edge.Points[0];
        var next = edge.Points[1];
        var previous = edge.Points[^2];
        var end = edge.Points[^1];

        static Pt Move(Pt point, Pt toward, double amount)
        {
            var dx = toward.X - point.X;
            var dy = toward.Y - point.Y;
            var length = Math.Sqrt(dx * dx + dy * dy);
            return length <= 1e-9
                ? point
                : new Pt(point.X + dx / length * amount, point.Y + dy / length * amount);
        }

        // Source interior is opposite the outgoing segment; target interior is
        // beyond the incoming segment. Detached ends remain exact.
        if (!string.IsNullOrWhiteSpace(edge.SourceNodeId))
            start = Move(start, new Pt(start.X - (next.X - start.X), start.Y - (next.Y - start.Y)), overlapCm);
        if (!string.IsNullOrWhiteSpace(edge.TargetNodeId))
            end = Move(end, new Pt(end.X + (end.X - previous.X), end.Y + (end.Y - previous.Y)), overlapCm);

        return (start.X, start.Y, end.X, end.Y);
    }

    private ConnectionShape BuildDiagramConnector(uint id, double x1, double y1, double x2, double y2,
                                                  string color, bool arrowAtEnd, bool dashed = false,
                                                  IReadOnlyList<string>? factRefs = null, string? diagramId = null,
                                                  uint? sourceShapeId = null, uint? targetShapeId = null,
                                                  uint startConnectionIndex = 0, uint endConnectionIndex = 0,
                                                  bool bent = false)
    {
        long ox = (long)Math.Round(Math.Min(x1, x2) * CmToEmu);
        long oy = (long)Math.Round(Math.Min(y1, y2) * CmToEmu);
        long cx = (long)Math.Round(Math.Abs(x2 - x1) * CmToEmu);
        long cy = (long)Math.Round(Math.Abs(y2 - y1) * CmToEmu);

        // A StraightConnector1 with no flip is drawn from the top-left corner
        // (off) to the bottom-right (off+ext). Flip so the connector's START is
        // (x1,y1) and its END is (x2,y2) for ALL four diagonal directions —
        // otherwise a right-and-up (or left-and-down) segment draws the wrong
        // diagonal AND puts the arrowhead on the wrong end. With the flips set,
        // the arrow is ALWAYS TailEnd (the (x2,y2)=target end).
        var xfrm = new Drawing.Transform2D(
            new Drawing.Offset { X = ox, Y = oy },
            new Drawing.Extents { Cx = cx, Cy = cy });
        if (x2 < x1) xfrm.HorizontalFlip = true;
        if (y2 < y1) xfrm.VerticalFlip = true;
        var connectionProperties = new NonVisualConnectorShapeDrawingProperties();
        if (sourceShapeId is not null)
            connectionProperties.StartConnection = new Drawing.StartConnection { Id = sourceShapeId.Value, Index = startConnectionIndex };
        if (targetShapeId is not null)
            connectionProperties.EndConnection = new Drawing.EndConnection { Id = targetShapeId.Value, Index = endConnectionIndex };
        var connector = new ConnectionShape
        {
            NonVisualConnectionShapeProperties = new NonVisualConnectionShapeProperties(
                new NonVisualDrawingProperties { Id = id, Name = $"DiagramEdge {id}",
                    Description = DiagramMetadata(diagramId, factRefs) },
                connectionProperties,
                new ApplicationNonVisualDrawingProperties()),
            ShapeProperties = new ShapeProperties(
                xfrm,
                new Drawing.PresetGeometry(new Drawing.AdjustValueList())
                    { Preset = bent ? Drawing.ShapeTypeValues.BentConnector3 : Drawing.ShapeTypeValues.StraightConnector1 }),
        };
        var outline = new Drawing.Outline(BuildSolidFill(color)) { Width = 12700 }; // 1pt
        if (dashed) // schema order: fill → prstDash → line-ends
            outline.AppendChild(new Drawing.PresetDash { Val = Drawing.PresetLineDashValues.Dash });
        if (arrowAtEnd)
            outline.AppendChild(new Drawing.TailEnd { Type = Drawing.LineEndValues.Triangle });
        connector.ShapeProperties!.AppendChild(outline);
        return connector;
    }

    private static string? DiagramMetadata(string? diagramId, IReadOnlyList<string>? factRefs)
    {
        var fields = new List<string>();
        if (!string.IsNullOrWhiteSpace(diagramId)) fields.Add("DiagramId:" + diagramId);
        if (factRefs is { Count: > 0 }) fields.Add("DiagramFactRefs:" + string.Join(",", factRefs));
        return fields.Count == 0 ? null : string.Join(";", fields);
    }

    private static double DiagramLabelWidthCm(string text)
    {
        double w = 0;
        foreach (var c in text) w += c > 0x2E80 ? 0.58 : 0.30;
        return Math.Min(w, 5.0) + 0.4;
    }

    // PowerPoint refuses to open a deck whose slide edge exceeds 56 inches
    // (=142.24cm =51206400 EMU). poster sizing clamps to this.
    private const double MaxSlideEdgeCm = 142.24;

    private void SetSlideSizeCm(double wCm, double hCm)
    {
        var pres = _doc?.PresentationPart?.Presentation;
        if (pres == null) return;
        // Clamp each edge to PowerPoint's maximum so an oversized poster (a very
        // long flowchart grown to its natural size) still yields an openable file.
        wCm = Math.Min(wCm, MaxSlideEdgeCm);
        hCm = Math.Min(hCm, MaxSlideEdgeCm);
        pres.SlideSize ??= new SlideSize();
        pres.SlideSize.Cx = (int)Math.Round(wCm * CmToEmu);
        pres.SlideSize.Cy = (int)Math.Round(hCm * CmToEmu);
    }
}
