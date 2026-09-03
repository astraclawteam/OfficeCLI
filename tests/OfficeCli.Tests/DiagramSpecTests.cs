using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeCli.Handlers;
using OfficeCli.Core.Diagram;
using Xunit;

namespace OfficeCli.Tests;

public sealed class DiagramSpecTests
{
    [Fact]
    public void DiagramNodeTextUsesAccessibleContrastAndPortableLatinFont()
    {
        var theme = new DiagramTheme { Text = "000000", Accent = "7C3AED" };
        Assert.Equal("FFFFFF", DiagramStyles.TextColorFor(theme.Accent, theme));
        Assert.Equal("Arial", DiagramTheme.Default.MinorLatinFont);
    }

    [Fact]
    public void NativeConnectorGeometryOverlapsAttachedNodeBoundaries()
    {
        var edge = new RoutedEdge
        {
            SourceNodeId = "source",
            TargetNodeId = "target",
            Points = new List<Pt> { new(2, 1), new(2, 3), new(2, 5) },
        };

        var anchors = PowerPointHandler.NativeConnectorAnchorPoints(edge);

        Assert.Equal(2, anchors.X1, 6);
        Assert.Equal(0.92, anchors.Y1, 6);
        Assert.Equal(2, anchors.X2, 6);
        Assert.Equal(5.08, anchors.Y2, 6);
    }

    private static string CompactDiagramText(string text) =>
        string.Concat(text.Where(character => !char.IsWhiteSpace(character)));

    private const string ArchitectureSpec = """
    {
      "schemaVersion": 1,
      "diagramId": "delivery-architecture",
      "type": "architecture",
      "title": "智能交付架构",
      "direction": "top-down",
      "facts": [
        {"factId":"fact-local","sourceId":"office-project","locator":"privacy.rawFilesUploaded","summary":"原文件不上传","confidence":1}
      ],
      "nodes": [
        {"id":"materials","label":"Word / Excel / PPT / PDF / MarkItDown 材料","factRefs":["fact-local"]},
        {"id":"planner","label":"LLMRouter 语义规划\n类型、节点、关系、事实引用","shape":"process","factRefs":[]},
        {"id":"spec","label":"DiagramSpec","shape":"database","factRefs":[]},
        {"id":"office","label":"Word / PPT 原生图形","factRefs":[]}
      ],
      "edges": [
        {"id":"e1","from":"materials","to":"planner","label":"结构化摘要","dashed":false,"factRefs":["fact-local"]},
        {"id":"e2","from":"planner","to":"spec","label":"图形语义","dashed":false,"factRefs":[]},
        {"id":"e3","from":"spec","to":"office","label":"确定性布局","dashed":false,"factRefs":[]}
      ]
    }
    """;

    [Fact]
    public void ParseCompileAndSvgPreserveFactBindings()
    {
        var spec = DiagramSpec.Parse(ArchitectureSpec);
        var graph = DiagramCompiler.Compile(spec);
        var svg = DiagramSvgRenderer.Render(spec, graph, DiagramTheme.Default);
        var evidence = DiagramSvgRenderer.Evidence(spec, graph);

        Assert.Equal(4, graph.Nodes.Count);
        Assert.True(graph.Nodes.Single(node => node.Id == "planner").H >= 2.5,
            "deterministic mixed Latin/CJK wrapping must contribute to the native node height");
        var planner = graph.Nodes.Single(node => node.Id == "planner");
        var plannerLines = DiagramTextMetrics.Wrap(planner.Label, Math.Max(0.8, planner.W - 0.6));
        Assert.Contains("LLMRouter 语义规划", plannerLines);
        Assert.DoesNotContain(plannerLines, line => line == "划");
        var materials = graph.Nodes.Single(node => node.Id == "materials");
        var materialLines = DiagramTextMetrics.Wrap(materials.Label, Math.Max(0.8, materials.W - 0.6));
        Assert.DoesNotContain(materialLines, line => line is "P" or "PT");
        Assert.Equal(CompactDiagramText(materials.Label), CompactDiagramText(string.Join(' ', materialLines)));
        Assert.Empty(evidence.NodeOverlaps);
        Assert.True(evidence.FactBindingsComplete);
        Assert.Contains("data-fact-refs=\"fact-local\"", svg);
        var svgVisibleText = CompactDiagramText(string.Concat(
            System.Xml.Linq.XDocument.Parse(svg).DescendantNodes()
                .OfType<System.Xml.Linq.XText>().Select(item => item.Value)));
        Assert.Contains(CompactDiagramText("LLMRouter 语义规划"), svgVisibleText);
    }

    [Fact]
    public void CommunicationIntentBindingSurvivesDeterministicDiagramRendering()
    {
        var json = ArchitectureSpec.Replace("\"facts\": [", "\"communication\": {\"intentId\":\"monthly-review\",\"representationChoiceId\":\"decision-flow\",\"purpose\":\"解释经营判断链路\",\"audience\":\"管理层\",\"desiredResponse\":\"批准纠偏行动\",\"coreMessage\":\"收入事实必须驱动行动\",\"maxNodesPerDiagram\":4},\n      \"facts\": [");
        var spec = DiagramSpec.Parse(json);
        var graph = DiagramCompiler.Compile(spec);
        var svg = DiagramSvgRenderer.Render(spec, graph, DiagramTheme.Default);
        var evidence = DiagramSvgRenderer.Evidence(spec, graph);

        Assert.True(evidence.CommunicationBindingComplete);
        Assert.Equal("monthly-review", evidence.CommunicationIntentId);
        Assert.Equal("decision-flow", evidence.RepresentationChoiceId);
        Assert.Contains("data-communication-intent-id=\"monthly-review\"", svg);
        Assert.Contains("data-representation-choice-id=\"decision-flow\"", svg);
    }

    [Fact]
    public void CommunicationIntentDensityBudgetRequiresSemanticSplitBeforeLayout()
    {
        var json = ArchitectureSpec.Replace("\"facts\": [", "\"communication\": {\"intentId\":\"monthly-review\",\"representationChoiceId\":\"decision-flow\",\"purpose\":\"解释经营判断链路\",\"audience\":\"管理层\",\"desiredResponse\":\"批准纠偏行动\",\"coreMessage\":\"收入事实必须驱动行动\",\"maxNodesPerDiagram\":3},\n      \"facts\": [");
        var error = Assert.Throws<ArgumentException>(() => DiagramSpec.Parse(json));
        Assert.Contains("node-density budget", error.Message);
    }

    [Theory]
    [InlineData("flowchart")]
    [InlineData("mindmap")]
    [InlineData("relationship")]
    [InlineData("architecture")]
    [InlineData("timeline")]
    public void GraphKindsUseOneDeterministicLayoutContract(string type)
    {
        var json = ArchitectureSpec.Replace("\"architecture\"", $"\"{type}\"");
        var graph = DiagramCompiler.Compile(DiagramSpec.Parse(json));
        Assert.Equal(4, graph.Nodes.Count);
        Assert.True(graph.SlideWidthCm > 0);
        Assert.True(graph.SlideHeightCm > 0);
    }

    [Fact]
    public void SequenceUsesTheSameSpecInsteadOfMermaidOrOoxml()
    {
        var json = ArchitectureSpec.Replace("\"architecture\"", "\"sequence\"");
        var graph = DiagramCompiler.Compile(DiagramSpec.Parse(json));
        Assert.Equal(4, graph.Nodes.Count);
        Assert.True(graph.Edges.Count >= 7); // 4 lifelines + 3 messages
    }

    [Fact]
    public void UnknownFieldsFailClosed()
    {
        var json = ArchitectureSpec.Replace("\"type\": \"architecture\",", "\"type\": \"architecture\", \"ooxml\": \"forbidden\",");
        var error = Assert.Throws<ArgumentException>(() => DiagramSpec.Parse(json));
        Assert.Contains("invalid DiagramSpec JSON", error.Message);
    }

    [Fact]
    public void UnknownFactReferenceFailsClosed()
    {
        var json = ArchitectureSpec.Replace("\"fact-local\"]}", "\"missing\"]}");
        var error = Assert.Throws<ArgumentException>(() => DiagramSpec.Parse(json));
        Assert.Contains("unknown fact", error.Message);
    }

    [Fact]
    public void OversizedDiagramRequiresSemanticSplit()
    {
        var nodes = string.Join(',', Enumerable.Range(1, 25)
            .Select(index => $"{{\"id\":\"n{index}\",\"label\":\"节点 {index}\",\"factRefs\":[]}}"));
        var json = $"{{\"schemaVersion\":1,\"diagramId\":\"oversized\",\"type\":\"flowchart\",\"facts\":[],\"nodes\":[{nodes}],\"edges\":[]}}";
        var error = Assert.Throws<ArgumentException>(() => DiagramSpec.Parse(json));
        Assert.Contains("split larger diagrams", error.Message);
    }

    [Fact]
    public void DenseButSchemaValidDiagramRequiresSemanticSplitEvidence()
    {
        var nodes = string.Join(',', Enumerable.Range(1, 19)
            .Select(index => $"{{\"id\":\"n{index}\",\"label\":\"节点 {index}\",\"factRefs\":[]}}"));
        var json = $"{{\"schemaVersion\":1,\"diagramId\":\"dense\",\"type\":\"relationship\",\"facts\":[],\"nodes\":[{nodes}],\"edges\":[]}}";
        var spec = DiagramSpec.Parse(json);
        var graph = DiagramCompiler.Compile(spec);
        var evidence = DiagramSvgRenderer.Evidence(spec, graph);

        Assert.True(evidence.RequiresSemanticSplit);
        Assert.True(evidence.SuggestedPageCount >= 2);
        Assert.True(evidence.EstimatedManualRepairActions >= 1);
    }

    [Fact]
    public void NativeOfficeDiagramsUseRealConnectorAttachmentsAndChineseThemeFonts()
    {
        using var temp = new TempDirectory();
        var specPath = temp.File("diagram.json");
        var themePath = temp.File("theme.json");
        File.WriteAllText(specPath, ArchitectureSpec.Replace("确定性布局", "确定性布局与中文长标签验收"));
        File.WriteAllText(themePath, """
        {"schemaVersion":1,"colors":{"dk1":"172033","dk2":"2F3E56","lt1":"FFFFFF","lt2":"F3F6FA","accent1":"2563EB","accent2":"0EA5E9","accent3":"14B8A6","accent4":"F59E0B","accent5":"8B5CF6","accent6":"EF4444"},"fonts":{"majorLatin":"Aptos Display","majorEastAsia":"微软雅黑","minorLatin":"Aptos","minorEastAsia":"微软雅黑"}}
        """);
        var spec = DiagramSpec.Load(specPath);
        var graph = DiagramCompiler.Compile(spec);
        var evidence = DiagramSvgRenderer.Evidence(spec, graph, DiagramTheme.Load(themePath));
        Assert.Empty(evidence.NodeOverlaps);
        Assert.Empty(evidence.TextOverflows);
        Assert.True(evidence.StructureComplete);
        Assert.True(evidence.ConnectorAttachmentsComplete);
        Assert.True(evidence.TypeSelectionPassed);
        Assert.False(evidence.RequiresSemanticSplit);
        Assert.Equal(0, evidence.EstimatedManualRepairActions);
        Assert.Equal("微软雅黑", evidence.BodyFont);

        var pptx = temp.File("diagram.pptx");
        BlankDocCreator.Create(pptx, "zh-CN");
        using (var handler = new PowerPointHandler(pptx, editable: true))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "diagram", null, new Dictionary<string, string>
                { ["spec"] = specPath, ["themeFile"] = themePath });
        }
        using (var package = PresentationDocument.Open(pptx, false))
        {
            Assert.Empty(new OpenXmlValidator().Validate(package));
            var group = Assert.Single(package.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!.Elements<GroupShape>());
            Assert.Equal(spec.Edges.Count, group.Elements<ConnectionShape>().Count());
            Assert.All(group.Elements<ConnectionShape>(), connector =>
            {
                Assert.NotNull(connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties?.StartConnection);
                Assert.NotNull(connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties?.EndConnection);
            });
            Assert.All(group.Elements<Shape>().Where(shape => shape.TextBody is not null), shape =>
                Assert.Contains(shape.TextBody!.Descendants<DocumentFormat.OpenXml.Drawing.EastAsianFont>(),
                    font => font.Typeface?.Value == "微软雅黑"));
        }

        var docx = temp.File("diagram.docx");
        OpenXmlFixture.CreateDocument(docx);
        using (var handler = new WordHandler(docx, editable: true))
            handler.Add("/body", "diagram", null, new Dictionary<string, string>
                { ["spec"] = specPath, ["themeFile"] = themePath });
        using (var package = WordprocessingDocument.Open(docx, false))
        {
            Assert.Empty(new OpenXmlValidator().Validate(package));
            var xml = package.MainDocumentPart!.Document!.OuterXml;
            Assert.Equal(spec.Edges.Count, System.Text.RegularExpressions.Regex.Matches(xml, "<a:stCxn\\b").Count);
            Assert.Equal(spec.Edges.Count, System.Text.RegularExpressions.Regex.Matches(xml, "<a:endCxn\\b").Count);
            Assert.Matches("<a:prstGeom prst=\"(?:downArrow|upArrow|rightArrow|leftArrow|rect)\"", xml);
            Assert.Contains("w:eastAsia=\"微软雅黑\"", xml);
            Assert.True(System.Text.RegularExpressions.Regex.Matches(xml, "<w:br\\b").Count >= 4,
                "deterministic Word wrapping must use native w:br elements rather than raw newline text");
        }
    }

    [Fact]
    public void NativeDiagramUsesRemainingCanvasBelowAnExistingConclusionTitle()
    {
        using var temp = new TempDirectory();
        var specPath = temp.File("diagram.json");
        File.WriteAllText(specPath, ArchitectureSpec.Replace("\"top-down\"", "\"left-right\""));
        var pptx = temp.File("conclusion-led-diagram.pptx");
        BlankDocCreator.Create(pptx, "zh-CN");

        using (var handler = new PowerPointHandler(pptx, editable: true))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
            {
                ["text"] = "收入增长但风险仍需在本期关闭",
                ["x"] = "1.2cm",
                ["y"] = "0.6cm",
                ["width"] = "31.4cm",
                ["height"] = "1.4cm",
                ["fill"] = "none",
                ["line"] = "none",
            });
            handler.Add("/slide[1]", "diagram", null, new Dictionary<string, string> { ["spec"] = specPath });
        }

        using var package = PresentationDocument.Open(pptx, false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var shapeTree = package.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!;
        var title = Assert.Single(shapeTree.Elements<Shape>());
        var diagram = Assert.Single(shapeTree.Elements<GroupShape>());
        var titleTransform = title.ShapeProperties!.Transform2D!;
        var diagramTransform = diagram.GroupShapeProperties!.TransformGroup!;
        var titleBottom = titleTransform.Offset!.Y!.Value + titleTransform.Extents!.Cy!.Value;
        Assert.True(diagramTransform.Offset!.Y!.Value > titleBottom,
            "the diagram must be placed below the existing business conclusion instead of overlapping it");

        var slideSize = package.PresentationPart.Presentation.SlideSize!;
        Assert.True(diagramTransform.Extents!.Cx!.Value >= slideSize.Cx!.Value * 0.45,
            "the editable diagram should use a meaningful share of the remaining slide canvas");
    }

    [Fact]
    public void NativeDiagramIgnoresEmptyBackgroundPanelAndReservesBottomNote()
    {
        using var temp = new TempDirectory();
        var specPath = temp.File("diagram.json");
        File.WriteAllText(specPath, ArchitectureSpec.Replace("\"top-down\"", "\"left-right\""));
        var pptx = temp.File("panel-diagram.pptx");
        BlankDocCreator.Create(pptx, "zh-CN");

        using (var handler = new PowerPointHandler(pptx, editable: true))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
            {
                ["text"] = "经营事实到管理动作", ["x"] = "1.2cm", ["y"] = "0.6cm",
                ["width"] = "31.4cm", ["height"] = "1.4cm", ["fill"] = "none", ["line"] = "none",
            });
            handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
            {
                ["x"] = "2.1cm", ["y"] = "4.2cm", ["width"] = "29.6cm", ["height"] = "11.5cm",
                ["fill"] = "#F7F9FC", ["geometry"] = "roundRect",
            });
            handler.Add("/slide[1]", "shape", null, new Dictionary<string, string>
            {
                ["text"] = "来源：经营事实台账", ["x"] = "2.1cm", ["y"] = "17.0cm",
                ["width"] = "29.6cm", ["height"] = "0.7cm", ["fill"] = "none", ["line"] = "none",
            });
            handler.Add("/slide[1]", "diagram", null, new Dictionary<string, string> { ["spec"] = specPath });
        }

        using var package = PresentationDocument.Open(pptx, false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var tree = package.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!;
        var diagram = Assert.Single(tree.Elements<GroupShape>());
        var transform = diagram.GroupShapeProperties!.TransformGroup!;
        var slide = package.PresentationPart.Presentation.SlideSize!;
        Assert.True(transform.Extents!.Cx!.Value >= slide.Cx!.Value * 0.45,
            "an empty background panel must not collapse the diagram width");
        Assert.True(transform.Extents.Cy!.Value >= slide.Cy!.Value * 0.28,
            "an empty background panel must not collapse the diagram height");
        Assert.True(transform.Offset!.Y!.Value + transform.Extents.Cy.Value < (long)(17.0 * 360000),
            "the diagram must stay above the existing bottom source note");
    }

    [Fact]
    public void ShapeConnectorDiagramCanBeRefreshedInPlaceByStableDiagramId()
    {
        using var temp = new TempDirectory();
        var specPath = temp.File("diagram.json");
        File.WriteAllText(specPath, ArchitectureSpec);
        var pptx = temp.File("refresh.pptx");
        BlankDocCreator.Create(pptx, "zh-CN");
        using (var handler = new PowerPointHandler(pptx, editable: true))
        {
            handler.Add("/", "slide", null, new Dictionary<string, string>());
            handler.Add("/slide[1]", "diagram", null, new Dictionary<string, string> { ["spec"] = specPath });
        }
        var updated = DiagramSpec.Parse(ArchitectureSpec.Replace("LLMRouter 语义规划", "本地材料变更后的语义规划"));
        var receipt = global::OfficeCli.CommandBuilder.RefreshNativeDiagram(pptx, updated, null);
        Assert.True(receipt.Ok);
        using var package = PresentationDocument.Open(pptx, false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        var group = Assert.Single(package.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!.Elements<GroupShape>());
        var visibleText = CompactDiagramText(string.Concat(
            group.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(item => item.Text)));
        Assert.Contains(CompactDiagramText("本地材料变更后的语义规划"), visibleText);
        Assert.DoesNotContain(CompactDiagramText("LLMRouter 语义规划"), visibleText);
    }

    [Fact]
    public void WordShapeConnectorDiagramRefreshPreservesUnrelatedContentAndPlacement()
    {
        using var temp = new TempDirectory();
        var specPath = temp.File("diagram.json");
        File.WriteAllText(specPath, ArchitectureSpec);
        var docx = temp.File("refresh.docx");
        OpenXmlFixture.CreateDocument(docx,
            new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                new DocumentFormat.OpenXml.Wordprocessing.Run(
                    new DocumentFormat.OpenXml.Wordprocessing.Text("用户保留的前置正文"))));
        using (var handler = new WordHandler(docx, editable: true))
            handler.Add("/body", "diagram", null, new Dictionary<string, string> { ["spec"] = specPath });
        using (var package = WordprocessingDocument.Open(docx, true))
        {
            package.MainDocumentPart!.Document!.Body!.AppendChild(
                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text("用户保留的后置正文"))));
            package.MainDocumentPart.Document.Save();
        }

        string[] preservedBefore;
        using (var package = WordprocessingDocument.Open(docx, false))
            preservedBefore = package.MainDocumentPart!.Document!.Body!.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                .Where(paragraph => paragraph.InnerText.StartsWith("用户保留的", StringComparison.Ordinal))
                .Select(paragraph => paragraph.InnerText).ToArray();

        var updated = DiagramSpec.Parse(ArchitectureSpec.Replace("LLMRouter 语义规划", "只更新受影响的智能规划节点"));
        var receipt = global::OfficeCli.CommandBuilder.RefreshNativeDiagram(docx, updated, null);
        Assert.True(receipt.Ok);
        using var refreshed = WordprocessingDocument.Open(docx, false);
        Assert.Empty(new OpenXmlValidator().Validate(refreshed));
        var body = refreshed.MainDocumentPart!.Document!.Body!;
        var preservedAfter = body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
            .Where(paragraph => paragraph.InnerText.StartsWith("用户保留的", StringComparison.Ordinal))
            .Select(paragraph => paragraph.InnerText).ToArray();
        Assert.Equal(preservedBefore, preservedAfter);
        Assert.Single(body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>(), paragraph =>
            paragraph.OuterXml.Contains("DiagramId:delivery-architecture", StringComparison.Ordinal));
        var visibleText = CompactDiagramText(body.InnerText);
        Assert.Contains(CompactDiagramText("只更新受影响的智能规划节点"), visibleText);
        Assert.DoesNotContain(CompactDiagramText("LLMRouter 语义规划"), visibleText);
        Assert.Equal(updated.Edges.Count, System.Text.RegularExpressions.Regex.Matches(body.OuterXml, "<a:stCxn\\b").Count);
        Assert.Equal(updated.Edges.Count, System.Text.RegularExpressions.Regex.Matches(body.OuterXml, "<a:endCxn\\b").Count);
    }
}
