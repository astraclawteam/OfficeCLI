// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    private sealed record LayoutBox(
        OpenXmlElement Element, string Path, long X, long Y, long Width, long Height,
        string Text, PlaceholderValues? Placeholder);

    private void AppendOptInLayoutIssues(
        ShapeTree shapeTree, int slideNum, long slideWidth, long slideHeight,
        List<DocumentIssue> issues, ref int issueNum)
    {
        const long quarterInch = 228600;
        const long tightGap = 137160;
        const long alignmentTolerance = 91440;

        var boxes = new List<LayoutBox>();
        int shapeIndex = 0;
        foreach (var shape in shapeTree.Elements<Shape>())
        {
            shapeIndex++;
            var xfrm = shape.ShapeProperties?.Transform2D;
            if (xfrm?.Offset?.X == null || xfrm.Offset.Y == null
                || xfrm.Extents?.Cx == null || xfrm.Extents.Cy == null)
                continue;
            var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<PlaceholderShape>()?.Type?.Value;
            boxes.Add(new LayoutBox(
                shape,
                $"/slide[{slideNum}]/{BuildElementPathSegment("shape", shape, shapeIndex)}",
                xfrm.Offset.X.Value, xfrm.Offset.Y.Value,
                xfrm.Extents.Cx.Value, xfrm.Extents.Cy.Value,
                GetShapeText(shape), ph));
        }

        static bool IsFooter(LayoutBox box) =>
            box.Placeholder == PlaceholderValues.Footer
            || box.Placeholder == PlaceholderValues.SlideNumber
            || box.Placeholder == PlaceholderValues.DateAndTime;
        static long Intersection(long a0, long aLength, long b0, long bLength) =>
            Math.Max(0, Math.Min(a0 + aLength, b0 + bLength) - Math.Max(a0, b0));
        static double VisualUnits(string text)
        {
            double units = 0;
            foreach (var c in text)
                units += c > 0x2E80 ? 1.0 : char.IsWhiteSpace(c) ? 0.35 : 0.55;
            return units;
        }

        foreach (var box in boxes.Where(b => !string.IsNullOrWhiteSpace(b.Text) && !IsFooter(b)))
        {
            var edge = Math.Min(Math.Min(box.X, box.Y),
                Math.Min(slideWidth - (box.X + box.Width), slideHeight - (box.Y + box.Height)));
            if (edge >= 0 && edge < quarterInch)
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"L{++issueNum}", Type = IssueType.Format,
                    Subtype = IssueSubtypes.PptLayout, Severity = IssueSeverity.Warning,
                    Path = box.Path,
                    Message = $"Text box is only {edge / 914400.0:F2}in from the nearest slide edge",
                    Suggestion = "Move it inward or deliberately mark it as full-bleed decoration."
                });
            }

            var normalized = string.Join(' ', box.Text.Split((char[]?)null,
                StringSplitOptions.RemoveEmptyEntries));
            var capacity = Math.Max(1.0, box.Width / 914400.0 * 8.0);
            if (!box.Text.Contains('\n') && normalized.Length >= 24
                && VisualUnits(normalized) / capacity >= 4.0)
            {
                issues.Add(new DocumentIssue
                {
                    Id = $"L{++issueNum}", Type = IssueType.Format,
                    Subtype = IssueSubtypes.PptLayout, Severity = IssueSeverity.Warning,
                    Path = box.Path,
                    Message = "Narrow text box is likely to wrap into four or more short lines",
                    Suggestion = "Widen the box, shorten the copy, or reduce the font size after rendering review."
                });
            }
        }

        foreach (var footer in boxes.Where(IsFooter))
        foreach (var content in boxes.Where(b => !ReferenceEquals(b, footer) && !IsFooter(b)))
        {
            if (Intersection(footer.X, footer.Width, content.X, content.Width) <= 0
                || Intersection(footer.Y, footer.Height, content.Y, content.Height) <= 0)
                continue;
            issues.Add(new DocumentIssue
            {
                Id = $"L{++issueNum}", Type = IssueType.Format,
                Subtype = IssueSubtypes.PptLayout, Severity = IssueSeverity.Warning,
                Path = footer.Path,
                Message = $"Footer/date/slide-number box intersects {content.Path}",
                Suggestion = "Reserve a footer band or move the content above it."
            });
            break;
        }

        var contentBoxes = boxes.Where(b => !IsFooter(b) && !string.IsNullOrWhiteSpace(b.Text)).ToList();
        for (int i = 0; i < contentBoxes.Count; i++)
        for (int j = i + 1; j < contentBoxes.Count; j++)
        {
            var a = contentBoxes[i];
            var b = contentBoxes[j];
            var verticalOverlap = Intersection(a.Y, a.Height, b.Y, b.Height);
            var horizontalOverlap = Intersection(a.X, a.Width, b.X, b.Width);
            var horizontalGap = Math.Max(b.X - (a.X + a.Width), a.X - (b.X + b.Width));
            var verticalGap = Math.Max(b.Y - (a.Y + a.Height), a.Y - (b.Y + b.Height));
            bool tightHorizontal = horizontalGap > 0 && horizontalGap < tightGap
                && verticalOverlap * 2 >= Math.Min(a.Height, b.Height);
            bool tightVertical = verticalGap > 0 && verticalGap < tightGap
                && horizontalOverlap * 2 >= Math.Min(a.Width, b.Width);
            if (!tightHorizontal && !tightVertical) continue;
            var gap = tightHorizontal ? horizontalGap : verticalGap;
            issues.Add(new DocumentIssue
            {
                Id = $"L{++issueNum}", Type = IssueType.Format,
                Subtype = IssueSubtypes.PptLayout, Severity = IssueSeverity.Warning,
                Path = a.Path,
                Message = $"Only {gap / 914400.0:F2}in gap to {b.Path}",
                Suggestion = "Increase the gap or confirm the elements are intentionally joined."
            });
        }

        var rows = contentBoxes
            .GroupBy(b => (long)Math.Round((b.Y + b.Height / 2.0) / (alignmentTolerance * 4)))
            .Select(g => g.OrderBy(b => b.X).ToList())
            .Where(g => g.Count >= 3);
        foreach (var row in rows)
        {
            var topSpread = row.Max(b => b.Y) - row.Min(b => b.Y);
            var gaps = row.Zip(row.Skip(1), (a, b) => b.X - (a.X + a.Width))
                .Where(g => g >= 0).ToList();
            var gapSpread = gaps.Count >= 2 ? gaps.Max() - gaps.Min() : 0;
            if (topSpread <= alignmentTolerance && gapSpread <= alignmentTolerance) continue;
            issues.Add(new DocumentIssue
            {
                Id = $"L{++issueNum}", Type = IssueType.Format,
                Subtype = IssueSubtypes.PptLayout, Severity = IssueSeverity.Warning,
                Path = row[0].Path,
                Message = "Repeated row elements have inconsistent baselines or gaps",
                Suggestion = "Align tops/centres and distribute the repeated elements evenly."
            });
        }

        int connectorIndex = 0;
        foreach (var connector in shapeTree.Elements<ConnectionShape>())
        {
            connectorIndex++;
            var nv = connector.NonVisualConnectionShapeProperties?.NonVisualConnectorShapeDrawingProperties;
            if (nv?.StartConnection == null || nv.EndConnection == null) continue;
            var outline = connector.ShapeProperties?.GetFirstChild<Drawing.Outline>();
            var head = outline?.GetFirstChild<Drawing.HeadEnd>()?.Type?.Value;
            var tail = outline?.GetFirstChild<Drawing.TailEnd>()?.Type?.Value;
            bool hasHead = head != null && head != Drawing.LineEndValues.None;
            bool hasTail = tail != null && tail != Drawing.LineEndValues.None;
            if (hasHead || hasTail) continue;
            issues.Add(new DocumentIssue
            {
                Id = $"L{++issueNum}", Type = IssueType.Format,
                Subtype = IssueSubtypes.PptLayout, Severity = IssueSeverity.Warning,
                Path = $"/slide[{slideNum}]/connector[{connectorIndex}]",
                Message = "Connector links two shapes but has no arrowhead",
                Suggestion = "Add a head or tail arrow when the line represents directional flow."
            });
        }
    }
}
