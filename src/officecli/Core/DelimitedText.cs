// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Quote-aware grid parser for the table `data=` property, in both of its
/// forms: inline ("H1,H2;R1C1,R1C2" — semicolon rows) and CSV pulled from a
/// file/URL/data-URI (newline rows).
///
/// Splitting on the separator alone loses any cell that contains one, which is
/// exactly what quoting exists for: "Doe, John",30 is two fields, not three.
///
/// Excel import keeps its own parser (ExcelHandler.ParseCsv) — it has to
/// preserve blank source lines to keep row numbers aligned with the sheet,
/// while a table built from `data=` wants empty rows dropped.
/// </summary>
public static class DelimitedText
{
    /// <summary>
    /// Parse the lossless table form accepted by <c>dataJson=</c>. The root
    /// value must be an array of equally-sized row arrays. Cells may be JSON
    /// scalars, or an object with a required <c>value</c> member and optional
    /// <c>display</c> member. When present, <c>display</c> is the exact text
    /// written to Office; otherwise the scalar value is rendered invariantly.
    /// This keeps values such as "12,480" in one cell without depending on
    /// delimiter quoting and gives project planners a typed, auditable input.
    /// </summary>
    public static string[][] ParseJsonGrid(string content)
    {
        if (string.IsNullOrWhiteSpace(content))
            throw new ArgumentException("Table 'dataJson' is empty.");

        using var document = JsonDocument.Parse(content);
        if (document.RootElement.ValueKind != JsonValueKind.Array)
            throw new ArgumentException("Table 'dataJson' must be a JSON array of row arrays.");

        var rows = new List<string[]>();
        foreach (var rowElement in document.RootElement.EnumerateArray())
        {
            if (rowElement.ValueKind != JsonValueKind.Array)
                throw new ArgumentException("Every table 'dataJson' row must be a JSON array.");
            rows.Add(rowElement.EnumerateArray().Select(JsonCellText).ToArray());
        }

        EnsureRectangular(rows, "dataJson");
        return rows.ToArray();
    }

    /// <summary>
    /// Reject ragged tables instead of padding them and silently shifting
    /// business values into the wrong columns.
    /// </summary>
    public static void EnsureRectangular(IReadOnlyList<string[]> rows, string propertyName = "data")
    {
        if (rows.Count == 0 || rows.All(row => row.Length == 0))
            throw new ArgumentException($"Table '{propertyName}' is empty — provide at least one cell.");

        var expected = rows[0].Length;
        if (expected == 0)
            throw new ArgumentException($"Table '{propertyName}' first row is empty.");
        for (var index = 1; index < rows.Count; index++)
        {
            if (rows[index].Length != expected)
                throw new ArgumentException(
                    $"Table '{propertyName}' row {index + 1} has {rows[index].Length} cell(s), but row 1 has {expected}. "
                    + "Quote delimiter-containing cells or use dataJson with explicit row arrays; ragged rows are not padded because that can corrupt business data.");
        }
    }

    /// <summary>
    /// Split <paramref name="content"/> into rows of fields. A field wrapped in
    /// double quotes may contain the field separator, the row separator and
    /// doubled quotes ("" → a literal "). Unquoted fields are trimmed; quoted
    /// fields keep their interior verbatim. Rows whose fields are all empty are
    /// dropped, so a trailing newline does not add a phantom row. When
    /// <paramref name="rowSeparator"/> is '\n', CRLF is handled too.
    /// </summary>
    public static string[][] ParseGrid(string content, char fieldSeparator, char rowSeparator)
    {
        var rows = new List<string[]>();
        if (string.IsNullOrEmpty(content)) return rows.ToArray();
        if (content[0] == '﻿') content = content[1..];

        var row = new List<string>();
        var field = new StringBuilder();
        var inQuotes = false;
        var quoted = false;

        void EndField()
        {
            row.Add(quoted ? field.ToString() : field.ToString().Trim());
            field.Clear();
            quoted = false;
        }

        void EndRow()
        {
            EndField();
            if (row.Exists(c => c.Length > 0)) rows.Add(row.ToArray());
            row.Clear();
        }

        for (int i = 0; i < content.Length; i++)
        {
            var c = content[i];

            if (inQuotes)
            {
                if (c != '"') { field.Append(c); continue; }
                // "" inside a quoted field is one literal quote; a lone quote closes it.
                if (i + 1 < content.Length && content[i + 1] == '"') { field.Append('"'); i++; }
                else inQuotes = false;
                continue;
            }

            if (c == '"' && IsBlank(field))
            {
                inQuotes = true;
                quoted = true;
                field.Clear();          // drop the whitespace that preceded the quote
                continue;
            }
            if (c == fieldSeparator) { EndField(); continue; }
            if (c == rowSeparator) { EndRow(); continue; }
            if (rowSeparator == '\n' && c == '\r') continue;
            field.Append(c);
        }

        EndRow();
        return rows.ToArray();
    }

    private static string JsonCellText(JsonElement cell)
    {
        if (cell.ValueKind == JsonValueKind.Object)
        {
            if (!cell.TryGetProperty("value", out var value))
                throw new ArgumentException("A table dataJson cell object must contain 'value'.");
            if (cell.TryGetProperty("display", out var display))
            {
                if (display.ValueKind != JsonValueKind.String)
                    throw new ArgumentException("A table dataJson cell 'display' must be a string.");
                return display.GetString() ?? "";
            }
            return JsonScalarText(value);
        }
        return JsonScalarText(cell);
    }

    private static string JsonScalarText(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.String => value.GetString() ?? "",
        JsonValueKind.Number => value.GetRawText(),
        JsonValueKind.True => "true",
        JsonValueKind.False => "false",
        JsonValueKind.Null => "",
        _ => throw new ArgumentException("Table dataJson cells must be scalars or { value, display? } objects."),
    };

    private static bool IsBlank(StringBuilder sb)
    {
        for (int i = 0; i < sb.Length; i++)
            if (!char.IsWhiteSpace(sb[i])) return false;
        return true;
    }
}
