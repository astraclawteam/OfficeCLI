// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeCli.Core.Diagram;

/// <summary>Small projection of the shared Office theme used by diagram emitters.</summary>
public sealed class DiagramTheme
{
    public string Background { get; init; } = "FFFFFF";
    public string Text { get; init; } = "172033";
    public string MutedText { get; init; } = "2F3E56";
    public string Surface { get; init; } = "F3F6FA";
    public string Primary { get; init; } = "2563EB";
    public string Secondary { get; init; } = "0EA5E9";
    public string Positive { get; init; } = "14B8A6";
    public string Warning { get; init; } = "F59E0B";
    public string Accent { get; init; } = "8B5CF6";
    public string Danger { get; init; } = "EF4444";
    public string MajorLatinFont { get; init; } = "Aptos Display";
    public string MajorEastAsiaFont { get; init; } = "Microsoft YaHei";
    public string MinorLatinFont { get; init; } = "Aptos";
    public string MinorEastAsiaFont { get; init; } = "Microsoft YaHei";

    public static DiagramTheme Default { get; } = new();

    public static DiagramTheme Load(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return Default;
        if (!File.Exists(path)) throw new ArgumentException($"Office theme file not found: '{path}'.");
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        var root = document.RootElement;
        if (!root.TryGetProperty("schemaVersion", out var schema) || schema.GetInt32() != 1 ||
            !root.TryGetProperty("colors", out var colors) || colors.ValueKind != JsonValueKind.Object)
            throw new ArgumentException("Office theme must use schemaVersion 1 and contain colors.");
        string C(string key, string fallback)
        {
            if (!colors.TryGetProperty(key, out var value) || value.ValueKind != JsonValueKind.String) return fallback;
            var color = value.GetString()?.TrimStart('#') ?? "";
            if (!Regex.IsMatch(color, "^[0-9A-Fa-f]{6}$")) throw new ArgumentException($"Office theme color '{key}' is invalid.");
            return color.ToUpperInvariant();
        }
        string F(string key, string fallback)
        {
            if (!root.TryGetProperty("fonts", out var fonts) || fonts.ValueKind != JsonValueKind.Object ||
                !fonts.TryGetProperty(key, out var value) || value.ValueKind != JsonValueKind.String)
                return fallback;
            var font = value.GetString()?.Trim() ?? "";
            if (font.Length is < 1 or > 120 || font.IndexOfAny(['\r', '\n', '<', '>']) >= 0)
                throw new ArgumentException($"Office theme font '{key}' is invalid.");
            return font;
        }
        return new DiagramTheme
        {
            Background = C("lt1", "FFFFFF"), Text = C("dk1", "172033"), MutedText = C("dk2", "2F3E56"),
            Surface = C("lt2", "F3F6FA"), Primary = C("accent1", "2563EB"), Secondary = C("accent2", "0EA5E9"),
            Positive = C("accent3", "14B8A6"), Warning = C("accent4", "F59E0B"), Accent = C("accent5", "8B5CF6"),
            Danger = C("accent6", "EF4444"),
            MajorLatinFont = F("majorLatin", "Aptos Display"),
            MajorEastAsiaFont = F("majorEastAsia", "Microsoft YaHei"),
            MinorLatinFont = F("minorLatin", "Aptos"),
            MinorEastAsiaFont = F("minorEastAsia", "Microsoft YaHei"),
        };
    }
}
