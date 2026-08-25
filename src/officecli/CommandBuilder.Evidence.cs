// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildBrandExtractCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Existing DOCX, XLSX or PPTX used as local brand evidence" };
        var outOption = new Option<FileInfo>("--out") { Description = "Output brand profile JSON", Required = true };
        var themeOutOption = new Option<FileInfo>("--theme-out") { Description = "Output cross-format Office theme JSON", Required = true };
        var assetsOption = new Option<DirectoryInfo>("--assets") { Description = "Local directory for extracted media candidates", Required = true };
        var idOption = new Option<string>("--id") { Description = "Stable lowercase brand profile id", Required = true };
        var nameOption = new Option<string>("--name") { Description = "Brand display name", Required = true };
        var command = new Command("brand-extract", "Extract a local cross-format brand profile without uploading the source file");
        command.Add(fileArg); command.Add(outOption); command.Add(themeOutOption); command.Add(assetsOption); command.Add(idOption); command.Add(nameOption); command.Add(jsonOption);
        command.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var id = result.GetValue(idOption)!;
            if (!System.Text.RegularExpressions.Regex.IsMatch(id, "^[a-z][a-z0-9._-]{0,63}$"))
                throw new ArgumentException("brand profile id must match ^[a-z][a-z0-9._-]{0,63}$");
            var output = result.GetValue(outOption)!.FullName;
            var themeOutput = result.GetValue(themeOutOption)!.FullName;
            var assets = result.GetValue(assetsOption)!.FullName;
            var (profile, theme) = OfficePackageEvidence.ExtractBrand(result.GetValue(fileArg)!.FullName, id, result.GetValue(nameOption)!, assets);
            OfficePackageEvidence.WriteJson(output, profile);
            OfficePackageEvidence.WriteJson(themeOutput, theme);
            var payload = new OfficeEvidenceReceipt(output, themeOutput, assets, profile.Assets.Count, profile.Source.Sha256);
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelope(JsonSerializer.Serialize(payload, OfficeEvidenceJsonContext.Default.OfficeEvidenceReceipt)));
            else Console.WriteLine($"Brand profile written: {output}");
            return 0;
        }, json); });
        return command;
    }

    private static Command BuildFidelitySnapshotCommand(Option<bool> jsonOption)
    {
        var fileArg = new Argument<FileInfo>("file") { Description = "Existing DOCX, XLSX or PPTX" };
        var outOption = new Option<FileInfo>("--out") { Description = "Output fidelity snapshot JSON", Required = true };
        var command = new Command("fidelity-snapshot", "Capture package and format evidence before editing an existing Office file");
        command.Add(fileArg); command.Add(outOption); command.Add(jsonOption);
        command.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var snapshot = OfficePackageEvidence.Snapshot(result.GetValue(fileArg)!.FullName);
            var output = result.GetValue(outOption)!.FullName;
            OfficePackageEvidence.WriteJson(output, snapshot);
            var payload = new OfficeEvidenceReceipt(output, SourceSha256: snapshot.SourceSha256, PartCount: snapshot.Parts.Count, Features: snapshot.Features);
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelope(JsonSerializer.Serialize(payload, OfficeEvidenceJsonContext.Default.OfficeEvidenceReceipt)));
            else Console.WriteLine($"Fidelity snapshot written: {output}");
            return 0;
        }, json); });
        return command;
    }

    private static Command BuildFidelityDiffCommand(Option<bool> jsonOption)
    {
        var snapshotArg = new Argument<FileInfo>("snapshot") { Description = "Snapshot created before the edit" };
        var fileArg = new Argument<FileInfo>("file") { Description = "Edited DOCX, XLSX or PPTX" };
        var outOption = new Option<FileInfo>("--out") { Description = "Output change manifest JSON", Required = true };
        var command = new Command("fidelity-diff", "Produce a machine-readable list of modified, preserved and lost Office structures");
        command.Add(snapshotArg); command.Add(fileArg); command.Add(outOption); command.Add(jsonOption);
        command.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var before = OfficePackageEvidence.ReadSnapshot(result.GetValue(snapshotArg)!.FullName);
            var manifest = OfficePackageEvidence.Diff(before, result.GetValue(fileArg)!.FullName);
            var output = result.GetValue(outOption)!.FullName;
            OfficePackageEvidence.WriteJson(output, manifest);
            var payload = new OfficeEvidenceReceipt(output, Passed: manifest.Passed, FormatRetentionRate: manifest.FormatRetentionRate, BytePreservationRate: manifest.BytePreservationRate, Modified: manifest.ModifiedParts.Count, Preserved: manifest.PreservedParts.Count, Removed: manifest.RemovedParts.Count);
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelope(JsonSerializer.Serialize(payload, OfficeEvidenceJsonContext.Default.OfficeEvidenceReceipt)));
            else Console.WriteLine($"Change manifest written: {output} (passed={manifest.Passed})");
            return manifest.Passed ? 0 : 2;
        }, json); });
        return command;
    }
}
