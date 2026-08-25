// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;
using OfficeCli.Core;
using OfficeCli.Core.Diagram;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildSmartArtCommand(Option<bool> jsonOption)
    {
        var command = new Command("smartart", "Read and deeply update native Office SmartArt through DiagramSpec");

        var inspectFileArg = new Argument<FileInfo>("file") { Description = "DOCX, XLSX or PPTX containing SmartArt" };
        var inspectOutOption = new Option<FileInfo?>("--out") { Description = "Optional structured inspection JSON output" };
        var inspect = new Command("inspect", "Read SmartArt nodes, relationships and owning document parts")
        {
            inspectFileArg, inspectOutOption, jsonOption,
        };
        inspect.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var inspection = NativeSmartArtCodec.Inspect(result.GetValue(inspectFileArg)!.FullName);
            var serialized = JsonSerializer.Serialize(inspection, DiagramJsonContext.Default.SmartArtPackageInspection);
            if (result.GetValue(inspectOutOption)?.FullName is { Length: > 0 } output)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(output) ?? ".");
                File.WriteAllText(output, serialized, new System.Text.UTF8Encoding(false));
            }
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelope(serialized));
            else Console.WriteLine($"SmartArt diagrams: {inspection.DiagramCount}");
            return 0;
        }, json); });
        command.Add(inspect);

        var updateFileArg = new Argument<FileInfo>("file") { Description = "DOCX, XLSX or PPTX to update in place" };
        var dataPartArg = new Argument<string>("data-part") { Description = "SmartArt data part listed by inspect, e.g. /ppt/diagrams/data1.xml" };
        var specOption = new Option<FileInfo>("--spec") { Description = "Validated DiagramSpec v1 JSON", Required = true };
        var update = new Command("update", "Replace one SmartArt data model while preserving its native layout, theme and host frame")
        {
            updateFileArg, dataPartArg, specOption, jsonOption,
        };
        update.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var spec = DiagramSpec.Load(result.GetValue(specOption)!.FullName);
            var receipt = NativeSmartArtCodec.Update(result.GetValue(updateFileArg)!.FullName,
                result.GetValue(dataPartArg)!, spec);
            var serialized = JsonSerializer.Serialize(receipt, DiagramJsonContext.Default.SmartArtUpdateReceipt);
            if (json) Console.WriteLine(OutputFormatter.WrapEnvelope(serialized));
            else Console.WriteLine($"SmartArt updated: {receipt.DataPart} ({receipt.NodeCount} nodes)");
            return 0;
        }, json); });
        command.Add(update);

        return command;
    }
}
