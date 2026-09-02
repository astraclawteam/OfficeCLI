// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildChartSpecCommand(Option<bool> jsonOption)
    {
        var root = new Command("chart-spec", "Select and produce information-design charts from a traceable ChartSpec");
        var list = new Command("list", "List the eight first-phase information chart types");
        list.Add(jsonOption);
        list.SetAction(result => SafeRun(() =>
        {
            var charts = InformationChartEngine.List();
            if (result.GetValue(jsonOption)) Console.WriteLine(JsonSerializer.Serialize(
                new InformationChartListResponse(true, charts.Count, charts), InformationChartJsonContext.Default.InformationChartListResponse));
            else foreach (var chart in charts) Console.WriteLine($"{chart.ChartType}\t{chart.SemanticIntent}\tfallback={chart.FallbackRepresentation}");
            return 0;
        }, result.GetValue(jsonOption)));
        root.Add(list);

        var apply = new Command("apply", "Insert a native editable chart or a truthful native fallback");
        var file = new Argument<FileInfo>("file");
        var spec = new Argument<FileInfo>("chart-spec");
        var events = new Option<bool>("--events-jsonl");
        var taskId = new Option<string?>("--task-id");
        apply.Add(file); apply.Add(spec); apply.Add(events); apply.Add(taskId); apply.Add(jsonOption);
        apply.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            AgentEventStream.Configure(result.GetValue(events), result.GetValue(taskId), "officecli-chart-spec");
            return SafeRun(() =>
            {
                var target = result.GetValue(file)!;
                var chart = InformationChartEngine.Parse(result.GetValue(spec)!.FullName);
                AgentEventStream.TaskStarted($"已接收信息设计图表 {chart.ChartType}", "chart_selection");
                AgentEventStream.StartStage("chart_selection", "图表语义选择", chart.ChartId,
                    "正在校验 claim、数据结构、轴策略和降级条件", "chart_composition", chart.Items.Count);
                ResidentClient.SendClose(target.FullName);
                using var handler = DocumentHandlerFactory.Open(target.FullName, editable: true);
                var receipt = InformationChartEngine.Apply(handler, target.FullName, chart);
                AgentEventStream.Progress(chart.Items.Count, chart.Items.Count, receipt.NativeObjectPath,
                    receipt.FallbackReason is null ? "已生成原生可编辑图表" : $"已按真实数据条件降级：{receipt.FallbackReason}");
                AgentEventStream.StageCompleted("图表构图与结构回读完成", $"cp-{chart.ChartId}");
                AgentEventStream.CheckpointSaved("图表语义、数据和对象路径已保存", $"cp-{chart.ChartId}");
                AgentEventStream.TaskValidating("正在回读图表标题、分类、系列、单位和追溯绑定", "chart_readback");
                AgentEventStream.ArtifactReady(receipt.NativeObjectPath, "信息设计图表已就绪", $"cp-{chart.ChartId}");
                AgentEventStream.TaskCompleted("ChartSpec 生产完成");
                if (json) Console.WriteLine(JsonSerializer.Serialize(receipt, InformationChartJsonContext.Default.InformationChartReceipt));
                else Console.WriteLine($"{receipt.Representation}: {receipt.NativeObjectPath}");
                return 0;
            }, json);
        });
        root.Add(apply);

        var read = new Command("read", "Read native chart structure from an Office file");
        var readFile = new Argument<FileInfo>("file");
        read.Add(readFile); read.Add(jsonOption);
        read.SetAction(result => SafeRun(() =>
        {
            var target = result.GetValue(readFile)!;
            ResidentClient.SendSave(target.FullName);
            using var handler = DocumentHandlerFactory.Open(target.FullName, editable: false);
            var charts = InformationChartEngine.Read(handler);
            if (result.GetValue(jsonOption)) Console.WriteLine(JsonSerializer.Serialize(
                new InformationChartReadResponse(true, charts.Count, charts), InformationChartJsonContext.Default.InformationChartReadResponse));
            else foreach (var chart in charts) Console.WriteLine($"{chart.NativeObjectPath}\t{chart.Title}\t{chart.RequestedChartType}");
            return charts.Count > 0 ? 0 : 2;
        }, result.GetValue(jsonOption)));
        root.Add(read);
        return root;
    }
}
