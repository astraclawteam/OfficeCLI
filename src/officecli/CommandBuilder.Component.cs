// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildComponentCommand(Option<bool> jsonOption)
    {
        var root = new Command("component", "List, inspect, insert, update and read professional native Office components");

        var list = new Command("list", "List the 12 semantic professional components");
        list.Add(jsonOption);
        list.SetAction(result => SafeRun(() =>
        {
            var json = result.GetValue(jsonOption);
            var values = ProfessionalComponentCatalog.List();
            if (json) Console.WriteLine(JsonSerializer.Serialize(
                new ProfessionalComponentListResponse(true, values.Count, values),
                ProfessionalComponentJsonContext.Default.ProfessionalComponentListResponse));
            else foreach (var item in values) Console.WriteLine($"{item.ComponentId}\t{item.Category}\t{item.SemanticIntent}");
            return 0;
        }, result.GetValue(jsonOption)));
        root.Add(list);

        var componentId = new Argument<string>("component-id") { Description = "Component identifier returned by component list" };
        var describe = new Command("describe", "Show semantic intent, slots, unsupported conditions and adaptive rules");
        describe.Add(componentId);
        describe.Add(jsonOption);
        describe.SetAction(result => SafeRun(() =>
        {
            var definition = ProfessionalComponentCatalog.Get(result.GetValue(componentId)!);
            if (result.GetValue(jsonOption)) Console.WriteLine(JsonSerializer.Serialize(
                new ProfessionalComponentDescribeResponse(true, definition),
                ProfessionalComponentJsonContext.Default.ProfessionalComponentDescribeResponse));
            else Console.WriteLine(JsonSerializer.Serialize(
                new ProfessionalComponentDescribeResponse(true, definition),
                ProfessionalComponentJsonContext.Default.ProfessionalComponentDescribeResponse));
            return 0;
        }, result.GetValue(jsonOption)));
        root.Add(describe);

        root.Add(BuildComponentMutation("insert", "Insert one editable native component", jsonOption, update: false));
        root.Add(BuildComponentMutation("update", "Replace one component instance without rebuilding the containing page or sheet", jsonOption, update: true));

        var read = new Command("read", "Read component identities, bindings and native object paths from an Office file");
        var readFile = new Argument<FileInfo>("file");
        var instance = new Option<string?>("--instance-id") { Description = "Optional component instance filter" };
        read.Add(readFile); read.Add(instance); read.Add(jsonOption);
        read.SetAction(result => SafeRun(() =>
        {
            var file = result.GetValue(readFile)!;
            ResidentClient.SendSave(file.FullName);
            using var handler = DocumentHandlerFactory.Open(file.FullName, editable: false);
            var items = ProfessionalComponentCatalog.Read(handler, file.FullName, result.GetValue(instance));
            if (result.GetValue(jsonOption)) Console.WriteLine(JsonSerializer.Serialize(
                new ProfessionalComponentReadResponse(true, items.Count, items),
                ProfessionalComponentJsonContext.Default.ProfessionalComponentReadResponse));
            else foreach (var item in items) Console.WriteLine($"{item.InstanceId}\t{item.ComponentId}\t{item.NativeObjectPath}");
            return items.Count > 0 ? 0 : 2;
        }, result.GetValue(jsonOption)));
        root.Add(read);
        return root;
    }

    private static Command BuildComponentMutation(string name, string description, Option<bool> jsonOption, bool update)
    {
        var command = new Command(name, description);
        var file = new Argument<FileInfo>("file");
        var spec = new Argument<FileInfo>("component-spec");
        var events = new Option<bool>("--events-jsonl") { Description = "Emit structured Agent progress events on stderr" };
        var taskId = new Option<string?>("--task-id");
        command.Add(file); command.Add(spec); command.Add(events); command.Add(taskId); command.Add(jsonOption);
        command.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            AgentEventStream.Configure(result.GetValue(events), result.GetValue(taskId), $"officecli-component-{name}");
            return SafeRun(() =>
            {
                var target = result.GetValue(file)!;
                var component = ProfessionalComponentCatalog.Parse(result.GetValue(spec)!.FullName);
                AgentEventStream.TaskStarted($"已接收专业组件 {component.ComponentId}", "component_validation");
                AgentEventStream.StartStage("component_validation", "组件语义与绑定校验", component.InstanceId,
                    "正在校验组件适用条件、必需槽位和追溯绑定", "component_composition", 1);
                _ = ProfessionalComponentCatalog.Get(component.ComponentId);
                AgentEventStream.Progress(1, 1, component.InstanceId, "组件协议校验完成");
                AgentEventStream.StageCompleted("组件协议校验完成", $"cp-{component.InstanceId}-validated");
                AgentEventStream.CheckpointSaved("组件语义与绑定已保存", $"cp-{component.InstanceId}-validated");
                AgentEventStream.StartStage("component_composition", "专业原生组件构图", target.Name,
                    "正在依据数据量、宿主格式和密度选择原生构图", "component_readback", component.Items.Count);
                ResidentClient.SendClose(target.FullName);
                using var handler = DocumentHandlerFactory.Open(target.FullName, editable: true);
                var receipt = ProfessionalComponentCatalog.Apply(handler, target.FullName, component, update);
                AgentEventStream.Progress(component.Items.Count, component.Items.Count, receipt.NativeObjectPath,
                    $"已形成 {receipt.ItemCount} 项可编辑业务对象");
                AgentEventStream.StageCompleted("原生组件已写入", $"cp-{component.InstanceId}-composed");
                AgentEventStream.TaskValidating("正在回读组件身份、绑定和原生对象路径", "component_readback");
                AgentEventStream.ArtifactReady(receipt.NativeObjectPath, "组件已完成结构回读绑定", $"cp-{component.InstanceId}-readback");
                AgentEventStream.TaskCompleted("专业原生组件生产完成");
                if (json) Console.WriteLine(JsonSerializer.Serialize(receipt, ProfessionalComponentJsonContext.Default.ProfessionalComponentReceipt));
                else Console.WriteLine($"{receipt.Operation}: {receipt.NativeObjectPath}");
                return 0;
            }, json);
        });
        return command;
    }
}
