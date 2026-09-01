// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text.Json;
using OfficeCli.Core;
using OfficeCli.Handlers;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildComposeCommand(Option<bool> jsonOption)
    {
        var command = new Command("compose", "Compile a format-neutral professional PageSpec into differentiated Word, Excel or PowerPoint native objects");
        var file = new Argument<FileInfo>("file");
        var spec = new Argument<FileInfo>("page-spec");
        var events = new Option<bool>("--events-jsonl");
        var taskId = new Option<string?>("--task-id");
        command.Add(file); command.Add(spec); command.Add(events); command.Add(taskId); command.Add(jsonOption);
        command.SetAction(result =>
        {
            var json = result.GetValue(jsonOption);
            AgentEventStream.Configure(result.GetValue(events), result.GetValue(taskId), "officecli-compose");
            return SafeRun(() =>
            {
                var target = result.GetValue(file)!;
                var pageSpec = ProfessionalPageCompiler.Parse(result.GetValue(spec)!.FullName);
                AgentEventStream.TaskStarted($"已接收 {pageSpec.Format} 专业 PageSpec", "page_spec_validation");
                AgentEventStream.StartStage("page_spec_validation", "页面故事板与格式约束校验", pageSpec.DocumentId,
                    "正在校验页面任务、内容槽位、追溯绑定和格式差异", "format_composition", pageSpec.Pages.Count);
                using var handler = DocumentHandlerFactory.Open(target.FullName, editable: true);
                var receipt = ProfessionalPageCompiler.Compile(handler, target.FullName, pageSpec);
                AgentEventStream.Progress(pageSpec.Pages.Count, pageSpec.Pages.Count, target.Name,
                    $"{receipt.Composer} 已完成 {receipt.Pages.Count} 个页面或章节");
                AgentEventStream.StageCompleted("差异化原生编译完成", $"cp-{pageSpec.DocumentId}-composed");
                AgentEventStream.CheckpointSaved("PageSpec 与格式化产物哈希已保存", $"cp-{pageSpec.DocumentId}-composed");
                AgentEventStream.TaskValidating("正在回读页面、组件、图表和格式差异", "composition_readback");
                AgentEventStream.ArtifactReady(target.Name, "PageSpec 已编译为可编辑原生对象", $"cp-{pageSpec.DocumentId}-readback");
                AgentEventStream.TaskCompleted("专业格式编译完成");
                if (json) Console.WriteLine(JsonSerializer.Serialize(receipt, ProfessionalPageJsonContext.Default.ProfessionalCompositionReceipt));
                else Console.WriteLine($"{receipt.Composer}: {receipt.Pages.Count} pages/sections");
                return 0;
            }, json);
        });
        return command;
    }
}
