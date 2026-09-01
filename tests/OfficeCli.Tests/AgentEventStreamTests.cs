using System.Text.Json;
using OfficeCli.Core;
using Xunit;

namespace OfficeCli.Tests;

public sealed class AgentEventStreamTests
{
    [Fact]
    public void JsonlProtocolReportsRealCountsWithoutPercentages()
    {
        var original = Console.Error;
        using var output = new StringWriter();
        Console.SetError(output);
        try
        {
            AgentEventStream.Configure(true, "task-test", "officecli-batch");
            AgentEventStream.TaskStarted("开始", "batch_execution");
            AgentEventStream.StartStage("batch_execution", "Office 原生对象批处理", "report.docx", "执行三个操作", "artifact_validation", 3);
            AgentEventStream.Progress(2, 3, "set", "完成两个操作");
            AgentEventStream.StageCompleted("三个操作完成", "cp-officecli-batch");
            AgentEventStream.TaskCompleted("完成");
        }
        finally
        {
            Console.SetError(original);
        }

        var lines = output.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(5, lines.Length);
        using var progress = JsonDocument.Parse(lines[2]);
        Assert.Equal("stage.progress", progress.RootElement.GetProperty("event").GetString());
        Assert.Equal(2, progress.RootElement.GetProperty("completedItems").GetInt32());
        Assert.Equal(3, progress.RootElement.GetProperty("totalItems").GetInt32());
        Assert.False(progress.RootElement.TryGetProperty("percent", out _));
    }

    [Fact]
    public void ProtocolCanRepresentResumeReviewWarningRetryCheckpointAndValidation()
    {
        var original = Console.Error;
        using var output = new StringWriter();
        Console.SetError(output);
        try
        {
            AgentEventStream.Configure(true, "task-review", "officecli-compose");
            AgentEventStream.TaskResumed("从内容稿继续", "cp-professional-content", "visual_direction");
            AgentEventStream.StartStage("visual_direction", "视觉创意方向", "art-direction", "形成视觉方向", "storyboard", 1);
            AgentEventStream.Warning("品牌来源仍需确认", ["brand-uncertain"]);
            AgentEventStream.RetryStarted("重新选择对比度", 2, 3);
            AgentEventStream.CheckpointSaved("视觉方向已保存", "cp-visual-direction");
            AgentEventStream.ReviewRequired("请审阅视觉方向", "cp-visual-direction", ["brand-uncertain"]);
            AgentEventStream.TaskValidating("正在执行结构回读", "final_delivery");
        }
        finally { Console.SetError(original); }

        var events = output.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)
            .Select(line => JsonDocument.Parse(line).RootElement.GetProperty("event").GetString()).ToArray();
        Assert.Equal(new[] { "task.resumed", "stage.started", "warning.detected", "retry.started", "checkpoint.saved", "review.required", "task.validating" }, events);
        using var review = JsonDocument.Parse(output.ToString().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries)[5]);
        Assert.True(review.RootElement.GetProperty("needsUserReview").GetBoolean());
        Assert.Equal("brand-uncertain", review.RootElement.GetProperty("riskCodes")[0].GetString());
    }
}
