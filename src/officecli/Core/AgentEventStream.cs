// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>Shared Office Agent JSONL progress protocol on stderr.</summary>
internal static class AgentEventStream
{
    private static readonly object Gate = new();
    private static bool _enabled;
    private static string _taskId = "";
    private static string _operation = "";
    private static string _stageId = "";
    private static string _stageName = "";
    private static string _artifact = "";
    private static string _summary = "";
    private static string _nextStage = "";
    private static int _completed;
    private static int _total;
    private static long _sequence;
    private static DateTimeOffset _startedAt;
    private static Timer? _heartbeat;

    internal static void Configure(bool requested, string? taskId, string operation)
    {
        lock (Gate)
        {
            _heartbeat?.Dispose();
            _heartbeat = null;
            _enabled = requested || EnvironmentEnabled("OFFICE_AGENT_EVENTS");
            _taskId = string.IsNullOrWhiteSpace(taskId) ? $"{operation}-{Guid.NewGuid():N}" : taskId.Trim();
            _operation = operation;
            _stageId = "";
            _stageName = "";
            _artifact = "";
            _summary = "";
            _nextStage = "";
            _completed = 0;
            _total = 0;
            _startedAt = DateTimeOffset.UtcNow;
            _sequence = 0;
        }
    }

    internal static void TaskStarted(string summary, string nextStage)
        => Emit("task.started", "running", summary: summary, nextStage: nextStage);

    internal static void TaskResumed(string summary, string checkpointId, string nextStage)
        => Emit("task.resumed", "running", summary: summary, checkpointId: checkpointId, nextStage: nextStage);

    internal static void StartStage(string stageId, string stageName, string artifact, string summary, string nextStage, int total)
    {
        lock (Gate)
        {
            if (!_enabled) return;
            _stageId = stageId;
            _stageName = stageName;
            _artifact = artifact;
            _summary = summary;
            _nextStage = nextStage;
            _completed = 0;
            _total = total;
            _heartbeat?.Dispose();
            var interval = HeartbeatInterval();
            _heartbeat = new Timer(_ => Heartbeat(), null, interval, interval);
        }
        Emit("stage.started", "running", completed: 0, total: total, artifact: artifact, summary: summary, nextStage: nextStage);
    }

    internal static void Progress(int completed, int total, string artifact, string summary)
    {
        string nextStage;
        lock (Gate)
        {
            if (!_enabled) return;
            _completed = completed;
            _total = total;
            _artifact = artifact;
            _summary = summary;
            nextStage = _nextStage;
        }
        Emit("stage.progress", "running", completed: completed, total: total, artifact: artifact, summary: summary, nextStage: nextStage);
    }

    internal static void StageCompleted(string summary, string checkpointId)
    {
        StopHeartbeat();
        Emit("stage.completed", "completed", completed: _total, total: _total, artifact: _artifact, summary: summary, checkpointId: checkpointId, nextStage: _nextStage);
    }

    internal static void ArtifactReady(string artifact, string summary, string checkpointId)
        => Emit("artifact.ready", "ready", artifact: artifact, summary: summary, checkpointId: checkpointId, nextStage: _nextStage);

    internal static void ReviewRequired(string summary, string checkpointId, IReadOnlyList<string>? riskCodes = null)
        => Emit("review.required", "paused", artifact: _artifact, summary: summary, checkpointId: checkpointId,
            nextStage: _nextStage, needsUserReview: true, riskCodes: riskCodes);

    internal static void Warning(string summary, IReadOnlyList<string>? riskCodes = null)
        => Emit("warning.detected", "warning", artifact: _artifact, summary: summary, nextStage: _nextStage, riskCodes: riskCodes);

    internal static void RetryStarted(string summary, int attempt, int maximumAttempts)
        => Emit("retry.started", "running", completed: attempt - 1, total: maximumAttempts, artifact: _artifact,
            summary: summary, nextStage: _nextStage, attempt: attempt);

    internal static void CheckpointSaved(string summary, string checkpointId)
        => Emit("checkpoint.saved", "saved", artifact: _artifact, summary: summary, checkpointId: checkpointId, nextStage: _nextStage);

    internal static void TaskValidating(string summary, string nextStage)
        => Emit("task.validating", "running", artifact: _artifact, summary: summary, nextStage: nextStage);

    internal static void TaskCompleted(string summary)
    {
        StopHeartbeat();
        Emit("task.completed", "completed", summary: summary);
    }

    internal static void TaskFailed(string summary)
    {
        StopHeartbeat();
        Emit("task.failed", "failed", artifact: _artifact, summary: summary);
    }

    private static void Heartbeat()
    {
        int completed;
        int total;
        string artifact;
        string summary;
        string nextStage;
        lock (Gate)
        {
            if (!_enabled) return;
            completed = _completed;
            total = _total;
            artifact = _artifact;
            summary = _summary;
            nextStage = _nextStage;
        }
        Emit("stage.progress", "running", completed: completed, total: total, artifact: artifact, summary: summary, nextStage: nextStage);
    }

    private static void StopHeartbeat()
    {
        lock (Gate)
        {
            _heartbeat?.Dispose();
            _heartbeat = null;
        }
    }

    private static bool EnvironmentEnabled(string name)
        => Environment.GetEnvironmentVariable(name)?.Trim().ToLowerInvariant() is "1" or "true" or "yes" or "on";

    private static TimeSpan HeartbeatInterval()
    {
        var value = Environment.GetEnvironmentVariable("OFFICE_AGENT_HEARTBEAT_SECONDS");
        return int.TryParse(value, out var seconds) && seconds is >= 15 and <= 30
            ? TimeSpan.FromSeconds(seconds)
            : TimeSpan.FromSeconds(20);
    }

    private static void Emit(
        string eventName, string status, int? completed = null, int? total = null,
        string? artifact = null, string? summary = null, string? checkpointId = null,
        string? nextStage = null, bool needsUserReview = false, IReadOnlyList<string>? riskCodes = null, int? attempt = null)
    {
        lock (Gate)
        {
            if (!_enabled) return;
            using var buffer = new MemoryStream();
            using (var json = new Utf8JsonWriter(buffer))
            {
                json.WriteStartObject();
                json.WriteNumber("protocolVersion", 1);
                json.WriteString("event", eventName);
                json.WriteString("taskId", _taskId);
                json.WriteString("operation", _operation);
                if (!string.IsNullOrEmpty(_stageId)) json.WriteString("stageId", _stageId);
                if (!string.IsNullOrEmpty(_stageName)) json.WriteString("stageName", _stageName);
                json.WriteString("status", status);
                if (completed.HasValue) json.WriteNumber("completedItems", completed.Value);
                if (total.HasValue) json.WriteNumber("totalItems", total.Value);
                if (!string.IsNullOrEmpty(artifact)) json.WriteString("currentArtifact", artifact);
                if (!string.IsNullOrEmpty(summary)) json.WriteString("summary", summary);
                if (!string.IsNullOrEmpty(checkpointId)) json.WriteString("checkpointId", checkpointId);
                json.WriteNumber("elapsedSeconds", Math.Round((DateTimeOffset.UtcNow - _startedAt).TotalSeconds, 3));
                if (!string.IsNullOrEmpty(nextStage)) json.WriteString("nextStage", nextStage);
                json.WriteBoolean("needsUserReview", needsUserReview);
                if (riskCodes is { Count: > 0 })
                {
                    json.WriteStartArray("riskCodes");
                    foreach (var code in riskCodes) json.WriteStringValue(code);
                    json.WriteEndArray();
                }
                if (attempt.HasValue) json.WriteNumber("attempt", attempt.Value);
                json.WriteNumber("sequence", ++_sequence);
                json.WriteString("timestamp", DateTimeOffset.UtcNow.ToString("O"));
                json.WriteEndObject();
            }
            Console.Error.WriteLine(System.Text.Encoding.UTF8.GetString(buffer.ToArray()));
        }
    }
}
