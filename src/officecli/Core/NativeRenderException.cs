// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Preserves the native Office automation stage that failed. Callers may
/// distinguish an unavailable application from a document-open or export
/// failure instead of reporting every COM exception as "Office not installed".
/// </summary>
internal sealed class NativeRenderException : Exception
{
    public string ApplicationName { get; }
    public string Stage { get; }

    public NativeRenderException(string applicationName, string stage, Exception inner)
        : base($"{applicationName} native render failed during {stage}: {inner.Message}", inner)
    {
        ApplicationName = applicationName;
        Stage = stage;
        HResult = inner.HResult;
    }
}
