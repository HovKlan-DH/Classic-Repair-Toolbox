// The thin "drive the minipro binary" seam. The real runner shells out; the mock
// returns canned output so the UI and the test project work with NO hardware and
// NO minipro installed. Everything above this (catalogue, parsing, orchestration)
// is hardware-agnostic and unit-testable.

using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.IcTesting;

public sealed record MiniproRunResult(bool Started, int ExitCode, string StandardOutput, string StandardError)
{
    public static MiniproRunResult NotStarted(string why) => new(false, -1, "", why);
}

public interface IMiniproRunner
{
    /// <summary>The resolved minipro binary path/name, or null if none can be found.</summary>
    string? ResolveBinary();

    /// <summary>
    /// Run minipro with <paramref name="args"/>, streaming stdout lines to
    /// <paramref name="output"/>. <see cref="MiniproRunResult.Started"/> is false when the
    /// binary could not be launched (e.g. not installed) — distinct from a non-zero exit.
    /// </summary>
    Task<MiniproRunResult> RunAsync(IReadOnlyList<string> args, IProgress<string>? output, CancellationToken ct);
}
