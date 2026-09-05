// Orchestrates one IC test: catalogue entry -> minipro invocation -> typed result.
// Hardware-agnostic (drives any IMiniproRunner, real or mock). Logic/PLA tests feed
// our OWN exhaustive vectors to `minipro --logicic` so coverage is verified, not
// minipro's sample.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.IcTesting;

public sealed class IcTestService
{
    private readonly IMiniproRunner _runner;

    public IcTestService(IMiniproRunner runner) => _runner = runner;

    public async Task<IcTestResult> RunAsync(IcTestEntry entry, IcTestMode? mode, IProgress<string>? output, CancellationToken ct)
    {
        if (!entry.IsTestable)
        {
            return new IcTestResult
            {
                Outcome = TestOutcome.Unsupported,
                Connection = MiniproConnectionState.Ok,
                CoverageLabel = entry.Coverage,
                Headline = entry.IsFunctionalOnly
                    ? "Functional-only part — vector testing is a functional check, not exhaustive; no automated test is run."
                    : "This part is not vector-testable.",
            };
        }

        return await RunLogicAsync(entry, mode, output, ct).ConfigureAwait(false);
    }

    // ----- logic / PLA -----

    private async Task<IcTestResult> RunLogicAsync(IcTestEntry entry, IcTestMode? mode, IProgress<string>? output, CancellationToken ct)
    {
        string xml = PrepareLogicicXml(entry, mode);
        try
        {
            var device = string.IsNullOrWhiteSpace(entry.DeviceName) ? entry.Id : entry.DeviceName;
            var args = new List<string> { "-p", device, "--logicic", xml, "-T" };
            var run = await _runner.RunAsync(args, output, ct).ConfigureAwait(false);
            if (!run.Started)
                return IcTestResult.Connectionless(MiniproConnectionState.NotInstalled,
                    "minipro could not be launched — is it installed/bundled?", run.StandardError);

            var p = MiniproOutputParser.Parse(run.StandardOutput, run.StandardError);
            int total = mode is { VectorCount: > 0 } ? mode.VectorCount
                : entry.VectorCount > 0 ? entry.VectorCount : (entry.Vectors?.Count ?? 0);
            string cov = mode?.Coverage ?? entry.Coverage;
            string coverage = cov == "exhaustive"
                ? (entry.Kind == "pla" ? $"exhaustive — all {total} PLA vectors"
                                       : $"exhaustive — all {total} input combinations")
                : $"{cov} — {total} vectors";

            if (p.Passed == true)
                return new IcTestResult
                {
                    Outcome = TestOutcome.Pass,
                    Connection = MiniproConnectionState.Ok,
                    CoverageLabel = coverage,
                    TotalVectors = total,
                    Headline = $"PASS — All {total} vectors matched.\nLogic verified (static test only — no timing tested).",
                    RawOutput = Raw(run),
                };

            if (p.Passed == false)
                return new IcTestResult
                {
                    Outcome = TestOutcome.Fail,
                    Connection = MiniproConnectionState.Ok,
                    CoverageLabel = coverage,
                    TotalVectors = total,
                    FailingVectors = p.ErrorCount ?? p.Failures.Count,
                    FailingPins = p.FailingPins,
                    Headline = p.FailingPins.Count > 0
                        ? $"FAIL — Logic fault on pin {string.Join(", ", p.FailingPins)}."
                        : "FAIL — the chip did not match the truth table.",
                    RawOutput = Raw(run),
                };

            // No pass/fail summary -> the test never ran; classify why.
            return IcTestResult.Connectionless(p.State,
                StateMessage(p.State, "the logic test did not run"), Raw(run));
        }
        finally
        {
            TryDelete(xml);
        }
    }

    /// <summary>Produce the minipro logicic XML for an entry (decompress the PLA's bundled
    /// file, or build one from the entry's inline vectors) and return its temp path.</summary>
    private static string PrepareLogicicXml(IcTestEntry entry, IcTestMode? mode)
    {
        Directory.CreateDirectory(TempDir);
        var dest = Path.Combine(TempDir, $"crt_{entry.Id}_{Guid.NewGuid():N}.logicic.xml");

        string? vectorsFile = mode?.VectorsFile ?? entry.VectorsFile;
        bool compressed = mode is not null ? mode.VectorsCompressed : entry.VectorsCompressed;
        if (!string.IsNullOrEmpty(vectorsFile))
        {
            var root = Directory.GetParent(IcTestCatalogue.VectorsDir)!.FullName;
            var src = Path.Combine(root, vectorsFile);
            if (compressed)
            {
                using var raw = File.OpenRead(src);
                using var gz = new GZipStream(raw, CompressionMode.Decompress);
                using var outFile = File.Create(dest);
                gz.CopyTo(outFile);
            }
            else
            {
                File.Copy(src, dest, overwrite: true);
            }
            return dest;
        }

        File.WriteAllText(dest, BuildLogicicXml(entry, mode), Encoding.UTF8);
        return dest;
    }

    private static string BuildLogicicXml(IcTestEntry entry, IcTestMode? mode)
    {
        var name = string.IsNullOrWhiteSpace(entry.DeviceName) ? entry.Id : entry.DeviceName;
        int volts = entry.Voltage > 0 ? entry.Voltage : 5;
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<logicic>\n");
        sb.Append("  <database type=\"LOGIC\" device=\"LOGIC\">\n    <manufacturer name=\"Logic Ic\">\n");
        sb.Append($"      <ic name=\"{name}\" pins=\"{entry.PinCount}\" voltage=\"{volts}V\" type=\"5\">\n");
        var vectors = mode?.Vectors ?? entry.Vectors ?? new List<string>();
        for (int i = 0; i < vectors.Count; i++)
            sb.Append($"        <vector id=\"{i}\"> {string.Join(' ', vectors[i].ToCharArray())} </vector>\n");
        sb.Append("      </ic>\n    </manufacturer>\n  </database>\n</logicic>\n");
        return sb.ToString();
    }

    // ----- helpers -----

    private static string StateMessage(MiniproConnectionState s, string fallback) => s switch
    {
        MiniproConnectionState.NoProgrammer => "No programmer found — connect the T48 and try again.",
        MiniproConnectionState.NoChip => "No chip detected — seat the chip in the ZIF socket (pin 1 aligned).",
        MiniproConnectionState.Overcurrent => "Overcurrent — check the chip is the right way round and correctly seated. Remove power.",
        MiniproConnectionState.NotInstalled => "minipro is not installed/bundled.",
        _ => fallback + ".",
    };

    private static string Raw(MiniproRunResult r) =>
        (r.StandardOutput + (string.IsNullOrEmpty(r.StandardError) ? "" : "\n" + r.StandardError)).Trim();

    private static string TempDir => Path.Combine(Path.GetTempPath(), "crt-ic-test");

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* best effort */ }
    }
}
