// Canned IMiniproRunner: the default when no T48/minipro is present, and the
// engine behind the test project and hardware-free UI demos. Recognises the
// command from its args (--logicic / --version) and returns scripted stdout/stderr
// that the real parser consumes, so the whole pipeline above the process boundary
// is exercised without hardware.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.IcTesting;

public enum MockScenario
{
    GoodChip,       // logic: pass
    FaultyChip,     // logic: one failing pin
    NoProgrammer,   // T48 not connected
    NoChip,         // socket empty / device not found
    Overcurrent,    // insertion fault
}

public sealed class MockMiniproRunner : IMiniproRunner
{
    public MockScenario Scenario { get; set; } = MockScenario.GoodChip;

    /// <summary>The pin (1-based) the FaultyChip scenario reports failing.</summary>
    public int FaultyPin { get; set; } = 3;

    public string? ResolveBinary() => "minipro (mock)";

    public Task<MiniproRunResult> RunAsync(
        IReadOnlyList<string> args, IProgress<string>? output, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();

        if (args.Contains("--version"))
            return Done(output, true, 0, "minipro 0.7 (mock)\n", "");

        switch (Scenario)
        {
            case MockScenario.NoProgrammer:
                return Done(output, true, 1, "", "Error: No programmer found.\n");
            case MockScenario.Overcurrent:
                return Done(output, true, 1, "", "Error: Overcurrent protection triggered.\n");
            case MockScenario.NoChip:
                return Done(output, true, 1, "", "Error: device not found (Chip ID: 0x000000).\n");
        }

        // Logic test (--logicic ... -T): emit a parseable grid.
        if (Scenario == MockScenario.FaultyChip)
        {
            // One failing vector: a '-' marker on FaultyPin (3-char pin fields).
            var body = new System.Text.StringBuilder("0001: ");
            for (int p = 1; p <= 14; p++)
                body.Append(p == FaultyPin ? "L- " : "0  ");
            var grid = "      " + string.Concat(System.Linq.Enumerable.Range(1, 14).Select(p => $"{p,-3}")) + "\n"
                       + body + "\n";
            return Done(output, true, 1, grid, "Logic test failed: 1 errors encountered.\n");
        }

        // GoodChip: a clean large run prints only the header; the pass is in stderr.
        var header = "      " + string.Concat(System.Linq.Enumerable.Range(1, 14).Select(p => $"{p,-3}")) + "\n";
        return Done(output, true, 0, header, "Logic test successful.\n");
    }

    private static Task<MiniproRunResult> Done(IProgress<string>? output, bool started, int code, string so, string se)
    {
        foreach (var line in so.Split('\n', StringSplitOptions.RemoveEmptyEntries))
            output?.Report(line);
        return Task.FromResult(new MiniproRunResult(started, code, so, se));
    }
}
