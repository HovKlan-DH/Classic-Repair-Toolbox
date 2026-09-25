// IC-test domain model — the catalogue entries CRT loads, and the typed results
// the runner/service produce. Deliberately UI-free (no Avalonia) so it is unit-
// testable without a display and reusable from the (future) test project.
//
// Used by: IcTestCatalogue (loads entries), IcTestService (orchestrates a test),
// MiniproOutputParser (fills the result), and the Schematics Test UI (renders it).

using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace Handlers.IcTesting;

// ----- catalogue (mirrors trl-sharedlibs build_catalogue.py output; case-insensitive) -----

public sealed class IcTestEntry
{
    public int Schema { get; init; }
    public string Id { get; init; } = "";
    public List<string> Match { get; init; } = new();
    public string Description { get; init; } = "";
    public string Kind { get; init; } = "";       // pla | logic-combinational | logic-sequential
    public string Support { get; init; } = "";     // testable | functional-only | unsupported
    public string Coverage { get; init; } = "";     // exhaustive | functional | sample

    // logic / pla
    public int Inputs { get; init; }
    public int PinCount { get; init; }
    public int Voltage { get; init; }
    public string? DeviceName { get; init; }
    public List<string>? Vectors { get; init; }            // inline (small logic parts)
    public string? VectorsFile { get; init; }              // external (the PLA)
    public bool VectorsCompressed { get; init; }
    public int VectorCount { get; init; }
    public List<IcTestMode>? Modes { get; init; }   // selectable test depths (e.g. the PLA: quick/standard/full)

    public string? Note { get; init; }

    [JsonIgnore] public bool IsTestable => string.Equals(Support, "testable", System.StringComparison.OrdinalIgnoreCase);
    [JsonIgnore] public bool IsFunctionalOnly => string.Equals(Support, "functional-only", System.StringComparison.OrdinalIgnoreCase);
    [JsonIgnore] public bool HasModes => Modes is { Count: > 0 };
}

// A selectable test depth for a part (e.g. the PLA: Quick 25 / Standard 512 / Full 65536).
public sealed class IcTestMode
{
    public string Name { get; init; } = "";
    public string Label { get; init; } = "";
    public string Coverage { get; init; } = "";      // exhaustive | sample
    public int VectorCount { get; init; }
    public List<string>? Vectors { get; init; }       // inline (small)
    public string? VectorsFile { get; init; }         // external (gzipped logicic XML)
    public bool VectorsCompressed { get; init; }
}

// ----- results -----

public enum TestOutcome { NotRun, Pass, Fail, Unsupported, Error }

// Connection / fixture states surfaced as first-class results (design §12).
public enum MiniproConnectionState { Ok, NoProgrammer, NoChip, ChipReversed, Overcurrent, NotInstalled, Unknown }

public sealed class IcTestResult
{
    public TestOutcome Outcome { get; init; }
    public MiniproConnectionState Connection { get; init; } = MiniproConnectionState.Unknown;
    public string CoverageLabel { get; init; } = "";        // honest label, e.g. "exhaustive (65536 vectors)"
    public string Headline { get; init; } = "";
    public IReadOnlyList<int> FailingPins { get; init; } = System.Array.Empty<int>();
    public int FailingVectors { get; init; }
    public int TotalVectors { get; init; }
    public string RawOutput { get; init; } = "";

    public static IcTestResult Connectionless(MiniproConnectionState state, string headline, string raw = "") =>
        new() { Outcome = TestOutcome.Error, Connection = state, Headline = headline, RawOutput = raw };
}
