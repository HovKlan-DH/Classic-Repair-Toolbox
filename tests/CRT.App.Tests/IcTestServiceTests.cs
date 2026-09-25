using Handlers.IcTesting;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for IcTestService - catalogue entry in, typed IcTestResult out.
//
// This is what MockMiniproRunner was written for: the whole pipeline above the process
// boundary runs with NO T48 programmer and NO chip in the socket. The mock's five scenarios
// (GoodChip / FaultyChip / NoProgrammer / NoChip / Overcurrent) stand in for the hardware.
public class IcTestServiceTests
{
    private static IcTestService Service(MockScenario scenario, int faultyPin = 3) =>
        new(new MockMiniproRunner { Scenario = scenario, FaultyPin = faultyPin });

    private static IcTestEntry TestableEntry(
        string support = "testable",
        string coverage = "exhaustive",
        string kind = "logic-combinational",
        int vectorCount = 16) =>
        new()
        {
            Id = "74LS00",
            DeviceName = "74LS00",
            Support = support,
            Coverage = coverage,
            Kind = kind,
            PinCount = 14,
            Voltage = 5,
            VectorCount = vectorCount,
            Vectors = new List<string> { "LLH", "LHH", "HLH", "HHL" }
        };

    private static Task<IcTestResult> RunAsync(IcTestService service, IcTestEntry entry, IcTestMode? mode = null) =>
        service.RunAsync(entry, mode, output: null, ct: CancellationToken.None);

    // ------------------------------------------------------------------ unsupported

    [Fact]
    public async Task An_unsupported_part_is_reported_without_touching_the_programmer()
    {
        IcTestResult result = await RunAsync(
            Service(MockScenario.NoProgrammer),          // would fail if it reached the runner
            TestableEntry(support: "unsupported"));

        Assert.Equal(TestOutcome.Unsupported, result.Outcome);
        Assert.Equal(MiniproConnectionState.Ok, result.Connection);
        Assert.Contains("not vector-testable", result.Headline);
    }

    [Fact]
    public async Task A_functional_only_part_says_so_honestly()
    {
        // The wording matters: it must not imply the part was verified.
        IcTestResult result = await RunAsync(
            Service(MockScenario.GoodChip),
            TestableEntry(support: "functional-only"));

        Assert.Equal(TestOutcome.Unsupported, result.Outcome);
        Assert.Contains("Functional-only", result.Headline);
        Assert.Contains("not exhaustive", result.Headline);
    }

    // ------------------------------------------------------------------------- pass

    [Fact]
    public async Task A_good_chip_passes()
    {
        IcTestResult result = await RunAsync(Service(MockScenario.GoodChip), TestableEntry());

        Assert.Equal(TestOutcome.Pass, result.Outcome);
        Assert.Equal(MiniproConnectionState.Ok, result.Connection);
        Assert.Empty(result.FailingPins);
        Assert.StartsWith("PASS", result.Headline);
    }

    [Fact]
    public async Task A_pass_says_the_test_is_static_only()
    {
        // Honesty guard: a logic test proves the truth table, not timing. If this wording is
        // dropped, users may believe a marginal part was fully verified.
        IcTestResult result = await RunAsync(Service(MockScenario.GoodChip), TestableEntry());

        Assert.Contains("no timing tested", result.Headline);
    }

    [Fact]
    public async Task A_pass_reports_the_vector_total_from_the_entry()
    {
        IcTestResult result = await RunAsync(
            Service(MockScenario.GoodChip), TestableEntry(vectorCount: 256));

        Assert.Equal(256, result.TotalVectors);
        Assert.Contains("256", result.Headline);
    }

    [Fact]
    public async Task An_exhaustive_logic_part_is_labelled_by_input_combinations()
    {
        IcTestResult result = await RunAsync(
            Service(MockScenario.GoodChip),
            TestableEntry(kind: "logic-combinational", coverage: "exhaustive", vectorCount: 16));

        Assert.Equal("exhaustive — all 16 input combinations", result.CoverageLabel);
    }

    [Fact]
    public async Task An_exhaustive_pla_is_labelled_by_pla_vectors()
    {
        IcTestResult result = await RunAsync(
            Service(MockScenario.GoodChip),
            TestableEntry(kind: "pla", coverage: "exhaustive", vectorCount: 65536));

        Assert.Equal("exhaustive — all 65536 PLA vectors", result.CoverageLabel);
    }

    [Fact]
    public async Task A_sampled_run_is_never_labelled_exhaustive()
    {
        // The label is the user's only signal about how much was actually checked.
        IcTestResult result = await RunAsync(
            Service(MockScenario.GoodChip),
            TestableEntry(coverage: "sample", vectorCount: 25));

        Assert.Equal("sample — 25 vectors", result.CoverageLabel);
        Assert.DoesNotContain("exhaustive", result.CoverageLabel);
    }

    // ------------------------------------------------------------------------- fail

    [Fact]
    public async Task A_faulty_chip_fails_and_names_the_failing_pin()
    {
        IcTestResult result = await RunAsync(Service(MockScenario.FaultyChip, faultyPin: 3), TestableEntry());

        Assert.Equal(TestOutcome.Fail, result.Outcome);
        Assert.Equal(MiniproConnectionState.Ok, result.Connection);
        Assert.Equal(new[] { 3 }, result.FailingPins);
        Assert.Contains("pin 3", result.Headline);
    }

    [Fact]
    public async Task The_reported_failing_pin_follows_the_actual_fault()
    {
        IcTestResult result = await RunAsync(Service(MockScenario.FaultyChip, faultyPin: 7), TestableEntry());

        Assert.Equal(new[] { 7 }, result.FailingPins);
        Assert.Contains("pin 7", result.Headline);
    }

    [Fact]
    public async Task A_failure_reports_the_error_count_from_the_summary()
    {
        IcTestResult result = await RunAsync(Service(MockScenario.FaultyChip), TestableEntry());

        Assert.Equal(1, result.FailingVectors);
    }

    // ------------------------------------------------------- hardware/fixture states

    [Theory]
    [InlineData(MockScenario.NoProgrammer, MiniproConnectionState.NoProgrammer, "No programmer found")]
    [InlineData(MockScenario.NoChip, MiniproConnectionState.NoChip, "No chip detected")]
    [InlineData(MockScenario.Overcurrent, MiniproConnectionState.Overcurrent, "Overcurrent")]
    public async Task A_fixture_problem_is_surfaced_as_an_error_with_actionable_advice(
        MockScenario scenario, MiniproConnectionState expectedState, string expectedText)
    {
        IcTestResult result = await RunAsync(Service(scenario), TestableEntry());

        Assert.Equal(TestOutcome.Error, result.Outcome);
        Assert.Equal(expectedState, result.Connection);
        Assert.Contains(expectedText, result.Headline);
    }

    [Fact]
    public async Task A_fixture_problem_is_never_reported_as_a_pass_or_a_fail()
    {
        // "No chip in the socket" must not look like "this chip is faulty".
        foreach (MockScenario scenario in new[]
                 {
                     MockScenario.NoProgrammer, MockScenario.NoChip, MockScenario.Overcurrent
                 })
        {
            IcTestResult result = await RunAsync(Service(scenario), TestableEntry());

            Assert.NotEqual(TestOutcome.Pass, result.Outcome);
            Assert.NotEqual(TestOutcome.Fail, result.Outcome);
        }
    }

    [Fact]
    public async Task An_overcurrent_result_tells_the_user_to_remove_power()
    {
        IcTestResult result = await RunAsync(Service(MockScenario.Overcurrent), TestableEntry());

        Assert.Contains("Remove power", result.Headline);
    }

    [Fact]
    public async Task A_runner_that_cannot_start_is_reported_as_not_installed()
    {
        var service = new IcTestService(new NotStartingRunner());

        IcTestResult result = await RunAsync(service, TestableEntry());

        Assert.Equal(TestOutcome.Error, result.Outcome);
        Assert.Equal(MiniproConnectionState.NotInstalled, result.Connection);
        Assert.Contains("could not be launched", result.Headline);
    }

    // ------------------------------------------------------------------------ modes

    [Fact]
    public async Task A_selected_mode_overrides_the_entrys_coverage_and_vector_count()
    {
        // The PLA offers quick/standard/full depths; the label must follow the chosen one.
        var mode = new IcTestMode
        {
            Name = "quick",
            Coverage = "sample",
            VectorCount = 25,
            Vectors = new List<string> { "LLH", "HHL" }
        };

        IcTestResult result = await RunAsync(
            Service(MockScenario.GoodChip),
            TestableEntry(coverage: "exhaustive", vectorCount: 65536),
            mode);

        Assert.Equal("sample — 25 vectors", result.CoverageLabel);
        Assert.Equal(25, result.TotalVectors);
    }

    // ------------------------------------------------------------------- plumbing

    [Fact]
    public async Task The_service_passes_the_expected_arguments_to_minipro()
    {
        var recording = new RecordingRunner();
        var service = new IcTestService(recording);

        await RunAsync(service, TestableEntry());

        Assert.NotNull(recording.LastArgs);
        Assert.Equal("-p", recording.LastArgs![0]);
        Assert.Equal("74LS00", recording.LastArgs[1]);
        Assert.Contains("--logicic", recording.LastArgs);
        Assert.Contains("-T", recording.LastArgs);
    }

    [Fact]
    public async Task The_generated_logicic_xml_is_deleted_after_the_run()
    {
        var recording = new RecordingRunner();
        var service = new IcTestService(recording);

        await RunAsync(service, TestableEntry());

        string xmlPath = recording.LastArgs![recording.LastArgs.ToList().IndexOf("--logicic") + 1];
        Assert.False(File.Exists(xmlPath), "the temporary logicic XML should be cleaned up");
    }

    [Fact]
    public async Task The_generated_logicic_xml_describes_the_part_and_its_vectors()
    {
        var capturing = new XmlCapturingRunner();
        var service = new IcTestService(capturing);

        await RunAsync(service, TestableEntry());

        Assert.Contains("<logicic>", capturing.Xml);
        Assert.Contains("name=\"74LS00\"", capturing.Xml);
        Assert.Contains("pins=\"14\"", capturing.Xml);
        Assert.Contains("voltage=\"5V\"", capturing.Xml);
        Assert.Contains("<vector id=\"0\">", capturing.Xml);
        Assert.Contains("L L H", capturing.Xml);   // one char per pin, space separated
    }

    [Fact]
    public async Task A_cancelled_run_does_not_report_a_pass()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            Service(MockScenario.GoodChip).RunAsync(TestableEntry(), null, null, cts.Token));
    }

    // ------------------------------------------------------------------ test doubles

    private sealed class NotStartingRunner : IMiniproRunner
    {
        public string? ResolveBinary() => null;

        public Task<MiniproRunResult> RunAsync(
            IReadOnlyList<string> args, IProgress<string>? output, CancellationToken ct) =>
            Task.FromResult(MiniproRunResult.NotStarted("minipro not found"));
    }

    private sealed class RecordingRunner : IMiniproRunner
    {
        public IReadOnlyList<string>? LastArgs { get; private set; }

        public string? ResolveBinary() => "minipro (recording)";

        public Task<MiniproRunResult> RunAsync(
            IReadOnlyList<string> args, IProgress<string>? output, CancellationToken ct)
        {
            this.LastArgs = args;
            return Task.FromResult(new MiniproRunResult(true, 0, "", "Logic test successful.\n"));
        }
    }

    private sealed class XmlCapturingRunner : IMiniproRunner
    {
        public string Xml { get; private set; } = "";

        public string? ResolveBinary() => "minipro (capturing)";

        public Task<MiniproRunResult> RunAsync(
            IReadOnlyList<string> args, IProgress<string>? output, CancellationToken ct)
        {
            int index = args.ToList().IndexOf("--logicic");
            if (index >= 0 && index + 1 < args.Count && File.Exists(args[index + 1]))
            {
                this.Xml = File.ReadAllText(args[index + 1]);
            }

            return Task.FromResult(new MiniproRunResult(true, 0, "", "Logic test successful.\n"));
        }
    }
}
