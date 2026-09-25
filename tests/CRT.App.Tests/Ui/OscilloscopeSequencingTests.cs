using CRT;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The oscilloscope tab's SCPI command sequencing, driven against a fake scope.
//
// TabOscilloscope.axaml.cs was the second-worst file in the codebase (3,237 lines at 7.5%),
// and the reason was structural: every sequencing path took a concrete ScopeScpiClient, which
// opens a TCP socket, so none of it could run without a scope on the network.
//
// ScopeScpiClient now implements IScopeClient and the tab's methods take the interface - the same
// arrangement IMiniproRunner already provides for the programmer, which is why Handlers/MiniPro
// is testable while the process-spawning half is not. FakeScopeClient below answers with canned
// responses and records what was asked, so the tests can assert on the exact SCPI conversation.
//
// WHAT THESE COVER. Which commands a palette sends and in what order; which of them are queries
// versus fire-and-forget writes; how a query response is parsed into the tab's cached values; and
// how a Set command turns a cached value back into SCPI text. That is the logic that decides what
// actually goes down the wire to someone's oscilloscope.
//
// WHAT THEY DO NOT. ScopeScpiClient itself - the socket, the binary-block reader, the connection
// lifecycle - stays uncovered on purpose; .claude/CLAUDE.md lists real TCP as an I/O boundary and
// the abstraction beneath it is the thing to test.
// ###########################################################################################
[Collection("HeadlessUi")]
public class OscilloscopeSequencingTests
{
    // ###########################################################################################
    // A scope that never existed: records every command it is asked to send, and answers queries
    // from a caller-supplied table (falling back to a benign "0" so a palette under test does not
    // have to stub every single query it happens to include).
    // ###########################################################################################
    // ###########################################################################################
    // *** EVERY METHOD HERE YIELDS, AND THAT FIXES A REAL INTERMITTENT FAILURE (2026-09-23). ***
    //
    // The fake used to answer with Task.CompletedTask / Task.FromResult - already-completed
    // tasks. ExecutePaletteAsync awaits this client with ConfigureAwait(false), which is CORRECT
    // for production: the scope I/O has no business capturing the UI context.
    //
    // Awaiting an ALREADY-COMPLETED task usually continues synchronously, so the rest of
    // ExecutePaletteAsync - which touches the tab - normally stayed on the dispatcher thread and
    // the test passed. "Usually" is the whole problem: the runtime is free to resume that
    // continuation elsewhere, and roughly one run in six it did, failing with
    //
    //     InvalidOperationException: The calling thread cannot access this object because a
    //     different thread owns it.
    //
    // It presented as a flake in the "HeadlessUi" collection - a different test each time, which
    // is what made it look like shared state rather than a threading contract.
    //
    // Yielding makes the behaviour DETERMINISTIC rather than hiding it: after a real await the
    // continuation resumes off the dispatcher every time, which is exactly what a real
    // ScopeScpiClient does over TCP. If production code touches a control after one of these
    // awaits, this fake now fails that test every run instead of one run in six.
    // ###########################################################################################
    private sealed class FakeScopeClient : IScopeClient
    {
        private readonly Dictionary<string, string> thisResponses;

        public FakeScopeClient(Dictionary<string, string>? responses = null)
        {
            this.thisResponses = responses ?? new Dictionary<string, string>();
        }

        // Every command sent, in order, whether or not it expected a response.
        public List<string> SentCommands { get; } = new();

        // Only the commands sent as queries - the ones the tab expects to read back from.
        public List<string> QueriedCommands { get; } = new();

        public async Task SendAsync(string commandText, CancellationToken cancellationToken)
        {
            this.SentCommands.Add(commandText);

            // Yield BEFORE returning, so the caller's continuation behaves as it would against a
            // real scope - see this class's header.
            await Task.Yield();
        }

        public async Task<string> QueryLineAsync(string commandText, CancellationToken cancellationToken)
        {
            this.SentCommands.Add(commandText);
            this.QueriedCommands.Add(commandText);

            await Task.Yield();

            return this.thisResponses.TryGetValue(commandText, out string? response) ? response : "0";
        }

        public async Task<byte[]> QueryBinaryBlockAsync(string commandText, CancellationToken cancellationToken)
        {
            this.SentCommands.Add(commandText);
            this.QueriedCommands.Add(commandText);

            await Task.Yield();

            return Array.Empty<byte>();
        }
    }

    // A scope definition with a distinct SCPI string per command, so an assertion can name exactly
    // which command was sent rather than only counting them. These are Rigol-shaped but the values
    // are arbitrary - what is under test is the tab's sequencing, not any real instrument.
    private static OscilloscopeEntry Scope() => new()
    {
        Brand = "TestBrand",
        SeriesOrModel = "TestModel",
        Port = "5555",
        Identify = "*IDN?",
        DrainErrorQueue = ":SYST:ERR?",
        OperationComplete = "*OPC?",
        ClearStatistics = ":MEAS:CLE",
        QueryActiveTrigger = ":TRIG:STAT?",
        Stop = ":STOP",
        Single = ":SING",
        Run = ":RUN",
        QueryTriggerMode = ":TRIG:MODE?",
        QueryTriggerLevel = ":TRIG:EDGE:LEV?",
        SetTriggerLevel = ":TRIG:EDGE:LEV {0}",
        QueryTimeDiv = ":TIM:SCAL?",
        SetTimeDiv = ":TIM:SCAL {0}",
        QueryVoltsDiv = ":CHAN1:SCAL?",
        SetVoltsDiv = ":CHAN1:SCAL {0}",
    };

    // -----------------------------------------------------------------------------------------
    // Which commands a palette sends
    // -----------------------------------------------------------------------------------------

    // The simplest palette: one command, sent as a query because Identify expects a response.
    [Fact]
    public async Task The_identify_palette_sends_the_identify_query()
    {
        await RunOnUiThreadAsync(async tab =>
        {
            var scope = new FakeScopeClient();

            await tab.ExecutePaletteForTestsAsync(scope, Scope(), ScopeCommandPalette.Identify);

            Assert.Equal(new[] { "*IDN?" }, scope.SentCommands);
            Assert.Equal(new[] { "*IDN?" }, scope.QueriedCommands);
        });
    }

    // A write-only command must NOT be sent as a query - waiting for a line back from a scope that
    // will never send one is how these sessions hang.
    [Fact]
    public async Task A_write_only_command_is_sent_without_expecting_a_response()
    {
        await RunOnUiThreadAsync(async tab =>
        {
            var scope = new FakeScopeClient();

            await tab.ExecutePaletteForTestsAsync(scope, Scope(), ScopeCommandPalette.Stop);

            Assert.Contains(":STOP", scope.SentCommands);
            Assert.DoesNotContain(":STOP", scope.QueriedCommands);
        });
    }

    // A multi-command palette sends its commands in the palette's declared ORDER. Order matters to
    // a scope: setting a value before the *OPC? that confirms it is not interchangeable.
    [Fact]
    public async Task A_multi_command_palette_sends_its_commands_in_order()
    {
        await RunOnUiThreadAsync(async tab =>
        {
            var scope = new FakeScopeClient();
            tab.SetLastTriggerLevelVoltsForTests(1.5);

            await tab.ExecutePaletteForTestsAsync(scope, Scope(), ScopeCommandPalette.SetTriggerLevel);

            // "1.5", not "1.5E+000": FormatScpiNumber uses G15, which only switches to
            // exponential notation for values that actually need it.
            Assert.Equal(
                new[] { ":TRIG:EDGE:LEV 1.5", "*OPC?", ":SYST:ERR?" },
                scope.SentCommands);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Reading values back
    // -----------------------------------------------------------------------------------------

    // A query response is parsed and cached - this is what the keyboard stepping and the Set
    // commands then build on.
    [Fact]
    public async Task A_trigger_level_response_is_parsed_and_cached()
    {
        await RunOnUiThreadAsync(async tab =>
        {
            var scope = new FakeScopeClient(new Dictionary<string, string>
            {
                [":TRIG:EDGE:LEV?"] = "2.5",
            });

            await tab.ExecutePaletteForTestsAsync(scope, Scope(), ScopeCommandPalette.QueryTriggerLevel);

            Assert.Equal(2.5, tab.LastTriggerLevelVoltsForTests);
        });
    }

    [Fact]
    public async Task A_time_div_response_is_parsed_and_cached()
    {
        await RunOnUiThreadAsync(async tab =>
        {
            var scope = new FakeScopeClient(new Dictionary<string, string>
            {
                [":TIM:SCAL?"] = "0.001",
            });

            await tab.ExecutePaletteForTestsAsync(scope, Scope(), ScopeCommandPalette.QueryTimeDiv);

            Assert.Equal(0.001, tab.LastTimeDivSecondsForTests);
        });
    }

    // Scopes answer in scientific notation as a matter of course, and always with a '.' decimal
    // separator regardless of the operator's locale - the parse is pinned to InvariantCulture for
    // exactly that reason. On a comma-decimal machine a culture-sensitive parse would fail here.
    [Fact]
    public async Task A_scientific_notation_response_is_parsed()
    {
        await RunOnUiThreadAsync(async tab =>
        {
            var scope = new FakeScopeClient(new Dictionary<string, string>
            {
                [":CHAN1:SCAL?"] = "5.0E-02",
            });

            await tab.ExecutePaletteForTestsAsync(scope, Scope(), ScopeCommandPalette.QueryVoltsDiv);

            Assert.Equal(0.05, tab.LastVoltsDivVoltsForTests!.Value, 10);
        });
    }

    // A scope that answers with something unparseable must leave the cache alone rather than
    // storing a garbage value that a later Set command would send straight back to it.
    [Fact]
    public async Task An_unparseable_response_leaves_the_cached_value_alone()
    {
        await RunOnUiThreadAsync(async tab =>
        {
            var scope = new FakeScopeClient(new Dictionary<string, string>
            {
                [":TRIG:EDGE:LEV?"] = "not a number",
            });

            tab.SetLastTriggerLevelVoltsForTests(null);

            await tab.ExecutePaletteForTestsAsync(scope, Scope(), ScopeCommandPalette.QueryTriggerLevel);

            Assert.Null(tab.LastTriggerLevelVoltsForTests);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Writing values back
    // -----------------------------------------------------------------------------------------

    // The cached value is substituted into the command's "{0}" placeholder, formatted as SCPI
    // expects. The exact text here IS the thing that reaches the instrument.
    [Fact]
    public async Task A_cached_value_is_formatted_into_the_set_command()
    {
        await RunOnUiThreadAsync(async tab =>
        {
            var scope = new FakeScopeClient();
            tab.SetLastTimeDivSecondsForTests(0.002);

            await tab.ExecutePaletteForTestsAsync(scope, Scope(), ScopeCommandPalette.SetTimeDiv);

            Assert.Equal(":TIM:SCAL 0.002", scope.SentCommands[0]);
        });
    }

    // With nothing cached there is no value to send, and the run is refused rather than sending a
    // malformed command (or worse, one with a literal "{0}" in it) to the scope.
    [Fact]
    public async Task A_set_command_with_no_cached_value_throws_rather_than_sending_nonsense()
    {
        await RunOnUiThreadAsync(async tab =>
        {
            var scope = new FakeScopeClient();
            tab.SetLastTriggerLevelVoltsForTests(null);

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                tab.ExecutePaletteForTestsAsync(scope, Scope(), ScopeCommandPalette.SetTriggerLevel));

            Assert.Empty(scope.SentCommands);
        });
    }

    // A scope definition missing the SCPI text for a command it is asked to run is a data problem,
    // and must fail loudly rather than sending an empty line down the socket.
    [Fact]
    public async Task A_missing_command_definition_throws()
    {
        await RunOnUiThreadAsync(async tab =>
        {
            var scope = new FakeScopeClient();
            var incompleteScope = new OscilloscopeEntry { Brand = "X", SeriesOrModel = "Y" };

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                tab.ExecutePaletteForTestsAsync(scope, incompleteScope, ScopeCommandPalette.Identify));

            Assert.Empty(scope.SentCommands);
        });
    }

    // -----------------------------------------------------------------------------------------
    // The session log
    // -----------------------------------------------------------------------------------------

    // Every command and response is logged - this is the debug output an operator reads when a
    // scope is not behaving, so it has to show both directions of the conversation.
    [Fact]
    public async Task Both_the_sent_command_and_its_response_are_logged()
    {
        await RunOnUiThreadAsync(async tab =>
        {
            var scope = new FakeScopeClient(new Dictionary<string, string>
            {
                [":TRIG:EDGE:LEV?"] = "2.5",
            });

            // Recorded rather than read off the pending-flush buffer, which the tab drains and
            // clears on a 40ms timer - see StartRecordingOutputLinesForTests.
            tab.StartRecordingOutputLinesForTests();

            await tab.ExecutePaletteForTestsAsync(scope, Scope(), ScopeCommandPalette.QueryTriggerLevel);

            var log = tab.RecordedOutputLinesForTests;

            Assert.Contains(log, line => line.Contains("SCPI >> :TRIG:EDGE:LEV?", StringComparison.Ordinal));
            Assert.Contains(log, line => line.Contains("SCPI << 2.5", StringComparison.Ordinal));
        });
    }

    // The *IDN? response carries the instrument's serial number. It is masked before being logged,
    // because these logs get pasted into bug reports.
    [Fact]
    public async Task The_identify_response_is_masked_in_the_log()
    {
        await RunOnUiThreadAsync(async tab =>
        {
            const string serial = "DS1ZA160800000";
            var scope = new FakeScopeClient(new Dictionary<string, string>
            {
                ["*IDN?"] = $"RIGOL,DS1054Z,{serial},00.04.04",
            });

            tab.StartRecordingOutputLinesForTests();

            await tab.ExecutePaletteForTestsAsync(scope, Scope(), ScopeCommandPalette.Identify);

            var log = tab.RecordedOutputLinesForTests;
            string scpiResponseLine = log.First(line => line.Contains("SCPI << ", StringComparison.Ordinal));

            Assert.DoesNotContain(serial, scpiResponseLine, StringComparison.Ordinal);
        });
    }

    // -----------------------------------------------------------------------------------------
    // Harness
    // -----------------------------------------------------------------------------------------

    // The tab must be constructed on the UI thread like any other control, and the body awaited
    // THERE rather than blocked on - see UiTest.RunAsync, which explains why blocking the
    // dispatcher thread inside UiTest.Run would deadlock anything under test that awaits a
    // dispatcher round-trip.
    private static Task RunOnUiThreadAsync(Func<TabOscilloscope, Task> body) =>
        UiTest.RunAsync(async () =>
        {
            var tab = new TabOscilloscope();
            await body(tab);
        });
}
