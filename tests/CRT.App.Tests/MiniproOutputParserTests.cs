using Handlers.IcTesting;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for MiniproOutputParser - the text that comes back from
// minipro.exe after a logic test. No T48 programmer and no chip required: these feed the
// parser the same strings the real process writes.
//
// The subtle rule this file exists to protect: a clean pass on a large run prints NO
// per-vector lines, so "no grid + a successful summary" IS a pass. Anyone optimising the
// grid parsing needs that to stay true.
public class MiniproOutputParserTests
{
    private const string Esc = "";

    // ------------------------------------------------------------------------ StripAnsi

    [Fact]
    public void StripAnsi_removes_colour_codes()
    {
        string coloured = Esc + "[32mLogic test successful" + Esc + "[0m";
        Assert.Equal("Logic test successful", MiniproOutputParser.StripAnsi(coloured));
    }

    [Fact]
    public void StripAnsi_leaves_bracket_text_that_is_not_an_escape_sequence()
    {
        // The ESC byte is required by the regex precisely so legitimate content survives.
        const string text = "Chip [74LS00] ready 1[0m";
        Assert.Equal(text, MiniproOutputParser.StripAnsi(text));
    }

    [Fact]
    public void StripAnsi_treats_null_as_empty()
    {
        Assert.Equal("", MiniproOutputParser.StripAnsi(null));
    }

    // ------------------------------------------------------------- AlignVectorTableHeader

    [Fact]
    public void AlignVectorTableHeader_rebuilds_the_header_to_match_the_data_columns()
    {
        // minipro writes pin numbers at their natural width, so pins 10+ drift out of line
        // with their 3-char data column. The header is rebuilt from the data row's layout.
        string input = string.Join('\n',
            "  1 2 3 10 11",
            "   1: H  H  L  H  L ");

        string aligned = MiniproOutputParser.AlignVectorTableHeader(input);
        string[] lines = aligned.Split('\n');

        // The data row is untouched...
        Assert.Equal("   1: H  H  L  H  L ", lines[1]);

        // ...and the header now starts at the same column the data body starts at.
        int bodyStart = lines[1].IndexOf("H", StringComparison.Ordinal);
        Assert.Equal(bodyStart, lines[0].IndexOf("1", StringComparison.Ordinal));

        // Every pin number still appears, in order.
        Assert.Contains("1", lines[0]);
        Assert.Contains("10", lines[0]);
        Assert.Contains("11", lines[0]);
    }

    [Fact]
    public void AlignVectorTableHeader_returns_the_text_unchanged_when_no_grid_was_printed()
    {
        // A clean pass on a large run prints no per-vector lines at all.
        const string text = "Found T48\nLogic test successful";
        Assert.Equal(text, MiniproOutputParser.AlignVectorTableHeader(text));
    }

    [Fact]
    public void AlignVectorTableHeader_returns_the_text_unchanged_when_the_header_is_unrecognisable()
    {
        string input = string.Join('\n',
            "pin pin pin",
            "   1: H  H  L ");

        Assert.Equal(input, MiniproOutputParser.AlignVectorTableHeader(input));
    }

    [Fact]
    public void AlignVectorTableHeader_treats_null_as_empty()
    {
        Assert.Equal("", MiniproOutputParser.AlignVectorTableHeader(null));
    }

    // ---------------------------------------------------------------------------- Parse

    [Fact]
    public void Parse_treats_a_successful_summary_with_no_grid_as_a_pass()
    {
        // THE case this parser exists for. Do not "fix" this to require a grid.
        var result = MiniproOutputParser.Parse(stdout: "", stderr: "Logic test successful");

        Assert.True(result.Passed);
        Assert.Equal(0, result.VectorsSeen);
        Assert.Empty(result.Failures);
        Assert.Empty(result.FailingPins);
        Assert.Equal(MiniproConnectionState.Ok, result.State);
        Assert.Null(result.ErrorCount);
    }

    [Fact]
    public void Parse_reads_the_error_count_out_of_a_failure_summary()
    {
        var result = MiniproOutputParser.Parse(stdout: "", stderr: "Logic test failed: 7 errors");

        Assert.False(result.Passed);
        Assert.Equal(7, result.ErrorCount);
    }

    [Fact]
    public void Parse_handles_a_failure_summary_with_no_count()
    {
        var result = MiniproOutputParser.Parse(stdout: "", stderr: "Logic test failed");

        Assert.False(result.Passed);
        Assert.Null(result.ErrorCount);
    }

    [Fact]
    public void Parse_is_case_insensitive_about_the_summary_text()
    {
        Assert.True(MiniproOutputParser.Parse("", "LOGIC TEST SUCCESSFUL").Passed);
        Assert.False(MiniproOutputParser.Parse("", "logic test failed: 2").Passed);
    }

    [Fact]
    public void Parse_counts_vectors_and_locates_failing_pins_from_the_grid()
    {
        // Each pin field is 3 chars: symbol, marker, separator. A '-' marker means that pin
        // failed on that vector. Here pin 2 fails on vector 1, pin 3 on vector 2.
        string stdout = string.Join('\n',
            "   1: H  L- L  ",
            "   2: H  H  L- ");

        var result = MiniproOutputParser.Parse(stdout, stderr: "");

        Assert.Equal(2, result.VectorsSeen);
        Assert.Equal(new[] { (1, 2), (2, 3) }, result.Failures);
        Assert.Equal(new[] { 2, 3 }, result.FailingPins);
    }

    [Fact]
    public void Parse_reports_failing_pins_sorted_and_deduplicated()
    {
        string stdout = string.Join('\n',
            "   1: H  H  L- ",
            "   2: H- H  L- ",
            "   3: H  H  L- ");

        var result = MiniproOutputParser.Parse(stdout, stderr: "");

        Assert.Equal(new[] { 1, 3 }, result.FailingPins);   // sorted, each pin once
        Assert.Equal(4, result.Failures.Count);             // but every occurrence kept
    }

    [Fact]
    public void Parse_fails_when_the_grid_shows_errors_even_if_the_summary_is_missing()
    {
        var result = MiniproOutputParser.Parse(stdout: "   1: H  L- L  ", stderr: "");

        Assert.False(result.Passed);
    }

    [Fact]
    public void Parse_fails_when_the_grid_shows_errors_and_the_summary_wrongly_says_success()
    {
        // The grid is the more specific evidence, so it wins.
        var result = MiniproOutputParser.Parse(stdout: "   1: H  L- L  ", stderr: "Logic test successful");

        Assert.False(result.Passed);
    }

    [Fact]
    public void Parse_returns_an_unknown_result_when_there_is_nothing_to_go_on()
    {
        var result = MiniproOutputParser.Parse(stdout: "", stderr: "");

        Assert.Null(result.Passed);
        Assert.Equal(MiniproConnectionState.Unknown, result.State);
    }

    [Fact]
    public void Parse_strips_ansi_colour_before_reading_the_grid()
    {
        string stdout = "   1: " + Esc + "[31mH  L- L  " + Esc + "[0m";

        var result = MiniproOutputParser.Parse(stdout, stderr: "");

        Assert.Equal(1, result.VectorsSeen);
        Assert.Equal(new[] { 2 }, result.FailingPins);
    }

    [Fact]
    public void Parse_tolerates_windows_line_endings()
    {
        string stdout = "   1: H  L- L  \r\n   2: H  H  L  \r\n";

        var result = MiniproOutputParser.Parse(stdout, stderr: "");

        Assert.Equal(2, result.VectorsSeen);
        Assert.Equal(new[] { 2 }, result.FailingPins);
    }

    [Fact]
    public void Parse_treats_nulls_as_empty()
    {
        var result = MiniproOutputParser.Parse(null, null);

        Assert.Null(result.Passed);
        Assert.Equal(0, result.VectorsSeen);
    }

    // -------------------------------------------------------------------- ClassifyState

    [Theory]
    [InlineData("No programmer found", MiniproConnectionState.NoProgrammer)]
    [InlineData("programmer not found", MiniproConnectionState.NoProgrammer)]
    [InlineData("Overcurrent detected!", MiniproConnectionState.Overcurrent)]
    [InlineData("over-current on VCC", MiniproConnectionState.Overcurrent)]
    [InlineData("over current", MiniproConnectionState.Overcurrent)]
    [InlineData("Device not found", MiniproConnectionState.NoChip)]
    [InlineData("Chip ID 0x000000", MiniproConnectionState.NoChip)]
    [InlineData("no device", MiniproConnectionState.NoChip)]
    public void ClassifyState_recognises_each_hardware_condition(string stderr, MiniproConnectionState expected)
    {
        Assert.Equal(expected, MiniproOutputParser.ClassifyState(stderr, passed: null));
    }

    [Fact]
    public void ClassifyState_is_case_insensitive()
    {
        Assert.Equal(
            MiniproConnectionState.NoProgrammer,
            MiniproOutputParser.ClassifyState("NO PROGRAMMER", passed: null));
    }

    [Fact]
    public void ClassifyState_reports_Ok_when_a_verdict_was_reached_and_no_fault_was_seen()
    {
        Assert.Equal(MiniproConnectionState.Ok, MiniproOutputParser.ClassifyState("", passed: true));
        Assert.Equal(MiniproConnectionState.Ok, MiniproOutputParser.ClassifyState("", passed: false));
    }

    [Fact]
    public void ClassifyState_reports_Unknown_when_no_verdict_was_reached()
    {
        Assert.Equal(MiniproConnectionState.Unknown, MiniproOutputParser.ClassifyState("", passed: null));
    }

    [Fact]
    public void ClassifyState_prefers_a_hardware_fault_over_a_verdict()
    {
        Assert.Equal(
            MiniproConnectionState.NoProgrammer,
            MiniproOutputParser.ClassifyState("no programmer", passed: true));
    }
}
