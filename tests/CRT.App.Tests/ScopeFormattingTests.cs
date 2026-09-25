using System.Globalization;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for ScopeFormatting - the number, voltage, time and string formatting
// the oscilloscope panel does before text reaches the output log, a window title or a file name.
//
// This logic used to live as private members of TabOscilloscope, where no test could reach it.
// These lock in CURRENT behaviour; if one fails after a change, decide deliberately whether the
// behaviour change is intended before editing the expectation.
public class ScopeFormattingTests
{
    // -------------------------------------------------------------- FormatScpiNumber

    [Theory]
    [InlineData(1.0, "1")]
    [InlineData(0.5, "0.5")]
    [InlineData(-2.25, "-2.25")]
    [InlineData(1000.0, "1000")]
    public void FormatScpiNumber_writes_plain_decimals_without_an_exponent(double value, string expected)
    {
        Assert.Equal(expected, ScopeFormatting.FormatScpiNumber(value));
    }

    // Some scope firmware rejects a capital "E" in an exponent, so the formatter lower-cases it.
    // This is the whole reason the method exists rather than calling ToString("G15") at the site.
    [Fact]
    public void FormatScpiNumber_lower_cases_the_exponent_marker()
    {
        string formatted = ScopeFormatting.FormatScpiNumber(0.0000001);

        Assert.Contains("e", formatted, StringComparison.Ordinal);
        Assert.DoesNotContain("E", formatted, StringComparison.Ordinal);
    }

    // Guards against a machine with a comma decimal separator emitting "0,5" into an SCPI command.
    [Fact]
    public void FormatScpiNumber_is_invariant_of_the_current_culture()
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("da-DK");
            Assert.Equal("0.5", ScopeFormatting.FormatScpiNumber(0.5));
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }

    // -------------------------------------------------------------- FormatVoltage

    [Theory]
    [InlineData(1.0, "1V")]
    [InlineData(2.5, "2.5V")]
    [InlineData(-3.0, "-3V")]
    [InlineData(0.5, "500mV")]
    [InlineData(0.001, "1mV")]
    [InlineData(0.0005, "500uV")]
    public void FormatVoltage_picks_the_unit_from_the_magnitude(double volts, string expected)
    {
        Assert.Equal(expected, ScopeFormatting.FormatVoltage(volts));
    }

    // The thresholds compare the ABSOLUTE value, so a negative volt-level scale still reads "mV"
    // rather than falling through to the smallest unit.
    [Fact]
    public void FormatVoltage_selects_the_unit_from_the_absolute_value_for_negatives()
    {
        Assert.Equal("-500mV", ScopeFormatting.FormatVoltage(-0.5));
        Assert.Equal("-500uV", ScopeFormatting.FormatVoltage(-0.0005));
    }

    // Zero is below every threshold, so it falls all the way through to microvolts.
    [Fact]
    public void FormatVoltage_renders_zero_as_microvolts()
    {
        Assert.Equal("0uV", ScopeFormatting.FormatVoltage(0.0));
    }

    // "0.###" truncates to three decimals, so very fine values lose precision by design - the
    // output panel is for a human reading a value, not for round-tripping it back to the scope.
    [Fact]
    public void FormatVoltage_rounds_to_three_decimal_places()
    {
        Assert.Equal("1.235V", ScopeFormatting.FormatVoltage(1.23456));
    }

    // -------------------------------------------------------------- FormatTime

    [Theory]
    [InlineData(1.0, "1S")]
    [InlineData(2.5, "2.5S")]
    [InlineData(0.5, "500mS")]
    [InlineData(0.001, "1mS")]
    [InlineData(0.0005, "500uS")]
    [InlineData(0.000001, "1uS")]
    [InlineData(0.0000005, "500nS")]
    public void FormatTime_picks_the_unit_from_the_magnitude(double seconds, string expected)
    {
        Assert.Equal(expected, ScopeFormatting.FormatTime(seconds));
    }

    // Deliberate quirk: time uses a CAPITAL "S" for seconds ("500mS", not "500ms"). It reads
    // oddly against SI but matches what the scope panel has always shown.
    [Fact]
    public void FormatTime_capitalises_the_second_marker()
    {
        Assert.EndsWith("S", ScopeFormatting.FormatTime(0.5), StringComparison.Ordinal);
        Assert.Equal("500mS", ScopeFormatting.FormatTime(0.5));
    }

    // Nanoseconds is the last branch, so anything smaller collapses toward "0nS" rather than
    // gaining a picosecond unit.
    [Fact]
    public void FormatTime_bottoms_out_at_nanoseconds()
    {
        Assert.Equal("0nS", ScopeFormatting.FormatTime(1e-15));
    }

    // -------------------------------------------------------------- GetNextSnappedTriggerLevelVolts

    // A level already sitting exactly on the 0.25V grid moves one whole step.
    [Theory]
    [InlineData(0.0, 1, 0.25)]
    [InlineData(0.25, 1, 0.5)]
    [InlineData(0.0, -1, -0.25)]
    [InlineData(-0.5, -1, -0.75)]
    public void A_trigger_level_on_the_grid_steps_a_whole_division(
        double current, int direction, double expected)
    {
        Assert.Equal(expected, ScopeFormatting.GetNextSnappedTriggerLevelVolts(current, direction), precision: 10);
    }

    // A level BETWEEN grid lines snaps outward to the next line rather than jumping a full step
    // past it - so the first keypress after reading an arbitrary level from the scope lands on
    // the grid instead of overshooting.
    [Theory]
    [InlineData(0.30, 1, 0.50)]
    [InlineData(0.30, -1, 0.25)]
    [InlineData(-0.30, 1, -0.25)]
    [InlineData(-0.30, -1, -0.50)]
    public void A_trigger_level_between_grid_lines_snaps_outward_to_the_next_line(
        double current, int direction, double expected)
    {
        Assert.Equal(expected, ScopeFormatting.GetNextSnappedTriggerLevelVolts(current, direction), precision: 10);
    }

    // A level a hair off the grid (floating-point noise from the scope) counts as ON the grid,
    // otherwise a reported 0.2499999 would only creep back to 0.25 instead of advancing.
    [Fact]
    public void A_trigger_level_within_tolerance_of_the_grid_is_treated_as_on_it()
    {
        Assert.Equal(0.5, ScopeFormatting.GetNextSnappedTriggerLevelVolts(0.25 + 1e-9, 1), precision: 10);
    }

    // Direction 0 is not "no change": only > 0 steps up, so zero falls into the down branch.
    [Fact]
    public void A_zero_direction_steps_down_rather_than_holding()
    {
        Assert.Equal(0.0, ScopeFormatting.GetNextSnappedTriggerLevelVolts(0.25, 0), precision: 10);
    }

    // -------------------------------------------------------------- GetMainWindowTitleBase

    [Theory]
    [InlineData("Classic Repair Toolbox (oscilloscope connected)", "Classic Repair Toolbox")]
    [InlineData("Classic Repair Toolbox (oscilloscope disconnected)", "Classic Repair Toolbox")]
    [InlineData("Classic Repair Toolbox", "Classic Repair Toolbox")]
    [InlineData("", "")]
    public void GetMainWindowTitleBase_strips_a_connection_suffix_once(string title, string expected)
    {
        Assert.Equal(expected, ScopeFormatting.GetMainWindowTitleBase(title));
    }

    // Only ONE suffix is removed per call. The caller strips before re-appending, so a title can
    // never legitimately carry two - but if it somehow did, the inner one survives.
    [Fact]
    public void GetMainWindowTitleBase_removes_only_the_outermost_suffix()
    {
        Assert.Equal(
            "CRT (oscilloscope connected)",
            ScopeFormatting.GetMainWindowTitleBase("CRT (oscilloscope connected) (oscilloscope connected)"));
    }

    // The comparison is ordinal, so a differently-cased suffix is left alone rather than stripped.
    [Fact]
    public void GetMainWindowTitleBase_is_case_sensitive()
    {
        const string title = "CRT (Oscilloscope Connected)";
        Assert.Equal(title, ScopeFormatting.GetMainWindowTitleBase(title));
    }

    // -------------------------------------------------------------- BuildOscilloscopeWindowTitle

    [Theory]
    [InlineData(true, "CRT (oscilloscope connected)")]
    [InlineData(false, "CRT (oscilloscope disconnected)")]
    public void A_reported_session_state_appends_the_matching_suffix(bool hasEstablishedSession, string expected)
    {
        Assert.Equal(
            expected,
            ScopeFormatting.BuildOscilloscopeWindowTitle("CRT", true, true, hasEstablishedSession));
    }

    // Nothing worth reporting yet: no session has existed and (for the main window) auto-connect is
    // off, so the title stays clean rather than announcing a scope the user never asked about.
    [Fact]
    public void Nothing_to_report_leaves_the_base_title_alone()
    {
        Assert.Equal("CRT", ScopeFormatting.BuildOscilloscopeWindowTitle("CRT", true, false, false));
        Assert.Equal("CRT", ScopeFormatting.BuildOscilloscopeWindowTitle("CRT", true, false, true));
    }

    // The tab switch beats everything, INCLUDING a session that is still established. Hiding the
    // oscilloscope tab tears the session down and stops auto-connect, so a window still claiming
    // "(oscilloscope disconnected)" would be reporting on a feature the user has switched off.
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void A_disabled_oscilloscope_tab_suppresses_the_suffix_whatever_the_session_state(
        bool shouldReportSessionState,
        bool hasEstablishedSession)
    {
        Assert.Equal(
            "CRT",
            ScopeFormatting.BuildOscilloscopeWindowTitle(
                "CRT",
                isOscilloscopeTabEnabled: false,
                shouldReportSessionState,
                hasEstablishedSession));
    }

    // Round-trips with GetMainWindowTitleBase, which is what stops suffixes stacking up: the tab
    // strips before it rebuilds, so the pair must agree on the exact suffix text.
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void A_built_title_is_stripped_back_to_its_base(bool hasEstablishedSession)
    {
        string built = ScopeFormatting.BuildOscilloscopeWindowTitle("CRT", true, true, hasEstablishedSession);

        Assert.Equal("CRT", ScopeFormatting.GetMainWindowTitleBase(built));
    }

    // -------------------------------------------------------------- SanitizeCapturedOscilloscopeImageFileNamePart

    [Fact]
    public void A_blank_file_name_part_becomes_Unknown_so_the_name_never_collapses()
    {
        Assert.Equal("Unknown", ScopeFormatting.SanitizeCapturedOscilloscopeImageFileNamePart(""));
        Assert.Equal("Unknown", ScopeFormatting.SanitizeCapturedOscilloscopeImageFileNamePart("   "));
    }

    [Fact]
    public void A_file_name_part_is_trimmed_and_kept_when_already_valid()
    {
        Assert.Equal("U1", ScopeFormatting.SanitizeCapturedOscilloscopeImageFileNamePart("  U1  "));
    }

    // Every character the platform rejects becomes an underscore, so a board label like "A/B"
    // cannot escape into a path separator.
    [Fact]
    public void Invalid_file_name_characters_are_replaced_with_underscores()
    {
        string sanitized = ScopeFormatting.SanitizeCapturedOscilloscopeImageFileNamePart("A/B:C*D?");

        foreach (char invalid in Path.GetInvalidFileNameChars())
        {
            Assert.DoesNotContain(invalid.ToString(), sanitized, StringComparison.Ordinal);
        }

        Assert.StartsWith("A_B", sanitized, StringComparison.Ordinal);
    }

    // -------------------------------------------------------------- MaskScopeSerial

    // The mask preserves LENGTH rather than emitting a fixed token, so a truncated or empty
    // response still looks visibly different in the log.
    [Theory]
    [InlineData("ABC123", "******")]
    [InlineData("X", "*")]
    [InlineData("", "")]
    public void MaskScopeSerial_replaces_every_character_but_keeps_the_length(string serial, string expected)
    {
        Assert.Equal(expected, ScopeFormatting.MaskScopeSerial(serial));
    }

    // -------------------------------------------------------------- MaskIdentifyResponseSerial

    // A *IDN? response is "brand,model,serial,firmware". Only field 3 is masked so the rest stays
    // readable for debugging.
    [Fact]
    public void Only_the_serial_field_of_an_identify_response_is_masked()
    {
        Assert.Equal(
            "RIGOL,DS1054Z,**********,00.04.04",
            ScopeFormatting.MaskIdentifyResponseSerial("RIGOL,DS1054Z,DS1ZA12345 ,00.04.04"));
    }

    // A response too short to carry a serial is passed through untouched rather than throwing.
    [Theory]
    [InlineData("RIGOL,DS1054Z")]
    [InlineData("RIGOL")]
    [InlineData("")]
    public void An_identify_response_without_a_serial_field_is_left_alone(string response)
    {
        Assert.Equal(response, ScopeFormatting.MaskIdentifyResponseSerial(response));
    }

    // The serial is trimmed BEFORE masking, so surrounding spaces do not inflate the star count.
    // That is why the mask above is 10 stars for "DS1ZA12345 " rather than 11.
    [Fact]
    public void The_serial_is_trimmed_before_the_mask_length_is_chosen()
    {
        string masked = ScopeFormatting.MaskIdentifyResponseSerial("A,B,  1234  ,D");

        Assert.Equal("A,B,****,D", masked);
    }

    [Fact]
    public void A_null_identify_response_is_treated_as_empty()
    {
        Assert.Equal(string.Empty, ScopeFormatting.MaskIdentifyResponseSerial(null!));
    }

    // -------------------------------------------------------------- NormalizeScopeOverlayValue

    // The rule is positional, not unit-aware: uppercase the last letter, lowercase the one before
    // it. That turns "1US" into "1uS" and "5MV" into "5mV".
    [Theory]
    [InlineData("1US", "1uS")]
    [InlineData("5MV", "5mV")]
    [InlineData("1.5v", "1.5V")]
    [InlineData("100NS", "100nS")]
    public void NormalizeScopeOverlayValue_uppercases_the_unit_and_lowercases_its_prefix(
        string input, string expected)
    {
        Assert.Equal(expected, ScopeFormatting.NormalizeScopeOverlayValue(input));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_blank_overlay_value_normalises_to_empty(string? input)
    {
        Assert.Equal(string.Empty, ScopeFormatting.NormalizeScopeOverlayValue(input));
    }

    [Fact]
    public void An_overlay_value_is_trimmed_before_normalising()
    {
        Assert.Equal("1uS", ScopeFormatting.NormalizeScopeOverlayValue("  1US  "));
    }

    // Quirk worth knowing: the two rules are INDEPENDENT. The second-last character is
    // lower-cased whether or not the last one turned out to be a unit letter, so a value ending
    // in a digit still gets its preceding letter changed - "US1" becomes "Us1", not "US1".
    // That is nonsense as a unit, but these values are display-only and never parsed back.
    [Fact]
    public void An_overlay_value_ending_in_a_digit_still_has_its_preceding_letter_lowercased()
    {
        Assert.Equal("Us1", ScopeFormatting.NormalizeScopeOverlayValue("US1"));
    }

    // A single character has no "second last" position to lowercase.
    [Fact]
    public void A_single_character_overlay_value_is_only_uppercased()
    {
        Assert.Equal("V", ScopeFormatting.NormalizeScopeOverlayValue("v"));
    }
}
