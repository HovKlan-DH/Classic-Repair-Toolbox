using System.Globalization;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for ScopeValueMapper - the translation between the engineering
// values written in board Excel files ("5ms", "500mV") and the SCPI numbers sent to a
// real oscilloscope. Nothing here needs a scope: every function under test is pure.
//
// These lock in CURRENT behaviour. If one fails after a change, decide deliberately
// whether the behaviour change is intended before editing the expectation.
public class ScopeValueMapperTests
{
    // -------------------------------------------------------------- TryParseTimeValue

    [Theory]
    [InlineData("1s", 1.0)]
    [InlineData("5ms", 5e-3)]
    [InlineData("100us", 100e-6)]
    [InlineData("5ns", 5e-9)]
    [InlineData("2.5ms", 2.5e-3)]
    [InlineData("0.5s", 0.5)]
    public void TryParseTimeValue_converts_each_supported_unit_to_seconds(string text, double expected)
    {
        Assert.True(ScopeValueMapper.TryParseTimeValue(text, out double seconds));
        Assert.Equal(expected, seconds, precision: 15);
    }

    [Theory]
    [InlineData("5US")]      // upper case
    [InlineData("5uS")]      // mixed case
    [InlineData("5µs")] // micro sign U+00B5
    [InlineData("  5us  ")]  // surrounding whitespace
    [InlineData("5 us")]     // space between number and unit
    public void TryParseTimeValue_normalises_unit_spelling_and_whitespace(string text)
    {
        Assert.True(ScopeValueMapper.TryParseTimeValue(text, out double seconds));
        Assert.Equal(5e-6, seconds, precision: 15);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("banana")]
    [InlineData("5")]        // no unit at all
    [InlineData("ms")]       // no number at all
    [InlineData("5kg")]      // unrecognised unit
    [InlineData("5V")]       // a voltage is not a time
    public void TryParseTimeValue_rejects_unparseable_text(string text)
    {
        Assert.False(ScopeValueMapper.TryParseTimeValue(text, out double seconds));
        Assert.Equal(0, seconds);
    }

    [Fact]
    public void TryParseTimeValue_does_not_understand_exponent_notation()
    {
        // The numeric scan stops at the first character that is not a digit/./-/+, so "e-7"
        // becomes the unit and fails to match. Documented so nobody "fixes" a board file to
        // use 5e-7 expecting it to work.
        Assert.False(ScopeValueMapper.TryParseTimeValue("5e-7s", out _));
    }

    [Fact]
    public void TryParseTimeValue_accepts_a_negative_value()
    {
        Assert.True(ScopeValueMapper.TryParseTimeValue("-5ms", out double seconds));
        Assert.Equal(-5e-3, seconds, precision: 15);
    }

    // ----------------------------------------------------------- TryParseVoltageValue

    [Theory]
    [InlineData("1V", 1.0)]
    [InlineData("500mV", 0.5)]
    [InlineData("0.5V", 0.5)]
    [InlineData("2mV", 2e-3)]
    [InlineData("20uV", 20e-6)]
    [InlineData("20µV", 20e-6)]
    public void TryParseVoltageValue_converts_each_supported_unit_to_volts(string text, double expected)
    {
        Assert.True(ScopeValueMapper.TryParseVoltageValue(text, out double volts));
        Assert.Equal(expected, volts, precision: 15);
    }

    [Theory]
    [InlineData("5ms")]  // a time is not a voltage
    [InlineData("5A")]
    [InlineData("")]
    public void TryParseVoltageValue_rejects_unparseable_text(string text)
    {
        Assert.False(ScopeValueMapper.TryParseVoltageValue(text, out _));
    }

    // ------------------------------------------------------------------ TryMapTimeDiv

    private static OscilloscopeEntry Scope(string timeDivList = "", string voltsDivList = "") =>
        new() { TimeDivList = timeDivList, VoltsDivList = voltsDivList };

    [Fact]
    public void TryMapTimeDiv_matches_a_supported_value_and_reports_the_list_spelling()
    {
        var component = new ComponentImageEntry { TimeDiv = "5ms" };
        var scope = Scope(timeDivList: "1ms, 2ms, 5ms, 10ms");

        Assert.True(ScopeValueMapper.TryMapTimeDiv(component, scope, out ScopeMappedValue mapped));

        Assert.Equal("5ms", mapped.RawValue);
        Assert.Equal("5ms", mapped.MatchedDisplayValue);
        Assert.Equal(5e-3, mapped.NumericValue, precision: 15);
        Assert.Equal("0.005", mapped.ScpiValue);
    }

    [Fact]
    public void TryMapTimeDiv_matches_across_different_unit_spellings()
    {
        // The board file says 1000us; the scope list says 1ms. They are the same timebase.
        var component = new ComponentImageEntry { TimeDiv = "1000us" };
        var scope = Scope(timeDivList: "500us, 1ms, 2ms");

        Assert.True(ScopeValueMapper.TryMapTimeDiv(component, scope, out ScopeMappedValue mapped));

        Assert.Equal("1000us", mapped.RawValue);
        Assert.Equal("1ms", mapped.MatchedDisplayValue);   // the scope's own spelling wins
    }

    [Fact]
    public void TryMapTimeDiv_does_not_match_an_adjacent_nanosecond_value()
    {
        // Regression guard: the tolerance is relative, not a flat 1e-9. A flat tolerance made
        // 1ns and 2ns compare equal - see the comment on AreEquivalent.
        var component = new ComponentImageEntry { TimeDiv = "1ns" };
        var scope = Scope(timeDivList: "2ns, 5ns, 10ns");

        Assert.False(ScopeValueMapper.TryMapTimeDiv(component, scope, out _));
    }

    [Fact]
    public void TryMapTimeDiv_returns_false_when_the_value_is_not_supported_by_the_scope()
    {
        var component = new ComponentImageEntry { TimeDiv = "3ms" };
        var scope = Scope(timeDivList: "1ms, 2ms, 5ms");

        Assert.False(ScopeValueMapper.TryMapTimeDiv(component, scope, out _));
    }

    [Theory]
    [InlineData("", "1ms, 2ms")]   // component has no value
    [InlineData("5ms", "")]        // scope has no supported list
    [InlineData("banana", "1ms")]  // component value is unparseable
    public void TryMapTimeDiv_returns_false_on_missing_or_unparseable_input(string timeDiv, string list)
    {
        var component = new ComponentImageEntry { TimeDiv = timeDiv };
        Assert.False(ScopeValueMapper.TryMapTimeDiv(component, Scope(timeDivList: list), out _));
    }

    [Fact]
    public void TryMapTimeDiv_skips_malformed_entries_in_the_scope_list()
    {
        var component = new ComponentImageEntry { TimeDiv = "5ms" };
        var scope = Scope(timeDivList: "1ms, , nonsense, 5ms");

        Assert.True(ScopeValueMapper.TryMapTimeDiv(component, scope, out ScopeMappedValue mapped));
        Assert.Equal("5ms", mapped.MatchedDisplayValue);
    }

    // ----------------------------------------------------------------- TryMapVoltsDiv

    [Fact]
    public void TryMapVoltsDiv_matches_a_supported_value()
    {
        var component = new ComponentImageEntry { VoltsDiv = "500mV" };
        var scope = Scope(voltsDivList: "100mV, 200mV, 500mV, 1V");

        Assert.True(ScopeValueMapper.TryMapVoltsDiv(component, scope, out ScopeMappedValue mapped));

        Assert.Equal("500mV", mapped.MatchedDisplayValue);
        Assert.Equal(0.5, mapped.NumericValue, precision: 15);
        Assert.Equal("0.5", mapped.ScpiValue);
    }

    // -------------------------------------------------------------- TryMapTriggerLevel

    [Fact]
    public void TryMapTriggerLevel_does_not_need_a_supported_list()
    {
        // Trigger level is a free numeric value, unlike T/DIV and V/DIV which must snap
        // to a value the scope actually offers.
        var component = new ComponentImageEntry { TriggerLevelVolts = "1.65V" };

        Assert.True(ScopeValueMapper.TryMapTriggerLevel(component, out ScopeMappedValue mapped));

        Assert.Equal("1.65V", mapped.RawValue);
        Assert.Equal(1.65, mapped.NumericValue, precision: 15);
        Assert.Equal("1.65", mapped.ScpiValue);
    }

    [Fact]
    public void TryMapTriggerLevel_returns_false_when_empty()
    {
        Assert.False(ScopeValueMapper.TryMapTriggerLevel(new ComponentImageEntry(), out _));
    }

    // ------------------------------------------------------- TryGetAdjacentTimeDivValue

    [Fact]
    public void TryGetAdjacentTimeDivValue_steps_forward_and_back_in_the_scopes_own_order()
    {
        var scope = Scope(timeDivList: "1ms, 2ms, 5ms, 10ms");

        Assert.True(ScopeValueMapper.TryGetAdjacentTimeDivValue(scope, 2e-3, +1, out ScopeMappedValue next));
        Assert.Equal("5ms", next.MatchedDisplayValue);

        Assert.True(ScopeValueMapper.TryGetAdjacentTimeDivValue(scope, 2e-3, -1, out ScopeMappedValue prev));
        Assert.Equal("1ms", prev.MatchedDisplayValue);
    }

    [Theory]
    [InlineData(1e-3, -1)]   // already at the first entry
    [InlineData(10e-3, +1)]  // already at the last entry
    public void TryGetAdjacentTimeDivValue_returns_false_at_the_ends_of_the_list(double current, int offset)
    {
        var scope = Scope(timeDivList: "1ms, 2ms, 5ms, 10ms");
        Assert.False(ScopeValueMapper.TryGetAdjacentTimeDivValue(scope, current, offset, out _));
    }

    [Fact]
    public void TryGetAdjacentTimeDivValue_returns_false_for_a_zero_offset()
    {
        var scope = Scope(timeDivList: "1ms, 2ms");
        Assert.False(ScopeValueMapper.TryGetAdjacentTimeDivValue(scope, 1e-3, 0, out _));
    }

    [Fact]
    public void TryGetAdjacentTimeDivValue_returns_false_when_the_current_value_is_not_in_the_list()
    {
        var scope = Scope(timeDivList: "1ms, 2ms, 5ms");
        Assert.False(ScopeValueMapper.TryGetAdjacentTimeDivValue(scope, 3e-3, +1, out _));
    }

    // ----------------------------------------------------- TryGetSupportedVoltsDivValue

    [Fact]
    public void TryGetSupportedVoltsDivValue_matches_a_numeric_voltage_against_the_list()
    {
        var scope = Scope(voltsDivList: "100mV, 500mV, 1V, 2V");

        Assert.True(ScopeValueMapper.TryGetSupportedVoltsDivValue(scope, 1.0, out ScopeMappedValue mapped));

        Assert.Equal("1V", mapped.MatchedDisplayValue);
        Assert.Equal("1", mapped.ScpiValue);
    }

    [Fact]
    public void TryGetSupportedVoltsDivValue_returns_false_for_an_unsupported_voltage()
    {
        var scope = Scope(voltsDivList: "100mV, 500mV, 1V");
        Assert.False(ScopeValueMapper.TryGetSupportedVoltsDivValue(scope, 0.75, out _));
    }

    // ------------------------------------------------------------------- SCPI numbers

    [Theory]
    [InlineData("1V", "1")]
    [InlineData("1.65V", "1.65")]
    [InlineData("0.5V", "0.5")]
    public void ScpiValue_uses_an_invariant_decimal_point(string triggerLevel, string expected)
    {
        // Guards against a comma decimal separator on a Danish/German machine reaching the wire.
        var component = new ComponentImageEntry { TriggerLevelVolts = triggerLevel };

        Assert.True(ScopeValueMapper.TryMapTriggerLevel(component, out ScopeMappedValue mapped));

        Assert.Equal(expected, mapped.ScpiValue);
        Assert.DoesNotContain(",", mapped.ScpiValue);
    }

    [Fact]
    public void ScpiValue_uses_a_lower_case_exponent_for_very_small_values()
    {
        var scope = Scope(timeDivList: "5ns");

        Assert.True(ScopeValueMapper.TryMapTimeDiv(
            new ComponentImageEntry { TimeDiv = "5ns" }, scope, out ScopeMappedValue mapped));

        Assert.Equal("5e-09", mapped.ScpiValue);
        Assert.DoesNotContain("E", mapped.ScpiValue);
    }

    [Fact]
    public void Micro_sign_is_normalised_but_greek_mu_is_not()
    {
        // NormalizeUnit only rewrites U+00B5 MICRO SIGN. A Greek small letter mu (U+03BC),
        // which some editors and fonts substitute silently, is NOT recognised. Worth knowing
        // if a board file's timebase mysteriously fails to map.
        Assert.True(ScopeValueMapper.TryParseTimeValue("5µs", out _));   // micro sign
        Assert.False(ScopeValueMapper.TryParseTimeValue("5μs", out _));  // greek mu
    }
}
