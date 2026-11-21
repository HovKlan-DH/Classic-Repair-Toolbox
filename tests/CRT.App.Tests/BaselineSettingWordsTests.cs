using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BaselineSettingWords - what the oscilloscope says it is set to, written as a baseline stores it (owner decision, 2026-11-18:
// the settings read back are what the baseline stores). The spellings are the published data's, counted on the converted trees
// that day: "50us", "500ns", "1ms", "1V", "500mV", "1.5V", "0.75V".
// ###########################################################################################
public sealed class BaselineSettingWordsTests
{
    [Theory]
    [InlineData(5e-05, "50us")]
    [InlineData(1e-06, "1us")]
    [InlineData(5e-07, "500ns")]
    [InlineData(0.001, "1ms")]
    [InlineData(0.002, "2ms")]
    [InlineData(1.0, "1s")]
    [InlineData(10.0, "10s")]
    // A scope's answer a hair off its step is the step - never "1000us" for a millisecond.
    [InlineData(4.9999999999e-05, "50us")]
    [InlineData(0.000999999999999, "1ms")]
    [InlineData(2.5e-05, "25us")]
    public void A_time_per_division_is_written_in_the_largest_unit_it_is_one_of(double seconds, string expected)
    {
        Assert.Equal(expected, BaselineSettingWords.TimeDiv(seconds));
    }

    [Theory]
    [InlineData(1.0, "1V")]
    [InlineData(2.0, "2V")]
    [InlineData(0.5, "500mV")]
    [InlineData(0.02, "20mV")]
    [InlineData(1.25, "1.25V")]
    public void A_volts_per_division_is_written_in_volts_or_millivolts(double volts, string expected)
    {
        Assert.Equal(expected, BaselineSettingWords.VoltsDiv(volts));
    }

    // A trigger level is always in volts, as every published one is - below a volt, at 0 and below it too.
    [Theory]
    [InlineData(1.5, "1.5V")]
    [InlineData(0.75, "0.75V")]
    [InlineData(1.2000000001, "1.2V")]
    [InlineData(0.0, "0V")]
    [InlineData(-0.0001, "0V")]
    [InlineData(-0.25, "-0.25V")]
    public void A_trigger_level_is_always_in_volts(double volts, string expected)
    {
        Assert.Equal(expected, BaselineSettingWords.TriggerLevel(volts));
    }

    // A value the oscilloscope did not answer - or one no setting can be - is none, so the baseline keeps what it had.
    [Fact]
    public void A_value_not_answered_or_impossible_is_none()
    {
        Assert.Null(BaselineSettingWords.TimeDiv(null));
        Assert.Null(BaselineSettingWords.TimeDiv(0));
        Assert.Null(BaselineSettingWords.TimeDiv(-1e-6));
        Assert.Null(BaselineSettingWords.TimeDiv(double.NaN));
        Assert.Null(BaselineSettingWords.VoltsDiv(null));
        Assert.Null(BaselineSettingWords.VoltsDiv(double.PositiveInfinity));
        Assert.Null(BaselineSettingWords.TriggerLevel(null));
        Assert.Null(BaselineSettingWords.TriggerLevel(double.NaN));
    }

    // ###########################################################################################
    // The baseline a capture makes - ONE rule for what the window shows before it is accepted and what goes into the draft (code
    // review, 2026-11-21: shown as "All regions", put in PAL): a baseline it had keeps every value but its image and the settings the
    // oscilloscope said; a new one is the pin's, named, in the machine state and in the region a new baseline goes in.
    // ###########################################################################################
    [Fact]
    public void A_capture_makes_its_baseline_or_a_new_one_in_the_region_new_baselines_go_in()
    {
        var settings = new ScopeSettingsReading(0.0001, null, 0.8);
        var had = new Handlers.DataHandling.ComponentImageEntry
        {
            BoardLabel = "U27", Region = "NTSC", Pin = "3", Name = "1Y", File = "old.png", Note = "**Ringing**",
            TimeDiv = "1us", VoltsDiv = "1V", TriggerLevelVolts = "1.5V", MachineState = "READY", ExpectedOscilloscopeReading = "Pulsing",
        };

        Handlers.DataHandling.ComponentImageEntry again = BaselineSettingWords.Captured(had, settings, "new.png", "9", "X", "Running", "PAL");
        Assert.Equal(
            ("U27", "NTSC", "3", "1Y", "new.png", "**Ringing**", "100us", "1V", "0.8V", "READY", "Pulsing"),
            (again.BoardLabel, again.Region, again.Pin, again.Name, again.File, again.Note, again.TimeDiv, again.VoltsDiv, again.TriggerLevelVolts, again.MachineState, again.ExpectedOscilloscopeReading));

        Handlers.DataHandling.ComponentImageEntry added = BaselineSettingWords.Captured(null, settings, "new.png", "5", "2A", " Running ", "PAL");
        Assert.Equal(
            ("PAL", "5", "2A", "new.png", "100us", "", "0.8V", "Running"),
            (added.Region, added.Pin, added.Name, added.File, added.TimeDiv, added.VoltsDiv, added.TriggerLevelVolts, added.MachineState));
    }
}
