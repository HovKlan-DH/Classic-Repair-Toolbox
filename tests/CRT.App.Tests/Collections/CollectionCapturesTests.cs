using Handlers.Collections;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// A capture as a collection keeps it (CollectionCaptures): the pin and baseline shown, the settings read back from the
// oscilloscope in a baseline's words.
// ###########################################################################################
public sealed class CollectionCapturesTests
{
    private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

    private static readonly string BaselineSha = new('c', 64);

    private static ComponentImageEntry Baseline(string pin = "12") => new()
    {
        BoardLabel = " U7 ",
        Pin = pin,
        Name = "/RAS",
        ExpectedOscilloscopeReading = "Pulsing",
        File = "Scope baseline/U7_12_RAS.png",
        TimeDiv = "100us",
        VoltsDiv = "2V",
        TriggerLevelVolts = "1.5V"
    };

    [Fact]
    public void A_capture_of_a_baseline_keeps_its_pin_the_settings_read_back_and_the_baseline()
    {
        LocalCapture capture = CollectionCaptures.CaptureOf(Baseline(), new ScopeSettingsReading(50e-6, 2, 1.4), BaselineSha, Now);

        Assert.Equal(("U7", "12", "/RAS", "Pulsing"), (capture.ComponentLabel, capture.Pin, capture.SignalName, capture.ExpectedReading));
        Assert.Equal(("50us", "2V", "1.4V"), (capture.TimePerDivision, capture.VoltsPerDivision, capture.TriggerLevel));
        Assert.Equal((BaselineSha, "Scope baseline/U7_12_RAS.png"), (capture.BaselineSha256, capture.BaselineFileName));
        Assert.Equal(("100us", "2V", "1.5V"), (capture.BaselineTimePerDivision, capture.BaselineVoltsPerDivision, capture.BaselineTriggerLevel));
        Assert.Equal(Now, capture.CapturedUtc);
    }

    // The oscilloscope's model cannot ask, or it did not answer: the value is left out, never guessed.
    [Fact]
    public void A_setting_not_read_back_is_left_out()
    {
        LocalCapture capture = CollectionCaptures.CaptureOf(Baseline(), new ScopeSettingsReading(null, 9.9e37, null), BaselineSha, Now);

        Assert.Null(capture.TimePerDivision);
        Assert.Null(capture.VoltsPerDivision);
        Assert.Null(capture.TriggerLevel);
        Assert.Null(CollectionCaptures.CaptureOf(Baseline(), null, BaselineSha, Now).TimePerDivision);
    }

    // A component's picture (no pin) shown as the capture was taken is no baseline of anything.
    [Fact]
    public void A_capture_taken_over_a_picture_has_no_baseline()
    {
        ComponentImageEntry picture = Baseline(pin: string.Empty);

        LocalCapture capture = CollectionCaptures.CaptureOf(picture, null, BaselineSha, Now);

        Assert.Null(capture.Pin);
        Assert.Null(capture.BaselineSha256);
        Assert.Null(capture.BaselineTimePerDivision);
    }

    [Fact]
    public void A_baseline_checksum_that_is_none_is_left_out() =>
        Assert.Null(CollectionCaptures.CaptureOf(Baseline(), null, "not a checksum", Now).BaselineSha256);
}
