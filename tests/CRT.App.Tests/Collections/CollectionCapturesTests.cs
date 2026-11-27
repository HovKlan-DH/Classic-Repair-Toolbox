using Handlers.Collections;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// A capture as a collection keeps it (CollectionCaptures): the pin shown, the settings read back from the oscilloscope in a
// baseline's words, and what the component info window showed of its component (owner request, 2026-11-27).
// ###########################################################################################
public sealed class CollectionCapturesTests
{
    private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

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

    private static readonly CollectionComponent Vic = new("U7", "U7 | VIC-II | 6569", null, null, null, null);

    [Fact]
    public void A_capture_keeps_its_pin_the_settings_read_back_and_its_component()
    {
        LocalCapture capture = CollectionCaptures.CaptureOf(Baseline(), new ScopeSettingsReading(50e-6, 2, 1.4), Vic, Now);

        Assert.Equal(("U7", "12", "/RAS", "Pulsing"), (capture.ComponentLabel, capture.Pin, capture.SignalName, capture.ExpectedReading));
        Assert.Equal(("50us", "2V", "1.4V"), (capture.TimePerDivision, capture.VoltsPerDivision, capture.TriggerLevel));
        Assert.Same(Vic, capture.Component);
        Assert.Equal(Now, capture.CapturedUtc);
    }

    // The oscilloscope's model cannot ask, or it did not answer: the value is left out, never guessed.
    [Fact]
    public void A_setting_not_read_back_is_left_out()
    {
        LocalCapture capture = CollectionCaptures.CaptureOf(Baseline(), new ScopeSettingsReading(null, 9.9e37, null), null, Now);

        Assert.Null(capture.TimePerDivision);
        Assert.Null(capture.VoltsPerDivision);
        Assert.Null(capture.TriggerLevel);
        Assert.Null(CollectionCaptures.CaptureOf(Baseline(), null, null, Now).TimePerDivision);
    }

    // A component's picture (no pin) shown as the capture was taken is of no pin.
    [Fact]
    public void A_capture_taken_over_a_picture_is_of_no_pin() =>
        Assert.Null(CollectionCaptures.CaptureOf(Baseline(pin: string.Empty), null, null, Now).Pin);

    // The window's facts and pin list as they are shown - blank fields as none, the purpose as its words, the pins listed.
    [Fact]
    public void A_component_is_told_as_the_window_shows_it()
    {
        CollectionComponent? component = CollectionCaptures.ComponentOf(
            " U7 ",
            "U7 | VIC-II | 6569",
            "IC | 6569 | DIP-40",
            "  ",
            "Makes **the** picture",
            [new ComponentPopupPin("12", "/RAS", "Output", "RAS", Reading: "Pulsing"), new ComponentPopupPin(" ", "x", "Input")]);

        Assert.NotNull(component);
        Assert.Equal(("U7", "U7 | VIC-II | 6569", "IC | 6569 | DIP-40", (string?)null, "Makes the picture"), (component.Label, component.Title, component.CategoryLine, component.Description, component.Purpose));
        Assert.Equal([new CollectionPin("12", "/RAS", "Output", "RAS", "Pulsing")], component.Pins);
    }

    // Nothing to tell is none; what passes the server's limits is cut, never refused.
    [Fact]
    public void A_component_with_nothing_to_tell_is_none_and_one_too_long_is_cut()
    {
        Assert.Null(CollectionCaptures.ComponentOf("U7", "", null, null, null, []));

        CollectionComponent long_ = CollectionCaptures.ComponentOf(
            "U7",
            new string('T', 500),
            null,
            null,
            null,
            Enumerable.Range(1, CollectionRules.MaximumPinsPerComponent + 5).Select(pin => new ComponentPopupPin($"{pin}", "n", "Input")))!;

        Assert.Equal(CollectionRules.MaximumComponentTextLength, long_.Title!.Length);
        Assert.Equal(CollectionRules.MaximumPinsPerComponent, long_.Pins!.Count);
        Assert.Empty(CollectionRules.ProblemsOf(
            new CollectionContentRequest("C64", "250407", null, null, null, null, Now.AddMonths(1), [new(new string('a', 64), "U7", "12", null, null, null, null, null, null, null)], [long_]),
            Now));
    }
}
