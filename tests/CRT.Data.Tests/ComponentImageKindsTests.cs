using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// The one rule for which "Component images" rows are oscilloscope baselines - the popup's, and
// the rule the CRT 4.0 converter splits the sheet by into pictures and baselines.
public sealed class ComponentImageKindsTests
{
    private static ComponentImageEntry Row(string pin = "", string timeDiv = "", string voltsDiv = "", string trigger = "") => new()
    {
        BoardLabel = "U1",
        Pin = pin,
        TimeDiv = timeDiv,
        VoltsDiv = voltsDiv,
        TriggerLevelVolts = trigger,
    };

    [Theory]
    [InlineData("1uS", "", "")]
    [InlineData("", "2V", "")]
    [InlineData("", "", "1.4V")]
    public void A_row_with_a_pin_and_any_scope_setting_is_a_baseline(string timeDiv, string voltsDiv, string trigger)
    {
        Assert.True(ComponentImageKinds.IsScopeBaseline(Row("7", timeDiv, voltsDiv, trigger)));
    }

    // A per-pin photo with no scope setting is a picture of that pin, not a reading.
    [Fact]
    public void A_row_with_a_pin_and_no_scope_setting_is_a_picture()
    {
        Assert.False(ComponentImageKinds.IsScopeBaseline(Row("7")));
    }

    [Fact]
    public void A_row_without_a_pin_is_a_picture_whatever_its_settings()
    {
        Assert.False(ComponentImageKinds.IsScopeBaseline(Row("", "1uS", "2V", "1.4V")));
    }

    [Fact]
    public void Blank_text_counts_as_nothing_and_no_row_is_no_baseline()
    {
        Assert.False(ComponentImageKinds.IsScopeBaseline(Row("  ", "1uS")));
        Assert.False(ComponentImageKinds.IsScopeBaseline(Row("7", "  ", " ", "\t")));
        Assert.False(ComponentImageKinds.IsScopeBaseline(null));
    }
}
