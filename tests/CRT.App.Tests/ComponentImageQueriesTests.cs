using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for ComponentImageQueries - the rules deciding which component images
// and entries the popup window shows, and which of them carry an oscilloscope baseline.
//
// These were private members of ComponentInfoWindow. Each rule has an edge case invisible from
// the UI: a blank region means "shared", an image with no File is not displayable at all, and an
// entry counts as an oscilloscope baseline only with a pin AND at least one scope setting.
public class ComponentImageQueriesTests
{
    private static ComponentImageEntry Image(
        string boardLabel = "U1", string region = "", string file = "img.png",
        string pin = "", string name = "",
        string timeDiv = "", string voltsDiv = "", string triggerLevel = "")
        => new()
        {
            BoardLabel = boardLabel,
            Region = region,
            File = file,
            Pin = pin,
            Name = name,
            TimeDiv = timeDiv,
            VoltsDiv = voltsDiv,
            TriggerLevelVolts = triggerLevel
        };

    // -------------------------------------------------------------- IsImageVisibleInRegion

    // The shared-region rule: an image with no region shows under both PAL and NTSC.
    [Fact]
    public void An_image_without_a_region_is_visible_everywhere()
    {
        Assert.True(ComponentImageQueries.IsImageVisibleInRegion(Image(region: ""), "PAL"));
        Assert.True(ComponentImageQueries.IsImageVisibleInRegion(Image(region: ""), "NTSC"));
    }

    [Fact]
    public void An_image_naming_a_region_is_visible_only_there()
    {
        Assert.True(ComponentImageQueries.IsImageVisibleInRegion(Image(region: "PAL"), "PAL"));
        Assert.False(ComponentImageQueries.IsImageVisibleInRegion(Image(region: "PAL"), "NTSC"));
    }

    [Theory]
    [InlineData("  pal  ")]
    [InlineData("PAL")]
    public void Region_visibility_ignores_case_and_whitespace(string region)
    {
        Assert.True(ComponentImageQueries.IsImageVisibleInRegion(Image(region: region), "PAL"));
    }

    // Whitespace-only is blank, so it is shared rather than a region nothing matches.
    [Fact]
    public void A_whitespace_only_image_region_is_shared()
    {
        Assert.True(ComponentImageQueries.IsImageVisibleInRegion(Image(region: "   "), "NTSC"));
    }

    // -------------------------------------------------------------- HasDisplayableImageFile

    [Theory]
    [InlineData("img.png", true)]
    [InlineData("", false)]
    [InlineData("   ", false)]
    public void An_image_is_displayable_only_when_it_names_a_file(string file, bool expected)
    {
        Assert.Equal(expected, ComponentImageQueries.HasDisplayableImageFile(Image(file: file)));
    }

    // -------------------------------------------------------------- CountImagesForRegion

    [Fact]
    public void Only_images_for_the_requested_board_label_are_counted()
    {
        var images = new[] { Image(boardLabel: "U1"), Image(boardLabel: "U2") };

        Assert.Equal(1, ComponentImageQueries.CountImagesForRegion(images, "U1", "PAL"));
    }

    [Fact]
    public void The_board_label_match_ignores_case()
    {
        var images = new[] { Image(boardLabel: "u1") };

        Assert.Equal(1, ComponentImageQueries.CountImagesForRegion(images, "U1", "PAL"));
    }

    // Shared images count toward BOTH region totals, which is why the two counters can add up to
    // more than the number of images.
    [Fact]
    public void A_shared_image_counts_toward_every_region()
    {
        var images = new[] { Image(region: ""), Image(region: "PAL") };

        Assert.Equal(2, ComponentImageQueries.CountImagesForRegion(images, "U1", "PAL"));
        Assert.Equal(1, ComponentImageQueries.CountImagesForRegion(images, "U1", "NTSC"));
    }

    // An entry with no File is excluded even when its board label and region both match, so the
    // counter never promises a thumbnail that cannot be rendered.
    [Fact]
    public void An_image_without_a_file_is_not_counted()
    {
        var images = new[] { Image(file: ""), Image(file: "img.png") };

        Assert.Equal(1, ComponentImageQueries.CountImagesForRegion(images, "U1", "PAL"));
    }

    [Fact]
    public void An_empty_image_set_counts_zero()
    {
        Assert.Equal(0, ComponentImageQueries.CountImagesForRegion(Array.Empty<ComponentImageEntry>(), "U1", "PAL"));
    }

    // -------------------------------------------------------------- BuildImageLabel

    // Pin wins over name, so a baseline capture is always labelled by its pin.
    [Fact]
    public void An_image_label_prefers_the_pin_over_the_name()
    {
        Assert.Equal("Pin 7", ComponentImageQueries.BuildImageLabel(Image(pin: "7", name: "Clock")));
    }

    [Fact]
    public void An_image_label_falls_back_to_the_name_when_there_is_no_pin()
    {
        Assert.Equal("Clock", ComponentImageQueries.BuildImageLabel(Image(pin: "", name: "Clock")));
    }

    [Fact]
    public void An_image_with_neither_pin_nor_name_has_an_empty_label()
    {
        Assert.Equal(string.Empty, ComponentImageQueries.BuildImageLabel(Image()));
    }

    [Fact]
    public void An_image_label_trims_its_source_text()
    {
        Assert.Equal("Pin 7", ComponentImageQueries.BuildImageLabel(Image(pin: "  7  ")));
        Assert.Equal("Clock", ComponentImageQueries.BuildImageLabel(Image(name: "  Clock  ")));
    }

    // -------------------------------------------------------------- IsOscilloscopeImage

    // A pin alone is not enough - without a scope setting there is nothing to apply to the scope.
    [Fact]
    public void An_entry_with_a_pin_but_no_scope_settings_is_not_an_oscilloscope_image()
    {
        Assert.False(ComponentImageQueries.IsOscilloscopeImage(Image(pin: "7")));
    }

    // Equally, scope settings without a pin are not enough - there is no probe point.
    [Fact]
    public void An_entry_with_scope_settings_but_no_pin_is_not_an_oscilloscope_image()
    {
        Assert.False(ComponentImageQueries.IsOscilloscopeImage(Image(pin: "", timeDiv: "1ms")));
    }

    // Any ONE of the three settings is sufficient once a pin is present.
    [Theory]
    [InlineData("1ms", "", "")]
    [InlineData("", "500mV", "")]
    [InlineData("", "", "1.5")]
    public void A_pin_plus_any_single_scope_setting_makes_an_oscilloscope_image(
        string timeDiv, string voltsDiv, string triggerLevel)
    {
        Assert.True(ComponentImageQueries.IsOscilloscopeImage(
            Image(pin: "7", timeDiv: timeDiv, voltsDiv: voltsDiv, triggerLevel: triggerLevel)));
    }

    [Fact]
    public void A_null_entry_is_not_an_oscilloscope_image()
    {
        Assert.False(ComponentImageQueries.IsOscilloscopeImage(null));
    }

    // -------------------------------------------------------------- PickComponentEntry

    private static ComponentEntry Entry(string region, string friendly = "")
        => new() { Region = region, FriendlyName = friendly };

    // Preference order is: exact region -> generic (blank region) -> first available.
    [Fact]
    public void An_exact_region_match_wins()
    {
        var entries = new[] { Entry("", "generic"), Entry("PAL", "pal"), Entry("NTSC", "ntsc") };

        Assert.Equal("pal", ComponentImageQueries.PickComponentEntry(entries, "PAL")!.FriendlyName);
    }

    [Fact]
    public void A_generic_entry_is_used_when_no_region_matches()
    {
        var entries = new[] { Entry("NTSC", "ntsc"), Entry("", "generic") };

        Assert.Equal("generic", ComponentImageQueries.PickComponentEntry(entries, "PAL")!.FriendlyName);
    }

    // Last resort: rather than showing nothing, the first entry is used even though its region
    // is wrong. That is a deliberate "show something" choice, not a bug.
    [Fact]
    public void The_first_entry_is_used_when_there_is_no_match_and_no_generic()
    {
        var entries = new[] { Entry("NTSC", "ntsc"), Entry("SECAM", "secam") };

        Assert.Equal("ntsc", ComponentImageQueries.PickComponentEntry(entries, "PAL")!.FriendlyName);
    }

    [Fact]
    public void An_empty_entry_list_picks_nothing()
    {
        Assert.Null(ComponentImageQueries.PickComponentEntry(Array.Empty<ComponentEntry>(), "PAL"));
    }

    [Fact]
    public void Region_matching_when_picking_an_entry_ignores_case_and_whitespace()
    {
        var entries = new[] { Entry("  pal  ", "pal") };

        Assert.Equal("pal", ComponentImageQueries.PickComponentEntry(entries, "PAL")!.FriendlyName);
    }

    // A whitespace-only region trims to empty, so it is found by the GENERIC lookup rather than
    // the exact-region one - it does not match "PAL" but is still preferred over an NTSC entry.
    [Fact]
    public void A_whitespace_only_entry_region_is_treated_as_generic()
    {
        var entries = new[] { Entry("NTSC", "ntsc"), Entry("   ", "blank") };

        Assert.Equal("blank", ComponentImageQueries.PickComponentEntry(entries, "PAL")!.FriendlyName);
    }
}
