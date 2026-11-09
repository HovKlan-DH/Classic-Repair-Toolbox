using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentPopupContent - what the component popup shows of CRT 4.0's new component data (step 2,
// slice 3c). The popup shows ONE region at a time, as the rest of it does, so the pins and parts it
// lists are the component's in that region, a row with no region being in every one.
// ###########################################################################################
public sealed class ComponentPopupContentTests
{
    [Fact]
    public void The_category_line_names_category_part_number_and_package_that_are_filled()
    {
        Assert.Equal("IC | 6510 | DIP-40", ComponentPopupContent.CategoryLine(new ComponentEntry { Category = "IC", PartNumber = "6510", Package = " DIP-40 " }));
        Assert.Equal("IC | DIP-40", ComponentPopupContent.CategoryLine(new ComponentEntry { Category = "IC", Package = "DIP-40" }));
        Assert.Equal(string.Empty, ComponentPopupContent.CategoryLine(new ComponentEntry()));
        Assert.Equal(string.Empty, ComponentPopupContent.CategoryLine(null));
    }

    // The component's pins of the region shown, in number order - "2" before "10".
    [Fact]
    public void The_pins_are_the_components_in_the_region_shown_in_number_order()
    {
        ComponentPinEntry[] pins =
        [
            new() { BoardLabel = "U8", Pin = "10", Name = "A0" },
            new() { BoardLabel = "u8", Pin = "2", Name = "PHI0", Direction = "Input" },
            new() { BoardLabel = "U8", Region = "NTSC", Pin = "3", Name = "NTSC only" },
            new() { BoardLabel = "U9", Pin = "1", Name = "Another chip" },
        ];

        List<ComponentPopupPin> shown = ComponentPopupContent.PinsOf(pins, "U8", "PAL");

        Assert.Equal([new ComponentPopupPin("2", "PHI0", "Input"), new ComponentPopupPin("10", "A0", string.Empty)], shown);
        Assert.Equal(3, ComponentPopupContent.PinsOf(pins, "U8", "NTSC").Count);
        Assert.Equal(3, ComponentPopupContent.PinEntriesOf(pins, "U8", "NTSC").Count);
        Assert.Empty(ComponentPopupContent.PinsOf(pins, "", "PAL"));
    }

    // Each pin beside its net on this board, by its number - a pin the KiCad data has no net for, none.
    [Fact]
    public void A_pin_shows_its_net_on_this_board_by_its_number()
    {
        ComponentPinEntry[] pins =
        [
            new() { BoardLabel = "U8", Pin = " 2 ", Name = "PHI0", Direction = "Input" },
            new() { BoardLabel = "U8", Pin = "10", Name = "A0" },
        ];

        var nets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["2"] = "PHI0" };

        Assert.Equal(
            [new ComponentPopupPin("2", "PHI0", "Input", "PHI0"), new ComponentPopupPin("10", "A0", string.Empty, string.Empty)],
            ComponentPopupContent.PinsOf(pins, "U8", "PAL", nets));
    }

    // "8500 - Drop-in replacement" on the first line, its note under it.
    [Fact]
    public void A_replacement_part_reads_part_number_and_kind_then_its_note()
    {
        ComponentReplacementEntry[] parts =
        [
            new() { BoardLabel = "U8", PartNumber = "8500", Kind = "Drop-in replacement" },
            new() { BoardLabel = "U8", PartNumber = "6502", Kind = "Needs modification", Note = "Add a 6522 for the port" },
            new() { BoardLabel = "U8", Kind = "Needs modification" },
            new() { BoardLabel = "U8", Region = "NTSC", PartNumber = "NTSC part" },
            new() { BoardLabel = "U8" },
        ];

        List<ComponentPopupReplacement> shown = ComponentPopupContent.ReplacementsOf(parts, "U8", "PAL");

        Assert.Equal(3, shown.Count);
        Assert.Equal(new ComponentPopupReplacement("8500 - Drop-in replacement", string.Empty), shown[0]);
        Assert.False(shown[0].HasNote);
        Assert.Equal("6502 - Needs modification", shown[1].Title);
        Assert.Equal("Add a 6522 for the port", shown[1].Note);
        Assert.True(shown[1].HasNote);
        Assert.Equal("Needs modification", shown[2].Title);
    }

    [Fact]
    public void A_baselines_conditions_name_its_machine_state()
    {
        Assert.Equal(
            "Machine state: Running BASIC",
            ComponentPopupContent.BaselineConditions(new ComponentImageEntry { MachineState = " Running BASIC " }));
        Assert.Null(ComponentPopupContent.BaselineConditions(new ComponentImageEntry { Pin = "1", TimeDiv = "1ms" }));
        Assert.Null(ComponentPopupContent.BaselineConditions(null));
    }

    // A pin clicked goes to its first image - "2 [64K]" is pin 2, as the checks read it (owner request,
    // 2026-11-01: "go to that pin and see what images are available for that pin").
    [Theory]
    [InlineData("2", 1)]
    [InlineData(" 2 ", 1)]
    [InlineData("7", 3)]
    [InlineData("A1", 4)]
    [InlineData("a1", 4)]
    [InlineData("9", -1)]
    [InlineData("", -1)]
    [InlineData(null, -1)]
    public void A_pin_finds_its_first_image(string? pin, int index)
    {
        string[] imagePins = ["1", "2 [64K]", "2", "7", "A1", ""];

        Assert.Equal(index, ComponentPopupContent.FirstImageOfPin(imagePins, pin));
    }

    [Fact]
    public void A_pin_with_no_image_says_so()
    {
        Assert.Equal("No image of pin 9", ComponentPopupContent.NoImageOfPin(" 9 "));
    }

    // ###########################################################################################
    // The pin of the image shown says so on its row (owner request, 2026-11-02: "if I am viewing a
    // specific pin, then do highlight it") - by its number, as "2 [64K]" is pin 2.
    // ###########################################################################################
    [Fact]
    public void The_row_of_the_pin_shown_says_so()
    {
        ComponentPinEntry[] pins = [new() { BoardLabel = "U8", Pin = "1", Name = "VCC" }, new() { BoardLabel = "U8", Pin = "2", Name = "CLK" }];

        Assert.Equal([false, true], ComponentPopupContent.PinsOf(pins, "U8", "PAL", viewedPin: "2 [64K]").Select(pin => pin.IsViewed));
        Assert.Equal([false, false], ComponentPopupContent.PinsOf(pins, "U8", "PAL", viewedPin: null).Select(pin => pin.IsViewed));
        Assert.Equal([false, false], ComponentPopupContent.PinsOf(pins, "U8", "PAL", viewedPin: "9").Select(pin => pin.IsViewed));
    }

    // ###########################################################################################
    // The pin the pointer is on in the drawing says so on its row (owner request, 2026-11-09: "hovering the IC image
    // pin will not highlight the table pin ... do highlight both places") - by its number, as the drawing knows pins.
    // ###########################################################################################
    [Fact]
    public void The_row_of_the_pin_pointed_at_in_the_drawing_says_so()
    {
        ComponentPinEntry[] pins = [new() { BoardLabel = "U8", Pin = "1", Name = "VCC" }, new() { BoardLabel = "U8", Pin = " 2 ", Name = "CLK" }];

        Assert.Equal([false, true], ComponentPopupContent.PinsOf(pins, "U8", "PAL", pointedPin: 2).Select(pin => pin.IsPointedAt));
        Assert.Equal([false, false], ComponentPopupContent.PinsOf(pins, "U8", "PAL", pointedPin: null).Select(pin => pin.IsPointedAt));
        Assert.Equal([false, false], ComponentPopupContent.PinsOf(pins, "U8", "PAL", pointedPin: 9).Select(pin => pin.IsPointedAt));

        // The pin shown and the one pointed at are two things - one row can be both.
        ComponentPopupPin both = ComponentPopupContent.PinsOf(pins, "U8", "PAL", viewedPin: "2", pointedPin: 2)[1];
        Assert.True(both.IsViewed);
        Assert.True(both.IsPointedAt);
    }

    // ###########################################################################################
    // EACH PIN SAYS WHAT THE SCOPE SHOULD READ ON IT (owner request, 2026-11-09: "add another column to it named
    // 'Expected oscilloscope reading' and have the value from the oscilloscope pin") - from the component's oscilloscope
    // baselines in the region shown, a baseline with no region being in every one, "2 [64K]" being pin 2. A picture
    // (no pin) is no baseline, and another component's baselines say nothing.
    // ###########################################################################################
    [Fact]
    public void Each_pin_shows_what_its_baselines_say_the_scope_should_read()
    {
        ComponentPinEntry[] pins =
        [
            new() { BoardLabel = "U8", Pin = "1", Name = "VCC" },
            new() { BoardLabel = "U8", Pin = "2", Name = "CLK" },
            new() { BoardLabel = "U8", Pin = "3", Name = "D0" },
            new() { BoardLabel = "U8", Pin = "4", Name = "NC" },
        ];

        ComponentImageEntry[] images =
        [
            new() { BoardLabel = "u8", Pin = "1", ExpectedOscilloscopeReading = " 5.1V (+5VDC power rail) " },
            new() { BoardLabel = "U8", Pin = "2 [64K]", ExpectedOscilloscopeReading = "1.02272 MHz", Region = "NTSC" },
            new() { BoardLabel = "U8", Pin = "2", ExpectedOscilloscopeReading = "985.232 kHz", Region = "PAL" },
            new() { BoardLabel = "U8", Pin = "3", ExpectedOscilloscopeReading = "   " },
            new() { BoardLabel = "U8", ExpectedOscilloscopeReading = "A picture, no baseline" },
            new() { BoardLabel = "U9", Pin = "4", ExpectedOscilloscopeReading = "Another chip's" },
        ];

        Assert.Equal(
            ["5.1V (+5VDC power rail)", "985.232 kHz", "", ""],
            ComponentPopupContent.PinsOf(pins, "U8", "PAL", images: images).Select(pin => pin.Reading));
        Assert.Equal(
            ["5.1V (+5VDC power rail)", "1.02272 MHz", "", ""],
            ComponentPopupContent.PinsOf(pins, "U8", "NTSC", images: images).Select(pin => pin.Reading));

        Assert.True(ComponentPopupContent.PinsOf(pins, "U8", "PAL", images: images)[0].HasReading);
        Assert.False(ComponentPopupContent.PinsOf(pins, "U8", "PAL", images: images)[2].HasReading);

        // Without images, no reading - as before the column came.
        Assert.All(ComponentPopupContent.PinsOf(pins, "U8", "PAL"), pin => Assert.Equal(string.Empty, pin.Reading));
    }

    // ###########################################################################################
    // A pin whose baselines read differently - the C16's "2" HIGH and "2 [64K]" Pulsing, the only kind measured on the
    // published boards (2026-11-09) - lists each reading on a line of its own, with what its pin says beside its number.
    // A reading also seen on the plain pin needs no such word, and one reading said twice, capitals aside, is said once.
    // ###########################################################################################
    [Fact]
    public void Readings_that_differ_on_one_pin_are_each_a_line_saying_which_baseline()
    {
        static string Reading(params (string Pin, string Reading)[] baselines) =>
            ComponentPopupContent.ReadingOf(baselines.Select(baseline => new ComponentImageEntry { Pin = baseline.Pin, ExpectedOscilloscopeReading = baseline.Reading }), "2");

        Assert.Equal("HIGH\nPulsing [64K]", Reading(("2", "HIGH"), ("2 [64K]", "Pulsing")));
        Assert.Equal("HIGH", Reading(("2", "HIGH"), ("2 [64K]", "high")));
        Assert.Equal("HIGH\nPulsing", Reading(("2", "HIGH"), ("2 [64K]", "Pulsing"), ("2", "Pulsing")));
        Assert.Equal("HIGH [16K]\nPulsing [64K]", Reading(("2 [16K]", "HIGH"), ("2 [64K]", "Pulsing")));
        Assert.Equal("LOW\nPulsing [16K] [64K]", Reading(("2", "LOW"), ("2 [16K]", "Pulsing"), ("2 [64K]", "Pulsing")));
        Assert.Equal(string.Empty, Reading(("3", "HIGH")));
        Assert.Equal(string.Empty, Reading());
    }

    // A file or a link of one region shows in that region alone, one for every region in each.
    [Fact]
    public void The_files_and_links_are_the_components_in_the_region_shown()
    {
        ComponentLocalFileEntry[] files =
        [
            new() { BoardLabel = "U8", Name = "All", File = "a.pdf" },
            new() { BoardLabel = "u8", Name = "PAL", File = "b.pdf", Region = "PAL" },
            new() { BoardLabel = "U8", Name = "NTSC", File = "c.pdf", Region = "NTSC" },
            new() { BoardLabel = "U9", Name = "Other", File = "d.pdf" },
        ];

        ComponentLinkEntry[] links =
        [
            new() { BoardLabel = "U8", Name = "All", Url = "https://example.org/a" },
            new() { BoardLabel = "U8", Name = "NTSC", Url = "https://example.org/b", Region = "NTSC" },
        ];

        Assert.Equal(["All", "PAL"], ComponentPopupContent.FilesOf(files, "U8", "PAL").Select(row => row.Name));
        Assert.Equal(["All", "NTSC"], ComponentPopupContent.FilesOf(files, "U8", "NTSC").Select(row => row.Name));
        Assert.Equal(["All"], ComponentPopupContent.LinksOf(links, "U8", "PAL").Select(row => row.Name));
        Assert.Equal(["All", "NTSC"], ComponentPopupContent.LinksOf(links, "U8", "NTSC").Select(row => row.Name));
    }
}
