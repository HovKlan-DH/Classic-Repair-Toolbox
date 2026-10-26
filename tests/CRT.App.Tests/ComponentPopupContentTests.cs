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

    // "8500 - Drop-in" on the first line, what to change and where it comes from under it.
    [Fact]
    public void A_replacement_part_reads_part_number_and_kind_then_changes_and_source()
    {
        ComponentReplacementEntry[] parts =
        [
            new() { BoardLabel = "U8", PartNumber = "8500", Kind = "Drop-in", Source = "Service manual" },
            new() { BoardLabel = "U8", PartNumber = "6502", Kind = "Needs changes", Changes = "Add a 6522 for the port" },
            new() { BoardLabel = "U8", Kind = "Needs changes" },
            new() { BoardLabel = "U8", Region = "NTSC", PartNumber = "NTSC part" },
            new() { BoardLabel = "U8" },
        ];

        List<ComponentPopupReplacement> shown = ComponentPopupContent.ReplacementsOf(parts, "U8", "PAL");

        Assert.Equal(3, shown.Count);
        Assert.Equal(new ComponentPopupReplacement("8500 - Drop-in", string.Empty, "Comes from: Service manual"), shown[0]);
        Assert.True(shown[0].HasSource);
        Assert.False(shown[0].HasChanges);
        Assert.Equal("6502 - Needs changes", shown[1].Title);
        Assert.Equal("Add a 6522 for the port", shown[1].Changes);
        Assert.Equal("Needs changes", shown[2].Title);
    }

    [Fact]
    public void A_baselines_conditions_name_its_machine_state_and_where_it_comes_from()
    {
        Assert.Equal(
            "Machine state: Running BASIC\nComes from: Known-good board",
            ComponentPopupContent.BaselineConditions(new ComponentImageEntry { MachineState = " Running BASIC ", Source = "Known-good board" }));
        Assert.Equal("Comes from: Datasheet", ComponentPopupContent.BaselineConditions(new ComponentImageEntry { Source = "Datasheet" }));
        Assert.Null(ComponentPopupContent.BaselineConditions(new ComponentImageEntry { Pin = "1", TimeDiv = "1ms" }));
        Assert.Null(ComponentPopupContent.BaselineConditions(null));
    }
}
