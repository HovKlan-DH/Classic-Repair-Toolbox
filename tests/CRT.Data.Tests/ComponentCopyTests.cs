using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentCopy - "Copy from another board" (CRT 4.0 step 2, slice 5): the components of the other
// boards with the same part number, what each has, and what a copy puts into the component shown -
// its package, description and repair notes, its pins, replacement parts and links, filed under this
// component's label and region, and nothing board-specific.
// ###########################################################################################
public sealed class ComponentCopyTests
{
    // The other board: U7, the PLA, with everything filled - and another chip.
    private static BoardData Other() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U7", FriendlyName = "PLA", PartNumber = "906114-01", Category = "IC", Package = "DIP-28", Description = "Glue logic", RepairNotes = "Fails often" },
            new ComponentEntry { BoardLabel = "U6", FriendlyName = "CPU", PartNumber = "6510", Category = "IC" },
        ],
        ComponentPins =
        [
            new ComponentPinEntry { BoardLabel = "U7", Pin = "1", Name = "FE", Direction = "Input" },
            new ComponentPinEntry { BoardLabel = "U7", Pin = "28", Name = "VCC", Direction = "Power" },
        ],
        ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "U7", PartNumber = "PLAnkton", Kind = "Drop-in" }],
        ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U7", Name = "Datasheet", Url = "https://example.org/pla" }],
        ComponentImages = [new ComponentImageEntry { BoardLabel = "U7", Pin = "1", File = "Scope baseline/U7_1.png", TimeDiv = "1us" }],
    };

    // This board: U17, the same PLA, with nothing but its label, part number and a pin of its own.
    private static BoardData This() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U17", FriendlyName = "PLA", PartNumber = " 906114-01 ", Category = "IC", Region = "PAL" },
        ],
        ComponentPins = [new ComponentPinEntry { BoardLabel = "U17", Region = "PAL", Pin = "1", Name = "old" }],
    };

    [Fact]
    public void The_components_of_other_boards_with_the_same_part_number_are_offered()
    {
        List<ComponentCopySource> sources = ComponentCopy.Sources([("Commodore 64 - 250425", Other()), ("Commodore 64 - 250469", This())], "906114-01", "PAL");

        Assert.Equal(["Commodore 64 - 250425: U7 - PLA", "Commodore 64 - 250469: U17 - PLA"], sources.Select(source => source.DisplayText));
        Assert.Empty(ComponentCopy.Sources([("x", Other())], "  ", "PAL"));
        Assert.Empty(ComponentCopy.Sources([("x", Other())], "nothing", "PAL"));
    }

    // What a source has - and how each is named, with what it holds.
    // ###########################################################################################
    // A component of two regions with a chip of its own in each (code review, 2026-10-26): the row
    // offered is one that MATCHED the part number - the region's when both did, else the first that
    // did. The region's row whatever its part number copied the other chip's data, and "Also on other
    // boards" wrote this chip's change onto the other region's chip.
    // ###########################################################################################
    [Fact]
    public void The_row_offered_is_one_with_the_part_number_never_the_other_regions_chip()
    {
        var other = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U19", FriendlyName = "VIC-II", PartNumber = "6569", Region = "PAL", Package = "DIP-40" },
                new ComponentEntry { BoardLabel = "U19", FriendlyName = "VIC-II", PartNumber = "6567", Region = "NTSC", Package = "DIP-40 NTSC" },
                new ComponentEntry { BoardLabel = "U20", PartNumber = "4066", Region = "PAL" },
                new ComponentEntry { BoardLabel = "U20", PartNumber = "4066", Region = "NTSC", RepairNotes = "NTSC's" },
            ],
        };

        ComponentCopySource ntsc = Assert.Single(ComponentCopy.Sources([("Other", other)], "6567", "PAL"));
        Assert.Equal(("6567", "NTSC"), (ntsc.Row.PartNumber, ntsc.Row.Region));

        ComponentCopySource pal = Assert.Single(ComponentCopy.Sources([("Other", other)], "6569", "NTSC"));
        Assert.Equal(("6569", "PAL"), (pal.Row.PartNumber, pal.Row.Region));

        // Both regions the same chip: the region's own row.
        ComponentCopySource both = Assert.Single(ComponentCopy.Sources([("Other", other)], "4066", "NTSC"));
        Assert.Equal("NTSC's", both.Row.RepairNotes);
    }

    [Fact]
    public void A_source_offers_what_it_holds()
    {
        ComponentCopySource source = ComponentCopy.Sources([("Other", Other())], "906114-01", "PAL")[0];

        Assert.Equal(
            [ComponentCopyPart.Package, ComponentCopyPart.Description, ComponentCopyPart.RepairNotes, ComponentCopyPart.Pins, ComponentCopyPart.Replacements, ComponentCopyPart.Links],
            ComponentCopy.Available(source));
        Assert.Equal("Package (DIP-28)", ComponentCopy.Describe(source, ComponentCopyPart.Package));
        Assert.Equal("Pins (2)", ComponentCopy.Describe(source, ComponentCopyPart.Pins));

        ComponentCopySource bare = ComponentCopy.Sources([("Other", Other())], "6510", "PAL")[0];
        Assert.Empty(ComponentCopy.Available(bare));
    }

    // ###########################################################################################
    // A copy fills the fields, and replaces the component's pins, parts and links with the source's -
    // under THIS component's label and region. Nothing board-specific comes along: not the friendly
    // name, the category, the scope baselines.
    // ###########################################################################################
    [Fact]
    public void A_copy_takes_what_is_picked_under_this_components_label_and_region()
    {
        BoardData board = This();
        ComponentCopySource source = ComponentCopy.Sources([("Other", Other())], "906114-01", "PAL")[0];

        BoardData copied = ComponentCopy.Apply(board, "U17", "PAL", source, ComponentCopy.Available(source));

        ComponentEntry row = Assert.Single(copied.Components);
        Assert.Equal(("DIP-28", "Glue logic", "Fails often", "PAL", "IC"), (row.Package, row.Description, row.RepairNotes, row.Region, row.Category));
        Assert.Equal(["1:FE", "28:VCC"], copied.ComponentPins.Select(pin => $"{pin.Pin}:{pin.Name}"));
        Assert.All(copied.ComponentPins, pin => Assert.Equal(("U17", "PAL"), (pin.BoardLabel, pin.Region)));
        Assert.Equal(("U17", "PLAnkton", "PAL"), (copied.ComponentReplacements[0].BoardLabel, copied.ComponentReplacements[0].PartNumber, copied.ComponentReplacements[0].Region));
        Assert.Equal(("U17", "Datasheet"), (copied.ComponentLinks[0].BoardLabel, copied.ComponentLinks[0].Name));
        Assert.Empty(copied.ComponentImages);
    }

    // ###########################################################################################
    // "Also on other boards": what the draft changed of the component against the published board -
    // only parts it holds, only parts that differ; for a component the published board lacks, all.
    // A source can know the board it is on, for that board's own draft.
    // ###########################################################################################
    [Fact]
    public void What_the_draft_changed_is_what_can_go_to_the_other_boards()
    {
        BoardData published = Other();
        BoardData draft = ComponentPaneEdits.WithPins(
            ComponentPaneEdits.WithField(published, "U7", "", ComponentField.RepairNotes, "Fails very often"),
            "U7",
            "",
            [new ComponentPinEntry { Pin = "1", Name = "FE", Direction = "Input" }]);

        Assert.Equal([ComponentCopyPart.RepairNotes, ComponentCopyPart.Pins], ComponentCopy.Changed(draft, published, "U7", "PAL"));
        Assert.Empty(ComponentCopy.Changed(published, published, "U7", "PAL"));
        Assert.Equal(ComponentCopy.Available(ComponentCopy.Sources([("x", draft)], "906114-01", "PAL")[0]), ComponentCopy.Changed(draft, null, "U7", "PAL"));
        Assert.Empty(ComponentCopy.Changed(draft, published, "U99", "PAL"));

        ComponentCopySource known = ComponentCopy.Sources([("Commodore 64 - 250425", "Commodore/C64/250425/board.json", Other())], "906114-01", "PAL")[0];
        Assert.Equal("Commodore/C64/250425/board.json", known.BoardIdentity);
    }

    // Only what is picked; nothing picked, or a component the board does not have, changes nothing.
    [Fact]
    public void Only_the_parts_picked_are_copied_and_nothing_picked_is_the_same_board()
    {
        BoardData board = This();
        ComponentCopySource source = ComponentCopy.Sources([("Other", Other())], "906114-01", "PAL")[0];

        BoardData copied = ComponentCopy.Apply(board, "U17", "PAL", source, [ComponentCopyPart.Package]);

        Assert.Equal("DIP-28", copied.Components[0].Package);
        Assert.Equal("old", Assert.Single(copied.ComponentPins).Name);
        Assert.Same(board, ComponentCopy.Apply(board, "U17", "PAL", source, []));
        Assert.Same(board, ComponentCopy.Apply(board, "U99", "PAL", source, [ComponentCopyPart.Package]));
    }
}
