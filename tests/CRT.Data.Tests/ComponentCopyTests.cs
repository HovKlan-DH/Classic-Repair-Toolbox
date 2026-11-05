using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentCopy - "Copy from another board" (CRT 4.0 step 2, slice 5): the components of the other
// boards with the same part number, what each has, and what a copy puts into the component shown -
// its package and description, its pins, replacement parts and links, filed under this component's
// label and region, and nothing board-specific.
// ###########################################################################################
public sealed class ComponentCopyTests
{
    // The other board: U7, the PLA, with everything filled - and another chip.
    private static BoardData Other() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U7", FriendlyName = "PLA", PartNumber = "906114-01", Category = "IC", Package = "DIP-28", Description = "Glue logic" },
            new ComponentEntry { BoardLabel = "U6", FriendlyName = "CPU", PartNumber = "6510", Category = "IC" },
        ],
        ComponentPins =
        [
            new ComponentPinEntry { BoardLabel = "U7", Pin = "1", Name = "FE", Direction = "Input" },
            new ComponentPinEntry { BoardLabel = "U7", Pin = "28", Name = "VCC", Direction = "Power" },
        ],
        ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "U7", PartNumber = "PLAnkton", Kind = "Drop-in replacement" }],
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
                new ComponentEntry { BoardLabel = "U20", PartNumber = "4066", Region = "NTSC", Description = "NTSC's" },
            ],
        };

        ComponentCopySource ntsc = Assert.Single(ComponentCopy.Sources([("Other", other)], "6567", "PAL"));
        Assert.Equal(("6567", "NTSC"), (ntsc.Row.PartNumber, ntsc.Row.Region));

        ComponentCopySource pal = Assert.Single(ComponentCopy.Sources([("Other", other)], "6569", "NTSC"));
        Assert.Equal(("6569", "PAL"), (pal.Row.PartNumber, pal.Row.Region));

        // Both regions the same chip: the region's own row.
        ComponentCopySource both = Assert.Single(ComponentCopy.Sources([("Other", other)], "4066", "NTSC"));
        Assert.Equal("NTSC's", both.Row.Description);
    }

    [Fact]
    public void A_source_offers_what_it_holds()
    {
        ComponentCopySource source = ComponentCopy.Sources([("Other", Other())], "906114-01", "PAL")[0];

        Assert.Equal(
            [ComponentCopyPart.Package, ComponentCopyPart.Description, ComponentCopyPart.Pins, ComponentCopyPart.Replacements, ComponentCopyPart.Links],
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
        Assert.Equal(("DIP-28", "Glue logic", "PAL", "IC"), (row.Package, row.Description, row.Region, row.Category));
        Assert.Equal(["1:FE", "28:VCC"], copied.ComponentPins.Select(pin => $"{pin.Pin}:{pin.Name}"));
        Assert.All(copied.ComponentPins, pin => Assert.Equal(("U17", "PAL"), (pin.BoardLabel, pin.Region)));
        Assert.Equal(("U17", "PLAnkton", "PAL"), (copied.ComponentReplacements[0].BoardLabel, copied.ComponentReplacements[0].PartNumber, copied.ComponentReplacements[0].Region));
        Assert.Equal(("U17", "Datasheet", "PAL"), (copied.ComponentLinks[0].BoardLabel, copied.ComponentLinks[0].Name, copied.ComponentLinks[0].Region));
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
            ComponentPaneEdits.WithField(published, "U7", "", ComponentField.Description, "Glue logic, and more"),
            "U7",
            "",
            [new ComponentPinEntry { Pin = "1", Name = "FE", Direction = "Input" }]);

        Assert.Equal([ComponentCopyPart.Description, ComponentCopyPart.Pins], ComponentCopy.Changed(draft, published, "U7", "PAL"));
        Assert.Empty(ComponentCopy.Changed(published, published, "U7", "PAL"));
        Assert.Equal(ComponentCopy.Available(ComponentCopy.Sources([("x", draft)], "906114-01", "PAL")[0]), ComponentCopy.Changed(draft, null, "U7", "PAL"));
        Assert.Empty(ComponentCopy.Changed(draft, published, "U99", "PAL"));

        ComponentCopySource known = ComponentCopy.Sources([("Commodore 64 - 250425", "Commodore/C64/250425/board.json", Other())], "906114-01", "PAL")[0];
        Assert.Equal("Commodore/C64/250425/board.json", known.BoardIdentity);
    }

    // ###########################################################################################
    // *** A ROW OF ONE REGION ON A COMPONENT OF EVERY REGION (code review, 2026-11-02) - the usual shape
    // on a regional board. *** Its rows were read by the region of the component's ROW, blank, which
    // takes in only rows of every region: the page showed U7's PAL link while Copy never offered,
    // counted or copied it, and "Also on other boards" saw no change when only such a link was added.
    // ###########################################################################################
    private static BoardData Regional() => new()
    {
        Regions = ["PAL", "NTSC"],
        Components = [new ComponentEntry { BoardLabel = "U7", PartNumber = "906114-01", Category = "IC" }],
        ComponentLinks =
        [
            new ComponentLinkEntry { BoardLabel = "U7", Name = "Datasheet", Url = "https://example.org/pla" },
            new ComponentLinkEntry { BoardLabel = "U7", Name = "PAL notes", Url = "https://example.org/pal", Region = "PAL" },
            new ComponentLinkEntry { BoardLabel = "U7", Name = "NTSC notes", Url = "https://example.org/ntsc", Region = "NTSC" },
        ],
        ComponentPins = [new ComponentPinEntry { BoardLabel = "U7", Pin = "1", Name = "PAL's", Region = "PAL" }],
        ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "U7", PartNumber = "PLAnkton", Region = "PAL" }],
    };

    [Fact]
    public void Rows_of_the_region_shown_are_offered_from_a_component_of_every_region()
    {
        ComponentCopySource pal = ComponentCopy.Sources([("Other", Regional())], "906114-01", "PAL")[0];

        Assert.Equal([ComponentCopyPart.Pins, ComponentCopyPart.Replacements, ComponentCopyPart.Links], ComponentCopy.Available(pal));
        Assert.Equal("Links (2)", ComponentCopy.Describe(pal, ComponentCopyPart.Links));
        Assert.Equal("Pins (1)", ComponentCopy.Describe(pal, ComponentCopyPart.Pins));

        ComponentCopySource ntsc = ComponentCopy.Sources([("Other", Regional())], "906114-01", "NTSC")[0];

        Assert.Equal([ComponentCopyPart.Links], ComponentCopy.Available(ntsc));
    }

    // Copied, a row of one region stays of one region - the one shown - where the board has regions at
    // all: "PAL" on a board with none would turn its PAL/NTSC buttons on.
    [Fact]
    public void A_row_of_one_region_stays_of_one_region_where_the_board_has_regions()
    {
        ComponentCopySource pal = ComponentCopy.Sources([("Other", Regional())], "906114-01", "PAL")[0];

        var regional = new BoardData { Regions = ["PAL", "NTSC"], Components = [new ComponentEntry { BoardLabel = "U17", Category = "IC" }] };
        BoardData copied = ComponentCopy.Apply(regional, "U17", "PAL", pal, [ComponentCopyPart.Links]);

        Assert.Equal(["Datasheet|", "PAL notes|PAL"], copied.ComponentLinks.Select(link => $"{link.Name}|{link.Region}"));

        var none = new BoardData { Components = [new ComponentEntry { BoardLabel = "U17", Category = "IC" }] };
        BoardData plain = ComponentCopy.Apply(none, "U17", "PAL", pal, [ComponentCopyPart.Links]);

        Assert.Equal(["Datasheet|", "PAL notes|"], plain.ComponentLinks.Select(link => $"{link.Name}|{link.Region}"));
        Assert.False(BoardRegions.Has(plain));
    }

    // ###########################################################################################
    // *** ROWS OF ONE REGION REPLACE ONLY THAT REGION'S (code review, 2026-11-02). *** Filed under the
    // region shown, they also took the place of the target's rows for every region - which the other
    // region shows too, so it lost its pins and links without a word.
    // ###########################################################################################
    [Fact]
    public void Rows_of_one_region_leave_the_other_region_what_it_showed()
    {
        ComponentCopySource pal = ComponentCopy.Sources([("Other", Regional())], "906114-01", "PAL")[0];

        var target = new BoardData
        {
            Regions = ["PAL", "NTSC"],
            Components = [new ComponentEntry { BoardLabel = "U17", Category = "IC" }],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U17", Pin = "9", Name = "Old" }],
            ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "U17", PartNumber = "Old part" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U17", Name = "Old link", Url = "https://example.org/old" }],
        };

        BoardData copied = ComponentCopy.Apply(target, "U17", "PAL", pal, [ComponentCopyPart.Pins, ComponentCopyPart.Replacements, ComponentCopyPart.Links]);

        Assert.Equal(["1"], ComponentPaneEdits.PinsOf(copied, "U17", "PAL").Select(pin => pin.Pin));
        Assert.Equal(["9"], ComponentPaneEdits.PinsOf(copied, "U17", "NTSC").Select(pin => pin.Pin));

        Assert.Equal(["PLAnkton"], ComponentPaneEdits.ReplacementsOf(copied, "U17", "PAL").Select(row => row.PartNumber));
        Assert.Equal(["Old part"], ComponentPaneEdits.ReplacementsOf(copied, "U17", "NTSC").Select(row => row.PartNumber));

        // A link for every region among what is copied stays with the region shown here: NTSC goes on
        // showing what it showed, nothing beside it.
        Assert.Equal(["Datasheet", "PAL notes"], ComponentPaneEdits.LinksOf(copied, "U17", "PAL").Select(row => row.Name));
        Assert.Equal(["Old link"], ComponentPaneEdits.LinksOf(copied, "U17", "NTSC").Select(row => row.Name));
    }

    // A component of its own in each region: its row of the region shown says where the copy goes, and
    // the pins for every region stay for the other.
    [Fact]
    public void A_component_with_a_row_per_region_keeps_its_rows_for_every_region_in_the_other()
    {
        ComponentCopySource pal = ComponentCopy.Sources([("Other", Regional())], "906114-01", "PAL")[0];

        var target = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U17", Category = "IC", Region = "PAL" },
                new ComponentEntry { BoardLabel = "U17", Category = "IC", Region = "NTSC" },
            ],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U17", Pin = "9", Name = "Old" }],
        };

        BoardData copied = ComponentCopy.Apply(target, "U17", "PAL", pal, [ComponentCopyPart.Pins]);

        Assert.Equal(["1"], ComponentPaneEdits.PinsOf(copied, "U17", "PAL").Select(pin => pin.Pin));
        Assert.Equal(["9"], ComponentPaneEdits.PinsOf(copied, "U17", "NTSC").Select(pin => pin.Pin));
    }

    // A component only one region has: there is no other region to keep anything for.
    [Fact]
    public void A_component_of_one_region_has_no_other_region_to_keep_rows_for()
    {
        ComponentCopySource pal = ComponentCopy.Sources([("Other", Regional())], "906114-01", "PAL")[0];

        var target = new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U17", Category = "IC", Region = "PAL" }],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U17", Pin = "9", Name = "Old" }],
        };

        BoardData copied = ComponentCopy.Apply(target, "U17", "PAL", pal, [ComponentCopyPart.Pins]);

        Assert.Equal(["1|PAL"], copied.ComponentPins.Select(pin => $"{pin.Pin}|{pin.Region}"));
    }

    // Rows for every region replace rows for every region: the same chip, in both.
    [Fact]
    public void Rows_for_every_region_copied_replace_them_in_every_region()
    {
        var source = new BoardData
        {
            Regions = ["PAL", "NTSC"],
            Components = [new ComponentEntry { BoardLabel = "U7", PartNumber = "906114-01", Category = "IC" }],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U7", Pin = "1", Name = "New" }],
        };

        var target = new BoardData
        {
            Regions = ["PAL", "NTSC"],
            Components = [new ComponentEntry { BoardLabel = "U17", Category = "IC" }],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U17", Pin = "9", Name = "Old" }],
        };

        BoardData copied = ComponentCopy.Apply(target, "U17", "PAL", ComponentCopy.Sources([("Other", source)], "906114-01", "PAL")[0], [ComponentCopyPart.Pins]);

        Assert.Equal(["1|"], copied.ComponentPins.Select(pin => $"{pin.Pin}|{pin.Region}"));
    }

    // ###########################################################################################
    // *** ...WHEREVER THE TARGET TOLD THEM APART PER REGION TOO (code review, 2026-11-02). *** Copied in
    // PAL, the rows for every region replaced only what PAL listed, so the old NTSC rows stayed beside
    // them: NTSC showed pin 1 twice, the new one and its own. "Rows for every region copied replace them
    // in every region" is the promise; it now holds for the target's rows of one region as well.
    // ###########################################################################################
    private static ComponentCopySource ForEveryRegion() => ComponentCopy.Sources(
        [("Other", new BoardData
        {
            Regions = ["PAL", "NTSC"],
            Components = [new ComponentEntry { BoardLabel = "U7", PartNumber = "906114-01", Category = "IC" }],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U7", Pin = "1", Name = "New" }],
            ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "U7", PartNumber = "New part" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U7", Name = "New link", Url = "https://example.org/new" }],
        })],
        "906114-01",
        "PAL")[0];

    private static BoardData ToldApartPerRegion() => new()
    {
        Regions = ["PAL", "NTSC"],
        Components = [new ComponentEntry { BoardLabel = "U17", Category = "IC" }],
        ComponentPins =
        [
            new ComponentPinEntry { BoardLabel = "U17", Pin = "1", Name = "OldP", Region = "PAL" },
            new ComponentPinEntry { BoardLabel = "U17", Pin = "1", Name = "OldN", Region = "NTSC" },
        ],
        ComponentReplacements =
        [
            new ComponentReplacementEntry { BoardLabel = "U17", PartNumber = "OldP", Region = "PAL" },
            new ComponentReplacementEntry { BoardLabel = "U17", PartNumber = "OldN", Region = "NTSC" },
        ],
        ComponentLinks =
        [
            new ComponentLinkEntry { BoardLabel = "U17", Name = "OldP", Url = "https://example.org/p", Region = "PAL" },
            new ComponentLinkEntry { BoardLabel = "U17", Name = "OldN", Url = "https://example.org/n", Region = "NTSC" },
        ],
    };

    [Theory]
    [InlineData("PAL")]
    [InlineData("NTSC")]
    public void Rows_for_every_region_copied_replace_the_targets_rows_of_each_region(string shown)
    {
        BoardData copied = ComponentCopy.Apply(
            ToldApartPerRegion(),
            "U17",
            shown,
            ForEveryRegion(),
            [ComponentCopyPart.Pins, ComponentCopyPart.Replacements, ComponentCopyPart.Links]);

        Assert.Equal(["1:New:"], copied.ComponentPins.Select(pin => $"{pin.Pin}:{pin.Name}:{pin.Region}"));
        Assert.Equal(["New part:"], copied.ComponentReplacements.Select(row => $"{row.PartNumber}:{row.Region}"));
        Assert.Equal(["New link:"], copied.ComponentLinks.Select(row => $"{row.Name}:{row.Region}"));

        Assert.Equal(["New"], ComponentPaneEdits.PinsOf(copied, "U17", "PAL").Select(pin => pin.Name));
        Assert.Equal(["New"], ComponentPaneEdits.PinsOf(copied, "U17", "NTSC").Select(pin => pin.Name));
    }

    // A second copy changes nothing, so no empty step in the draft's history.
    [Fact]
    public void The_same_rows_for_every_region_copied_again_are_no_change()
    {
        List<ComponentCopyPart> parts = [ComponentCopyPart.Pins, ComponentCopyPart.Replacements, ComponentCopyPart.Links];

        BoardData copied = ComponentCopy.Apply(ToldApartPerRegion(), "U17", "PAL", ForEveryRegion(), parts);

        Assert.Same(copied, ComponentCopy.Apply(copied, "U17", "PAL", ForEveryRegion(), parts));
    }

    // Another component's rows of the other region are its own.
    [Fact]
    public void Rows_for_every_region_copied_leave_another_components_rows_alone()
    {
        BoardData target = ToldApartPerRegion().WithComponentSections(
            componentPins: [.. ToldApartPerRegion().ComponentPins, new ComponentPinEntry { BoardLabel = "U18", Pin = "1", Name = "Mine", Region = "NTSC" }]);

        BoardData copied = ComponentCopy.Apply(target, "U17", "PAL", ForEveryRegion(), [ComponentCopyPart.Pins]);

        Assert.Contains(copied.ComponentPins, pin => pin is { BoardLabel: "U18", Name: "Mine", Region: "NTSC" });
    }

    // ###########################################################################################
    // *** A COMPONENT RENAMED IS THE SAME COMPONENT (code review, 2026-11-02). *** No published row under
    // its new label, it read as a component the published board lacks: "What you changed here" offered
    // everything it holds, every part ticked, so pressing the button overwrote other boards' drafts with
    // parts nobody had touched. Found through its old label; what really changed is still offered.
    // ###########################################################################################
    [Fact]
    public void A_component_only_renamed_has_changed_nothing_the_other_boards_can_take()
    {
        BoardData published = Other();
        BoardData draft = ComponentPaneEdits.WithLabel(published, "U7", "U7X");

        Assert.Empty(ComponentCopy.Changed(draft, published, "U7X", "PAL"));
    }

    [Fact]
    public void What_else_changed_of_a_renamed_component_is_offered()
    {
        BoardData published = Other();
        BoardData draft = ComponentPaneEdits.WithLinks(
            ComponentPaneEdits.WithLabel(published, "U7", "U7X"),
            "U7X",
            "",
            [new ComponentLinkEntry { Name = "Datasheet", Url = "https://example.org/other" }]);

        Assert.Equal([ComponentCopyPart.Links], ComponentCopy.Changed(draft, published, "U7X", "PAL"));
    }

    // Published with no such label and no rename to find it by: every part it holds, as before.
    [Fact]
    public void A_component_the_published_board_lacks_still_offers_all_it_holds()
    {
        BoardData draft = ComponentPaneEdits.WithField(
            ComponentPaneEdits.WithNewComponent(Other(), "U99", "", "IC"), "U99", "", ComponentField.Package, "DIP-8");

        Assert.Equal([ComponentCopyPart.Package], ComponentCopy.Changed(draft, Other(), "U99", "PAL"));
    }

    [Fact]
    public void A_link_of_the_region_shown_added_alone_is_what_changed()
    {
        BoardData published = Regional();
        BoardData draft = ComponentPaneEdits.WithLinks(
            published,
            "U7",
            "PAL",
            [.. ComponentPaneEdits.LinksOf(published, "U7", "PAL"), new ComponentLinkEntry { Name = "More", Url = "https://example.org/more", Region = "PAL" }]);

        Assert.Equal([ComponentCopyPart.Links], ComponentCopy.Changed(draft, published, "U7", "PAL"));
        Assert.Empty(ComponentCopy.Changed(draft, published, "U7", "NTSC"));
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
