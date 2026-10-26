using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentPaneEdits - what the component pane changes in a draft (CRT 4.0 step 2, 2026-10-25).
// The pane edits a component AS SHOWN - the region CRT shows, by the component list's rule - so
// every edit must leave the other region's rows, and every other component's, exactly as they
// were; and an edit that changes nothing must return the very same board, which the draft's
// history takes as no step.
// ###########################################################################################
public sealed class ComponentPaneEditsTests
{
    // U1 in two regions, U8 in every region, with pins and parts of both regions and another component's.
    private static BoardData Board() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U1", Region = "PAL", FriendlyName = "VIC", TechnicalNameOrValue = "6569", Category = "IC" },
            new ComponentEntry { BoardLabel = "U1", Region = "NTSC", FriendlyName = "VIC", TechnicalNameOrValue = "6567", Category = "IC" },
            new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", Package = "DIP-40", Description = "The processor", RepairNotes = "Socketed" },
        ],
        ComponentPins =
        [
            new ComponentPinEntry { BoardLabel = "U1", Region = "PAL", Pin = "1", Name = "D6" },
            new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS", Direction = "Ground" },
            new ComponentPinEntry { BoardLabel = "U1", Region = "NTSC", Pin = "1", Name = "D6 (NTSC)" },
            new ComponentPinEntry { BoardLabel = "U8", Pin = "40", Name = "RES" },
        ],
        ComponentReplacements =
        [
            new ComponentReplacementEntry { BoardLabel = "U1", Region = "NTSC", PartNumber = "8562" },
            new ComponentReplacementEntry { BoardLabel = "U8", PartNumber = "8500", Kind = "Drop-in" },
        ],
        ComponentLinks =
        [
            new ComponentLinkEntry { BoardLabel = "U8", Name = "Datasheet", Url = "https://example.org/6510.pdf" },
            new ComponentLinkEntry { BoardLabel = "U1", Name = "VIC article", Url = "https://example.org/vic" },
        ],
    };

    // ---- the row shown ----------------------------------------------------------------------

    [Fact]
    public void The_row_shown_is_the_one_of_the_region_shown()
    {
        Assert.Equal("6567", ComponentPaneEdits.RowShown(Board(), "u1", "NTSC")!.TechnicalNameOrValue);
        Assert.Equal("6569", ComponentPaneEdits.RowShown(Board(), "U1", "PAL")!.TechnicalNameOrValue);
    }

    // A component in every region is shown in each; a region the component has no row of falls back
    // to its first row - what the component editor's window has always opened on.
    [Fact]
    public void A_component_in_every_region_is_shown_everywhere_and_a_missing_region_falls_back_to_the_first_row()
    {
        Assert.Equal("CPU", ComponentPaneEdits.RowShown(Board(), "U8", "PAL")!.FriendlyName);
        Assert.Equal("6569", ComponentPaneEdits.RowShown(Board(), "U1", "SECAM")!.TechnicalNameOrValue);
        Assert.Null(ComponentPaneEdits.RowShown(Board(), "U99", "PAL"));
        Assert.Null(ComponentPaneEdits.RowShown(Board(), "  ", "PAL"));
    }

    // ---- fields -----------------------------------------------------------------------------

    [Fact]
    public void A_field_changes_the_row_of_the_region_shown_and_nothing_else()
    {
        BoardData before = Board();

        BoardData after = ComponentPaneEdits.WithField(before, "U1", "NTSC", ComponentField.Description, "  Video chip  ");

        Assert.Equal("Video chip", after.Components[1].Description);
        Assert.Equal(string.Empty, after.Components[0].Description);

        // Every other field of the row, and every other row, as they were.
        Assert.Equal("6567", after.Components[1].TechnicalNameOrValue);
        Assert.Equal("NTSC", after.Components[1].Region);
        Assert.Same(before.Components[0], after.Components[0]);
        Assert.Same(before.Components[2], after.Components[2]);
        Assert.Same(before.ComponentPins, after.ComponentPins);
    }

    [Theory]
    [InlineData(ComponentField.FriendlyName, "Processor")]
    [InlineData(ComponentField.TechnicalNameOrValue, "6510")]
    [InlineData(ComponentField.PartNumber, "906122")]
    [InlineData(ComponentField.Category, "CPU")]
    [InlineData(ComponentField.Package, "DIP-28")]
    [InlineData(ComponentField.Description, "Runs the show")]
    [InlineData(ComponentField.RepairNotes, "Check the clock first")]
    public void Every_field_is_set_and_read_back(ComponentField field, string value)
    {
        BoardData after = ComponentPaneEdits.WithField(Board(), "U8", "PAL", field, value);

        Assert.Equal(value, ComponentPaneEdits.ValueOf(after.Components[2], field));
    }

    // The same value - spaces around it aside - is no change: the same board comes back, which the
    // draft's history takes as no step at all.
    [Fact]
    public void A_field_set_to_what_it_says_already_returns_the_same_board()
    {
        BoardData before = Board();

        Assert.Same(before, ComponentPaneEdits.WithField(before, "U8", "PAL", ComponentField.FriendlyName, " CPU "));
        Assert.Same(before, ComponentPaneEdits.WithField(before, "U99", "PAL", ComponentField.FriendlyName, "Anything"));
    }

    // The short description is one line; repair notes keep their lines, each ended by "\n" alone.
    [Fact]
    public void One_line_fields_lose_their_line_breaks_and_repair_notes_keep_theirs()
    {
        Assert.Equal("first second", ComponentPaneEdits.Normalised(ComponentField.Description, "first\r\nsecond\n"));
        Assert.Equal("first\nsecond\nthird", ComponentPaneEdits.Normalised(ComponentField.RepairNotes, "  first\r\nsecond\rthird \n"));
    }

    [Fact]
    public void Every_field_has_a_name_for_the_pane_and_the_history()
    {
        foreach (ComponentField field in Enum.GetValues<ComponentField>())
        {
            Assert.False(string.IsNullOrWhiteSpace(ComponentPaneEdits.NameOf(field)));
        }

        Assert.Equal("Short description", ComponentPaneEdits.NameOf(ComponentField.Description));
    }

    // ---- pins -------------------------------------------------------------------------------

    [Fact]
    public void The_pins_shown_are_the_components_in_the_region_shown()
    {
        Assert.Equal(["D6 (NTSC)"], ComponentPaneEdits.PinsOf(Board(), "U1", "NTSC").Select(pin => pin.Name));
        Assert.Equal(["VSS", "RES"], ComponentPaneEdits.PinsOf(Board(), "U8", "NTSC").Select(pin => pin.Name));
    }

    // ###########################################################################################
    // The pins of the region shown are replaced where the first of them stood - the other region's
    // pins and the other component's stay, in place. A new pin is filed under the component's own
    // spelling of its label, in the region it was given.
    // ###########################################################################################
    [Fact]
    public void Pins_replace_the_regions_pins_in_place_and_leave_the_rest()
    {
        BoardData after = ComponentPaneEdits.WithPins(Board(), "u1", "PAL",
        [
            new ComponentPinEntry { Region = "PAL", Pin = "1", Name = "D6" },
            new ComponentPinEntry { Region = "PAL", Pin = "2", Name = "D5", Direction = "Input/output" },
        ]);

        Assert.Equal(
            ["U1|PAL|1|D6", "U1|PAL|2|D5", "U8||1|VSS", "U1|NTSC|1|D6 (NTSC)", "U8||40|RES"],
            after.ComponentPins.Select(pin => $"{pin.BoardLabel}|{pin.Region}|{pin.Pin}|{pin.Name}"));
    }

    // A pin row with nothing in it is not kept; the same pins are no change.
    [Fact]
    public void An_empty_pin_is_left_out_and_the_same_pins_are_no_change()
    {
        BoardData before = Board();

        BoardData same = ComponentPaneEdits.WithPins(before, "U8", "PAL",
        [
            new ComponentPinEntry { Pin = " 1 ", Name = "VSS", Direction = "Ground" },
            new ComponentPinEntry(),
            new ComponentPinEntry { Pin = "40", Name = "RES" },
        ]);

        Assert.Same(before, same);
    }

    [Fact]
    public void Removing_every_pin_of_the_region_keeps_the_other_regions()
    {
        BoardData after = ComponentPaneEdits.WithPins(Board(), "U1", "PAL", []);

        Assert.DoesNotContain(after.ComponentPins, pin => pin.Region == "PAL");
        Assert.Contains(after.ComponentPins, pin => pin.Name == "D6 (NTSC)");
        Assert.Equal(3, after.ComponentPins.Count);
    }

    // A component with no pins yet gets its first ones at the end of the list.
    [Fact]
    public void A_components_first_pins_go_at_the_end()
    {
        BoardData board = Board();
        board.Components.Add(new ComponentEntry { BoardLabel = "U9", FriendlyName = "RAM" });

        BoardData after = ComponentPaneEdits.WithPins(board, "U9", "PAL", [new ComponentPinEntry { Pin = "1", Name = "A0" }]);

        Assert.Equal("U9", after.ComponentPins[^1].BoardLabel);
        Assert.Equal(string.Empty, after.ComponentPins[^1].Region);
    }

    // ###########################################################################################
    // "Add the pins of the case": every number of the case not listed yet, added with nothing but its
    // number, all in number order - "2" before "10", the pins already named kept as they are.
    // ###########################################################################################
    [Fact]
    public void Every_pin_of_the_case_is_added_in_number_order_keeping_the_named_ones()
    {
        List<ComponentPinEntry> pins = ComponentPaneEdits.WithEveryPinOf(
            [new ComponentPinEntry { Pin = "10", Name = "GND" }, new ComponentPinEntry { Pin = "1", Name = "VCC" }],
            14,
            "PAL");

        Assert.Equal(Enumerable.Range(1, 14).Select(number => number.ToString(System.Globalization.CultureInfo.InvariantCulture)), pins.Select(pin => pin.Pin));
        Assert.Equal("VCC", pins[0].Name);
        Assert.Equal("GND", pins[9].Name);
        Assert.Equal("PAL", pins[1].Region);
    }

    [Fact]
    public void Pins_are_ordered_by_number_then_by_name_with_unnumbered_ones_last()
    {
        string?[] pins = ["10", "", "A1", "2", null, "1", "B2"];

        Assert.Equal(["1", "2", "10", "A1", "B2", "", null], pins.Order(ComponentPaneEdits.PinOrder));
    }

    // ---- replacement parts and links ---------------------------------------------------------

    [Fact]
    public void Replacement_parts_replace_the_regions_parts_and_leave_the_rest()
    {
        BoardData after = ComponentPaneEdits.WithReplacements(Board(), "U1", "PAL",
        [
            new ComponentReplacementEntry { Region = "PAL", PartNumber = "8565", Kind = "Needs changes", Changes = "12 V to 9 V", Source = "Forum" },
        ]);

        Assert.Equal(
            ["U1|NTSC|8562", "U8||8500", "U1|PAL|8565"],
            after.ComponentReplacements.Select(row => $"{row.BoardLabel}|{row.Region}|{row.PartNumber}"));
        Assert.Equal("12 V to 9 V", after.ComponentReplacements[^1].Changes);
    }

    [Fact]
    public void The_same_replacement_parts_are_no_change()
    {
        BoardData before = Board();

        Assert.Same(before, ComponentPaneEdits.WithReplacements(before, "U8", "PAL", [new ComponentReplacementEntry { PartNumber = " 8500 ", Kind = "Drop-in" }]));
    }

    // A link has no region: it is the component's in every region.
    [Fact]
    public void Links_replace_the_components_links_in_place()
    {
        BoardData after = ComponentPaneEdits.WithLinks(Board(), "U8",
        [
            new ComponentLinkEntry { Name = "Datasheet", Url = "https://example.org/6510.pdf" },
            new ComponentLinkEntry { Name = "Pinout", Url = "https://example.org/pinout" },
            new ComponentLinkEntry(),
        ]);

        Assert.Equal(["U8|Datasheet", "U8|Pinout", "U1|VIC article"], after.ComponentLinks.Select(row => $"{row.BoardLabel}|{row.Name}"));
    }

    [Fact]
    public void The_same_links_are_no_change_and_a_blank_label_changes_nothing()
    {
        BoardData before = Board();

        Assert.Same(before, ComponentPaneEdits.WithLinks(before, "U8", [new ComponentLinkEntry { Name = "Datasheet", Url = "https://example.org/6510.pdf" }]));
        Assert.Same(before, ComponentPaneEdits.WithLinks(before, " ", [new ComponentLinkEntry { Name = "x", Url = "y" }]));
    }

    // ###########################################################################################
    // *** A LABEL SPELT IN OTHER CAPITALS IS NO CHANGE (code review, 2026-10-26). *** Rows filed under
    // "u8" for the component "U8" - typed so in the table, or an older board - are the component's,
    // and the pane hands them back at every commit. They were compared by exact text, so leaving any
    // box rewrote them all under "U8": an undo step, orange cells and changed rows nobody made. A real
    // change keeps the spelling the rows already use.
    // ###########################################################################################
    [Fact]
    public void Rows_filed_under_the_label_in_other_capitals_handed_back_unchanged_are_no_change()
    {
        BoardData before = new()
        {
            Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU" }],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "u8", Pin = "1", Name = "VSS" }],
            ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "u8", PartNumber = "8500" }],
            ComponentLinks =
            [
                new ComponentLinkEntry { BoardLabel = "u8", Name = "Datasheet", Url = "https://example.org/6510.pdf" },
                new ComponentLinkEntry { BoardLabel = "U8 ", Name = "Pinout", Url = "https://example.org/pinout" },
            ],
        };

        Assert.Same(before, ComponentPaneEdits.WithPins(before, "U8", "PAL", ComponentPaneEdits.PinsOf(before, "U8", "PAL")));
        Assert.Same(before, ComponentPaneEdits.WithReplacements(before, "U8", "PAL", ComponentPaneEdits.ReplacementsOf(before, "U8", "PAL")));
        Assert.Same(before, ComponentPaneEdits.WithLinks(before, "U8", ComponentPaneEdits.LinksOf(before, "U8")));

        BoardData changed = ComponentPaneEdits.WithPins(before, "U8", "PAL", [new ComponentPinEntry { Pin = "1", Name = "GND" }]);
        Assert.Equal("u8", Assert.Single(changed.ComponentPins).BoardLabel);
    }

    // Every edit keeps the board's other sections - its caption and date included.
    [Fact]
    public void An_edit_keeps_the_rest_of_the_board()
    {
        BoardData before = new()
        {
            HardwareName = "Commodore 64",
            BoardName = "250407",
            RevisionDate = "2026-October-4",
            Components = [new ComponentEntry { BoardLabel = "U8" }],
            Credits = [new CreditEntry { NameOrHandle = "Anna" }],
        };

        BoardData after = ComponentPaneEdits.WithField(before, "U8", "PAL", ComponentField.FriendlyName, "CPU");

        Assert.Equal("Commodore 64", after.HardwareName);
        Assert.Equal("250407", after.BoardName);
        Assert.Equal("2026-October-4", after.RevisionDate);
        Assert.Same(before.Credits, after.Credits);
    }
}
