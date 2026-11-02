using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftChangeSummary - what a draft changes, in words a person reads (CRT 4.0 step 1, 2026-10-25:
// "Review changes shows a readable summary, and the Submit dialog shows the same"). Grouped per
// component and for the board's own parts, counted by BoardDataDiffer's rules - so one change here
// is one row the table and the row count call changed.
// ###########################################################################################
public sealed class DraftChangeSummaryTests
{
    private static ComponentEntry Component(string label, string friendlyName = "", string description = "", string region = "", string partNumber = "") =>
        new() { BoardLabel = label, FriendlyName = friendlyName, Category = "IC", Description = description, Region = region, PartNumber = partNumber };

    private static ComponentImageEntry Picture(string label, string name, string file = "") =>
        new() { BoardLabel = label, Name = name, File = file.Length == 0 ? $"Component images/{label}-{name}.png" : file };

    private static ComponentImageEntry Baseline(string label, string pin) =>
        new() { BoardLabel = label, Pin = pin, Name = "Clock", TimeDiv = "1uS", File = $"Scope baseline/{label}-{pin}.png" };

    private static ComponentHighlightEntry Mark(string schematic, string label) =>
        new() { SchematicName = schematic, BoardLabel = label, X = "1", Y = "2", Width = "3", Height = "4" };

    private static BoardContent Content(BoardData board, params KiCadCalibrationEntry[] calibrations) => new(board, calibrations, []);

    private static BoardData Board(params ComponentEntry[] components) => new()
    {
        HardwareName = "Commodore 64",
        BoardName = "250407",
        Components = [.. components],
    };

    private static DraftChangeGroup Single(DraftChangeSummary summary, string label) =>
        Assert.Single(summary.Groups, group => group.Label == label);

    [Fact]
    public void A_draft_that_changes_nothing_has_nothing_to_say()
    {
        BoardData board = Board(Component("U8", "CPU"));

        Assert.True(DraftChangeSummary.Build(Content(board), Content(board)).IsEmpty);
    }

    // The fields in plain words, in the order the row has them.
    [Fact]
    public void A_changed_component_says_which_of_its_fields_changed_under_its_name()
    {
        DraftChangeSummary summary = DraftChangeSummary.Build(
            Content(Board(Component("U8", "CPU", description: "6510"))),
            Content(Board(Component("U8", "CPU", description: "6510 CPU", partNumber: "MOS 6510"))));

        DraftChangeGroup group = Single(summary, "U8");

        Assert.Equal("U8 - CPU", group.Title);
        Assert.Equal(DraftComponentChange.Changed, group.Change);
        Assert.Equal(["Part number and description changed"], group.Lines);
    }

    // A component with more than one region says which region a change is in.
    [Fact]
    public void A_change_to_one_region_of_a_component_names_the_region()
    {
        DraftChangeSummary summary = DraftChangeSummary.Build(
            Content(Board(Component("U1", "VIC", "PAL chip", "PAL"), Component("U1", "VIC", "NTSC chip", "NTSC"))),
            Content(Board(Component("U1", "VIC", "PAL chip", "PAL"), Component("U1", "VIC", "NTSC VIC-II", "NTSC"))));

        Assert.Equal(["Description changed (NTSC)"], Single(summary, "U1").Lines);
    }

    [Fact]
    public void A_new_component_says_so_and_counts_what_it_brought()
    {
        BoardData published = Board(Component("U8", "CPU"));
        BoardData draft = Board(Component("U8", "CPU"), Component("U9", "SID"));
        draft.ComponentImages.AddRange([Picture("U9", "Top"), Picture("U9", "Bottom"), Baseline("U9", "6")]);

        DraftChangeGroup group = Single(DraftChangeSummary.Build(Content(published), Content(draft)), "U9");

        Assert.Equal(DraftComponentChange.Added, group.Change);
        Assert.Equal(["New component", "2 pictures added", "1 oscilloscope baseline added"], group.Lines);
    }

    [Fact]
    public void A_removed_component_says_so_and_counts_what_went_with_it()
    {
        BoardData published = Board(Component("U8", "CPU"), Component("U9", "SID"));
        published.ComponentImages.Add(Picture("U9", "Top"));
        published.ComponentHighlights.Add(Mark("Top", "U9"));

        DraftChangeGroup group = Single(DraftChangeSummary.Build(Content(published), Content(Board(Component("U8", "CPU")))), "U9");

        Assert.Equal(DraftComponentChange.Removed, group.Change);
        Assert.Equal(["Removed", "1 picture removed", "1 mark on the schematics removed"], group.Lines);
    }

    // A relabelled component is ONE component (ComponentRelabel), its marks followed with it - not a
    // removal plus an addition.
    [Fact]
    public void A_relabelled_component_is_one_component_relabelled()
    {
        BoardData published = Board(Component("U8", "CPU"));
        published.ComponentHighlights.Add(Mark("Top", "U8"));

        BoardData draft = Board(Component("U08", "CPU"));
        draft.ComponentHighlights.Add(Mark("Top", "U08"));

        DraftChangeSummary summary = DraftChangeSummary.Build(Content(published), Content(draft));

        DraftChangeGroup group = Assert.Single(summary.Groups);
        Assert.Equal("U08", group.Label);
        Assert.Equal(["Relabelled from [U8]"], group.Lines);
    }

    // ###########################################################################################
    // *** RENAMED IN THE PANE, EVERY ROW OF IT MOVES (code review, 2026-11-02; ComponentPaneEdits.WithLabel) -
    // and a rename alone is all there is to say. *** Only its row and its marks were filed under the new
    // label, so it read "U08: Relabelled from [U8], 3 pictures added, 2 pins added" beside a group
    // "U8: 3 pictures removed, 2 pins removed", whose "Undo changes to [U8]" brought U8 back beside U08.
    // ###########################################################################################
    private static BoardData WithEverythingOf(string label)
    {
        BoardData board = Board(Component(label, "CPU"));
        board.ComponentImages.AddRange([Picture(label, "Top", "Component images/a.png"), Picture(label, "Bottom", "Component images/b.png"), Baseline(label, "6")]);
        board.ComponentHighlights.Add(Mark("Top", label));
        board.ComponentLocalFiles.Add(new ComponentLocalFileEntry { BoardLabel = label, Name = "Datasheet", File = "Datasheets/d.pdf" });
        board.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = label, Name = "Article", Url = "https://example.org/a" });
        board.ComponentPins.AddRange([new ComponentPinEntry { BoardLabel = label, Pin = "1", Name = "VSS" }, new ComponentPinEntry { BoardLabel = label, Pin = "2", Name = "VCC" }]);
        board.ComponentReplacements.Add(new ComponentReplacementEntry { BoardLabel = label, PartNumber = "6510" });
        return board;
    }

    [Fact]
    public void A_component_renamed_in_the_pane_says_only_that_with_no_group_under_its_old_label()
    {
        BoardData published = WithEverythingOf("U8");
        BoardData draft = ComponentPaneEdits.WithLabel(published, "U8", "U08");

        DraftChangeGroup group = Assert.Single(DraftChangeSummary.Build(Content(published), Content(draft)).Groups);

        Assert.Equal("U08", group.Label);
        Assert.Equal(DraftComponentChange.Changed, group.Change);
        Assert.Equal(["Relabelled from [U8]"], group.Lines);
    }

    // What else changed of it is still said - under its one group.
    [Fact]
    public void A_component_renamed_in_the_pane_and_given_another_pin_says_that_too()
    {
        BoardData published = WithEverythingOf("U8");
        BoardData draft = ComponentPaneEdits.WithLabel(published, "U8", "U08");
        draft.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U08", Pin = "3", Name = "D0" });

        DraftChangeGroup group = Assert.Single(DraftChangeSummary.Build(Content(published), Content(draft)).Groups);

        Assert.Equal(["Relabelled from [U8]", "1 pin added"], group.Lines);
    }

    // The table renames the component's own row and no other: those rows stay on the old label, compared
    // there - no change - and the rename alone is said, as before.
    [Fact]
    public void A_component_renamed_in_the_table_keeps_its_other_rows_on_the_old_label()
    {
        BoardData published = WithEverythingOf("U8");

        BoardData draft = WithEverythingOf("U8");
        draft = draft.WithComponentSections(components: [Component("U08", "CPU")]);

        DraftChangeGroup group = Assert.Single(DraftChangeSummary.Build(Content(published), Content(draft)).Groups);

        Assert.Equal("U08", group.Label);
        Assert.Equal(["Relabelled from [U8]"], group.Lines);
    }

    // Undoing the pane's rename from its one group brings it back whole, nothing beside it.
    [Fact]
    public void Undoing_the_one_group_of_a_component_renamed_in_the_pane_brings_it_back_whole()
    {
        BoardData published = WithEverythingOf("U8");
        BoardData draft = ComponentPaneEdits.WithLabel(published, "U8", "U08");

        BoardData reverted = ComponentRevert.ToPublished(draft, published, "U08");

        Assert.True(DraftChangeSummary.Build(Content(published), Content(reverted)).IsEmpty);
    }

    // ###########################################################################################
    // A second rectangle of a component on a schematic is work (code review, 2026-11-02): only the first
    // of them was compared. Their order alone is none.
    // ###########################################################################################
    [Fact]
    public void Every_rectangle_of_a_component_is_said_and_their_order_is_not()
    {
        static ComponentHighlightEntry Rectangle(string x) =>
            new() { SchematicName = "Top", BoardLabel = "U8", X = x, Y = "2", Width = "3", Height = "4" };

        BoardData Marked(params string[] xs)
        {
            BoardData board = Board(Component("U8", "CPU"));
            board.ComponentHighlights.AddRange(xs.Select(Rectangle));
            return board;
        }

        string[] Lines(BoardData draft) =>
            DraftChangeSummary.Build(Content(Marked("1", "10")), Content(draft)).Groups.SelectMany(group => group.Lines).ToArray();

        Assert.Equal(["1 mark on the schematics changed"], Lines(Marked("1", "20")));
        Assert.Equal(["2 marks on the schematics changed"], Lines(Marked("5", "20")));
        Assert.Equal(["1 mark on the schematics added"], Lines(Marked("1", "10", "20")));
        Assert.Equal(["1 mark on the schematics removed"], Lines(Marked("10")));
        Assert.Empty(Lines(Marked("10", "1")));
    }

    // ###########################################################################################
    // *** A RECTANGLE THAT ONLY FOLLOWED ITS SCHEMATIC IS THE BOARD'S CHANGE (code review, 2026-11-02). ***
    // A schematic renamed or deleted read "U8: 2 marks on the schematics changed" (or "removed") with a red
    // "Undo changes to [U8]" that did nothing and said nothing: ComponentRevert and BoardDraftSummary hold
    // it to be no change of U8. Said once, for the board - and every group left has something to undo.
    // ###########################################################################################
    private static ComponentHighlightEntry Rectangle(string schematic, string label, string x) =>
        new() { SchematicName = schematic, BoardLabel = label, X = x, Y = "2", Width = "3", Height = "4" };

    private static BoardData TopAndBack() => new()
    {
        HardwareName = "Commodore 64",
        BoardName = "250407",
        Schematics =
        [
            new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png" },
            new BoardSchematicEntry { SchematicName = "Back", SchematicImageFile = "back.png" },
        ],
        Components = [Component("U8", "CPU"), Component("U9", "SID")],
        ComponentHighlights = [Rectangle("Top", "U8", "1"), Rectangle("Top", "U8", "10"), Rectangle("Top", "U9", "5")],
    };

    private static void Every_component_group_has_something_to_undo(BoardData published, BoardData draft, DraftChangeSummary summary)
    {
        foreach (DraftChangeGroup group in summary.Groups.Where(group => group.IsComponent))
        {
            Assert.True(ComponentRevert.HasChanges(draft, published, group.Label), $"[{group.Label}] offers an undo with nothing to take back");
            Assert.NotNull(ComponentRevert.RowsToPublished(draft, published, group.Label));
        }
    }

    [Fact]
    public void A_schematic_renamed_is_said_once_and_no_component_has_a_change_to_undo()
    {
        BoardData published = TopAndBack();
        BoardData draft = TopAndBack()
            .WithSchematics(
            [
                new BoardSchematicEntry { SchematicName = "Front", SchematicImageFile = "top.png" },
                new BoardSchematicEntry { SchematicName = "Back", SchematicImageFile = "back.png" },
            ])
            .WithComponentHighlights([Rectangle("Front", "U8", "1"), Rectangle("Front", "U8", "10"), Rectangle("Front", "U9", "5")]);

        DraftChangeSummary summary = DraftChangeSummary.Build(Content(published), Content(draft));

        DraftChangeGroup board = Assert.Single(summary.Groups);
        Assert.False(board.IsComponent);
        Assert.Equal(["Schematic image [Top] renamed [Front]"], board.Lines);
    }

    [Fact]
    public void A_schematic_deleted_is_said_once_and_no_component_has_a_change_to_undo()
    {
        BoardData published = TopAndBack();
        BoardData draft = TopAndBack()
            .WithSchematics([new BoardSchematicEntry { SchematicName = "Back", SchematicImageFile = "back.png" }])
            .WithComponentHighlights([]);

        DraftChangeSummary summary = DraftChangeSummary.Build(Content(published), Content(draft));

        DraftChangeGroup board = Assert.Single(summary.Groups);
        Assert.False(board.IsComponent);
        Assert.Equal(["Schematic image [Top] deleted"], board.Lines);
    }

    // A rectangle really moved on the renamed schematic is its component's, and has its undo.
    [Fact]
    public void A_rectangle_moved_on_a_renamed_schematic_is_still_a_change_of_its_component()
    {
        BoardData published = TopAndBack();
        BoardData draft = TopAndBack()
            .WithSchematics(
            [
                new BoardSchematicEntry { SchematicName = "Front", SchematicImageFile = "top.png" },
                new BoardSchematicEntry { SchematicName = "Back", SchematicImageFile = "back.png" },
            ])
            .WithComponentHighlights([Rectangle("Front", "U8", "1"), Rectangle("Front", "U8", "99"), Rectangle("Front", "U9", "5")]);

        DraftChangeSummary summary = DraftChangeSummary.Build(Content(published), Content(draft));

        Assert.Equal(["Schematic image [Top] renamed [Front]"], summary.Groups[0].Lines);
        Assert.Equal(["1 mark on the schematics changed"], Single(summary, "U8").Lines);
        Assert.DoesNotContain(summary.Groups, group => group.Label == "U9");

        Every_component_group_has_something_to_undo(published, draft, summary);
    }

    // On the schematic left, deleted with another, a rectangle added elsewhere is still said - nothing hides.
    [Fact]
    public void A_rectangle_added_beside_a_schematic_deleted_is_still_a_change_of_its_component()
    {
        BoardData published = TopAndBack();
        BoardData draft = TopAndBack()
            .WithSchematics([new BoardSchematicEntry { SchematicName = "Back", SchematicImageFile = "back.png" }])
            .WithComponentHighlights([Rectangle("Back", "U9", "7")]);

        DraftChangeSummary summary = DraftChangeSummary.Build(Content(published), Content(draft));

        Assert.Equal(["1 mark on the schematics added"], Single(summary, "U9").Lines);
        Assert.DoesNotContain(summary.Groups, group => group.Label == "U8");

        Every_component_group_has_something_to_undo(published, draft, summary);
    }

    // ###########################################################################################
    // A file two components cite, or the board too, is the board's to say (code review, 2026-11-02): under a
    // component it had an "Undo changes to" that cannot take it back while the other still cites it.
    // ###########################################################################################
    [Fact]
    public void A_replaced_file_more_than_one_component_cites_is_said_for_the_board()
    {
        BoardData board = Board(Component("U8", "RAM"), Component("U9", "RAM"), Component("U10", "ROM"));
        board.ComponentLocalFiles.AddRange([
            new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "Datasheets/4164.pdf" },
            new ComponentLocalFileEntry { BoardLabel = "U9", Name = "Datasheet", File = "Datasheets/4164.pdf" },
            new ComponentLocalFileEntry { BoardLabel = "U10", Name = "Datasheet", File = "Datasheets/rom.pdf" }]);

        static BoardDocument.StoredFileRow File(string name, char sha) => new() { Name = name, Sha256 = new string(sha, 64), Size = 1 };

        var publishedFiles = new[] { File("Datasheets/4164.pdf", 'a'), File("Datasheets/rom.pdf", 'a') }.ToDictionary(file => file.Name, StringComparer.OrdinalIgnoreCase);
        var draftFiles = new[] { File("Datasheets/4164.pdf", 'b'), File("Datasheets/rom.pdf", 'b') }.ToDictionary(file => file.Name, StringComparer.OrdinalIgnoreCase);

        DraftChangeSummary summary = DraftChangeSummary.Build(Content(board), Content(board), publishedFiles, draftFiles);

        Assert.Equal(["File [Datasheets/4164.pdf] replaced"], Single(summary, string.Empty).Lines);
        Assert.Equal(["File [Datasheets/rom.pdf] replaced"], Single(summary, "U10").Lines);
        Assert.DoesNotContain(summary.Groups, group => group.Label is "U8" or "U9");
    }

    // ###########################################################################################
    // The board's own parts: a renamed schematic is said once (its calibration follows it, so that
    // is no change), a calibration changed is said, and credits, files and signals are named.
    // ###########################################################################################
    [Fact]
    public void The_boards_own_parts_are_said_by_name_and_a_rename_is_one_line()
    {
        BoardData published = Board(Component("U8", "CPU"));
        published.Schematics.AddRange([
            new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png" },
            new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "bottom.png" }]);

        BoardData draft = Board(Component("U8", "CPU"));
        draft.Schematics.AddRange([
            new BoardSchematicEntry { SchematicName = "Front", SchematicImageFile = "top.png" },
            new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "bottom.png", CadName = "PCB Bottom" }]);
        draft.Credits.Add(new CreditEntry { Category = "Data", NameOrHandle = "Anna" });

        var top = new KiCadCalibrationEntry { SchematicName = "Top", CadName = "PCB Top", ScaleX = 1, ScaleY = 1 };
        var bottom = new KiCadCalibrationEntry { SchematicName = "Bottom", CadName = "PCB Bottom", ScaleX = 1, ScaleY = 1 };

        DraftChangeSummary summary = DraftChangeSummary.Build(
            Content(published, top, bottom),
            Content(draft, top.WithSchematicName("Front"), new KiCadCalibrationEntry { SchematicName = "Bottom", CadName = "PCB Bottom", ScaleX = 2, ScaleY = 1 }));

        DraftChangeGroup board = Assert.Single(summary.Groups);
        Assert.False(board.IsComponent);
        Assert.Equal(DraftChangeSummary.BoardTitle, board.Title);
        Assert.Equal(
            [
                "Schematic image [Top] renamed [Front]",
                "Schematic image [Bottom]: CAD name changed",
                "Credit [Data / Anna] added",
                "KiCad calibration of [Bottom] changed",
            ],
            board.Lines);
    }

    // ###########################################################################################
    // A file replaced under its own name changes no row and is still work: said on the component
    // whose row cites it, or on the board. The board's KiCad data, which no row cites, is counted.
    // ###########################################################################################
    [Fact]
    public void Files_replaced_under_their_own_names_and_the_KiCad_data_are_said_too()
    {
        BoardData board = Board(Component("U8", "CPU"));
        board.ComponentImages.Add(Picture("U8", "Top", "Component images/u8.png"));
        board.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Manuals", Name = "Service", File = "Board local files/service.pdf" });

        static BoardDocument.StoredFileRow File(string name, char sha) => new() { Name = name, Sha256 = new string(sha, 64), Size = 1 };

        var publishedFiles = new[] { File("Component images/u8.png", 'a'), File("Board local files/service.pdf", 'b'), File("KiCad data/board.kicad_pcb", 'c') }
            .ToDictionary(file => file.Name, StringComparer.OrdinalIgnoreCase);
        var draftFiles = new[] { File("Component images/u8.png", 'd'), File("Board local files/service.pdf", 'e'), File("KiCad data/board.kicad_pcb", 'f'),
                File("KiCad data/Pages/a.kicad_sch", '1'), File("KiCad data/Pages/b.kicad_sch", '2') }
            .ToDictionary(file => file.Name, StringComparer.OrdinalIgnoreCase);

        DraftChangeSummary summary = DraftChangeSummary.Build(Content(board), Content(board), publishedFiles, draftFiles);

        Assert.Equal(["KiCad data: 2 files added, 1 replaced", "File [Board local files/service.pdf] replaced"], summary.Groups[0].Lines);
        Assert.Equal(["File [Component images/u8.png] replaced"], Single(summary, "U8").Lines);
    }

    // The board first, then the components as a person reads labels: C2 before C10.
    [Fact]
    public void The_board_comes_first_then_the_components_in_label_order()
    {
        BoardData published = Board();
        BoardData draft = Board(Component("C10"), Component("C2"), Component("U1"));
        draft.Credits.Add(new CreditEntry { Category = "Data", NameOrHandle = "Anna" });

        DraftChangeSummary summary = DraftChangeSummary.Build(Content(published), Content(draft));

        Assert.Equal(["", "C2", "C10", "U1"], summary.Groups.Select(group => group.Label));
    }

    // A board that exists only as a draft: every component is new.
    [Fact]
    public void A_board_only_drafted_has_every_component_new()
    {
        DraftChangeSummary summary = DraftChangeSummary.Build(null, Content(Board(Component("U1"), Component("U2"))));

        Assert.All(summary.Groups, group => Assert.Equal(DraftComponentChange.Added, group.Change));
        Assert.Equal(2, summary.Groups.Count);
    }

    [Fact]
    public void As_lines_each_line_names_its_group_and_a_long_list_ends_with_how_many_more()
    {
        DraftChangeSummary summary = DraftChangeSummary.Build(Content(Board()), Content(Board(Component("U1", "VIC"), Component("U2"), Component("U3"))));

        Assert.Equal(["U1 - VIC: New component", "U2: New component", "U3: New component"], summary.AsLines());
        Assert.Equal(["U1 - VIC: New component", "... and 2 more"], summary.AsLines(max: 2));
    }

    // ###########################################################################################
    // CRT 4.0's own data (2026-10-25), said under its component like everything else of it: a
    // package as a field changed, pins and replacement parts counted.
    // ###########################################################################################
    [Fact]
    public void A_components_package_pins_and_replacements_are_said_under_it()
    {
        BoardData before = Board(Component("U8", "CIA"));
        before.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS" });

        BoardData after = Board(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CIA", Category = "IC", Package = "DIP-40" });
        after.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "GND" });
        after.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U8", Pin = "2", Name = "PA0" });
        after.ComponentReplacements.Add(new ComponentReplacementEntry { BoardLabel = "U8", PartNumber = "6526A", Kind = "Drop-in" });

        DraftChangeGroup group = Single(DraftChangeSummary.Build(Content(before), Content(after)), "U8");

        Assert.Equal(
            ["Package changed", "1 pin added", "1 pin changed", "1 replacement part added"],
            group.Lines);
    }

    // ###########################################################################################
    // A list put in another order is said under the board (owner decision, 2026-10-30: order is a
    // change) - "Review changes" and the Submit dialog show it, so an order-only draft does not read
    // as empty.
    // ###########################################################################################
    // A component's own rows put in another order are said under the component too, so "Review changes"
    // gives it its "Undo changes to" (code review, 2026-11-02) - the board's line says the list changed.
    [Fact]
    public void A_components_own_rows_in_another_order_are_said_under_it()
    {
        BoardData published = Board(Component("U8", "CPU"), Component("U9", "SID"));
        published.ComponentLinks.AddRange(
        [
            new ComponentLinkEntry { BoardLabel = "U8", Name = "First", Url = "https://example.org/1" },
            new ComponentLinkEntry { BoardLabel = "U8", Name = "Second", Url = "https://example.org/2" },
        ]);

        BoardData draft = Board(Component("U8", "CPU"), Component("U9", "SID"));
        draft.ComponentLinks.AddRange([published.ComponentLinks[1], published.ComponentLinks[0]]);

        DraftChangeSummary summary = DraftChangeSummary.Build(Content(published), Content(draft));

        DraftChangeGroup group = Single(summary, "U8");
        Assert.True(group.IsComponent);
        Assert.Equal(["Order of the component links changed"], group.Lines);
        Assert.Equal(["Order of the component links changed"], Single(summary, string.Empty).Lines);
        Assert.DoesNotContain(summary.Groups, each => each.Label == "U9");
    }

    [Fact]
    public void A_list_in_another_order_is_said_under_the_board()
    {
        BoardData published = Board(Component("U8", "CPU"));
        published.BoardLocalFiles.AddRange(
        [
            new BoardLocalFileEntry { Category = "Resources", Name = "Manual", File = "Board local files/manual.pdf" },
            new BoardLocalFileEntry { Category = "Troubleshooting", Name = "Dead test", File = "Board local files/dead.pdf" },
        ]);

        BoardData draft = Board(Component("U8", "CPU"));
        draft.BoardLocalFiles.AddRange([published.BoardLocalFiles[1], published.BoardLocalFiles[0]]);

        DraftChangeSummary summary = DraftChangeSummary.Build(Content(published), Content(draft));

        Assert.Equal(["Order of the files changed"], Single(summary, string.Empty).Lines);
        Assert.Equal(["The board: Order of the files changed"], summary.AsLines());
    }
}
