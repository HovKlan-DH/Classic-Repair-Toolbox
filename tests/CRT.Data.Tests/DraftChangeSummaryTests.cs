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
        Assert.Equal(["New component", "2 pictures added", "1 scope baseline added"], group.Lines);
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
                "Picture [Top] renamed [Front]",
                "Picture [Bottom]: CAD name changed",
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
    // package or repair notes as fields changed, pins and replacement parts counted.
    // ###########################################################################################
    [Fact]
    public void A_components_package_notes_pins_and_replacements_are_said_under_it()
    {
        BoardData before = Board(Component("U8", "CIA"));
        before.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS" });

        BoardData after = Board(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CIA", Category = "IC", Package = "DIP-40", RepairNotes = "Socket it." });
        after.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "GND" });
        after.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U8", Pin = "2", Name = "PA0" });
        after.ComponentReplacements.Add(new ComponentReplacementEntry { BoardLabel = "U8", PartNumber = "6526A", Kind = "Drop-in" });

        DraftChangeGroup group = Single(DraftChangeSummary.Build(Content(before), Content(after)), "U8");

        Assert.Equal(
            ["Package and repair notes changed", "1 pin added", "1 pin changed", "1 replacement part added"],
            group.Lines);
    }
}
