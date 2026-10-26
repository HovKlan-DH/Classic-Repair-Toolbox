using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftEditRules - what follows an edit of a draft's board: highlights and calibrations with a
// renamed or deleted schematic, highlights with a relabelled component. Moved out of
// DraftWorkbookStore.EditCore (2026-10-25) so the draft held in memory (DraftSession) applies the very
// same rules; DraftWorkbookStoreTests keep holding the file writes to them.
// ###########################################################################################
public sealed class DraftEditRulesTests
{
    private static BoardSchematicEntry Schematic(string name, string file) => new() { SchematicName = name, SchematicImageFile = file };

    private static ComponentHighlightEntry Highlight(string schematic, string label) =>
        new() { SchematicName = schematic, BoardLabel = label, X = "1", Y = "2", Width = "3", Height = "4" };

    private static KiCadCalibrationEntry Calibration(string schematic) => new() { SchematicName = schematic, CadName = "PCB", ScaleX = 2 };

    private static BoardData Board(List<BoardSchematicEntry> schematics, params ComponentHighlightEntry[] highlights) => new()
    {
        Components = { new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", TechnicalNameOrValue = "6510" } },
        Schematics = schematics,
        ComponentHighlights = [.. highlights],
    };

    [Fact]
    public void A_schematic_renamed_on_the_same_image_takes_its_highlights_and_calibration_along()
    {
        BoardData before = Board([Schematic("Top", "top.png")], Highlight("Top", "U1"));
        BoardData after = before.WithSchematics([Schematic("Front", "top.png")]);

        (BoardData board, List<KiCadCalibrationEntry> calibrations) = DraftEditRules.Follow(before, after, [Calibration("Top")]);

        Assert.Equal("Front", Assert.Single(board.ComponentHighlights).SchematicName);
        Assert.Equal("Front", Assert.Single(calibrations).SchematicName);
        Assert.Equal(2, calibrations[0].ScaleX);
    }

    [Fact]
    public void A_deleted_schematic_drops_its_highlights_and_calibration()
    {
        BoardData before = Board([Schematic("Top", "top.png"), Schematic("Bottom", "bottom.png")], Highlight("Top", "U1"), Highlight("Bottom", "U1"));
        BoardData after = before.WithSchematics([Schematic("Bottom", "bottom.png")]);

        (BoardData board, List<KiCadCalibrationEntry> calibrations) = DraftEditRules.Follow(before, after, [Calibration("Top"), Calibration("Bottom")]);

        Assert.Equal("Bottom", Assert.Single(board.ComponentHighlights).SchematicName);
        Assert.Equal("Bottom", Assert.Single(calibrations).SchematicName);
    }

    // A label fixed alone ("U1" -> "U01") is the same component: its rectangles follow it.
    [Fact]
    public void A_relabelled_component_takes_its_highlights_along()
    {
        BoardData before = Board([Schematic("Top", "top.png")], Highlight("Top", "U1"));
        var after = new BoardData
        {
            Components = { new ComponentEntry { BoardLabel = "U01", FriendlyName = "CPU", TechnicalNameOrValue = "6510" } },
            Schematics = before.Schematics,
            ComponentHighlights = before.ComponentHighlights,
        };

        (BoardData board, _) = DraftEditRules.Follow(before, after, []);

        Assert.Equal("U01", Assert.Single(board.ComponentHighlights).BoardLabel);
    }

    // A calibration filed under a name the edit brings in is kept; one under neither name goes.
    [Fact]
    public void Calibrations_under_a_name_the_board_has_now_are_kept_and_strays_dropped()
    {
        BoardData before = Board([Schematic("Top", "top.png")]);
        BoardData after = before.WithSchematics([Schematic("Top", "top.png"), Schematic("Bottom", "bottom.png")]);

        (_, List<KiCadCalibrationEntry> calibrations) = DraftEditRules.Follow(
            before,
            after,
            [Calibration("Top"), Calibration("Bottom"), Calibration("Gone")]);

        Assert.Equal(["Top", "Bottom"], calibrations.Select(entry => entry.SchematicName));
        Assert.Equal(["Top", "Bottom"], DraftEditRules.NamesBeforeAndAfter(before, after));
    }

    // ###########################################################################################
    // What an edit took off the board with the schematic images it removed - said to the
    // contributor afterwards, with Ctrl+Z to bring it back (2026-10-25). A rename is no removal.
    // ###########################################################################################
    [Fact]
    public void What_a_removed_schematic_image_took_with_it_is_counted_and_a_rename_takes_nothing()
    {
        BoardData board = Board(
            [Schematic("Top", "top.png"), Schematic("Bottom", "bottom.png")],
            Highlight("Top", "U1"), Highlight("top", "U2"), Highlight("Bottom", "U1"));
        var before = new BoardContent(board, [Calibration("Top"), Calibration("Bottom")], []);

        DraftEditLoss removed = DraftEditRules.LostWith(before, before with { Board = board.WithSchematics([Schematic("Bottom", "bottom.png")]) });

        Assert.Equal(["Top"], removed.Schematics);
        Assert.Equal(2, removed.Highlights);
        Assert.Equal(1, removed.Calibrations);

        DraftEditLoss renamed = DraftEditRules.LostWith(before, before with { Board = board.WithSchematics([Schematic("Front", "top.png"), Schematic("Bottom", "bottom.png")]) });

        Assert.Same(DraftEditLoss.None, renamed);
    }
}
