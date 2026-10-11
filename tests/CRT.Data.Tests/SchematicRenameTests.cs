using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// SchematicRename - renaming a schematic image of a draft (owner request, 2026-10-10), and carrying
// along what the board's sidecar files under its name: the component highlights drawn on it and its
// KiCad calibration. DraftWorkbookStoreTests runs the same through a real draft on disk.
public sealed class SchematicRenameTests
{
    private static BoardData Board(params string[] schematicNames)
    {
        var board = new BoardData();
        board.Schematics.AddRange(schematicNames.Select(name => new BoardSchematicEntry
        {
            SchematicName = name,
            SchematicImageFile = $"Commodore/C64/250407/{name}.png",
            CadName = "Open128 - PCB Top",
        }));

        return board;
    }

    private static ComponentHighlightEntry Highlight(string schematicName, string label) =>
        new() { SchematicName = schematicName, BoardLabel = label, X = "1", Y = "2", Width = "3", Height = "4" };

    // ------------------------------------------------------------------ which names are allowed

    [Fact]
    public void A_blank_name_is_refused()
    {
        Assert.Equal(SchematicRename.NeedsAName, SchematicRename.Problem("   ", ["Top"]));
        Assert.Equal(SchematicRename.NeedsAName, SchematicRename.Problem(null, ["Top"]));
    }

    // The name is the sheet's key - two images of one name is an error the checks report.
    [Fact]
    public void A_name_another_image_has_is_refused_whatever_its_capitals()
    {
        Assert.Equal(SchematicRename.NameTaken("top"), SchematicRename.Problem(" top ", ["Top", "Bottom"]));
        Assert.Contains("[top]", SchematicRename.NameTaken("top"), StringComparison.Ordinal);
    }

    [Fact]
    public void A_name_no_other_image_has_is_allowed()
    {
        Assert.Null(SchematicRename.Problem("PCB; Top", ["PCB; Bottom", "Clock"]));
        Assert.Null(SchematicRename.Problem("Top", []));
    }

    // ------------------------------------------------------------------ what the rename carries

    [Fact]
    public void The_row_is_renamed_and_keeps_every_other_column()
    {
        (BoardData renamed, bool moved) = SchematicRename.Apply(Board("Top", "Bottom"), 0, " PCB; Top ");

        Assert.True(moved);
        Assert.Equal(["PCB; Top", "Bottom"], renamed.Schematics.Select(entry => entry.SchematicName));
        Assert.Equal("Commodore/C64/250407/Top.png", renamed.Schematics[0].SchematicImageFile);
        Assert.Equal("Open128 - PCB Top", renamed.Schematics[0].CadName);
    }

    // Left behind, they would point at a schematic that is gone - an error the server refuses.
    [Fact]
    public void The_highlights_drawn_on_the_image_move_with_it_and_no_others()
    {
        BoardData board = Board("Top", "Bottom");
        board.ComponentHighlights.AddRange([Highlight("Top", "U1"), Highlight("TOP", "U2"), Highlight("Bottom", "U1")]);

        (BoardData renamed, _) = SchematicRename.Apply(board, 0, "PCB; Top");

        Assert.Equal(
            ["PCB; Top/U1", "PCB; Top/U2", "Bottom/U1"],
            renamed.ComponentHighlights.Select(highlight => $"{highlight.SchematicName}/{highlight.BoardLabel}"));
        Assert.Equal("3", renamed.ComponentHighlights[0].Width);
    }

    // ------------------------------------------------------------------ following a whole edit
    //
    // Every draft write follows what hangs on a renamed or deleted schematic, with the server's
    // rule for a table edit (SubmissionRowsBoard.FollowSchematics): old name -> new name, or ->
    // null for a schematic deleted.

    private static readonly Dictionary<string, string?> TopRenamedBottomDeleted =
        new(StringComparer.OrdinalIgnoreCase) { ["Top"] = "PCB; Top", ["Bottom"] = null };

    [Fact]
    public void A_highlight_follows_its_schematic_renamed_goes_with_it_deleted_and_stays_otherwise()
    {
        List<ComponentHighlightEntry> followed = SchematicRename.FollowHighlights(
            [Highlight("top", "U1"), Highlight("Bottom", "U2"), Highlight("Clock", "U3")],
            TopRenamedBottomDeleted);

        Assert.Equal(["PCB; Top/U1", "Clock/U3"], followed.Select(highlight => $"{highlight.SchematicName}/{highlight.BoardLabel}"));
        Assert.Equal("3", followed[0].Width);
    }

    [Fact]
    public void A_calibration_follows_its_schematic_the_same_way()
    {
        List<KiCadCalibrationEntry> followed = SchematicRename.FollowCalibrations(
            [
                new() { SchematicName = "Top", CadName = "B.Cu", ScaleX = 3.5, MirrorY = true },
                new() { SchematicName = "Bottom", CadName = "F.Cu" },
                new() { SchematicName = "Clock" },
            ],
            TopRenamedBottomDeleted);

        Assert.Equal(["PCB; Top", "Clock"], followed.Select(calibration => calibration.SchematicName));
        Assert.Equal(3.5, followed[0].ScaleX);
        Assert.True(followed[0].MirrorY);
    }

    [Fact]
    public void The_renames_of_an_edit_leave_out_the_deletions()
    {
        Assert.Equal([("Top", "PCB; Top")], SchematicRename.RenamesIn(TopRenamedBottomDeleted));
        Assert.Empty(SchematicRename.RenamesIn(new Dictionary<string, string?>()));
    }

    // "top" to "Top" is a rename like any other: the sidecar files by the exact name, so a
    // calibration left under "top" would not be found for the row "Top".
    [Fact]
    public void A_change_of_capitals_alone_moves_what_hangs_on_the_name()
    {
        BoardData board = Board("top");
        board.ComponentHighlights.Add(Highlight("top", "U1"));

        (BoardData renamed, bool moved) = SchematicRename.Apply(board, 0, "Top");

        Assert.True(moved);
        Assert.Equal("Top", Assert.Single(renamed.Schematics).SchematicName);
        Assert.Equal("Top", Assert.Single(renamed.ComponentHighlights).SchematicName);
    }

    // ###########################################################################################
    // Two images of one name (an error), one of them renamed to fix it: nothing says which image
    // the highlights were drawn for, so they stay on the name that is still there.
    // ###########################################################################################
    [Fact]
    public void Renaming_one_of_two_images_sharing_a_name_leaves_the_highlights_on_the_other()
    {
        BoardData board = Board("Top", "Top");
        board.ComponentHighlights.Add(Highlight("Top", "U1"));

        (BoardData renamed, bool moved) = SchematicRename.Apply(board, 0, "Top (scan)");

        Assert.False(moved);
        Assert.Equal(["Top (scan)", "Top"], renamed.Schematics.Select(entry => entry.SchematicName));
        Assert.Equal("Top", Assert.Single(renamed.ComponentHighlights).SchematicName);
    }

    // ###########################################################################################
    // *** THE ROW RENAMED IS THE ONE ASKED FOR, NOT THE FIRST OF ITS NAME (code review,
    // 2026-10-10). *** Renaming the SECOND of two images named "Top" renamed the first, and the
    // window then showed each row under the other's name.
    // ###########################################################################################
    [Fact]
    public void Renaming_the_second_of_two_images_sharing_a_name_renames_that_one()
    {
        BoardData board = Board("Top", "Top");
        board.Schematics[1] = board.Schematics[1].WithCadName("Open128 - PCB Bottom");

        (BoardData renamed, _) = SchematicRename.Apply(board, 1, "Bottom");

        Assert.Equal(["Top", "Bottom"], renamed.Schematics.Select(entry => entry.SchematicName));
        Assert.Equal("Open128 - PCB Bottom", renamed.Schematics[1].CadName);
    }

    [Fact]
    public void Renaming_an_image_the_board_does_not_have_changes_nothing()
    {
        BoardData board = Board("Top");

        (BoardData renamed, bool moved) = SchematicRename.Apply(board, 1, "New");

        Assert.Same(board, renamed);
        Assert.False(moved);
        Assert.Same(board, SchematicRename.Apply(board, -1, "New").Board);
    }
}
