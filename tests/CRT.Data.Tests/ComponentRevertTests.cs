using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentRevert - "Undo changes to IC1": one component back to its published values (CRT 4.0
// step 1, 2026-10-25). Every row of it, in every section, as the published board has it; the rest
// of the draft untouched; the rows back where they were.
// ###########################################################################################
public sealed class ComponentRevertTests
{
    private static ComponentEntry Component(string label, string description = "") =>
        new() { BoardLabel = label, Category = "IC", Description = description };

    private static ComponentImageEntry Picture(string label, string name) =>
        new() { BoardLabel = label, Name = name, File = $"Component images/{label}-{name}.png" };

    private static ComponentHighlightEntry Mark(string label, string x = "1") =>
        new() { SchematicName = "Top", BoardLabel = label, X = x, Y = "2", Width = "3", Height = "4" };

    private static BoardData Board(params ComponentEntry[] components) => new()
    {
        HardwareName = "Commodore 64",
        BoardName = "250407",
        Components = [.. components],
    };

    private static string[] Labels(BoardData board) => board.Components.Select(component => component.BoardLabel).ToArray();

    [Fact]
    public void A_changed_component_goes_back_to_its_published_rows_and_nothing_else_moves()
    {
        BoardData published = Board(Component("U1", "VIC"), Component("U8", "CPU"), Component("U9", "SID"));
        published.ComponentImages.Add(Picture("U8", "Top"));
        published.ComponentHighlights.Add(Mark("U8"));

        BoardData draft = Board(Component("U1", "VIC-II"), Component("U8", "6510"), Component("U9", "SID"));
        draft.ComponentImages.AddRange([Picture("U8", "Bottom"), Picture("U1", "Top")]);
        draft.ComponentHighlights.Add(Mark("U8", x: "50"));
        draft.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U8", Name = "Datasheet", Url = "https://example.com" });

        Assert.True(ComponentRevert.HasChanges(draft, published, "U8"));

        BoardData reverted = ComponentRevert.ToPublished(draft, published, " u8 ");

        Assert.Equal(["U1", "U8", "U9"], Labels(reverted));
        Assert.Equal("CPU", reverted.Components[1].Description);
        Assert.Equal("VIC-II", reverted.Components[0].Description);
        Assert.Equal(["U8-Top", "U1-Top"], reverted.ComponentImages.Select(image => $"{image.BoardLabel}-{image.Name}"));
        Assert.Equal("1", Assert.Single(reverted.ComponentHighlights).X);
        Assert.Empty(reverted.ComponentLinks);

        Assert.False(ComponentRevert.HasChanges(reverted, published, "U8"));
        Assert.True(ComponentRevert.HasChanges(reverted, published, "U1"));
    }

    // A component the published board does not have goes altogether, with everything of it.
    [Fact]
    public void A_new_component_goes_altogether()
    {
        BoardData published = Board(Component("U8", "CPU"));
        BoardData draft = Board(Component("U8", "CPU"), Component("U99", "Mine"));
        draft.ComponentImages.Add(Picture("U99", "Top"));
        draft.ComponentHighlights.Add(Mark("U99"));

        BoardData reverted = ComponentRevert.ToPublished(draft, published, "U99");

        Assert.Equal(["U8"], Labels(reverted));
        Assert.Empty(reverted.ComponentImages);
        Assert.Empty(reverted.ComponentHighlights);
    }

    // A component the draft removed comes back where it was - after the row before it.
    [Fact]
    public void A_removed_component_comes_back_where_it_was()
    {
        BoardData published = Board(Component("U1"), Component("U8", "CPU"), Component("U9"));
        BoardData draft = Board(Component("U1"), Component("U9"), Component("U10"));

        BoardData reverted = ComponentRevert.ToPublished(draft, published, "U8");

        Assert.Equal(["U1", "U8", "U9", "U10"], Labels(reverted));
        Assert.Equal("CPU", reverted.Components[1].Description);
    }

    // ###########################################################################################
    // *** UNDOING A COMPONENT LEAVES THE IMAGES IN THEIR ORDER (code review, 2026-11-01). *** A board
    // file holds the pictures first and the scope baselines after them, so a component's two kinds of
    // rows do not stand together. Put back together, its baseline went before the baselines ahead of it,
    // and "Review changes" said "Order of the component pictures changed" under "The board", with
    // nothing there to take it back.
    // ###########################################################################################
    [Fact]
    public void Undoing_a_component_puts_its_pictures_and_baselines_back_among_their_own_kind()
    {
        static ComponentImageEntry Baseline(string label, string note = "") =>
            new() { BoardLabel = label, Pin = "1", Name = "Baseline", TimeDiv = "1mS", File = $"Scope baseline/{label}.png", Note = note };

        BoardData published = Board(Component("U7"), Component("U8"), Component("U9"));
        published.ComponentImages.AddRange([Picture("U7", "Top"), Picture("U8", "Top"), Picture("U9", "Top"), Baseline("U7"), Baseline("U8"), Baseline("U9")]);

        BoardData draft = Board(Component("U7"), Component("U8"), Component("U9"));
        draft.ComponentImages.AddRange([Picture("U7", "Top"), Picture("U8", "Top"), Picture("U9", "Top"), Baseline("U7"), Baseline("U8", note: "Mine"), Baseline("U9")]);

        BoardData reverted = ComponentRevert.ToPublished(draft, published, "U8");

        Assert.Equal(
            published.ComponentImages.Select(image => image.File),
            reverted.ComponentImages.Select(image => image.File));
        Assert.Empty(BoardRowOrder.ReorderedSheets(published, reverted));

        // With its picture removed in the draft, the picture comes back among the pictures all the same.
        BoardData pictureRemoved = Board(Component("U7"), Component("U8"), Component("U9"));
        pictureRemoved.ComponentImages.AddRange([Picture("U7", "Top"), Picture("U9", "Top"), Baseline("U7"), Baseline("U8", note: "Mine"), Baseline("U9")]);

        Assert.Equal(
            published.ComponentImages.Select(image => image.File),
            ComponentRevert.ToPublished(pictureRemoved, published, "U8").ComponentImages.Select(image => image.File));
    }

    // The same for any list: a component's rows that do not stand together go back where each was.
    [Fact]
    public void A_components_rows_that_do_not_stand_together_go_back_where_each_was()
    {
        static ComponentEntry InRegion(string label, string region, string description = "") =>
            new() { BoardLabel = label, Region = region, Category = "IC", Description = description };

        BoardData published = Board(InRegion("U1", "PAL"), InRegion("U2", "PAL"), InRegion("U1", "NTSC"));
        BoardData draft = Board(InRegion("U1", "PAL", "Mine"), InRegion("U2", "PAL"), InRegion("U1", "NTSC"));

        BoardData reverted = ComponentRevert.ToPublished(draft, published, "U1");

        Assert.Equal(["U1 PAL", "U2 PAL", "U1 NTSC"], reverted.Components.Select(row => $"{row.BoardLabel} {row.Region}"));
        Assert.Empty(BoardRowOrder.ReorderedSheets(published, reverted));
    }

    // ###########################################################################################
    // A component dragged elsewhere AND edited: its values and its place go back (code review, 2026-11-01).
    // The first row went where the draft had moved it, so "Order of the components changed" stayed with
    // nothing to take it back, and the draft still counted as holding work.
    // ###########################################################################################
    [Fact]
    public void A_component_dragged_elsewhere_goes_back_where_it_was()
    {
        BoardData published = Board(Component("U1"), Component("U8", "CPU"), Component("U9"));
        BoardData draft = Board(Component("U8", "6510"), Component("U1"), Component("U9"));

        BoardData reverted = ComponentRevert.ToPublished(draft, published, "U8");

        Assert.Equal(["U1", "U8", "U9"], Labels(reverted));
        Assert.Empty(BoardRowOrder.ReorderedSheets(published, reverted));

        // Published first, dragged down: first again.
        BoardData firstAgain = ComponentRevert.ToPublished(Board(Component("U8"), Component("U1", "Mine"), Component("U9")), published, "U1");
        Assert.Equal(["U1", "U8", "U9"], Labels(firstAgain));
    }

    // What the draft added or relabelled in front of it stays in front of it - nothing else moves.
    [Fact]
    public void A_row_the_published_board_lacks_stays_where_it_stood_beside_the_component()
    {
        BoardData published = Board(Component("U1"), Component("U8", "CPU"), Component("U9"));

        BoardData added = Board(Component("U1"), Component("U2"), Component("U8", "6510"), Component("U9"));
        Assert.Equal(["U1", "U2", "U8", "U9"], Labels(ComponentRevert.ToPublished(added, published, "U8")));

        BoardData relabelled = Board(Component("U1x"), Component("U8", "6510"), Component("U9"));
        Assert.Equal(["U1x", "U8", "U9"], Labels(ComponentRevert.ToPublished(relabelled, published, "U8")));
    }

    // Undoing a relabelled component brings its old label back, with its marks.
    [Fact]
    public void A_relabelled_component_goes_back_under_its_old_label()
    {
        BoardData published = Board(Component("U1"), Component("U8", "CPU"));
        published.ComponentHighlights.Add(Mark("U8"));

        BoardData draft = Board(Component("U1"), Component("U08", "CPU"));
        draft.ComponentHighlights.Add(Mark("U08"));

        BoardData reverted = ComponentRevert.ToPublished(draft, published, "U08");

        Assert.Equal(["U1", "U8"], Labels(reverted));
        Assert.Equal("U8", Assert.Single(reverted.ComponentHighlights).BoardLabel);
    }

    [Fact]
    public void A_blank_label_changes_nothing_and_a_component_like_the_published_one_has_nothing_to_undo()
    {
        BoardData board = Board(Component("U8", "CPU"));

        Assert.Same(board, ComponentRevert.ToPublished(board, board, "  "));
        Assert.False(ComponentRevert.HasChanges(board, board, "U8"));

        // A draft-only board: every component is new, so undoing one removes it.
        Assert.Empty(ComponentRevert.ToPublished(board, null, "U8").Components);
    }

    // The caption and the board's own sections are the draft's - only the component moves.
    [Fact]
    public void The_boards_own_sections_and_caption_stay_as_the_draft_has_them()
    {
        BoardData published = Board(Component("U8", "CPU"));
        BoardData draft = Board(Component("U8", "6510"));
        draft.Credits.Add(new CreditEntry { Category = "Data", NameOrHandle = "Anna" });
        draft.Schematics.Add(new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png" });

        BoardData reverted = ComponentRevert.ToPublished(draft, published, "U8");

        Assert.Same(draft.Credits, reverted.Credits);
        Assert.Same(draft.Schematics, reverted.Schematics);
        Assert.Equal(("Commodore 64", "250407"), (reverted.HardwareName, reverted.BoardName));
    }

    // CRT 4.0's pins and replacement parts (2026-10-25) are the component's too - back as published,
    // and another component's left alone.
    [Fact]
    public void The_components_pins_and_replacements_go_back_too()
    {
        BoardData published = Board(Component("U8"), Component("U9"));
        published.ComponentPins.Add(new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "VSS" });

        BoardData draft = Board(Component("U8"), Component("U9"));
        draft.ComponentPins.AddRange([
            new ComponentPinEntry { BoardLabel = "U8", Pin = "1", Name = "GND" },
            new ComponentPinEntry { BoardLabel = "U9", Pin = "1", Name = "CAP1A" }]);
        draft.ComponentReplacements.Add(new ComponentReplacementEntry { BoardLabel = "U8", PartNumber = "6526A" });

        Assert.True(ComponentRevert.HasChanges(draft, published, "U8"));

        BoardData reverted = ComponentRevert.ToPublished(draft, published, "U8");

        Assert.Equal(["U8 VSS", "U9 CAP1A"], reverted.ComponentPins.Select(pin => $"{pin.BoardLabel} {pin.Name}"));
        Assert.Empty(reverted.ComponentReplacements);
        Assert.False(ComponentRevert.HasChanges(reverted, published, "U8"));
    }

    // ###########################################################################################
    // *** A SCHEMATIC THE DRAFT RENAMED KEEPS ITS NEW NAME (code review, 2026-10-26). *** The
    // component's highlights went back under the published schematic name, which the draft no longer
    // has - an error at Submit that nothing in the table can fix. They go back under the draft's name
    // for that image (same image file - SubmissionRowsBoard.FollowSchematics, every draft edit's
    // rule), one on an image the draft removed goes, and the rename alone is no change of U8.
    // ###########################################################################################
    [Fact]
    public void Highlights_go_back_under_the_name_the_draft_gave_their_schematic()
    {
        BoardData published = Board(Component("U8"));
        published.Schematics.AddRange([
            new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "Images/sheet1.png" },
            new BoardSchematicEntry { SchematicName = "Back", SchematicImageFile = "Images/back.png" }]);
        published.ComponentHighlights.AddRange([
            new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" },
            new ComponentHighlightEntry { SchematicName = "Back", BoardLabel = "U8", X = "5", Y = "6", Width = "7", Height = "8" }]);

        BoardData draft = Board(Component("U8"));
        draft.Schematics.Add(new BoardSchematicEntry { SchematicName = "Main sheet", SchematicImageFile = "Images/sheet1.png" });
        draft.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Main sheet", BoardLabel = "U8", X = "50", Y = "2", Width = "3", Height = "4" });

        BoardData reverted = ComponentRevert.ToPublished(draft, published, "U8");

        ComponentHighlightEntry back = Assert.Single(reverted.ComponentHighlights);
        Assert.Equal(("Main sheet", "1"), (back.SchematicName, back.X));

        // The rename and the removal are the board's changes, not U8's.
        Assert.False(ComponentRevert.HasChanges(reverted, published, "U8"));
    }

    // ------------------------------------------------------------------ Files replaced under their own names

    private static Dictionary<string, BoardDocument.StoredFileRow> Files(params (string Name, char Sha)[] files) =>
        files.ToDictionary(
            file => file.Name,
            file => new BoardDocument.StoredFileRow { Name = file.Name, Sha256 = new string(file.Sha, 64), Size = 1 },
            StringComparer.OrdinalIgnoreCase);

    // ###########################################################################################
    // *** A PICTURE REPLACED UNDER ITS OWN NAME IS A CHANGE OF THE COMPONENT (code review,
    // 2026-10-26). *** Its rows are as published, so "Undo changes to [U8]" was not offered - and when
    // it was, it put the rows back and left the replaced bytes in the draft's list, to go out with the
    // next submission. The files the component cites with other bytes than the published board's are
    // named, for the draft's list to put back; another component's are not this one's.
    // ###########################################################################################
    [Fact]
    public void A_file_replaced_under_its_own_name_is_a_change_of_the_component_that_cites_it()
    {
        BoardData published = Board(Component("U8"), Component("U9"));
        published.ComponentImages.AddRange([Picture("U8", "Top"), Picture("U9", "Top")]);
        published.ComponentLocalFiles.Add(new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "Datasheets/6510.pdf" });
        BoardData draft = Board(Component("U8"), Component("U9"));
        draft.ComponentImages.AddRange([Picture("U8", "Top"), Picture("U9", "Top")]);
        draft.ComponentLocalFiles.Add(new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "Datasheets/6510.pdf" });

        var publishedFiles = Files(("Component images/U8-Top.png", 'a'), ("Component images/U9-Top.png", 'a'), ("Datasheets/6510.pdf", 'a'));
        var draftFiles = Files(("Component images/U8-Top.png", 'b'), ("Component images/U9-Top.png", 'a'), ("Datasheets/6510.pdf", 'a'));

        Assert.Equal(["Component images/U8-Top.png"], ComponentRevert.ReplacedFiles(draft, published, "U8", publishedFiles, draftFiles));
        Assert.Empty(ComponentRevert.ReplacedFiles(draft, published, "U9", publishedFiles, draftFiles));

        Assert.True(ComponentRevert.HasChanges(draft, published, "U8", publishedFiles, draftFiles));
        Assert.False(ComponentRevert.HasChanges(draft, published, "U9", publishedFiles, draftFiles));

        // Without the lists only the rows are compared - as they are the same, nothing to undo in them.
        Assert.False(ComponentRevert.HasChanges(draft, published, "U8"));
        Assert.Null(ComponentRevert.RowsToPublished(draft, published, "U8"));
    }

    // A relabelled component's files are the ones its OLD label's published rows cite.
    [Fact]
    public void A_relabelled_components_replaced_file_is_found_under_its_old_label()
    {
        BoardData published = Board(Component("U8", "CPU"));
        published.ComponentImages.Add(Picture("U8", "Top"));
        BoardData draft = Board(Component("U08", "CPU"));
        draft.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U08", Name = "Top", File = "Component images/U8-Top.png" });

        Assert.Equal(
            ["Component images/U8-Top.png"],
            ComponentRevert.ReplacedFiles(draft, published, "U08", Files(("Component images/U8-Top.png", 'a')), Files(("Component images/U8-Top.png", 'b'))));
    }

    // Rows to put back: the reverted board, or null when the rows already read as published.
    [Fact]
    public void Rows_to_put_back_are_none_when_the_rows_are_as_published()
    {
        BoardData published = Board(Component("U8", "CPU"));
        BoardData changed = Board(Component("U8", "6510"));

        Assert.Equal("CPU", ComponentRevert.RowsToPublished(changed, published, "U8")!.Components.Single().Description);
        Assert.Null(ComponentRevert.RowsToPublished(Board(Component("U8", "CPU")), published, "U8"));
        Assert.Null(ComponentRevert.RowsToPublished(changed, published, " "));
    }
}
