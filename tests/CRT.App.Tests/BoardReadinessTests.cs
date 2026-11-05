using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardReadiness - what a board still lacks ("Board status", BoardStatusPage), by what the board holds.
// ###########################################################################################
public sealed class BoardReadinessTests
{
    // ###########################################################################################
    // What an existing board lacks, each by its label once: no short description in any region, marked
    // on no picture, pins listed but no scope baseline.
    // ###########################################################################################
    [Fact]
    public void An_existing_board_says_what_it_lacks()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Description = "CPU" },
                new ComponentEntry { BoardLabel = "U2", Region = "PAL" },
                new ComponentEntry { BoardLabel = "U2", Region = "NTSC", Description = "VIC" },
                new ComponentEntry { BoardLabel = "U3" },
            ],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "u3", X = "0", Y = "0", Width = "1", Height = "1" }],
            ComponentPins =
            [
                new ComponentPinEntry { BoardLabel = "U1", Pin = "1" },
                new ComponentPinEntry { BoardLabel = "U3", Pin = "1" },
            ],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U1", Pin = "1", File = "Scope baseline/u1.png", TimeDiv = "1us" },
                new ComponentImageEntry { BoardLabel = "U3", Name = "Top", File = "Component images/u3.png" },
            ],
        };

        BoardMissing missing = BoardReadiness.Missing(board);

        Assert.Equal(["U3"], missing.WithoutDescription);
        Assert.Equal(["U1", "U2"], missing.Unmarked);
        Assert.Equal(["U3"], missing.PinsWithoutBaseline);

        // A board with nothing missing - its pictures and its credits there too (2026-10-30: the board's
        // own gaps count).
        var whole = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }],
            Credits = [new CreditEntry { Category = "Schematic image", NameOrHandle = "Anna" }],
        };

        BoardMissing none = BoardReadiness.Missing(whole);

        Assert.Empty(none.BoardGaps);
        Assert.Empty(none.WithoutDescription);
        Assert.Empty(none.Unmarked);
        Assert.Empty(none.PinsWithoutBaseline);
    }

    // ###########################################################################################
    // What a board lacks besides its components' own gaps: the lists with a row missing a field it needs
    // (BoardRowRequired, said by their panels), and with KiCad data read, the components it does not
    // have - their traces cannot light up. KiCad files CRT cannot read, or none, say nothing of that.
    // ###########################################################################################
    [Fact]
    public void Lists_with_rows_missing_a_field_and_components_KiCad_lacks_are_said()
    {
        var board = new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U1" }, new ComponentEntry { BoardLabel = "PLA/U17" }],
            Credits = [new CreditEntry { Category = "Schematic image" }],
            BoardLinks = [new BoardLinkEntry { Category = "Forums" }, new BoardLinkEntry { Category = "Forums" }],
        };

        var facts = new BoardKiCadFacts(["U1", "U17", "C5"], ["PCB top", "PCB bottom"]);
        BoardMissing missing = BoardReadiness.Missing(board, true, facts);

        Assert.Equal([BoardGapKind.IncompleteLinks, BoardGapKind.IncompleteCredits], missing.BoardGaps.Select(gap => gap.Kind));
        Assert.Equal(BoardReadiness.IncompleteGap(2, "link", "links"), missing.BoardGaps[0].Text);
        Assert.Equal(["PLA/U17"], missing.NotInKiCadLabels);

        Assert.Empty(BoardReadiness.Missing(board, true, new BoardKiCadFacts([], [], Read: false)).NotInKiCadLabels);
        Assert.Empty(BoardReadiness.Missing(board).NotInKiCadLabels);
    }

    // A group naming its first eight says how many more there are - nothing for eight or fewer.
    [Fact]
    public void A_long_list_says_how_many_more_there_are()
    {
        Assert.Equal(string.Empty, BoardReadiness.MoreThanListed(BoardReadiness.ListedAtMost));
        Assert.Equal("and 2 more", BoardReadiness.MoreThanListed(10));
    }

    // ###########################################################################################
    // The KiCad facts count the very components "Add N from KiCad" adds (KiCadComponentTable): KiCad's
    // placeholders - a logo's "REF**", a graphic's "G***" - and a footprint with no pad (a logo since
    // slice 3b) are no components; only the first PCB file counts, as only it lights up copper.
    // ###########################################################################################
    [Fact]
    public void The_KiCad_facts_leave_out_placeholders_and_other_PCB_files()
    {
        static KiCadPcbFootprint WithPad(string reference) => new() { Reference = reference, Pads = [new KiCadPcbPad { Number = "1" }] };

        var root = new KiCadProjectRoot
        {
            Project = new KiCadProjectInfo { Views = [new KiCadProjectView { DisplayName = "PCB top" }, new KiCadProjectView { DisplayName = " " }] },
            Pcb =
            [
                new KiCadPcb { Footprints = [WithPad(" U1 "), WithPad("REF**"), WithPad("G***"), WithPad(""), new KiCadPcbFootprint { Reference = "LOGO1" }] },
                new KiCadPcb { Footprints = [WithPad("U9")] },
            ],
        };

        BoardKiCadFacts facts = BoardKiCadFacts.Of(root);

        Assert.Equal(["U1"], facts.References);
        Assert.Equal(["PCB top"], facts.Views);
        Assert.Empty(BoardKiCadFacts.Of(null).References);
    }

    // What KiCad has that the board lacks is counted, capitals aside - the summary's button adds them.
    [Fact]
    public void What_KiCad_has_that_the_board_lacks_is_counted()
    {
        var board = new BoardData { Components = [new ComponentEntry { BoardLabel = "u1", Category = "IC" }] };

        Assert.Equal(2, BoardReadiness.NotOnBoard(board, ["U1", "U2", "C1"]));
        Assert.Equal(0, BoardReadiness.NotOnBoard(board, ["U1"]));
    }

    // The number in bold and a line per kind (owner request, 2026-11-03).
    [Fact]
    public void What_adding_from_KiCad_did_is_said_in_words()
    {
        Assert.Equal("Every component in the KiCad data is on the board already.", BoardReadiness.AddedFromKiCad(new KiCadComponentsAdded(0, 0, 0)));
        Assert.Equal("Added <b>1</b> component from KiCad project.", BoardReadiness.AddedFromKiCad(new KiCadComponentsAdded(1, 0, 0)));
        Assert.Equal("Added <b>1</b> component from KiCad project:\n- 1 with its pins", BoardReadiness.AddedFromKiCad(new KiCadComponentsAdded(1, 1, 0)));
        Assert.Equal(
            "Added <b>363</b> components from KiCad project:\n- 79 with their pins\n- 121 with a datasheet link",
            BoardReadiness.AddedFromKiCad(new KiCadComponentsAdded(363, 79, 121)));
        Assert.Equal("Added <b>3</b> components from KiCad project:\n- 2 with a datasheet link", BoardReadiness.AddedFromKiCad(new KiCadComponentsAdded(3, 0, 2)));
    }

    // ---- suggested matches (slice 3d) -----------------------------------------------------------------

    // ###########################################################################################
    // The facts say of every view what matching needs: a side of the board or a sheet, and for a
    // side the PCB file it comes from - as the loader names its views ("Issue 4 - PCB Top").
    // ###########################################################################################
    [Fact]
    public void The_KiCad_facts_tell_each_views_side_and_its_PCB_file()
    {
        var root = new KiCadProjectRoot
        {
            Project = new KiCadProjectInfo
            {
                Views =
                [
                    new KiCadProjectView { DisplayName = "Issue 4 - PCB Top", Type = "pcb_top", SourceFile = "Issue 4.kicad_pcb" },
                    new KiCadProjectView { DisplayName = "Issue 4 - PCB Bottom", Type = "PCB_BOTTOM", SourceFile = "Issue 4.kicad_pcb" },
                    new KiCadProjectView { DisplayName = " Sheet 1 ", Type = "schematic", SourceFile = "issue4.kicad_sch" },
                ],
            },
        };

        BoardKiCadFacts facts = BoardKiCadFacts.Of(root);

        Assert.Equal(
            [
                new KiCadViewToMatch("Issue 4 - PCB Top", KiCadViewSide.Top, "Issue 4"),
                new KiCadViewToMatch("Issue 4 - PCB Bottom", KiCadViewSide.Bottom, "Issue 4"),
                new KiCadViewToMatch("Sheet 1", KiCadViewSide.Sheet),
            ],
            facts.ToMatch);
        Assert.Equal(["Issue 4 - PCB Top", "Issue 4 - PCB Bottom"], BoardKiCadFacts.BoardViewsOf(root));
        Assert.Equal(["Issue 4 - PCB Top", "Issue 4 - PCB Bottom", "Sheet 1"], facts.Views);
    }

    [Fact]
    public void The_suggested_matches_are_said_in_words()
    {
        var top = new SchematicCadNameSuggestion("Top (replica)", "Issue 4.F.png", "Issue 4 - PCB Top");

        Assert.Equal("[Top (replica)] shows [Issue 4 - PCB Top]", BoardReadiness.SuggestedMatch(top));
        Assert.Equal("Give [Top (replica)] the KiCad view [Issue 4 - PCB Top]", BoardReadiness.UseMatchTip(top));
        Assert.Equal("Use all 3 suggestions", BoardReadiness.UseAllMatches(3));
        Assert.Equal("[Top (replica)] now shows [Issue 4 - PCB Top].", BoardReadiness.MatchesUsed([top]));
        Assert.Equal("2 pictures now show their KiCad views.", BoardReadiness.MatchesUsed([top, top with { SchematicName = "Bottom" }]));
        Assert.Equal("Those pictures, or their KiCad views, have been matched already.", BoardReadiness.MatchesUsed([]));
    }
}
