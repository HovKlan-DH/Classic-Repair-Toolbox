using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardReadiness - what a board still lacks ("Board status", BoardStatusPage) and what is left to do
// once its KiCad data is in (the KiCad import's summary), each by what the board holds.
// ###########################################################################################
public sealed class BoardReadinessTests
{
    private static readonly KiCadWorkLeft NothingLeft = new(Match: false, Align: false, Components: false, NotOnBoard: 0, Mark: false, Signals: false);

    private static BoardContent Content(BoardData board, params KiCadCalibrationEntry[] calibrations) => new(board, calibrations, []);

    [Fact]
    public void A_board_just_made_has_everything_left_to_do()
    {
        KiCadWorkLeft left = BoardReadiness.WorkLeftAfterKiCad(Content(new BoardData()), new BoardKiCadFacts([], ["PCB top"]));

        Assert.Equal(new KiCadWorkLeft(Match: true, Align: true, Components: true, NotOnBoard: 0, Mark: true, Signals: true), left);
    }

    // ###########################################################################################
    // A picture shows a view once its CAD name is one of the views; it is aligned once it has a
    // calibration; the components KiCad has but the board lacks are counted; a component marked on no
    // picture leaves marking to do.
    // ###########################################################################################
    [Fact]
    public void With_KiCad_data_what_is_left_is_judged_by_what_the_board_holds()
    {
        var board = new BoardData
        {
            Schematics =
            [
                new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png", CadName = "PCB top" },
                new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "Images/bottom.png", CadName = "PCB bottom" },
                new BoardSchematicEntry { SchematicName = "Photo", SchematicImageFile = "Images/photo.jpg", CadName = "Not a view" },
            ],
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Category = "IC" },
                new ComponentEntry { BoardLabel = "U2", Category = "IC", Region = "PAL" },
                new ComponentEntry { BoardLabel = "u2", Category = "IC", Region = "NTSC" },
            ],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "0", Y = "0", Width = "1", Height = "1" }],
        };

        var kiCad = new BoardKiCadFacts(["U1", "U2", "U3", "C1"], ["PCB top", "PCB bottom"]);

        // Matched - the photo shows no view, and that is nothing left - and one of the two aligned, the
        // calibration found whatever its capitals. U3 and C1 are not on the board, U2 is marked nowhere.
        Assert.Equal(
            new KiCadWorkLeft(Match: false, Align: true, Components: true, NotOnBoard: 2, Mark: true, Signals: true),
            BoardReadiness.WorkLeftAfterKiCad(Content(board, new KiCadCalibrationEntry { SchematicName = "top", CadName = "PCB top" }), kiCad));

        // None matched yet: to do.
        var unmatched = new BoardData { Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }] };

        Assert.True(BoardReadiness.WorkLeftAfterKiCad(Content(unmatched), kiCad).Match);
    }

    [Fact]
    public void A_board_with_everything_has_nothing_left_to_do()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png", CadName = "PCB top" }],
            Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "0", Y = "0", Width = "1", Height = "1" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "Clock", KiCadNetName = "PHI2" }],
        };

        Assert.Equal(
            NothingLeft,
            BoardReadiness.WorkLeftAfterKiCad(
                Content(board, new KiCadCalibrationEntry { SchematicName = "Top", CadName = "PCB top" }),
                new BoardKiCadFacts(["U1"], ["PCB top", "PCB bottom"])));
    }

    // ###########################################################################################
    // KiCad files that could not be read are not KiCad data with nothing in it (code review,
    // 2026-10-26: "0 found in KiCad - All 0 are on the board" was ticked done, and matching asked for
    // views there were none of): matching, aligning and the signals need data CRT can read, and the
    // components are judged on their own.
    // ###########################################################################################
    [Fact]
    public void KiCad_data_that_could_not_be_read_leaves_nothing_to_match_and_counts_no_components()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }],
            Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
        };

        Assert.Equal(NothingLeft with { Mark = true }, BoardReadiness.WorkLeftAfterKiCad(Content(board), BoardKiCadFacts.Of(null)));
        Assert.True(BoardReadiness.WorkLeftAfterKiCad(Content(new BoardData()), BoardKiCadFacts.Of(null)).Components);
    }

    // ###########################################################################################
    // A picture that shows no KiCad view never holds matching up (code review, 2026-10-26: one photo
    // of the ports beside six views - a PCB's two sides and four sheets - left matching to do for
    // ever), nor do the views no picture shows yet.
    // ###########################################################################################
    [Fact]
    public void A_picture_showing_no_view_never_holds_matching_up()
    {
        var board = new BoardData
        {
            Schematics =
            [
                new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png", CadName = "PCB top" },
                new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "Images/bottom.png", CadName = "PCB bottom" },
                new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "Images/sheet1.png", CadName = "Sheet 1" },
                new BoardSchematicEntry { SchematicName = "Sheet 2", SchematicImageFile = "Images/sheet2.png", CadName = "Sheet 2" },
                new BoardSchematicEntry { SchematicName = "Ports", SchematicImageFile = "Images/ports.jpg" },
            ],
            Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "0", Y = "0", Width = "1", Height = "1" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "Clock", KiCadNetName = "PHI2" }],
        };

        KiCadCalibrationEntry[] calibrations = [.. board.Schematics.Take(4).Select(row => new KiCadCalibrationEntry { SchematicName = row.SchematicName, CadName = row.CadName })];

        Assert.Equal(
            NothingLeft,
            BoardReadiness.WorkLeftAfterKiCad(
                Content(board, calibrations),
                new BoardKiCadFacts(["U1"], ["PCB top", "PCB bottom", "Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4"])));
    }

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

        KiCadWorkLeft withKiCad = BoardReadiness.WorkLeftAfterKiCad(Content(board), new BoardKiCadFacts(["U1", "U2", "C1"], []));

        Assert.Equal(2, withKiCad.NotOnBoard);
        Assert.True(withKiCad.Components);

        KiCadWorkLeft allThere = BoardReadiness.WorkLeftAfterKiCad(Content(board), new BoardKiCadFacts(["U1"], []));

        Assert.Equal(0, allThere.NotOnBoard);
        Assert.False(allThere.Components);
    }

    [Fact]
    public void What_adding_from_KiCad_did_is_said_in_words()
    {
        Assert.Equal("Every component in the KiCad data is on the board already.", BoardReadiness.AddedFromKiCad(new KiCadComponentsAdded(0, 0, 0)));
        Assert.Equal("Added 1 component from KiCad.", BoardReadiness.AddedFromKiCad(new KiCadComponentsAdded(1, 0, 0)));
        Assert.Equal("Added 1 component from KiCad - 1 with its pins.", BoardReadiness.AddedFromKiCad(new KiCadComponentsAdded(1, 1, 0)));
        Assert.Equal(
            "Added 214 components from KiCad - 40 with their pins, 12 with a datasheet link.",
            BoardReadiness.AddedFromKiCad(new KiCadComponentsAdded(214, 40, 12)));
        Assert.Equal("Added 3 components from KiCad - 2 with a datasheet link.", BoardReadiness.AddedFromKiCad(new KiCadComponentsAdded(3, 0, 2)));
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
