using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardReadiness - getting a new board ready (CRT 4.0 step 3: "Getting this board ready ... On an
// existing board, the same panel shows What's missing"): each step done by what the board holds,
// the KiCad project optional - matching and aligning skipped without it - the first step with
// something to do next, and "Check and submit" next once nothing else is.
// ###########################################################################################
public sealed class BoardReadinessTests
{
    private static BoardContent Content(BoardData board, params KiCadCalibrationEntry[] calibrations) => new(board, calibrations, []);

    private static Dictionary<ReadinessStepKind, ReadinessStep> ByKind(IEnumerable<ReadinessStep> steps) => steps.ToDictionary(step => step.Kind);

    [Fact]
    public void A_board_just_made_has_its_name_and_its_pictures_next()
    {
        List<ReadinessStep> steps = BoardReadiness.Steps(" Commodore 64 - 250407 ", Content(new BoardData()), hasKiCadFiles: false, kiCad: null);

        Assert.Equal(Enumerable.Range(1, 9), steps.Select(step => step.Number));
        Assert.Equal(
            [
                ReadinessState.Done, ReadinessState.Next, ReadinessState.Optional, ReadinessState.Skipped, ReadinessState.Skipped,
                ReadinessState.Later, ReadinessState.Later, ReadinessState.Skipped, ReadinessState.Later,
            ],
            steps.Select(step => step.State));
        Assert.Equal("Commodore 64 - 250407", steps[0].Detail);
        Assert.Equal("Needs the KiCad project.", steps[3].Detail);
    }

    // ###########################################################################################
    // With KiCad data: a picture shows a view once its CAD name is one of the views; it is aligned
    // once it has a calibration; the components KiCad has but the board lacks are counted; a component
    // marked on no picture is counted too.
    // ###########################################################################################
    [Fact]
    public void With_KiCad_data_each_step_counts_what_is_left()
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

        // Every view has its picture - the photo shows none, and that is no step left. Both views are
        // sides of the board.
        var kiCad = new BoardKiCadFacts(["U1", "U2", "U3", "C1"], ["PCB top", "PCB bottom"], BoardViews: ["PCB top", "PCB bottom"]);

        Dictionary<ReadinessStepKind, ReadinessStep> steps = ByKind(BoardReadiness.Steps(
            "C64",
            Content(board, new KiCadCalibrationEntry { SchematicName = "top", CadName = "PCB top" }),
            hasKiCadFiles: true,
            kiCad));

        Assert.Equal(ReadinessState.Done, steps[ReadinessStepKind.KiCad].State);
        Assert.Equal(ReadinessState.Done, steps[ReadinessStepKind.Match].State);
        Assert.Equal("2 of 3 pictures show a KiCad view.", steps[ReadinessStepKind.Match].Detail);

        // A view no picture shows yet, beside a picture showing none: worth finishing, never needed -
        // CRT cannot know whether that picture shows the view (code review, 2026-10-26).
        ReadinessStep partly = ByKind(BoardReadiness.Steps("C64", Content(board), hasKiCadFiles: true, kiCad with { Views = ["PCB top", "PCB bottom", "Main sheet"] }))[ReadinessStepKind.Match];
        Assert.Equal(ReadinessState.Optional, partly.State);
        Assert.Equal("2 of 3 pictures show a KiCad view - name the view of any other picture that shows one.", partly.Detail);

        // None matched yet: to do.
        Assert.Equal(
            "0 of 1 pictures show a KiCad view - name the view each one shows.",
            ByKind(BoardReadiness.Steps("C64", Content(new BoardData { Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }] }), hasKiCadFiles: true, kiCad))[ReadinessStepKind.Match].Detail);

        Assert.Equal(ReadinessState.Next, steps[ReadinessStepKind.Align].State);
        Assert.Equal("1 of 2 pictures aligned.", steps[ReadinessStepKind.Align].Detail);

        Assert.Equal("Components - 4 found in KiCad", steps[ReadinessStepKind.Components].Title);
        Assert.Equal("2 of them are not on the board yet.", steps[ReadinessStepKind.Components].Detail);
        Assert.Equal(ReadinessState.Later, steps[ReadinessStepKind.Components].State);

        // A picture is aligned - "Mark all automatically" is offered for the rest (slice 3c).
        Assert.Equal(
            "1 of 2 components are not marked on any picture yet. A picture aligned with KiCad draws them all with Mark all automatically.",
            steps[ReadinessStepKind.Mark].Detail);
        Assert.Equal(ReadinessState.Later, steps[ReadinessStepKind.Signals].State);
        Assert.Equal(ReadinessState.Later, steps[ReadinessStepKind.Submit].State);
    }

    // Everything in place: nothing left but to send it - "Check and submit" is next, never done.
    [Fact]
    public void A_board_with_everything_has_submitting_next()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png", CadName = "PCB top" }],
            Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "0", Y = "0", Width = "1", Height = "1" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "Clock", KiCadNetName = "PHI2" }],
        };

        List<ReadinessStep> steps = BoardReadiness.Steps(
            "C64",
            Content(board, new KiCadCalibrationEntry { SchematicName = "Top", CadName = "PCB top" }),
            hasKiCadFiles: true,
            new BoardKiCadFacts(["U1"], ["PCB top", "PCB bottom"]));

        Assert.All(steps.Take(8), step => Assert.Equal(ReadinessState.Done, step.State));
        Assert.Equal(ReadinessState.Next, steps[8].State);
        Assert.Equal("Everything is in place - send the board for review.", steps[8].Detail);
    }

    // ###########################################################################################
    // KiCad files that could not be read are not KiCad data with nothing in it (code review,
    // 2026-10-26: "0 found in KiCad - All 0 are on the board" was ticked done, and matching asked for
    // views there were none of): the KiCad step says so and waits, matching, aligning and the signals
    // need data CRT can read, and the components are counted on their own.
    // ###########################################################################################
    [Fact]
    public void KiCad_data_that_could_not_be_read_is_said_and_counts_no_components()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }],
            Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
        };

        Dictionary<ReadinessStepKind, ReadinessStep> steps = ByKind(BoardReadiness.Steps("C64", Content(board), hasKiCadFiles: true, BoardKiCadFacts.Of(null)));

        Assert.Equal(ReadinessState.Next, steps[ReadinessStepKind.KiCad].State);
        Assert.Equal(BoardReadiness.KiCadUnreadable, steps[ReadinessStepKind.KiCad].Detail);
        Assert.Equal(ReadinessState.Skipped, steps[ReadinessStepKind.Match].State);
        Assert.Equal(ReadinessState.Skipped, steps[ReadinessStepKind.Align].State);
        Assert.Equal(ReadinessState.Skipped, steps[ReadinessStepKind.Signals].State);
        Assert.Equal("Components (1)", steps[ReadinessStepKind.Components].Title);
        Assert.Equal(0, steps[ReadinessStepKind.Components].FromKiCad);
    }

    // ###########################################################################################
    // A picture that shows no KiCad view never holds matching up (code review, 2026-10-26: one photo
    // of the ports beside six views - a PCB's two sides and four sheets - left matching "next" for
    // ever, and "Check and submit" never came). Some matched but not all is worth finishing, never
    // needed.
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

        Dictionary<ReadinessStepKind, ReadinessStep> steps = ByKind(BoardReadiness.Steps(
            "C128",
            Content(board, calibrations),
            hasKiCadFiles: true,
            new BoardKiCadFacts(["U1"], ["PCB top", "PCB bottom", "Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4"])));

        Assert.Equal(ReadinessState.Optional, steps[ReadinessStepKind.Match].State);
        Assert.Equal("Match pictures", BoardReadiness.ActionOf(steps[ReadinessStepKind.Match]));
        Assert.Equal(ReadinessState.Done, steps[ReadinessStepKind.Align].State);
        Assert.Equal(ReadinessState.Next, steps[ReadinessStepKind.Submit].State);
    }

    // ###########################################################################################
    // "Mark all automatically" is offered only for a picture aligned with a SIDE OF THE BOARD (code
    // review, 2026-10-26): one aligned with a schematic sheet answered "shows a schematic sheet" -
    // and nothing is offered while the KiCad data is being read.
    // ###########################################################################################
    [Fact]
    public void Marking_automatically_is_offered_only_for_a_picture_aligned_with_the_board()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Sheet", SchematicImageFile = "Images/sheet.png", CadName = "Main sheet" }],
            Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
        };

        BoardContent content = Content(board, new KiCadCalibrationEntry { SchematicName = "Sheet", CadName = "Main sheet" });
        var facts = new BoardKiCadFacts(["U1"], ["PCB top", "Main sheet"], BoardViews: ["PCB top"]);

        Assert.Equal(
            "1 of 1 components are not marked on any picture yet.",
            ByKind(BoardReadiness.Steps("C64", content, hasKiCadFiles: true, facts))[ReadinessStepKind.Mark].Detail);

        Assert.Equal(
            "1 of 1 components are not marked on any picture yet.",
            ByKind(BoardReadiness.Steps("C64", content, hasKiCadFiles: true, kiCad: null))[ReadinessStepKind.Mark].Detail);
    }

    // While the KiCad data is still being read, matching says so and the components are counted alone.
    [Fact]
    public void While_the_KiCad_data_is_read_matching_says_so()
    {
        var board = new BoardData { Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }] };

        Dictionary<ReadinessStepKind, ReadinessStep> steps = ByKind(BoardReadiness.Steps("C64", Content(board), hasKiCadFiles: true, kiCad: null));

        Assert.Equal("Reading the KiCad data...", steps[ReadinessStepKind.Match].Detail);
        Assert.Equal(ReadinessState.Next, steps[ReadinessStepKind.Match].State);
        Assert.Equal("Components", steps[ReadinessStepKind.Components].Title);
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
        Assert.False(missing.IsEmpty);
        Assert.True(BoardReadiness.Missing(new BoardData()).IsEmpty);
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

    // ###########################################################################################
    // The components step offers to add, in one go, every component KiCad has that the board lacks
    // ("Components - 214 found in KiCad [Add]"); with all on the board, or no KiCad data read yet, its
    // button adds one by hand.
    // ###########################################################################################
    [Fact]
    public void The_components_step_offers_what_KiCad_has_that_the_board_lacks()
    {
        var board = new BoardData { Components = [new ComponentEntry { BoardLabel = "u1", Category = "IC" }] };

        ReadinessStep withKiCad = BoardReadiness.Steps("C64", Content(board), hasKiCadFiles: true, new BoardKiCadFacts(["U1", "U2", "C1"], []))
            .Single(step => step.Kind == ReadinessStepKind.Components);

        Assert.Equal(2, withKiCad.FromKiCad);
        Assert.Equal("Add 2 from KiCad", BoardReadiness.ActionOf(withKiCad));

        ReadinessStep allThere = BoardReadiness.Steps("C64", Content(board), hasKiCadFiles: true, new BoardKiCadFacts(["U1"], []))
            .Single(step => step.Kind == ReadinessStepKind.Components);

        Assert.Equal(0, allThere.FromKiCad);
        Assert.Equal("Add component", BoardReadiness.ActionOf(allThere));

        ReadinessStep whileReading = BoardReadiness.Steps("C64", Content(board), hasKiCadFiles: true, kiCad: null)
            .Single(step => step.Kind == ReadinessStepKind.Components);

        Assert.Equal("Add component", BoardReadiness.ActionOf(whileReading));
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
}
