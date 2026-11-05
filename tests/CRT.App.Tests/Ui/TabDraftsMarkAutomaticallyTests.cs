using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "MARK ALL AUTOMATICALLY" (TabDrafts.MarkAutomatically.cs - CRT 4.0 step 3, slice 3c: "on aligned
// PCB pictures, Mark all automatically draws each component's rectangle from its footprint").
//
// A picture card's button, on for a picture aligned with KiCad: every component not marked on the
// picture gets its footprint's rectangle as ONE step of the draft, saved; what came of it - or why
// nothing could - is said in the Pictures part. The numbers are KiCadFootprintMarksTests' board:
// pads over 9..51 x 9..31 mm on a 420 x 220 picture.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabDraftsMarkAutomaticallyTests : IDisposable
{
    private const string Identity = "Commodore/C64/999001/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    public TabDraftsMarkAutomaticallyTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static BoardData Board() => new()
    {
        Schematics =
        [
            new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png", CadName = "PCB top" },
            new BoardSchematicEntry { SchematicName = "Photo", SchematicImageFile = "Images/photo.jpg", CadName = "PCB top" },
        ],
        Components =
        [
            new ComponentEntry { BoardLabel = "U1", Category = "IC" },
            new ComponentEntry { BoardLabel = "U2", Category = "IC" },
        ],
        ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "1", Y = "1", Width = "5", Height = "5" }],
    };

    // The draft: the board (Board() unless said), "Top" aligned (the photo not) - or `aligned` - and a
    // KiCad file listed.
    private static void WriteDraft(BoardData? board = null, string aligned = "Top")
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, Identity);
        string folder = Path.GetDirectoryName(boardFile)!;
        Directory.CreateDirectory(folder);

        byte[] bytes = [1, 2, 3];
        string sha = Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(bytes));
        const string KiCad = "KiCad data/board.kicad_pcb";
        string stored = Path.Combine(folder, DataTreeLayout.StoredFilePath(sha, KiCad).Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(stored)!);
        File.WriteAllBytes(stored, bytes);

        BoardFiles.Write(boardFile, BoardDocumentMapping.ToDocument(
            new BoardContent(board ?? Board(), [new KiCadCalibrationEntry { SchematicName = aligned, CadName = "PCB top", ScaleX = 1, ScaleY = 1 }], []),
            name => name,
            [new BoardDocument.StoredFileRow { Name = KiCad, Sha256 = sha, Size = bytes.Length }]));

        DraftMarkerStore.Save(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, Identity), new DraftMarker { BoardKey = Identity });
    }

    private static KiCadPcbPad Pad(double x, double y) => new()
    {
        Number = "1",
        AbsoluteCenter = new KiCadPoint2D { X = x, Y = y },
        Size = new KiCadSize2D { X = 2, Y = 2 },
        Layers = ["*.Cu"],
    };

    private static KiCadProjectRoot KiCad() => new()
    {
        Project = new KiCadProjectInfo { Views = [new KiCadProjectView { DisplayName = "PCB top", Type = "pcb_top", SourceIndex = 0 }] },
        Pcb =
        [
            new KiCadPcb
            {
                Footprints =
                [
                    new KiCadPcbFootprint { Reference = "U1", Pads = [Pad(10, 10), Pad(20, 10)] },
                    new KiCadPcbFootprint { Reference = "U2", Pads = [Pad(30, 20), Pad(40, 20)] },
                    new KiCadPcbFootprint { Reference = "Q1", Pads = [Pad(50, 30)] },
                ],
            },
        ],
    };

    private static TabDrafts Tab(HardwareBoardEntry entry, KiCadProjectRoot? kiCad)
    {
        var tab = new TabDrafts
        {
            HardwareBoardsOverrideForTests = [entry],
            PublishedBoardOverrideForTests = _ => null,
            BoardKiCadOverrideForTests = _ => Task.FromResult(kiCad),
            PictureSizeOverrideForTests = _ => (420, 220),
            ReloadBoardOverrideForTests = _ => { },
        };

        tab.RefreshDrafts();
        tab.ShowBoard(entry, Board(), "PAL", null);
        return tab;
    }

    // Long enough for a busy machine, and failing HERE when it never comes - half a second, given up
    // silently, failed later as something else (2026-11-01).
    private static async Task Until(Func<bool> condition)
    {
        for (int attempt = 0; attempt < 3000 && !condition(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }

        Assert.True(condition(), "What this test waits for never came.");
    }

    private static HardwareBoardEntry Entry() => new() { HardwareName = "Commodore 64", BoardName = "999001", ExcelDataFile = Identity, IsDraftOnly = true };

    // ###########################################################################################
    // The aligned picture's card offers it, the photo's not; pressed, every component not marked on
    // the picture gets its rectangle - U1, marked by hand, is left as it is; Q1 is no component of the
    // board - as ONE step, saved, and said in the part.
    // ###########################################################################################
    [Fact]
    public async Task An_aligned_picture_gets_every_missing_rectangle_from_KiCad_as_one_step()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            ComponentPaneBoardPictureRow top = tab.ComponentPaneForTests.BoardPicturesForTests.Single(row => row.Name == "Top");
            ComponentPaneBoardPictureRow photo = tab.ComponentPaneForTests.BoardPicturesForTests.Single(row => row.Name == "Photo");
            Assert.True(top.CanMarkAutomatically);
            Assert.False(photo.CanMarkAutomatically);
            Assert.Equal(PictureToolWording.MarkAutomaticallyNeedsAlignTip, photo.MarkAutomaticallyTip);

            tab.ComponentPaneForTests.RequestMarkAutomatically(top);
            await Until(() => tab.ComponentPaneForTests.PicturesNoteForTests is not null);

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, Identity)!;
            Assert.Equal(
                [("Top", "U1", "1", "1", "5", "5"), ("Top", "U2", "195", "95", "130", "30")],
                session.Content.Board.ComponentHighlights.Select(row => (row.SchematicName, row.BoardLabel, row.X, row.Y, row.Width, row.Height)));
            Assert.Equal(DraftStepWording.ComponentsMarkedFromKiCad("Top", 1), session.UndoDescription);
            Assert.False(session.HasUnsavedChanges);
            Assert.Equal(PictureToolWording.MarkedAutomatically(1, "Top"), tab.ComponentPaneForTests.PicturesNoteForTests);

            // Pressed again: nothing left to mark, and no step.
            tab.ComponentPaneForTests.RequestMarkAutomatically(tab.ComponentPaneForTests.BoardPicturesForTests.Single(row => row.Name == "Top"));
            await Until(() => tab.ComponentPaneForTests.PicturesNoteForTests == PictureToolWording.MarkedAutomatically(0, "Top"));

            Assert.Equal(PictureToolWording.MarkedAutomatically(0, "Top"), tab.ComponentPaneForTests.PicturesNoteForTests);
            Assert.Equal(DraftStepWording.ComponentsMarkedFromKiCad("Top", 1), session.UndoDescription);
        });
    }

    // ###########################################################################################
    // *** A PICTURE WITH MARKS ON IT ASKS FIRST (owner request, 2026-11-01: "it should show a warning
    // popup"). *** Told how many are marked and how many would be; turned down, nothing changes.
    // ###########################################################################################
    [Fact]
    public async Task A_picture_with_marks_asks_first_and_turned_down_marks_nothing()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            var asked = new List<(string Picture, int Marked, int ToMark)>();
            bool answer = false;

            tab.AskMarkOverMarksOverrideForTests = (picture, marked, toMark) =>
            {
                asked.Add((picture, marked, toMark));
                return Task.FromResult(answer);
            };

            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);
            await tab.MarkAutomaticallyAsync("Top");

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, Identity)!;

            Assert.Equal(("Top", 1, 1), Assert.Single(asked));
            Assert.Equal(["U1"], session.Content.Board.ComponentHighlights.Select(row => row.BoardLabel));
            Assert.False(session.CanUndo);

            answer = true;
            await tab.MarkAutomaticallyAsync("Top");

            Assert.Equal(2, asked.Count);
            Assert.Equal(["U1", "U2"], session.Content.Board.ComponentHighlights.Select(row => row.BoardLabel));

            // U1's mark, drawn by hand, is where it was - what the question promises.
            Assert.Equal(("1", "1"), (session.Content.Board.ComponentHighlights[0].X, session.Content.Board.ComponentHighlights[0].Y));

            // Nothing left to mark: nothing to weigh, so no question, and the note says so as ever.
            await tab.MarkAutomaticallyAsync("Top");

            Assert.Equal(2, asked.Count);
            Assert.Equal(PictureToolWording.MarkedAutomatically(0, "Top"), tab.ComponentPaneForTests.PicturesNoteForTests);
        });
    }

    // With nothing marked yet there is nothing to weigh; and the marking after "Align KiCad" never asks.
    [Fact]
    public async Task A_picture_with_no_marks_asks_nothing()
    {
        BoardData board = Board();
        board.ComponentHighlights.Clear();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            tab.AskMarkOverMarksOverrideForTests = (_, _, _) => throw new InvalidOperationException("Asked, with nothing marked.");

            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);
            await tab.MarkAutomaticallyAsync("Top");

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, Identity)!;
            Assert.Equal(["U1", "U2"], session.Content.Board.ComponentHighlights.Select(row => row.BoardLabel));
        });
    }

    // What the question says is what happens, with its numbers right - one is not several.
    [Fact]
    public void The_question_says_what_stays_and_what_comes()
    {
        Assert.Equal("358 components are already marked on [Bottom].", PictureToolWording.MarkOverMarksQuestion("Bottom", 358));
        Assert.Equal("1 component is already marked on [Bottom].", PictureToolWording.MarkOverMarksQuestion("Bottom", 1));
        Assert.StartsWith("The 4 components with no mark on it get one from KiCad.", PictureToolWording.MarkOverMarksDetail(4));
        Assert.StartsWith("The 1 component with no mark on it gets one from KiCad.", PictureToolWording.MarkOverMarksDetail(1));
        Assert.Contains("stays exactly where it is", PictureToolWording.MarkOverMarksDetail(4));
    }

    // ###########################################################################################
    // *** AFTER "ALIGN KiCad"'s DONE, A PICTURE WITH NO MARKS IS MARKED BY ITSELF (owner request,
    // 2026-11-01: Done marks without asking, Undo takes it back). *** Every component seen on it gets
    // its rectangle as ONE step, saved, and the part says the alignment is in and the marks are made.
    // ###########################################################################################
    [Fact]
    public async Task After_aligning_a_picture_with_no_marks_every_component_is_marked_as_one_step()
    {
        BoardData board = Board();
        board.ComponentHighlights.Clear();
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            await tab.MarkAfterAligningAsync("Top", "Images/top.png");

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, Identity)!;
            Assert.Equal(["U1", "U2"], session.Content.Board.ComponentHighlights.Where(row => row.SchematicName == "Top").Select(row => row.BoardLabel));
            Assert.Equal(DraftStepWording.ComponentsMarkedFromKiCad("Top", 2), session.UndoDescription);
            Assert.False(session.HasUnsavedChanges);
            Assert.Equal(PictureToolWording.MarkedAfterAligning(2, "Top"), tab.ComponentPaneForTests.PicturesNoteForTests);

            // One Undo takes them all back.
            session.Undo();
            Assert.Empty(session.Content.Board.ComponentHighlights);
        });
    }

    // ###########################################################################################
    // A picture with ANY mark on it is left alone after aligning - its marks were drawn, by hand or
    // from KiCad, and another round is the contributor's to ask for ("Mark all automatically").
    // ###########################################################################################
    [Fact]
    public async Task After_aligning_a_picture_with_a_mark_on_it_is_left_alone()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            // "Top" has U1 marked by hand.
            await tab.MarkAfterAligningAsync("Top", "Images/top.png");

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, Identity)!;
            Assert.Equal(["U1"], session.Content.Board.ComponentHighlights.Select(row => row.BoardLabel));
            Assert.False(session.CanUndo);
            Assert.Null(tab.ComponentPaneForTests.PicturesNoteForTests);
        });
    }

    // ###########################################################################################
    // *** ALIGNED AGAIN, WHAT KiCad MARKED MOVES WITH IT (owner request, 2026-11-01: "if I then again do the
    // Align KiCad then it seems it will not readjust the components, which I think it should"). *** One step,
    // said, which Undo takes back; a mark drawn by hand stays where it is.
    // ###########################################################################################
    [Fact]
    public async Task Aligned_again_the_marks_KiCad_drew_move_and_those_by_hand_stay()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, Identity)!;
            KiCadCalibrationEntry now = session.Content.Calibrations.Single(row => row.SchematicName == "Top");

            // What KiCad drew under the alignment before: U2, where its footprint lay then. U1 was drawn by hand.
            var before = new KiCadCalibrationEntry { SchematicName = "Top", CadName = now.CadName, ScaleX = now.ScaleX, ScaleY = now.ScaleY, OffsetX = now.OffsetX + 40, OffsetY = now.OffsetY + 20, MirrorX = now.MirrorX, MirrorY = now.MirrorY };
            KiCadProjectRoot root = KiCad();
            KiCadMarkPlan plan = KiCadFootprintMarks.Plan(session.Content, "Top", root, (420, 220), "Images/top.png");
            ComponentHighlightEntry drawn = KiCadFootprintMarks
                .For(session.Content.Board, "Top", plan.Pcb!, plan.Bottom, KiCadFootprintMarks.ValuesOf(before), 420, 220, skipMarked: false)
                .Single(mark => mark.BoardLabel == "U2");

            Assert.True(session.Apply("KiCad's mark", content => content with { Board = KiCadFootprintMarks.WithMarks(content.Board, [drawn]) }));
            Assert.Equal(DraftWorkbookEditOutcome.Saved, await session.SaveAsync());

            ComponentHighlightEntry byHand = session.Content.Board.ComponentHighlights.Single(row => row.BoardLabel == "U1");
            ComponentHighlightEntry there = KiCadFootprintMarks
                .For(session.Content.Board, "Top", plan.Pcb!, plan.Bottom, plan.Calibration, 420, 220, skipMarked: false)
                .Single(mark => mark.BoardLabel == "U2");

            await tab.MarkAfterAligningAsync("Top", "Images/top.png", before);

            ComponentHighlightEntry moved = session.Content.Board.ComponentHighlights.Single(row => row.BoardLabel == "U2");
            Assert.Equal((there.X, there.Y, there.Width, there.Height), (moved.X, moved.Y, moved.Width, moved.Height));

            ComponentHighlightEntry stayed = session.Content.Board.ComponentHighlights.Single(row => row.BoardLabel == "U1");
            Assert.Equal((byHand.X, byHand.Y, byHand.Width, byHand.Height), (stayed.X, stayed.Y, stayed.Width, stayed.Height));

            Assert.Equal(DraftStepWording.ComponentsFollowedAlignment("Top", 1), session.UndoDescription);
            Assert.Equal(PictureToolWording.MarksFollowedAlignment(1, "Top"), tab.ComponentPaneForTests.PicturesNoteForTests);
            Assert.False(session.HasUnsavedChanges);

            session.Undo();
            ComponentHighlightEntry back = session.Content.Board.ComponentHighlights.Single(row => row.BoardLabel == "U2");
            Assert.Equal((drawn.X, drawn.Y), (back.X, back.Y));
        });
    }

    // A schematic sheet: U1 drawn in two units, U2 in one, and nothing else of the board's.
    private static KiCadProjectRoot KiCadWithSheet()
    {
        KiCadProjectRoot kiCad = KiCad();
        kiCad.Project.Views.Add(new KiCadProjectView { DisplayName = "Main sheet", Type = "schematic", SourceIndex = 0 });
        kiCad.Schematics.Add(new KiCadSchematic
        {
            Wires = [new KiCadSchematicPathItem { Points = [new KiCadPoint2D { X = 0, Y = 0 }, new KiCadPoint2D { X = 42, Y = 22 }] }],
            Symbols =
            [
                new KiCadSchematicSymbol { Reference = "U1", Extent = new KiCadSheetExtent { Left = 1, Top = 1, Right = 5, Bottom = 5 } },
                new KiCadSchematicSymbol { Reference = "U1", Extent = new KiCadSheetExtent { Left = 11, Top = 1, Right = 15, Bottom = 5 } },
                new KiCadSchematicSymbol { Reference = "U2", Extent = new KiCadSheetExtent { Left = 21, Top = 1, Right = 25, Bottom = 5 } },
            ],
        });

        return kiCad;
    }

    private static BoardData BoardOfSheet() => new()
    {
        Schematics = [new BoardSchematicEntry { SchematicName = "Sheet", SchematicImageFile = "Images/sheet.png", CadName = "Main sheet" }],
        Components =
        [
            new ComponentEntry { BoardLabel = "U1", Category = "IC" },
            new ComponentEntry { BoardLabel = "U2", Category = "IC" },
        ],
    };

    // ###########################################################################################
    // *** DONE ON A SCHEMATIC SHEET MARKS ITS COMPONENTS, AS ON A SIDE OF THE BOARD (owner report,
    // 2026-11-04: "nor has it automatically added any components when I aligned"). *** One step, which
    // "Undo" takes back; and a chip in two units is one component in what is said, with two rectangles.
    // ###########################################################################################
    [Fact]
    public async Task After_aligning_a_schematic_sheet_its_components_are_marked_from_their_symbols()
    {
        WriteDraft(BoardOfSheet(), aligned: "Sheet");

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCadWithSheet());
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            await tab.MarkAfterAligningAsync("Sheet", "Images/sheet.png");

            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, Identity)!;
            Assert.Equal(["U1", "U1", "U2"], session.Content.Board.ComponentHighlights.Select(row => row.BoardLabel));
            Assert.All(session.Content.Board.ComponentHighlights, row => Assert.Equal("Sheet", row.SchematicName));
            Assert.Equal(PictureToolWording.MarkedAfterAligning(2, "Sheet"), tab.ComponentPaneForTests.PicturesNoteForTests);

            Assert.True(session.CanUndo);
            session.Undo();
            Assert.Empty(session.Content.Board.ComponentHighlights);
        });
    }

    // "Mark all" counts components too, and on a sheet it says so in a sheet's words when nothing is there.
    [Fact]
    public async Task Mark_all_on_a_schematic_sheet_says_how_many_components_not_how_many_rectangles()
    {
        WriteDraft(BoardOfSheet(), aligned: "Sheet");

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCadWithSheet());
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            await tab.MarkAutomaticallyAsync("Sheet", "Images/sheet.png");

            Assert.Equal(PictureToolWording.MarkedAutomatically(2, "Sheet"), tab.ComponentPaneForTests.PicturesNoteForTests);
            Assert.Equal(3, DraftManager.Sessions.Find(DraftManager.DraftsRoot, Identity)!.Content.Board.ComponentHighlights.Count);
        });
    }

    // ###########################################################################################
    // THE CARD SAYS WHAT COMES NEXT (owner request, 2026-11-01): a picture with a KiCad view not aligned
    // yet, a side of the board or a schematic sheet, tells to align it, and that its Done marks the
    // components; aligned, it says no more. "Mark all automatically" is offered on all of them.
    // ###########################################################################################
    [Fact]
    public async Task A_card_says_to_align_first_and_offers_marking_on_a_schematic_sheet_too()
    {
        BoardData board = Board();
        board.Schematics.Add(new BoardSchematicEntry { SchematicName = "Sheet", SchematicImageFile = "Images/sheet.png", CadName = "Main sheet" });
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            tab.KiCadViewsOverrideForTests = _ => Task.FromResult<IReadOnlyList<string>>(["PCB top", "Main sheet"]);
            tab.KiCadBoardSidesOverrideForTests = _ => ["PCB top"];
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);
            await Until(() => tab.ComponentPaneForTests.BoardPicturesForTests.Any(row => row.AlignsFirst));

            ComponentPaneBoardPictureRow Row(string name) => tab.ComponentPaneForTests.BoardPicturesForTests.Single(row => row.Name == name);

            // Aligned: nothing more to say. Not aligned: align it first.
            Assert.False(Row("Top").AlignsFirst);
            Assert.True(Row("Photo").AlignsFirst);
            Assert.Equal(PictureToolWording.AlignFirst, Row("Photo").AlignFirstText);

            // A schematic sheet is told the same, and greyed out until it is aligned.
            Assert.True(Row("Sheet").AlignsFirst);
            Assert.True(Row("Sheet").ShowsMarkAutomatically);
            Assert.False(Row("Sheet").CanMarkAutomatically);
            Assert.True(Row("Photo").ShowsMarkAutomatically);
            Assert.True(Row("Top").ShowsMarkAutomatically);
        });
    }

    // ###########################################################################################
    // Nothing marked is said for what it is (code review, 2026-10-26: each read "marked already"): a
    // board with no components yet - aligned before "Add N from KiCad" - and one whose components
    // have no footprint on the picture.
    // ###########################################################################################
    [Fact]
    public async Task Nothing_to_mark_says_why()
    {
        WriteDraft(board: new BoardData { Schematics = Board().Schematics });

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            await tab.MarkAutomaticallyAsync("Top");
            Assert.Equal(PictureToolWording.NoComponentsYet("Top"), tab.ComponentPaneForTests.PicturesNoteForTests);
        });
    }

    [Fact]
    public async Task Components_KiCad_has_no_footprint_for_are_said_so()
    {
        WriteDraft(board: new BoardData { Schematics = Board().Schematics, Components = [new ComponentEntry { BoardLabel = "X9", Category = "Misc" }] });

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            await tab.MarkAutomaticallyAsync("Top");
            Assert.Equal(PictureToolWording.NoFootprintOnPicture("Top"), tab.ComponentPaneForTests.PicturesNoteForTests);
        });
    }

    // ###########################################################################################
    // Not while a picture tool is on (code review, 2026-10-26): "Mark components" on the picture, then
    // "Mark all automatically" from the Contribute tab - the editor's Done wrote the picture's
    // rectangles as it held them, and the marks went without a word.
    // ###########################################################################################
    [Fact]
    public async Task Nothing_is_marked_while_a_picture_tool_is_on()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            tab.PictureToolOnOverrideForTests = () => true;
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            await tab.MarkAutomaticallyAsync("Top");

            Assert.Equal(PictureToolWording.PictureToolOn, tab.ComponentPaneForTests.PicturesNoteForTests);
            Assert.Single(DraftManager.Sessions.Find(DraftManager.DraftsRoot, Identity)!.Content.Board.ComponentHighlights);
        });
    }

    // What was marked is said only while the draft is as it left it - an Undo takes the words with
    // it (code review, 2026-10-26; Board status's note has the same fix).
    [Fact]
    public async Task The_note_goes_once_the_draft_moves_on()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            await tab.MarkAutomaticallyAsync("Top");
            Assert.NotNull(tab.ComponentPaneForTests.PicturesNoteForTests);

            DraftManager.Sessions.Find(DraftManager.DraftsRoot, Identity)!.Undo();
            await Until(() => tab.ComponentPaneForTests.PicturesNoteForTests is null);

            Assert.Null(tab.ComponentPaneForTests.PicturesNoteForTests);
        });
    }

    // ###########################################################################################
    // A save that cannot be made changes nothing and says why: the draft written outside CRT while
    // a change of its own waited to be saved - the marks taken back out, the board on screen not read
    // again.
    // ###########################################################################################
    [Fact]
    public async Task A_save_that_fails_takes_the_marks_back_and_says_why()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            var reloaded = new List<string>();
            TabDrafts tab = Tab(Entry(), KiCad());
            tab.ReloadBoardOverrideForTests = reloaded.Add;
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            // A change waiting to be saved, then the file written by another program.
            DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, Identity)!;
            session.Apply("Waiting", content => content with { Board = content.Board.WithComponentSections(components: [.. content.Board.Components, new ComponentEntry { BoardLabel = "R1", Category = "Resistor" }]) });
            WriteDraft(board: new BoardData { Schematics = Board().Schematics, Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }] });

            await tab.MarkAutomaticallyAsync("Top");

            Assert.Equal(DraftSaveWording.EditNotSaved(DraftWorkbookEditOutcome.ChangedOnDisk), tab.ComponentPaneForTests.PicturesNoteForTests);
            Assert.DoesNotContain(session.Content.Board.ComponentHighlights, row => row.BoardLabel == "U2");
            Assert.Empty(reloaded);
        });
    }

    // KiCad data that cannot be read marks nothing and says so - the draft untouched.
    [Fact]
    public async Task Unreadable_KiCad_data_marks_nothing_and_says_so()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), kiCad: null);
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            await tab.MarkAutomaticallyAsync("Top");

            Assert.Equal(PictureToolWording.KiCadNotRead, tab.ComponentPaneForTests.PicturesNoteForTests);
            Assert.Single(DraftManager.Sessions.Find(DraftManager.DraftsRoot, Identity)!.Content.Board.ComponentHighlights);
        });
    }

    // ###########################################################################################
    // A card's "Traces on the preview" (2026-10-29): both sides' traces of an aligned image, from the
    // same plan "Mark all automatically" works with - a top picture's own side is the top's copper,
    // its other side the bottom's - and on an image not aligned, why not, in the contributor's words.
    // ###########################################################################################
    [Fact]
    public async Task The_traces_for_a_cards_preview_come_from_the_plan_mark_all_automatically_uses()
    {
        WriteDraft();
        KiCadProjectRoot kiCad = KiCad();
        kiCad.Pcb[0].Routing.Segments.AddRange([
            new KiCadPcbSegment { Start = new KiCadPoint2D { X = 10, Y = 10 }, End = new KiCadPoint2D { X = 40, Y = 20 }, Width = 0.5, Layer = "B.Cu" },
            new KiCadPcbSegment { Start = new KiCadPoint2D { X = 10, Y = 20 }, End = new KiCadPoint2D { X = 40, Y = 10 }, Width = 0.5, Layer = "F.Cu" },
        ]);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), kiCad);
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            (Handlers.Geometry.PreviewTraces traces, string? problem) = await tab.PreviewTracesAsync("Top");

            Assert.Null(problem);
            Handlers.Geometry.PreviewTrace other = Assert.Single(traces.OtherSide);
            Assert.True(other.From.X < other.To.X && other.From.Y < other.To.Y);
            Handlers.Geometry.PreviewTrace own = Assert.Single(traces.ThisSide);
            Assert.True(own.From.X < own.To.X && own.From.Y > own.To.Y);

            (traces, problem) = await tab.PreviewTracesAsync("Photo");

            Assert.Same(Handlers.Geometry.PreviewTraces.None, traces);
            Assert.Equal(PictureToolWording.NotAligned("Photo"), problem);
        });
    }

    // ###########################################################################################
    // THE CARDS' TRACES SHARE ONE READ OF THE KICAD PROJECT (code review, 2026-10-30): the cards are
    // made again at every undo, and each ticked card read the project by itself - four at once.
    // ###########################################################################################
    [Fact]
    public async Task Cards_asking_for_their_traces_at_once_share_one_read_of_the_KiCad_project()
    {
        WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            var answer = new TaskCompletionSource<KiCadProjectRoot?>();
            int reads = 0;
            tab.BoardKiCadOverrideForTests = _ =>
            {
                reads++;
                return answer.Task;
            };

            Task<(Handlers.Geometry.PreviewTraces, string?)> top = tab.PreviewTracesAsync("Top");
            Task<(Handlers.Geometry.PreviewTraces, string?)> photo = tab.PreviewTracesAsync("Photo");
            Assert.Equal(1, reads);

            answer.SetResult(KiCad());
            await Task.WhenAll(top, photo);
            Assert.Null((await top).Item2);
            Assert.Equal(PictureToolWording.NotAligned("Photo"), (await photo).Item2);

            // A read done is not kept: the next ask reads again (the KiCad files may have changed). The
            // part reads the KiCad views by itself meanwhile, so counted from here.
            tab.BoardKiCadOverrideForTests = _ =>
            {
                reads++;
                return Task.FromResult<KiCadProjectRoot?>(KiCad());
            };

            int before = reads;
            Task<(Handlers.Geometry.PreviewTraces, string?)> again = tab.PreviewTracesAsync("Top");
            Assert.Equal(before + 1, reads);
            await again;
        });
    }

    // ###########################################################################################
    // OF TWO PICTURES OF ONE NAME, "MARK ALL AUTOMATICALLY" PLANS FROM THE CARD PRESSED (code review,
    // 2026-10-30): it planned from the first of the name - its KiCad view and its size - so the second
    // card's rectangles landed off its components. The card is its name AND its image file, as its
    // traces are read.
    // ###########################################################################################
    [Fact]
    public async Task Of_two_pictures_of_one_name_mark_all_automatically_plans_from_the_card_pressed()
    {
        BoardData board = Board();
        board.Schematics.Clear();
        board.Schematics.AddRange([
            new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top-scan.png" },
            new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png", CadName = "PCB top" },
        ]);
        WriteDraft(board);

        await UiTest.RunAsync(async () =>
        {
            TabDrafts tab = Tab(Entry(), KiCad());
            var sized = new List<string>();
            tab.PictureSizeOverrideForTests = file =>
            {
                sized.Add(file);
                return (420, 220);
            };

            await tab.OpenBoardPartAsync(PaneBoardPart.Pictures);

            ComponentPaneBoardPictureRow second = tab.ComponentPaneForTests.BoardPicturesForTests.Single(row => row.File == "Images/top.png");
            Assert.True(second.CanMarkAutomatically);

            tab.ComponentPaneForTests.RequestMarkAutomatically(second);
            await Until(() => tab.ComponentPaneForTests.PicturesNoteForTests is not null);

            Assert.Equal(PictureToolWording.MarkedAutomatically(1, "Top"), tab.ComponentPaneForTests.PicturesNoteForTests);

            // The card's own file, every time it is read: the question before marking reads it too.
            Assert.Equal(["Images/top.png"], sized.Distinct());
        });
    }
}
