using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

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

    // The draft: the board (Board() unless said), "Top" aligned (the photo not), and a KiCad file listed.
    private static void WriteDraft(BoardData? board = null)
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
            new BoardContent(board ?? Board(), [new KiCadCalibrationEntry { SchematicName = "Top", CadName = "PCB top", ScaleX = 1, ScaleY = 1 }], []),
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

    private static async Task Until(Func<bool> condition)
    {
        for (int attempt = 0; attempt < 50 && !condition(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }
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
    // it (code review, 2026-10-26; the checklist's note had the same fix).
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
}
