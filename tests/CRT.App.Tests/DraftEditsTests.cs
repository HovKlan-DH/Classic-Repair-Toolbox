using Avalonia.Threading;
using ClassicRepairToolbox.Tests.Ui;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftEdits - one editor's change to a draft, through the draft held in memory (CRT 4.0 step 1):
// applied as one step, saved at once, taken back when its save cannot be made. Its rules are
// DraftSession's (CRT.Data); these cover what DraftEdits adds - where the work runs, and what an
// editor is told.
//
// In "HeadlessUi": the edits run on the UI thread, and DraftManager is shared static state the
// Drafts tab's tests point elsewhere too.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class DraftEditsTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    public DraftEditsTests()
    {
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));
    }

    public void Dispose()
    {
        DraftManager.Sessions.Close(DraftManager.DraftsRoot, DraftEditsTests.BoardKey);
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static void WriteDraft()
    {
        var board = new BoardData { RevisionDate = "2026-10-01" };
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU" });

        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, DraftEditsTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, board);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, DraftEditsTests.BoardKey),
            new DraftMarker { BoardKey = DraftEditsTests.BoardKey });
    }

    private static BoardData? OnDisk() => DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, DraftEditsTests.BoardKey);

    // ###########################################################################################
    // *** THE EDIT IS WORKED OUT OFF THE UI THREAD (code review, 2026-10-25). *** Opening the draft,
    // asking whether its file changed, and running the edit with what follows it all walk the whole
    // board - and ran on the UI thread before the editor's "please wait" could paint. The edit runs on
    // the pool; it is still one step of the draft's history, and saved.
    // ###########################################################################################
    [Fact]
    public async Task An_editors_change_is_worked_out_off_the_UI_thread_and_lands_as_one_saved_step()
    {
        DraftEditsTests.WriteDraft();
        var ranOnUiThread = new List<bool>();
        DraftWorkbookEditOutcome outcome = DraftWorkbookEditOutcome.WriteFailed;

        await UiTest.RunAsync(async () =>
        {
            outcome = await DraftEdits.ApplyAndSaveAsync(
                DraftEditsTests.BoardKey,
                DraftStepWording.ComponentSaved("U1"),
                board =>
                {
                    ranOnUiThread.Add(Dispatcher.UIThread.CheckAccess());
                    return board.WithRevisionDate("2026-10-25");
                });
        });

        Assert.Equal(DraftWorkbookEditOutcome.Saved, outcome);
        Assert.Equal([false], ranOnUiThread);
        Assert.Equal("2026-10-25", DraftEditsTests.OnDisk()!.RevisionDate);

        DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, DraftEditsTests.BoardKey)!;
        Assert.Equal(DraftStepWording.ComponentSaved("U1"), session.UndoDescription);
        Assert.False(session.HasUnsavedChanges);
    }

    // A change that changes nothing is no step, and nothing to save; no draft is said as such.
    [Fact]
    public async Task A_change_that_changes_nothing_is_saved_already_and_no_draft_says_so()
    {
        DraftWorkbookEditOutcome nothing = DraftWorkbookEditOutcome.WriteFailed;
        DraftWorkbookEditOutcome noDraft = DraftWorkbookEditOutcome.Saved;

        await UiTest.RunAsync(async () =>
        {
            noDraft = await DraftEdits.ApplyAndSaveAsync(DraftEditsTests.BoardKey, "Nothing", (BoardData board) => board);

            DraftEditsTests.WriteDraft();
            nothing = await DraftEdits.ApplyAndSaveAsync(DraftEditsTests.BoardKey, "Nothing", (BoardData board) => board);
        });

        Assert.Equal(DraftWorkbookEditOutcome.NoDraft, noDraft);
        Assert.Equal(DraftWorkbookEditOutcome.Saved, nothing);
        Assert.False(DraftManager.Sessions.Find(DraftManager.DraftsRoot, DraftEditsTests.BoardKey)!.CanUndo);
    }

    // ###########################################################################################
    // A picture replaced under the name its row already cites changes no row (code review,
    // 2026-10-26): "Save to draft" wrote nothing, and the new file lay in the draft folder until some
    // other editor's write took it in. Handed `takeInNewFiles`, a change that changes nothing still
    // writes the draft, so the file is taken in now - or the save says it could not be.
    // ###########################################################################################
    [Fact]
    public async Task A_file_replaced_under_the_name_its_row_cites_is_taken_in_by_the_save()
    {
        var board = new BoardData { RevisionDate = "2026-10-01" };
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU" });
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U1", Name = "Top", File = "Component images/u1.png" });
        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, DraftEditsTests.BoardKey);
        BoardFiles.Write(boardFile, BoardDocumentMapping.ToDocument(
            new BoardContent(board, [], []),
            name => name,
            [new BoardDocument.StoredFileRow { Name = "Component images/u1.png", Sha256 = new string('a', 64), Size = 3 }]));
        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, DraftEditsTests.BoardKey),
            new DraftMarker { BoardKey = DraftEditsTests.BoardKey });

        string replaced = this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "C64", "250407", "Component images", "u1.png"), "the new picture");
        string sha = DataTreeLayout.Sha256Of(replaced);
        DraftWorkbookEditOutcome outcome = DraftWorkbookEditOutcome.WriteFailed;

        await UiTest.RunAsync(async () =>
        {
            outcome = await DraftEdits.ApplyAndSaveAsync(DraftEditsTests.BoardKey, DraftStepWording.ComponentSaved("U1"), (BoardData unchanged) => unchanged, takeInNewFiles: true);
        });

        Assert.Equal(DraftWorkbookEditOutcome.Saved, outcome);
        Assert.Equal(sha, Assert.Single(BoardFiles.ReadDocument(boardFile)!.Files).Sha256);
        Assert.False(File.Exists(replaced));
    }

    // ###########################################################################################
    // "Undo changes to [U1]" PUTS A PICTURE REPLACED UNDER ITS OWN NAME BACK (code review,
    // 2026-10-26): the rows went back and the draft's list kept the contributor's bytes - the
    // component read as unchanged while the next submission sent the replaced picture.
    // ###########################################################################################
    [Fact]
    public async Task Undo_changes_puts_a_picture_replaced_under_its_own_name_back_as_published()
    {
        var board = new BoardData { RevisionDate = "2026-10-01" };
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU" });
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U1", Name = "Top", File = "Component images/u1.png" });

        string publishedBoardFile = Path.Combine(this.thisWorkspace.Root, "Data", "Commodore", "C64", "250407", DataTreeLayout.BoardFileName);
        Directory.CreateDirectory(Path.GetDirectoryName(publishedBoardFile)!);
        var theirs = new BoardDocument.StoredFileRow { Name = "Component images/u1.png", Sha256 = new string('a', 64), Size = 9 };
        BoardFiles.Write(publishedBoardFile, BoardDocumentMapping.ToDocument(new BoardContent(board, [], []), name => name, [theirs]));

        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, DraftEditsTests.BoardKey);
        string draftFolder = Path.GetDirectoryName(boardFile)!;
        BoardDocument.StoredFileRow mine = StoredFiles.Put(draftFolder, "Component images/u1.png", "the contributor's picture");
        BoardFiles.Write(boardFile, BoardDocumentMapping.ToDocument(new BoardContent(board, [], []), name => name, [mine]));
        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, DraftEditsTests.BoardKey),
            new DraftMarker { BoardKey = DraftEditsTests.BoardKey });

        DraftWorkbookEditOutcome outcome = DraftWorkbookEditOutcome.WriteFailed;

        await UiTest.RunAsync(async () =>
        {
            outcome = await DraftEdits.UndoComponentChangesAsync(DraftEditsTests.BoardKey, "U1", board, publishedBoardFile);
        });

        Assert.Equal(DraftWorkbookEditOutcome.Saved, outcome);
        Assert.Equal(theirs.Sha256, Assert.Single(BoardFiles.ReadDocument(boardFile)!.Files).Sha256);
        Assert.False(File.Exists(Path.Combine(draftFolder, DataTreeLayout.StoredFilePath(mine.Sha256, mine.Name).Replace('/', Path.DirectorySeparatorChar))));

        // And nothing of it reads as changed any more - the draft's list as the session holds it.
        DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, DraftEditsTests.BoardKey)!;
        Assert.False(ComponentRevert.HasChanges(session.Content.Board, board, "U1", DraftFiles.ListedIn(publishedBoardFile), session.ListedFiles));
    }

    // ###########################################################################################
    // A file another program wrote while nothing of the draft waited is read first, so the change
    // lands on the draft as it is on disk - and that reading is one undoable step of its own.
    // ###########################################################################################
    [Fact]
    public async Task A_file_written_by_another_program_is_read_first_and_the_change_lands_on_it()
    {
        DraftEditsTests.WriteDraft();

        await UiTest.RunAsync(async () =>
        {
            Assert.NotNull(DraftManager.Sessions.Get(DraftManager.DraftsRoot, DraftEditsTests.BoardKey));

            Assert.True(DraftWorkbookStore.Edit(DraftManager.DraftsRoot, DraftEditsTests.BoardKey, board => board.WithRevisionDate("2026-10-20")));

            Assert.Equal(
                DraftWorkbookEditOutcome.Saved,
                await DraftEdits.ApplyAndSaveAsync(DraftEditsTests.BoardKey, "Caption", board => board.WithCaption("Commodore 64", "250407")));
        });

        BoardData disk = DraftEditsTests.OnDisk()!;
        Assert.Equal("2026-10-20", disk.RevisionDate);
        Assert.Equal("Commodore 64", disk.HardwareName);

        DraftSession session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, DraftEditsTests.BoardKey)!;
        Assert.Equal("Caption", session.UndoDescription);
        Assert.True(session.Undo());
        Assert.Equal(DraftSaveWording.ReadAgainStep, session.UndoDescription);
    }
}
