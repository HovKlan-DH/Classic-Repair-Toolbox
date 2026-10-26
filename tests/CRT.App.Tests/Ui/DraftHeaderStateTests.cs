using System.ComponentModel;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// DraftHeaderState - the live part of a draft's row: its save state, Undo and Redo (CRT 4.0 step
// 1's last slice, 2026-10-25). Read off the draft's session as it changes and told to the row by
// property changes, never by rebuilding the row (a rebuilt row threw a pressed button away).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class DraftHeaderStateTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    public void Dispose() => this.thisWorkspace.Dispose();

    private DraftSession OpenDraft()
    {
        var board = new BoardData { RevisionDate = "2026-10-01" };
        board.Components.Add(new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU" });

        string boardFile = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftHeaderStateTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, board);

        return DraftSession.Open(this.DraftsRoot, DraftHeaderStateTests.BoardKey)!;
    }

    private static Func<BoardContent, BoardContent?> Revision(string date) =>
        content => content with { Board = content.Board.WithRevisionDate(date) };

    // A draft no editor touched this run has no session: Saved, with nothing to undo.
    [Fact]
    public void With_no_session_the_draft_is_saved_with_nothing_to_undo()
    {
        UiTest.Run(() =>
        {
            var header = new DraftHeaderState();
            header.Follow(null);

            Assert.Equal(DraftSaveWording.Saved, header.SaveStateText);
            Assert.False(header.IsFailed);
            Assert.False(header.CanUndo);
            Assert.False(header.CanRedo);
            Assert.Equal("Nothing to undo", header.UndoTip);
        });
    }

    // ###########################################################################################
    // It follows the session: a change waiting is "Saving...", Undo names the step, a save makes it
    // Saved, an undo makes Redo name it - each told as a property change, and only what changed.
    // ###########################################################################################
    [Fact]
    public void It_follows_the_sessions_steps_and_saves_and_tells_only_what_changed()
    {
        UiTest.Run(() =>
        {
            DraftSession session = this.OpenDraft();
            var header = new DraftHeaderState();
            header.Follow(session);

            var told = new List<string>();
            ((INotifyPropertyChanged)header).PropertyChanged += (_, e) => told.Add(e.PropertyName!);

            session.Apply("Changes to [U8]", Revision("2026-10-25"));

            Assert.Equal(DraftSaveWording.Saving, header.SaveStateText);
            Assert.True(header.CanUndo);
            Assert.Equal("Undo: Changes to [U8]", header.UndoTip);
            Assert.Contains(nameof(DraftHeaderState.SaveStateText), told);
            Assert.DoesNotContain(nameof(DraftHeaderState.CanRedo), told);

            Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());
            Assert.Equal(DraftSaveWording.Saved, header.SaveStateText);

            session.Undo();

            Assert.False(header.CanUndo);
            Assert.True(header.CanRedo);
            Assert.Equal("Redo: Changes to [U8]", header.RedoTip);
        });
    }

    // A save that could not be made says why, and offers Try again.
    [Fact]
    public void A_save_that_could_not_be_made_is_said_with_its_reason()
    {
        UiTest.Run(() =>
        {
            DraftSession session = this.OpenDraft();
            var header = new DraftHeaderState();
            header.Follow(session);

            session.Apply("Changes to [U8]", Revision("2026-10-25"));
            Assert.True(DraftWorkbookStore.Edit(this.DraftsRoot, DraftHeaderStateTests.BoardKey, board => board.WithRevisionDate("2026-10-20")));
            Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, session.Save());

            Assert.True(header.IsFailed);
            Assert.Equal(DraftSaveWording.CouldNotSave(DraftWorkbookEditOutcome.ChangedOnDisk), header.SaveStateText);
        });
    }

    // Following another session lets go of the first - a change there is no longer told.
    [Fact]
    public void Forgetting_the_session_lets_go_of_it()
    {
        UiTest.Run(() =>
        {
            DraftSession session = this.OpenDraft();
            var header = new DraftHeaderState();
            header.Follow(session);
            header.Forget();

            session.Apply("Changes to [U8]", Revision("2026-10-25"));

            Assert.Null(header.SessionForTests);
            Assert.False(header.CanUndo);
            Assert.Equal(DraftSaveWording.Saved, header.SaveStateText);
        });
    }
}
