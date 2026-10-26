using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftSession - a draft held once in memory, shared by every editor, saved by itself (CRT 4.0
// step 1, "Nothing is ever lost"), and DraftSessions, one per draft.
//
// The promises held here: an edit is seen at once and written a moment later; one history covers
// every editor's changes and undo puts back exactly what was there; saved means the file holds what
// is shown; a file changed outside CRT is never written over without being asked; a change made
// while a save runs is not lost; and what follows an edit (highlights, calibrations) follows it.
// ###########################################################################################
public sealed class DraftSessionTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();
    private DateTimeOffset thisNow = new(2026, 10, 25, 12, 0, 0, TimeSpan.Zero);

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string BoardFile => DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftSessionTests.BoardKey);

    public void Dispose() => this.thisWorkspace.Dispose();

    private static ComponentEntry Component(string label, string friendlyName) =>
        new() { BoardLabel = label, FriendlyName = friendlyName, TechnicalNameOrValue = $"part {label}" };

    private static BoardData Board(params ComponentEntry[] components) => new() { Components = [.. components] };

    // The draft's board file, holding `board` with `calibrations`.
    private void WriteDraft(BoardData board, params KiCadCalibrationEntry[] calibrations) =>
        BoardFiles.Write(this.BoardFile, BoardDocumentMapping.ToDocument(new BoardContent(board, calibrations, []), name => name, []));

    private DraftSession Open()
    {
        DraftSession? session = DraftSession.Open(this.DraftsRoot, DraftSessionTests.BoardKey, () => this.thisNow);
        Assert.NotNull(session);
        return session;
    }

    private DraftSession OpenWith(params ComponentEntry[] components)
    {
        this.WriteDraft(DraftSessionTests.Board(components));
        return this.Open();
    }

    // An edit renaming component `label`'s friendly name - a NEW board and component list, as every
    // draft writer builds them, so the content before it stays as it was.
    private static Func<BoardContent, BoardContent?> Name(string label, string friendlyName) =>
        content => content with
        {
            Board = new BoardData
            {
                Schematics = content.Board.Schematics,
                Components = content.Board.Components
                    .Select(component => component.BoardLabel == label ? DraftSessionTests.Component(label, friendlyName) : component)
                    .ToList(),
                ComponentHighlights = content.Board.ComponentHighlights,
            },
        };

    private static string NameOf(BoardContent content, string label) =>
        content.Board.Components.Single(component => component.BoardLabel == label).FriendlyName;

    private string NameOnDisk(string label) =>
        DraftSessionTests.NameOf(BoardFiles.ReadContent(this.BoardFile)!, label);

    // ------------------------------------------------------------------ Opening

    [Fact]
    public void A_draft_is_read_once_and_there_is_no_session_without_one()
    {
        Assert.Null(DraftSession.Open(this.DraftsRoot, DraftSessionTests.BoardKey));

        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));

        Assert.Equal("CPU", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal(DraftSaveState.Saved, session.SaveState);
        Assert.False(session.CanUndo);
    }

    // ------------------------------------------------------------------ Editing

    [Fact]
    public void An_edit_is_in_memory_at_once_seen_by_every_listener_and_on_disk_only_once_saved()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        var seen = new List<DraftSessionChange>();
        session.Changed += (_, change) => seen.Add(change);
        object table = new();

        Assert.True(session.Apply("Name of U1", DraftSessionTests.Name("U1", "CPU 6510"), origin: table));

        Assert.Equal("CPU 6510", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal("CPU", this.NameOnDisk("U1"));
        Assert.Equal(DraftSaveState.Pending, session.SaveState);

        DraftSessionChange change = Assert.Single(seen);
        Assert.Equal(DraftSessionChangeKind.Edit, change.Kind);
        Assert.Equal("Name of U1", change.Description);
        Assert.Same(table, change.Origin);

        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());
        Assert.Equal("CPU 6510", this.NameOnDisk("U1"));
        Assert.Equal(DraftSaveState.Saved, session.SaveState);
    }

    [Fact]
    public void An_edit_that_changes_nothing_is_no_step()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));

        Assert.False(session.Apply("nothing", content => null));
        Assert.False(session.Apply("the same", content => content));
        Assert.False(session.CanUndo);
        Assert.Equal(DraftSaveState.Saved, session.SaveState);
    }

    // ------------------------------------------------------------------ One history for every editor

    [Fact]
    public void Undo_and_redo_walk_every_editors_steps_and_say_what_each_was()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"), DraftSessionTests.Component("U2", "VIC"));

        session.Apply("Name of U1", DraftSessionTests.Name("U1", "CPU 6510"), origin: "table");
        session.Apply("Name of U2", DraftSessionTests.Name("U2", "VIC II"), origin: "component editor");

        Assert.Equal("Name of U2", session.UndoDescription);
        Assert.True(session.Undo());
        Assert.Equal("VIC", DraftSessionTests.NameOf(session.Content, "U2"));
        Assert.Equal("CPU 6510", DraftSessionTests.NameOf(session.Content, "U1"));

        Assert.Equal("Name of U1", session.UndoDescription);
        Assert.Equal("Name of U2", session.RedoDescription);

        Assert.True(session.Undo());
        Assert.False(session.CanUndo);
        Assert.Equal("CPU", DraftSessionTests.NameOf(session.Content, "U1"));

        Assert.True(session.Redo());
        Assert.True(session.Redo());
        Assert.False(session.CanRedo);
        Assert.Equal("VIC II", DraftSessionTests.NameOf(session.Content, "U2"));
    }

    [Fact]
    public void A_new_edit_after_an_undo_drops_what_could_be_redone()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));

        session.Apply("one", DraftSessionTests.Name("U1", "one"));
        session.Undo();
        session.Apply("two", DraftSessionTests.Name("U1", "two"));

        Assert.False(session.CanRedo);
        Assert.Equal("two", session.UndoDescription);
    }

    [Fact]
    public void The_history_keeps_the_latest_steps_up_to_its_bound()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));

        for (int i = 1; i <= DraftSession.MaxSteps + 5; i++)
        {
            session.Apply($"step {i}", DraftSessionTests.Name("U1", $"name {i}"));
        }

        int undone = 0;
        while (session.Undo())
        {
            undone++;
        }

        Assert.Equal(DraftSession.MaxSteps, undone);
        Assert.Equal("name 5", DraftSessionTests.NameOf(session.Content, "U1"));
    }

    // Undoing back to what the file holds is saved again - nothing to write.
    [Fact]
    public void Saved_means_the_file_holds_what_is_shown_whichever_way_it_got_there()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));

        session.Apply("one", DraftSessionTests.Name("U1", "one"));
        Assert.True(session.HasUnsavedChanges);

        session.Undo();
        Assert.False(session.HasUnsavedChanges);
        Assert.Equal(DraftSaveState.Saved, session.SaveState);

        session.Redo();
        session.Save();
        session.Undo();
        Assert.Equal(DraftSaveState.Pending, session.SaveState);
    }

    // ------------------------------------------------------------------ Saving by itself

    [Fact]
    public void A_save_is_due_a_second_after_the_last_change_and_not_before()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));

        Assert.False(session.IsSaveDue(this.thisNow.AddHours(1)));

        session.Apply("one", DraftSessionTests.Name("U1", "one"));
        Assert.False(session.IsSaveDue(this.thisNow.AddMilliseconds(900)));

        // A second change starts the second again.
        this.thisNow = this.thisNow.AddMilliseconds(800);
        session.Apply("two", DraftSessionTests.Name("U1", "two"));
        Assert.False(session.IsSaveDue(this.thisNow.AddMilliseconds(900)));
        Assert.True(session.IsSaveDue(this.thisNow + DraftSession.SaveDelay));

        session.Save();
        Assert.False(session.IsSaveDue(this.thisNow.AddHours(1)));
    }

    // A change made while the file is being written stays unsaved and is written next time.
    [Fact]
    public void A_change_made_while_a_save_runs_is_not_lost()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        session.Apply("one", DraftSessionTests.Name("U1", "one"));

        session.WhileWritingForTests = () =>
        {
            session.WhileWritingForTests = null;
            session.Apply("two", DraftSessionTests.Name("U1", "two"));
        };

        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

        Assert.Equal("one", this.NameOnDisk("U1"));
        Assert.Equal(DraftSaveState.Pending, session.SaveState);

        session.Save();
        Assert.Equal("two", this.NameOnDisk("U1"));
        Assert.Equal(DraftSaveState.Saved, session.SaveState);
    }

    // ###########################################################################################
    // A save stopped by an exception nobody expected ends as a FAILED save, never stuck in Saving
    // (code review, 2026-10-25): stuck, no save was ever due again and nothing said so. The
    // exception still reaches the host, and the next change is due as usual.
    // ###########################################################################################
    [Fact]
    public void A_save_stopped_by_an_unexpected_exception_ends_failed_and_the_next_change_is_due()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        session.Apply("one", DraftSessionTests.Name("U1", "one"));
        session.WhileWritingForTests = () => throw new ArgumentException("something nobody expected");

        Assert.Throws<ArgumentException>(() => session.Save());

        Assert.Equal(DraftSaveState.Failed, session.SaveState);
        Assert.Equal(DraftWorkbookEditOutcome.WriteFailed, session.SaveFailure);
        Assert.Equal("CPU", this.NameOnDisk("U1"));

        session.WhileWritingForTests = null;
        session.Apply("two", DraftSessionTests.Name("U1", "two"));

        Assert.True(session.IsSaveDue(this.thisNow.Add(DraftSession.SaveDelay)));
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());
        Assert.Equal("two", this.NameOnDisk("U1"));
    }

    // A rectangle drawn or moved in the label editor at any zoom lands between whole pixels
    // (code review, 2026-10-26): the 4.0 board file refused the fraction, so every such save failed.
    [Fact]
    public void A_label_editor_save_drawn_between_whole_pixels_is_saved_rounded()
    {
        this.WriteDraft(new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }],
            Components = [DraftSessionTests.Component("U1", "CPU")],
        });
        DraftSession session = this.Open();
        LabelEditorSaveRow[] drawn = [new() { SchematicName = "Top", BoardLabel = "U1", X = 123.4567, Y = 10.5, Width = 40.25, Height = 19.75 }];

        session.Apply("Highlights on [Top]", content => content with { Board = LabelEditorBoardWriter.ApplyLabelEditorSave(content.Board, "Top", drawn, region: "") });

        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());
        ComponentHighlightEntry saved = Assert.Single(BoardFiles.ReadContent(this.BoardFile)!.Board.ComponentHighlights);
        Assert.Equal(["123", "11", "40", "20"], new[] { saved.X, saved.Y, saved.Width, saved.Height });
    }

    // ###########################################################################################
    // SaveAsync - every host's save, the one way (code review, 2026-10-26: the table, the pane and the
    // header each had a copy): off the calling thread, and never throwing - a save stopped by an
    // exception nobody expected ends as a failed save, said by the session.
    // ###########################################################################################
    [Fact]
    public async Task SaveAsync_saves_off_the_calling_thread_and_never_throws()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        int callingThread = Environment.CurrentManagedThreadId;
        int? savingThread = null;

        session.Apply("one", DraftSessionTests.Name("U1", "one"));
        session.WhileWritingForTests = () => savingThread = Environment.CurrentManagedThreadId;

        Assert.Equal(DraftWorkbookEditOutcome.Saved, await session.SaveAsync());
        Assert.NotEqual(callingThread, savingThread);
        Assert.Equal("one", this.NameOnDisk("U1"));

        session.Apply("two", DraftSessionTests.Name("U1", "two"));
        session.WhileWritingForTests = () => throw new ArgumentException("something nobody expected");

        Assert.Equal(DraftWorkbookEditOutcome.WriteFailed, await session.SaveAsync());
        Assert.Equal(DraftSaveState.Failed, session.SaveState);
    }

    // ------------------------------------------------------------------ Files and the history

    // U1 with one picture citing `file`, or none - a new board and picture list, as every draft
    // writer builds them.
    private static Func<BoardContent, BoardContent?> Picture(string? file) =>
        content => content with
        {
            Board = new BoardData
            {
                Schematics = content.Board.Schematics,
                Components = content.Board.Components,
                ComponentImages = file is null ? [] : [new ComponentImageEntry { BoardLabel = "U1", Name = "Top", File = file }],
            },
        };

    private string DraftFolder => Path.GetDirectoryName(this.BoardFile)!;

    private string PutInDraftFolder(string name, string content)
    {
        string path = Path.Combine(this.DraftFolder, name.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, content);
        return path;
    }

    private List<BoardDocument.StoredFileRow> ListedOnDisk() => BoardFiles.ReadDocument(this.BoardFile)!.Files;

    // ###########################################################################################
    // A FILE WHOSE ROW WAS REMOVED COMES BACK WITH THE ROW (code review, 2026-10-26). The save after
    // the removal drops the file from the board file's list (no row cites it) and used to delete its
    // bytes from the draft's store - and the history holds rows, not the list - so the undo brought
    // the row back citing a file that was gone, for good.
    // ###########################################################################################
    [Fact]
    public void A_picked_file_whose_row_was_removed_comes_back_with_the_row_on_undo()
    {
        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "CPU")));
        string sha = DataTreeLayout.Sha256Of(this.PutInDraftFolder("Component images/u1.png", "the contributor's picture"));
        DraftSession session = this.Open();

        session.Apply("Picture of U1", DraftSessionTests.Picture("Component images/u1.png"));
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

        session.Apply("No picture of U1", DraftSessionTests.Picture(null));
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());
        Assert.Empty(this.ListedOnDisk());

        Assert.True(session.Undo());
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

        BoardDocument.StoredFileRow listed = Assert.Single(this.ListedOnDisk());
        Assert.Equal("Component images/u1.png", listed.Name);
        Assert.Equal(sha, listed.Sha256);
        Assert.True(File.Exists(Path.Combine(this.DraftFolder, DataTreeLayout.StoredFilePath(sha, listed.Name))));
    }

    // The same for a file the draft lists from the published board (its bytes are the tree's).
    [Fact]
    public void A_published_file_whose_row_was_removed_is_listed_again_on_undo()
    {
        BoardDocument.StoredFileRow published = StoredFiles.EntryFor("Component images/u1.png", "the published picture");
        BoardData board = DraftSessionTests.Board(DraftSessionTests.Component("U1", "CPU"));
        board.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U1", Name = "Top", File = "Component images/u1.png" });
        BoardFiles.Write(this.BoardFile, BoardDocumentMapping.ToDocument(new BoardContent(board, [], []), name => name, [published]));
        DraftSession session = this.Open();

        session.Apply("No picture of U1", DraftSessionTests.Picture(null));
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());
        Assert.Empty(this.ListedOnDisk());

        Assert.True(session.Undo());
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

        Assert.Equal(published.Sha256, Assert.Single(this.ListedOnDisk()).Sha256);
    }

    // What no row cites stays only while this run's history can reach it: the next run's first
    // save tidies its bytes away, as every save did before.
    [Fact]
    public void Bytes_no_row_cites_go_with_the_first_save_of_the_next_session()
    {
        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "CPU")));
        this.PutInDraftFolder("Component images/u1.png", "the contributor's picture");
        DraftSession first = this.Open();
        first.Apply("Picture of U1", DraftSessionTests.Picture("Component images/u1.png"));
        first.Save();
        first.Apply("No picture of U1", DraftSessionTests.Picture(null));
        first.Save();
        Assert.True(Directory.Exists(Path.Combine(this.DraftFolder, DataTreeLayout.FilesFolder)));

        DraftSession next = this.Open();
        next.Apply("Name of U1", DraftSessionTests.Name("U1", "CPU 6510"));
        Assert.Equal(DraftWorkbookEditOutcome.Saved, next.Save());

        Assert.False(Directory.Exists(Path.Combine(this.DraftFolder, DataTreeLayout.FilesFolder)));
    }

    // ------------------------------------------------------------------ A file changed outside

    // ###########################################################################################
    // *** ASKING WHETHER THE FILE CHANGED DOES NOT READ IT AGAIN WHILE ITS STAMP STANDS (code review,
    // 2026-10-26). *** The table asks at every switch to the Contribute tab and the pane at every
    // component opened, on the UI thread, and a board file of a few megabytes was hashed each time.
    // Once the file was seen unchanged with a stamp that had settled (two writes in one clock tick
    // share a stamp - StampSettling), the same stamp is the same file; a write that moves it is still
    // noticed.
    // ###########################################################################################
    [Fact]
    public void Asking_whether_the_file_changed_hashes_it_only_when_its_stamp_moved()
    {
        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "CPU")));
        File.SetLastWriteTimeUtc(this.BoardFile, DateTime.UtcNow.AddHours(-1));
        DraftSession session = this.Open();
        int before = session.FingerprintsTakenForTests;

        Assert.False(session.FileChangedOutside());
        Assert.False(session.FileChangedOutside());
        Assert.False(session.ReadFileIfChanged(DraftSaveWording.ReadAgainStep));
        Assert.Equal(before + 1, session.FingerprintsTakenForTests);

        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "changed outside")));

        Assert.True(session.FileChangedOutside());
        Assert.True(session.ReadFileIfChanged(DraftSaveWording.ReadAgainStep));
        Assert.Equal("changed outside", DraftSessionTests.NameOf(session.Content, "U1"));
    }

    // ###########################################################################################
    // *** ITS OWN SAVE IS NO OUTSIDE CHANGE (code review, 2026-10-26). *** Between a save's write and
    // the moment it records the new file as its own, another thread asking saw the new bytes against
    // the old fingerprint - and read the file back as a step of its own, Redo lost, "updated from
    // outside" said with nothing outside. While a write of the session's own runs, nothing is judged.
    // ###########################################################################################
    [Fact]
    public void A_save_still_running_is_not_taken_for_a_change_made_outside()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        session.Apply("one", DraftSessionTests.Name("U1", "one"));
        session.Apply("two", DraftSessionTests.Name("U1", "two"));
        session.Undo();

        // Asked from a thread of its own - a task waited for can run on the waiting thread, the one
        // writing, which may enter its own lock again.
        bool? changedWhileWriting = null;
        session.AfterWritingForTests = () =>
        {
            var asking = new System.Threading.Thread(() => changedWhileWriting = session.FileChangedOutside());
            asking.Start();
            asking.Join();
        };

        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());
        session.AfterWritingForTests = null;

        Assert.False(changedWhileWriting);
        Assert.False(session.ReadFileIfChanged(DraftSaveWording.ReadAgainStep));
        Assert.True(session.CanRedo);
    }

    // A file still settling - just written - is hashed each time: its stamp could be shared.
    [Fact]
    public void A_file_just_written_is_hashed_each_time_it_is_asked_about()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        int before = session.FingerprintsTakenForTests;

        Assert.False(session.FileChangedOutside());
        Assert.False(session.FileChangedOutside());

        Assert.Equal(before + 2, session.FingerprintsTakenForTests);
    }

    [Fact]
    public void A_file_changed_outside_CRT_is_not_written_over_without_being_asked()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        session.Apply("mine", DraftSessionTests.Name("U1", "mine"));

        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "theirs")));

        Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, session.Save());
        Assert.Equal("theirs", this.NameOnDisk("U1"));
        Assert.Equal(DraftSaveState.Failed, session.SaveState);
        Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, session.SaveFailure);

        // Not retried in a loop: due again only after a change, or when asked.
        Assert.False(session.IsSaveDue(this.thisNow.AddHours(1)));
        session.Retry();
        Assert.True(session.IsSaveDue(this.thisNow));

        // "Keep the version on screen".
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save(overwriteOutsideChange: true));
        Assert.Equal("mine", this.NameOnDisk("U1"));
        Assert.Null(session.SaveFailure);
    }

    // "Use the file's version": one undoable step, and saved - the file holds what is shown.
    [Fact]
    public void Reading_the_files_version_is_one_undoable_step()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        session.Apply("mine", DraftSessionTests.Name("U1", "mine"));
        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "theirs")));
        session.Save();

        Assert.True(session.ReloadFromFile("Read again from the draft file"));

        Assert.Equal("theirs", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal(DraftSaveState.Saved, session.SaveState);

        Assert.True(session.Undo());
        Assert.Equal("mine", DraftSessionTests.NameOf(session.Content, "U1"));

        // Kept on screen again, it is written over the file's - the user's own undo asked for it.
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());
        Assert.Equal("mine", this.NameOnDisk("U1"));
    }

    // ###########################################################################################
    // *** ANOTHER PART OF CRT WROTE THE FILE (ReadFileIfChanged, 2026-10-25). *** Read again as one
    // undoable step while nothing of the session's waits; nothing when the file is what the session
    // read or wrote; and nothing while a change of its own waits - that save refuses and the host
    // asks which version to keep, so neither is thrown away unasked.
    // ###########################################################################################
    [Fact]
    public void A_file_written_elsewhere_is_read_again_only_while_nothing_of_the_sessions_waits()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));

        Assert.False(session.ReadFileIfChanged("Written elsewhere"));
        Assert.False(session.CanUndo);

        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "component editor")));

        Assert.True(session.ReadFileIfChanged("Written elsewhere"));
        Assert.Equal("component editor", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal(DraftSaveState.Saved, session.SaveState);
        Assert.Equal("Written elsewhere", session.UndoDescription);

        // A change of the session's waiting: the file is left to its save, which refuses.
        session.Apply("mine", DraftSessionTests.Name("U1", "mine"));
        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "label editor")));

        Assert.False(session.ReadFileIfChanged("Written elsewhere"));
        Assert.Equal("mine", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, session.Save());
    }

    // The session's own clock says when a save is due - what the host's timer asks.
    [Fact]
    public void IsSaveDueNow_asks_the_sessions_own_clock()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        session.Apply("one", DraftSessionTests.Name("U1", "one"));

        Assert.False(session.IsSaveDueNow);

        this.thisNow += DraftSession.SaveDelay;
        Assert.True(session.IsSaveDueNow);
    }

    // ------------------------------------------------------------------ What follows an edit

    [Fact]
    public void A_renamed_schematic_takes_its_highlights_and_calibration_along()
    {
        var board = new BoardData
        {
            Components = { DraftSessionTests.Component("U1", "CPU") },
            Schematics = { new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png" } },
            ComponentHighlights = { new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "1", Y = "2", Width = "3", Height = "4" } },
        };

        this.WriteDraft(board, new KiCadCalibrationEntry { SchematicName = "Top", CadName = "PCB Top", ScaleX = 2 });
        DraftSession session = this.Open();

        session.Apply("Rename Top", content => content with
        {
            Board = content.Board.WithSchematics([new BoardSchematicEntry { SchematicName = "Front", SchematicImageFile = "top.png" }]),
        });

        Assert.Equal("Front", Assert.Single(session.Content.Board.ComponentHighlights).SchematicName);
        Assert.Equal("Front", Assert.Single(session.Content.Calibrations).SchematicName);

        session.Undo();
        Assert.Equal("Top", Assert.Single(session.Content.Board.ComponentHighlights).SchematicName);
        Assert.Equal("Top", Assert.Single(session.Content.Calibrations).SchematicName);
    }

    // ------------------------------------------------------------------ DraftSessions

    // A draft made before 4.0 still names a workbook; it is the same draft, so the same session.
    [Fact]
    public void Every_editor_gets_the_one_session_of_a_draft_whatever_its_identity_names()
    {
        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "CPU")));
        var sessions = new DraftSessions(() => this.thisNow);

        DraftSession? first = sessions.Get(this.DraftsRoot, DraftSessionTests.BoardKey);
        DraftSession? second = sessions.Get(this.DraftsRoot, "Commodore/C64/250407/Data C64 250407.xlsx");

        Assert.NotNull(first);
        Assert.Same(first, second);
        Assert.Same(first, sessions.Find(this.DraftsRoot, DraftSessionTests.BoardKey));
        Assert.Null(sessions.Get(this.DraftsRoot, "Commodore/C128/310378/board.json"));
    }

    [Fact]
    public void The_sessions_due_are_the_ones_a_second_past_their_last_change()
    {
        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "CPU")));
        var sessions = new DraftSessions(() => this.thisNow);
        DraftSession session = sessions.Get(this.DraftsRoot, DraftSessionTests.BoardKey)!;

        Assert.Empty(sessions.Due(this.thisNow.AddHours(1)));

        session.Apply("one", DraftSessionTests.Name("U1", "one"));
        Assert.Empty(sessions.Due(this.thisNow));
        Assert.Same(session, Assert.Single(sessions.Due(this.thisNow + DraftSession.SaveDelay)));

        sessions.Close(this.DraftsRoot, DraftSessionTests.BoardKey);
        Assert.Empty(sessions.Open);
        Assert.Null(sessions.Find(this.DraftsRoot, DraftSessionTests.BoardKey));
    }

    // ------------------------------------------------------------------ Every editor on one draft (2026-10-25)

    // ###########################################################################################
    // BEFORE EVERY CHANGE, WHOEVER HOLDS ONE NOT HANDED OVER YET IS ASKED FOR IT (ChangeComing): an
    // edit, an undo, a redo, the file read again, a write of the file about to leave the UI thread
    // (CollectHeldChanges). Raised while the content is still as it was, saying what comes and from
    // whom - so the Drafts tab's table can push what it holds and skip its own.
    // ###########################################################################################
    [Fact]
    public void Before_every_change_the_editors_holding_one_are_asked_to_hand_it_over()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        var coming = new List<(DraftSessionChangeKind Kind, string Description, object? Origin, string NameThen)>();
        session.ChangeComing += (_, change) => coming.Add((change.Kind, change.Description, change.Origin, DraftSessionTests.NameOf(session.Content, "U1")));

        session.Apply("Name of U1", DraftSessionTests.Name("U1", "CPU 6510"), origin: "label editor");
        session.Undo();
        session.Redo();
        session.ReadFileIfChanged("Written elsewhere");
        session.CollectHeldChanges("Brought up to date");

        Assert.Equal(
            [
                (DraftSessionChangeKind.Edit, "Name of U1", (object?)"label editor", "CPU"),
                (DraftSessionChangeKind.Undo, "Name of U1", "label editor", "CPU 6510"),
                (DraftSessionChangeKind.Redo, "Name of U1", "label editor", "CPU"),
                (DraftSessionChangeKind.ReadFromFile, "Written elsewhere", null, "CPU 6510"),
                (DraftSessionChangeKind.ReadFromFile, "Brought up to date", null, "CPU 6510"),
            ],
            coming);

        // Asking for what is held changes nothing by itself.
        Assert.Equal("Name of U1", session.UndoDescription);
        Assert.False(session.CanRedo);
    }

    // What an editor hands over when asked lands UNDER the change that asked - the table's typing,
    // then the component editor's save on top - each its own step, in that order.
    [Fact]
    public void A_change_handed_over_when_another_comes_lands_underneath_it()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"), DraftSessionTests.Component("U2", "VIC"));
        bool tableHolds = true;

        session.ChangeComing += (_, change) =>
        {
            if (tableHolds && !Equals(change.Origin, "table"))
            {
                tableHolds = false;
                session.Apply("Typed in the table", DraftSessionTests.Name("U1", "typed"), origin: "table");
            }
        };

        Assert.True(session.Apply("Changes to [U2]", DraftSessionTests.Name("U2", "saved"), origin: "component editor"));

        Assert.Equal("typed", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal("saved", DraftSessionTests.NameOf(session.Content, "U2"));

        Assert.Equal("Changes to [U2]", session.UndoDescription);
        session.Undo();
        Assert.Equal("Typed in the table", session.UndoDescription);
        Assert.Equal("typed", DraftSessionTests.NameOf(session.Content, "U1"));
    }

    // ###########################################################################################
    // AN EDITOR'S EXPLICIT SAVE EITHER HAPPENS OR CHANGES NOTHING (TakeBack): its change, the latest
    // step, is taken back out when its save could not be made - not redoable, it never happened as
    // far as that editor's "could not be saved" goes. Only its own, and only while nothing came since.
    // ###########################################################################################
    [Fact]
    public void A_change_whose_save_could_not_be_made_is_taken_back_and_cannot_be_redone()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        var seen = new List<DraftSessionChangeKind>();
        session.Changed += (_, change) => seen.Add(change.Kind);
        object componentEditor = new();

        session.Apply("Changes to [U1]", DraftSessionTests.Name("U1", "mine"), origin: componentEditor);
        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "theirs")));
        Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, session.Save());

        Assert.False(session.TakeBack(new object()));
        Assert.True(session.TakeBack(componentEditor));

        Assert.Equal("CPU", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.False(session.CanUndo);
        Assert.False(session.CanRedo);
        Assert.Equal(DraftSessionChangeKind.TakenBack, seen[^1]);

        Assert.False(session.TakeBack(componentEditor));
    }

    // ###########################################################################################
    // A save with nothing changed writes nothing - unless asked to (evenUnchanged): after a KiCad
    // project was copied in, the board file is written again so the new files lying in the draft
    // folder are taken into its store and listed (DraftFiles).
    // ###########################################################################################
    [Fact]
    public void A_save_of_an_unchanged_draft_writes_the_file_again_only_when_asked_and_takes_in_new_files()
    {
        var board = new BoardData
        {
            Components = { DraftSessionTests.Component("U1", "CPU") },
            ComponentImages = { new ComponentImageEntry { BoardLabel = "U1", Name = "Pinout", File = "Images/u1.png" } },
        };

        this.WriteDraft(board);
        DraftSession session = this.Open();

        string arrived = Path.Combine(DraftFolderLayout.GetBoardFolder(this.DraftsRoot, DraftSessionTests.BoardKey), "Images", "u1.png");
        Directory.CreateDirectory(Path.GetDirectoryName(arrived)!);
        File.WriteAllText(arrived, "the picture");

        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());
        Assert.True(File.Exists(arrived));

        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save(evenUnchanged: true));
        Assert.False(File.Exists(arrived));
        Assert.Equal(DraftSaveState.Saved, session.SaveState);

        // The file the session wrote is its own: the next change saves without a word about it.
        session.Apply("one", DraftSessionTests.Name("U1", "one"));
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());
    }

    // ###########################################################################################
    // A WRITE OF THE FILE THAT IS MORE THAN THE CONTENT (WriteFileDirectly - "Bring my draft up to
    // date", a KiCad file removed): what the session holds is saved first, the write runs, and the
    // file is read back as what the session holds and has saved. Bringing up to date drops the
    // history, whose steps were measured against the old base.
    // ###########################################################################################
    [Fact]
    public void A_write_of_the_file_runs_after_what_the_session_holds_is_saved_and_the_file_is_read_back()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        var seen = new List<DraftSessionChange>();
        session.Changed += (_, change) => seen.Add(change);
        session.Apply("mine", DraftSessionTests.Name("U1", "mine"));

        string onDiskWhenWritten = string.Empty;

        bool wrote = session.WriteFileDirectly(
            () =>
            {
                onDiskWhenWritten = this.NameOnDisk("U1");
                this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "brought up to date")));
                return true;
            },
            succeeded => succeeded,
            "Brought up to date",
            clearHistory: true);

        Assert.True(wrote);
        Assert.Equal("mine", onDiskWhenWritten);
        Assert.Equal("brought up to date", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal(DraftSaveState.Saved, session.SaveState);
        Assert.False(session.CanUndo);

        DraftSessionChange change = seen[^1];
        Assert.Equal(DraftSessionChangeKind.ReadFromFile, change.Kind);
        Assert.Equal("Brought up to date", change.Description);

        // The file read back is the session's own: a change after it saves without a word about it.
        session.Apply("after", DraftSessionTests.Name("U1", "after"));
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());
        Assert.Equal("after", this.NameOnDisk("U1"));
    }

    // Removing a KiCad file keeps the history; a write that wrote nothing reads nothing back.
    [Fact]
    public void A_write_of_the_file_keeps_the_history_unless_told_and_reads_nothing_back_when_it_did_not_write()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        session.Apply("one", DraftSessionTests.Name("U1", "one"));

        Assert.False(session.WriteFileDirectly(
            () =>
            {
                this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "not read")));
                return false;
            },
            succeeded => succeeded,
            "KiCad file removed",
            clearHistory: false));

        Assert.Equal("one", DraftSessionTests.NameOf(session.Content, "U1"));

        Assert.True(session.WriteFileDirectly(
            () =>
            {
                this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "one, file removed")));
                return true;
            },
            succeeded => succeeded,
            "KiCad file removed",
            clearHistory: false));

        Assert.Equal("one, file removed", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal("one", session.UndoDescription);
    }

    // What the session holds and cannot save first stops the write: it does not run, and why is said.
    [Fact]
    public void A_write_of_the_file_does_not_run_when_what_the_session_holds_cannot_be_saved_first()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));
        session.Apply("mine", DraftSessionTests.Name("U1", "mine"));
        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "theirs")));
        bool ran = false;

        DraftNotSavedException refused = Assert.Throws<DraftNotSavedException>(() => session.WriteFileDirectly(
            () => ran = true,
            succeeded => succeeded,
            "Brought up to date",
            clearHistory: true));

        Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, refused.Failure);
        Assert.False(ran);
        Assert.Equal("mine", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal("theirs", this.NameOnDisk("U1"));
    }

    // ###########################################################################################
    // An edit that works out its schematic renames on the draft as it is (ApplyWithFollow - a rename
    // decided by SchematicRename.Plan) hands them to what follows: here a rename that changed the
    // image file too, which the rule worked out from the two boards would read as a deletion, so the
    // highlight would go - the edit's own rename moves it instead.
    // ###########################################################################################
    [Fact]
    public void An_edit_that_works_out_its_own_renames_hands_them_to_what_follows_it()
    {
        var board = new BoardData
        {
            Components = { DraftSessionTests.Component("U1", "CPU") },
            Schematics = { new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png" } },
            ComponentHighlights = { new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "1", Y = "2", Width = "3", Height = "4" } },
        };

        this.WriteDraft(board);
        DraftSession session = this.Open();

        Assert.True(session.ApplyWithFollow("Rename Top", content =>
        (
            content with { Board = content.Board.WithSchematics([new BoardSchematicEntry { SchematicName = "Front", SchematicImageFile = "front.png" }]) },
            new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase) { ["Top"] = "Front" }
        )));

        Assert.Equal("Front", Assert.Single(session.Content.Board.ComponentHighlights).SchematicName);
    }

    // ------------------------------------------------------------------ Code review of 2026-10-25

    // ###########################################################################################
    // *** A CHANGE IS TAKEN BACK EVEN WHEN ANOTHER CAME AFTER IT (code review, 2026-10-25). *** The
    // component editor is a window of its own, so the table can push a change while the editor's
    // save runs; the editor then said "could not be saved" while its change stayed, to be saved by
    // the table's next save. Now only the editor's change goes: the table's later change stays, and
    // undoing the table's step cannot bring the editor's change back.
    // ###########################################################################################
    [Fact]
    public void A_change_is_taken_back_from_under_a_later_change_which_stays()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"), DraftSessionTests.Component("U2", "VIC"));
        object componentEditor = new();

        session.Apply("Changes to [U1]", DraftSessionTests.Name("U1", "mine"), origin: componentEditor);
        session.Apply("Typed in the table", DraftSessionTests.Name("U2", "typed"), origin: "table");

        Assert.True(session.TakeBack(componentEditor));

        Assert.Equal("CPU", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal("typed", DraftSessionTests.NameOf(session.Content, "U2"));
        Assert.Equal("Typed in the table", session.UndoDescription);

        Assert.True(session.Undo());
        Assert.Equal("CPU", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal("VIC", DraftSessionTests.NameOf(session.Content, "U2"));
        Assert.False(session.CanUndo);

        Assert.True(session.Redo());
        Assert.Equal("CPU", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal("typed", DraftSessionTests.NameOf(session.Content, "U2"));

        Assert.False(session.TakeBack(componentEditor));
    }

    // ###########################################################################################
    // *** A CHANGE HANDED OVER AS THE TAKE-BACK COMES (code review, 2026-10-26). *** TakeBack says
    // it is coming first, and an editor holding a change hands it over then - the table's open cell,
    // the pane's typed box - so the step to take back is no longer the latest. TakeBack gave up and
    // returned false: the editor said "could not be saved" while its change stayed, saved by the
    // next save. It is taken back from under the handed-over change, which stays.
    // ###########################################################################################
    [Fact]
    public void A_change_is_taken_back_even_when_another_is_handed_over_as_the_take_back_comes()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"), DraftSessionTests.Component("U2", "VIC"));
        object componentEditor = new();
        bool handedOver = false;

        session.Apply("Changes to [U1]", DraftSessionTests.Name("U1", "mine"), origin: componentEditor);

        session.ChangeComing += (_, change) =>
        {
            if (change.Kind == DraftSessionChangeKind.TakenBack && !handedOver)
            {
                handedOver = true;
                session.Apply("Typed in the table", DraftSessionTests.Name("U2", "typed"), origin: "table");
            }
        };

        Assert.True(session.TakeBack(componentEditor));

        Assert.True(handedOver);
        Assert.Equal("CPU", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal("typed", DraftSessionTests.NameOf(session.Content, "U2"));
        Assert.Equal("Typed in the table", session.UndoDescription);
    }

    // A row the later change changed too keeps the later change - it is what the user did last.
    [Fact]
    public void A_row_changed_again_after_the_change_taken_back_keeps_its_later_value()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"), DraftSessionTests.Component("U2", "VIC"));
        object componentEditor = new();

        session.Apply("Changes to [U1] and [U2]", content => DraftSessionTests.Name("U2", "mine too")(DraftSessionTests.Name("U1", "mine")(content)!), origin: componentEditor);
        session.Apply("Typed in the table", DraftSessionTests.Name("U1", "typed"), origin: "table");

        Assert.True(session.TakeBack(componentEditor));

        Assert.Equal("typed", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal("VIC", DraftSessionTests.NameOf(session.Content, "U2"));
    }

    // ###########################################################################################
    // AN EDIT WORKED OUT OFF THE UI THREAD (code review, 2026-10-25): Prepare runs the edit and what
    // follows it on the content it is handed, and ApplyPrepared applies it as one step - only while
    // the draft is still that content. A draft that moved meanwhile applies nothing, and the caller
    // works the edit out again on the draft as it is.
    // ###########################################################################################
    [Fact]
    public void A_prepared_edit_is_applied_only_on_the_content_it_was_worked_out_on()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"), DraftSessionTests.Component("U2", "VIC"));

        DraftPreparedEdit prepared = DraftSession.Prepare(session.Content, content => (DraftSessionTests.Name("U1", "prepared")(content), null))!;
        Assert.Null(DraftSession.Prepare(session.Content, content => (content, null)));

        session.Apply("Typed in the table", DraftSessionTests.Name("U2", "typed"), origin: "table");

        Assert.False(session.ApplyPrepared("Changes to [U1]", prepared));
        Assert.Equal("CPU", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal("Typed in the table", session.UndoDescription);

        DraftPreparedEdit again = DraftSession.Prepare(session.Content, content => (DraftSessionTests.Name("U1", "prepared")(content), null))!;

        Assert.True(session.ApplyPrepared("Changes to [U1]", again, origin: "component editor"));
        Assert.Equal("prepared", DraftSessionTests.NameOf(session.Content, "U1"));
        Assert.Equal("typed", DraftSessionTests.NameOf(session.Content, "U2"));
        Assert.Equal("Changes to [U1]", session.UndoDescription);
    }

    // What follows an edit is followed in the prepared one too - a renamed schematic's highlights.
    [Fact]
    public void A_prepared_edit_carries_what_follows_it()
    {
        var board = new BoardData
        {
            Components = { DraftSessionTests.Component("U1", "CPU") },
            Schematics = { new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png" } },
            ComponentHighlights = { new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "1", Y = "2", Width = "3", Height = "4" } },
        };

        this.WriteDraft(board);
        DraftSession session = this.Open();

        DraftPreparedEdit prepared = DraftSession.Prepare(session.Content, content =>
            (content with { Board = content.Board.WithSchematics([new BoardSchematicEntry { SchematicName = "Front", SchematicImageFile = "top.png" }]) }, null))!;

        Assert.Equal("Front", Assert.Single(prepared.After.Board.ComponentHighlights).SchematicName);
        Assert.True(session.ApplyPrepared("Rename Top", prepared));
        Assert.Equal("Front", Assert.Single(session.Content.Board.ComponentHighlights).SchematicName);
    }

    // Whether the file changed outside is a fingerprint, asked anywhere - what an editor asks off the
    // UI thread before reading the file again there.
    [Fact]
    public void FileChangedOutside_says_whether_the_file_is_still_what_the_session_read_or_wrote()
    {
        DraftSession session = this.OpenWith(DraftSessionTests.Component("U1", "CPU"));

        Assert.False(session.FileChangedOutside());

        this.WriteDraft(DraftSessionTests.Board(DraftSessionTests.Component("U1", "theirs")));
        Assert.True(session.FileChangedOutside());

        Assert.True(session.ReadFileIfChanged("Written elsewhere"));
        Assert.False(session.FileChangedOutside());
    }
}
