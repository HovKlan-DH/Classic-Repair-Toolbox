using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// One sitting in the Drafts tab's table editor: open a draft as a table, save it back
// (owner request, 2026-09-24).
//
// *** THE TEST THAT MATTERS MOST is the one that edits the workbook BEHIND the table's back and
// then saves. *** The table holds every sheet and a save replaces all of them, so a save built on
// a stale read would silently undo whatever Excel (or the label editor on another tab) wrote in
// the meantime. It must refuse instead, and leave the other edit on disk untouched.
//
// Shares the "BoardData" collection: these write and re-read workbooks, and BoardDataReader keeps
// a process-wide cache.
// ###########################################################################################
[Collection("BoardData")]
public sealed class DraftTableSessionTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private const string BoardKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string WorkbookPath => DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftTableSessionTests.BoardKey);

    public void Dispose() => this.thisWorkspace.Dispose();

    private static ComponentEntry Component(string label, string friendlyName = "") =>
        new() { BoardLabel = label, FriendlyName = friendlyName, TechnicalNameOrValue = "x" };

    private void WriteDraft(BoardData board)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(this.WorkbookPath)!);
        BoardWorkbookWriter.Write(this.WorkbookPath, board);
    }

    private BoardData ReadDraft() =>
        DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftTableSessionTests.BoardKey)!;

    private static BoardTableCell FriendlyNameOf(DraftTableSession session, string label)
    {
        BoardTableSheet sheet = session.Document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        int labelColumn = sheet.Columns.ToList().IndexOf(BoardWorkbookSchema.ColBoardLabel);
        int nameColumn = sheet.Columns.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);

        return sheet.Rows.Single(row => !row.IsDeleted && row.Cells[labelColumn].Text == label).Cells[nameColumn];
    }

    // ###########################################################################################
    // *** THE CHECKS LOOK FOR A ROW'S FILES WHERE A SUBMIT LOOKS (2026-10-02). *** Given the
    // downloaded data, the table finds a file in the draft's own folder or in the data, and marks
    // one in neither as an error; without it, files are not looked for at all.
    // ###########################################################################################
    [Fact]
    public void Given_the_downloaded_data_the_table_marks_a_file_that_is_nowhere()
    {
        this.WriteDraft(new BoardData
        {
            Schematics =
            [
                new BoardSchematicEntry { SchematicName = "Mine", SchematicImageFile = "Commodore/C64/250407/mine.png" },
                new BoardSchematicEntry { SchematicName = "Published", SchematicImageFile = "Commodore/C64/250407/published.png" },
                new BoardSchematicEntry { SchematicName = "Gone", SchematicImageFile = "Commodore/C64/250407/gone.png" }
            ]
        });

        this.thisWorkspace.WriteFile("Drafts/Commodore/C64/250407/mine.png", "x");
        this.thisWorkspace.WriteFile("Data/Commodore/C64/250407/published.png", "x");
        string dataRoot = Path.Combine(this.thisWorkspace.Root, "Data");

        DraftTableSession checkedSession = DraftTableSession.Open(this.DraftsRoot, DraftTableSessionTests.BoardKey, null, dataRoot)!;
        BoardTableSheet schematics = checkedSession.Document.FindSheet(BoardWorkbookSchema.SheetBoardSchematics)!;

        Assert.Equal([false, false, true], schematics.Rows.Select(row => row.HasErrors));
        Assert.Equal("file.missing", checkedSession.Document.Problems.Single(problem => problem.Level == BoardProblemLevel.Error).Code);

        DraftTableSession unchecked_ = DraftTableSession.Open(this.DraftsRoot, DraftTableSessionTests.BoardKey, null)!;
        Assert.Equal(0, unchecked_.Document.ErrorCount);
    }

    [Fact]
    public void Without_a_draft_nothing_opens()
    {
        Assert.Null(DraftTableSession.Open(this.DraftsRoot, DraftTableSessionTests.BoardKey, published: null));
    }

    [Fact]
    public void The_table_opens_on_the_drafts_own_rows_coloured_against_the_published_board()
    {
        this.WriteDraft(new BoardData { Components = [Component("U1", "CPU 6510")] });

        DraftTableSession session = DraftTableSession.Open(
            this.DraftsRoot,
            DraftTableSessionTests.BoardKey,
            published: new BoardData { Components = [Component("U1", "CPU")] })!;

        Assert.Equal(BoardTableCellState.Modified, FriendlyNameOf(session, "U1").State);
        Assert.False(session.Document.HasUnsavedChanges);
        Assert.False(session.HasChangedOnDisk());
    }

    // The Drafts tab names the source the published board was downloaded from (owner request,
    // 2026-10-05) - the label has to reach the changed cell's tooltip through the session.
    [Fact]
    public void A_changed_cell_names_the_source_the_session_was_opened_with()
    {
        this.WriteDraft(new BoardData { Components = [Component("U1", "CPU 6510")] });

        DraftTableSession session = DraftTableSession.Open(
            this.DraftsRoot,
            DraftTableSessionTests.BoardKey,
            published: new BoardData { Components = [Component("U1", "CPU")] },
            baselineLabel: BoardTableDocument.SourceBaselineLabel(betaSource: true))!;

        Assert.Equal("BETA source value: CPU", FriendlyNameOf(session, "U1").ToolTip);
    }

    [Fact]
    public void A_saved_edit_is_in_the_workbook_and_the_table_is_no_longer_unsaved()
    {
        this.WriteDraft(new BoardData { Components = [Component("U1", "CPU")] });
        DraftTableSession session = DraftTableSession.Open(this.DraftsRoot, DraftTableSessionTests.BoardKey, null)!;

        FriendlyNameOf(session, "U1").Text = "CPU 6510";

        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

        Assert.Equal("CPU 6510", this.ReadDraft().Components.Single().FriendlyName);
        Assert.False(session.Document.HasUnsavedChanges);
    }

    // ###########################################################################################
    // *** A SAVE IN THREE STEPS, THE SLOW ONE ON ANOTHER THREAD (2026-09-30). *** "Save changes"
    // takes the rows on the UI thread (PrepareSave), writes on the pool under "please wait", and
    // finishes on the UI thread (CompleteSave). The rows written are the ones there when the save
    // was PREPARED - the write reads nothing of the table - and it is the ordinary save otherwise:
    // on disk, and the table no longer unsaved.
    // ###########################################################################################
    [Fact]
    public async Task A_prepared_save_writes_the_rows_it_was_prepared_with_from_another_thread()
    {
        this.WriteDraft(new BoardData { Components = [Component("U1", "CPU"), Component("U2", "SID")] });
        DraftTableSession session = DraftTableSession.Open(this.DraftsRoot, DraftTableSessionTests.BoardKey, null)!;

        FriendlyNameOf(session, "U1").Text = "CPU 6510";

        Func<DraftWorkbookEditOutcome> write = session.PrepareSave();

        // After it was prepared: not in this save.
        FriendlyNameOf(session, "U2").Text = "SID 6581";

        DraftWorkbookEditOutcome outcome = await Task.Run(write);
        session.CompleteSave(outcome);

        Assert.Equal(DraftWorkbookEditOutcome.Saved, outcome);

        BoardData saved = this.ReadDraft();
        Assert.Equal("CPU 6510", saved.Components.Single(component => component.BoardLabel == "U1").FriendlyName);
        Assert.Equal("SID", saved.Components.Single(component => component.BoardLabel == "U2").FriendlyName);

        // Its fingerprint moved with the file, so the table's own write is not a change from outside.
        Assert.False(session.HasChangedOnDisk());
    }

    // ###########################################################################################
    // *** THE RE-READ'S FINGERPRINT, NOT A SECOND HASH (code review, 2026-10-01). *** The table
    // re-reads the saved draft on the pool, and that read hashes the file; CompleteSave then hashed
    // it again on the UI thread. Handed the re-read's fingerprint it takes it as is - and the
    // session is still in step with the file.
    // ###########################################################################################
    [Fact]
    public async Task A_completed_save_takes_the_fingerprint_it_is_handed_rather_than_hashing_again()
    {
        this.WriteDraft(new BoardData { Components = [Component("U1", "CPU")] });
        DraftTableSession session = DraftTableSession.Open(this.DraftsRoot, DraftTableSessionTests.BoardKey, null)!;

        FriendlyNameOf(session, "U1").Text = "CPU 6510";
        Func<DraftWorkbookEditOutcome> write = session.PrepareSave();

        DraftWorkbookEditOutcome outcome = await Task.Run(write);
        DraftTableSession reread = DraftTableSession.Open(this.DraftsRoot, DraftTableSessionTests.BoardKey, null)!;

        session.CompleteSave(outcome, reread.Fingerprint);

        Assert.Same(reread.Fingerprint, session.Fingerprint);
        Assert.False(session.HasChangedOnDisk());
        Assert.False(session.Document.HasUnsavedChanges);
    }

    // The three steps refuse a stale read exactly as Save does - the check is in the write.
    [Fact]
    public void A_prepared_save_is_still_refused_when_the_workbook_changed_since_the_table_was_read()
    {
        this.WriteDraft(new BoardData { Components = [Component("U1", "CPU")] });
        DraftTableSession session = DraftTableSession.Open(this.DraftsRoot, DraftTableSessionTests.BoardKey, null)!;

        FriendlyNameOf(session, "U1").Text = "CPU 6510";
        Func<DraftWorkbookEditOutcome> write = session.PrepareSave();

        this.WriteDraft(new BoardData { Components = [Component("U1", "Changed in Excel")] });

        DraftWorkbookEditOutcome outcome = write();
        session.CompleteSave(outcome);

        Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, outcome);
        Assert.Equal("Changed in Excel", this.ReadDraft().Components.Single().FriendlyName);
        Assert.True(session.Document.HasUnsavedChanges);
    }

    [Fact]
    public void A_save_is_REFUSED_when_the_workbook_changed_since_the_table_was_read()
    {
        this.WriteDraft(new BoardData
        {
            Components = [Component("U1", "CPU")],
            Credits = [new CreditEntry { Category = "Data", NameOrHandle = "Dennis" }],
        });

        DraftTableSession session = DraftTableSession.Open(this.DraftsRoot, DraftTableSessionTests.BoardKey, null)!;
        FriendlyNameOf(session, "U1").Text = "Edited in the table";

        // Meanwhile, in Excel: a different sheet is edited and saved.
        this.WriteDraft(new BoardData
        {
            Components = [Component("U1", "CPU")],
            Credits = [new CreditEntry { Category = "Data", NameOrHandle = "Dennis" }, new CreditEntry { Category = "Photos", NameOrHandle = "Excel edit" }],
        });

        Assert.True(session.HasChangedOnDisk());
        Assert.Equal(DraftWorkbookEditOutcome.ChangedOnDisk, session.Save());

        // The Excel edit survives, and the table's edit was not written over it.
        BoardData onDisk = this.ReadDraft();
        Assert.Equal(2, onDisk.Credits.Count);
        Assert.Equal("CPU", onDisk.Components.Single().FriendlyName);

        // Still unsaved - nothing of the table's reached the file.
        Assert.True(session.Document.HasUnsavedChanges);
    }

    [Fact]
    public void A_second_save_in_the_same_sitting_goes_through()
    {
        // Without re-fingerprinting after a save, the table's OWN write would read as a change
        // made behind its back, and every later save would be refused.
        this.WriteDraft(new BoardData { Components = [Component("U1", "CPU")] });
        DraftTableSession session = DraftTableSession.Open(this.DraftsRoot, DraftTableSessionTests.BoardKey, null)!;

        FriendlyNameOf(session, "U1").Text = "First";
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

        FriendlyNameOf(session, "U1").Text = "Second";
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

        Assert.Equal("Second", this.ReadDraft().Components.Single().FriendlyName);
    }

    [Fact]
    public void Saving_keeps_the_component_highlights_in_the_sidecar()
    {
        // Highlights are in the JSON beside the workbook, not in any sheet. A table save that
        // passed an empty list through would erase every rectangle the label editor drew.
        this.WriteDraft(new BoardData { Components = [Component("U1", "CPU")] });
        BoardSidecarWriter.Write(
            this.WorkbookPath,
            [new ComponentHighlightEntry { SchematicName = "Main", BoardLabel = "U1", X = "10", Y = "20", Width = "30", Height = "40" }],
            []);

        DraftTableSession session = DraftTableSession.Open(this.DraftsRoot, DraftTableSessionTests.BoardKey, null)!;
        FriendlyNameOf(session, "U1").Text = "CPU 6510";

        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

        ComponentHighlightEntry highlight = Assert.Single(this.ReadDraft().ComponentHighlights);
        Assert.Equal("U1", highlight.BoardLabel);
        Assert.Equal("30", highlight.Width);
    }

    [Fact]
    public void Blank_rows_and_deleted_ghosts_are_not_written()
    {
        this.WriteDraft(new BoardData { Components = [Component("U1"), Component("U2")] });

        DraftTableSession session = DraftTableSession.Open(
            this.DraftsRoot,
            DraftTableSessionTests.BoardKey,
            published: new BoardData { Components = [Component("U1"), Component("U2"), Component("U3")] })!;

        BoardTableSheet sheet = session.Document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        sheet.InsertRow(null);

        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

        Assert.Equal(["U1", "U2"], this.ReadDraft().Components.Select(component => component.BoardLabel));
    }

    // ------------------------------------------------------------------ Watching the file (2026-09-24)

    private DraftTableSession OpenSession(BoardData board)
    {
        this.WriteDraft(board);
        return DraftTableSession.Open(this.DraftsRoot, DraftTableSessionTests.BoardKey, published: null)!;
    }

    [Fact]
    public void A_file_nobody_touched_is_not_changed_and_not_open_elsewhere()
    {
        DraftTableSession session = this.OpenSession(new BoardData { Components = [Component("U1", "CPU")] });

        DraftFileStatus status = session.CheckFile();

        Assert.False(status.ChangedSinceRead);
        Assert.False(status.IsMissing);
        Assert.False(status.IsOpenElsewhere);
    }

    [Fact]
    public void A_file_saved_elsewhere_with_different_content_is_seen_as_changed()
    {
        // The Excel case: the file is written by something other than this table.
        DraftTableSession session = this.OpenSession(new BoardData { Components = [Component("U1", "CPU")] });
        Assert.False(session.CheckFile().ChangedSinceRead);

        this.WriteDraft(new BoardData { Components = [Component("U1", "Edited in Excel")] });

        Assert.True(session.CheckFile().ChangedSinceRead);
        Assert.True(session.CheckFile().ChangedSinceRead);
    }

    [Fact]
    public void A_file_rewritten_with_the_same_content_is_not_a_change()
    {
        // Excel saving without an edit rewrites the file; the table has nothing to catch up on.
        var board = new BoardData { Components = [Component("U1", "CPU")] };
        DraftTableSession session = this.OpenSession(board);
        byte[] bytes = File.ReadAllBytes(this.WorkbookPath);

        File.WriteAllBytes(this.WorkbookPath, bytes);
        File.SetLastWriteTimeUtc(this.WorkbookPath, DateTime.UtcNow.AddMinutes(1));

        Assert.False(session.CheckFile().ChangedSinceRead);
    }

    [Fact]
    public void The_tables_own_save_is_never_mistaken_for_an_outside_change()
    {
        DraftTableSession session = this.OpenSession(new BoardData { Components = [Component("U1", "CPU")] });
        Assert.False(session.CheckFile().ChangedSinceRead);

        FriendlyNameOf(session, "U1").Text = "Saved from the table";
        Assert.Equal(DraftWorkbookEditOutcome.Saved, session.Save());

        Assert.False(session.CheckFile().ChangedSinceRead);
    }

    [Fact]
    public void A_draft_file_that_is_gone_is_missing_not_changed()
    {
        DraftTableSession session = this.OpenSession(new BoardData { Components = [Component("U1", "CPU")] });

        File.Delete(this.WorkbookPath);
        DraftFileStatus status = session.CheckFile();

        Assert.True(status.IsMissing);
        Assert.False(status.ChangedSinceRead);
    }

    [Theory]
    [InlineData("~$Data C64 250407.xlsx")]
    [InlineData("~$ta C64 250407.xlsx")]
    [InlineData(".~lock.Data C64 250407.xlsx#")]
    public void The_lock_file_a_spreadsheet_program_leaves_beside_the_workbook_means_it_is_open_there(string lockFile)
    {
        // Excel's hidden owner file (both forms Office uses) and LibreOffice's lock file.
        DraftTableSession session = this.OpenSession(new BoardData { Components = [Component("U1", "CPU")] });
        string lockPath = Path.Combine(Path.GetDirectoryName(this.WorkbookPath)!, lockFile);

        File.WriteAllText(lockPath, "owner");
        Assert.True(session.CheckFile().IsOpenElsewhere);

        File.Delete(lockPath);
        Assert.False(session.CheckFile().IsOpenElsewhere);
    }

    [Fact]
    public void Another_workbooks_lock_file_does_not_count()
    {
        DraftTableSession session = this.OpenSession(new BoardData { Components = [Component("U1", "CPU")] });
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(this.WorkbookPath)!, "~$Some other board.xlsx"), "owner");

        Assert.False(session.CheckFile().IsOpenElsewhere);
    }

    [Fact]
    public void A_workbook_held_open_by_another_program_is_reported_held_only_while_it_is()
    {
        // Excel keeps its workbook open, so a save cannot replace it. File sharing is only
        // enforced on Windows - elsewhere nothing is ever found held, and a save goes ahead.
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File sharing is advisory outside Windows.");

        DraftTableSession session = this.OpenSession(new BoardData { Components = [Component("U1", "CPU")] });
        File.WriteAllText(Path.Combine(Path.GetDirectoryName(this.WorkbookPath)!, "~$" + Path.GetFileName(this.WorkbookPath)), "owner");

        using (new FileStream(this.WorkbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
        {
            DraftFileStatus held = session.CheckFile();
            Assert.True(held.IsOpenElsewhere);
            Assert.True(held.IsHeldOpenElsewhere);
        }

        // Excel gone, its lock file left behind (a crash): open by the lock file, but not held.
        DraftFileStatus stale = session.CheckFile();
        Assert.True(stale.IsOpenElsewhere);
        Assert.False(stale.IsHeldOpenElsewhere);
    }

    [Fact]
    public void Without_a_lock_file_the_workbook_is_never_probed_as_held()
    {
        // The exclusive probe runs only while a lock file says a program already has the file -
        // never while Excel might be opening it.
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File sharing is advisory outside Windows.");

        DraftTableSession session = this.OpenSession(new BoardData { Components = [Component("U1", "CPU")] });

        using (new FileStream(this.WorkbookPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read))
        {
            Assert.False(session.CheckFile().IsHeldOpenElsewhere);
        }
    }
}
