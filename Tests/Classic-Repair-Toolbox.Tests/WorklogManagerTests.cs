using System.Text.Json;
using Avalonia;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for WorklogManager - the local, never-synced "Workbooks" folder that backs
// the worklog bar. Same shape as UserSettingsTests: a static singleton whose LoadFrom() seam lets
// a test point it at a temporary folder instead of the user's real AppData folder. NOTHING here
// calls Load().
//
// The storage model has no central index: each workbook is its own subfolder (named after its id)
// holding its own index.json, and every query scans those subfolders fresh. That is deliberate -
// deleting a workbook is just deleting its folder, with no separate bookkeeping file that could go
// stale. One consequence, pinned down below: with no persisted id counter, the next id is simply
// the highest numbered subfolder on disk plus one, so deleting the highest-id workbook lets its
// number be reused by the next one created.
//
// The class is global mutable state, so this whole file is one xUnit collection: the tests run
// sequentially and each one re-loads a fresh workbook folder first.
[Collection("Worklog")]
public sealed class WorklogManagerTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose()
    {
        // Detach from the temp folder so nothing written later can reach the user's real one.
        this.LoadWorklog();
        this.thisWorkspace.Dispose();
    }

    /// <summary>Points WorklogManager at a fresh, uniquely-named workbook folder under the workspace.</summary>
    private string LoadWorklog()
    {
        string root = this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N"));
        WorklogManager.LoadFrom(root);
        return root;
    }

    /// <summary>
    /// Creates a workbook and asserts it succeeded. CreateWorkbook returns null when the root is
    /// unusable or the write fails, which no test here sets up - so a null is a real failure worth
    /// failing loudly on, rather than something to silence with "!" at every call site.
    /// </summary>
    private static WorkbookRecord CreateWorkbook(string boardKey, string title, string note = "")
    {
        var record = WorklogManager.CreateWorkbook(boardKey, title, note);
        Assert.NotNull(record);
        return record!;
    }

    /// <summary>
    /// Adds one entry and asserts it succeeded. Same reasoning as CreateWorkbook above: AddEntry
    /// returns null only when the workbook folder or the write fails, neither of which the delete
    /// tests set up, so a null is a real failure rather than something to "!" past at every call.
    /// The area is a fixed unit rect - none of these tests care where the entry was marked.
    /// </summary>
    private static WorklogEntryRecord AddEntry(
        int workbookId,
        string schematicName,
        string title,
        string category = "Issue",
        string state = "Open")
    {
        var entry = WorklogManager.AddEntry(
            workbookId, schematicName, new Rect(0, 0, 1, 1), title, "", category, state, Array.Empty<string>());
        Assert.NotNull(entry);
        return entry!;
    }

    [Fact]
    public void An_empty_workbook_folder_starts_with_no_workbooks_and_id_one()
    {
        this.LoadWorklog();

        Assert.Equal(1, WorklogManager.PeekNextId());
        Assert.Null(WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469"));
    }

    [Fact]
    public void Creating_a_workbook_allocates_id_one_and_makes_it_the_active_workbook_for_its_board()
    {
        this.LoadWorklog();

        var record = CreateWorkbook("Commodore 64|250469", "Mr. Jensens C64", "Bought at auction");

        Assert.Equal(1, record.Id);
        Assert.Equal("Open", record.Status);
        Assert.Equal(0, record.EntryCount);
        Assert.Equal("Mr. Jensens C64", record.Title);
        Assert.Equal("Bought at auction", record.Note);

        var active = WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469");
        Assert.NotNull(active);
        Assert.Equal(1, active!.Id);
    }

    [Fact]
    public void The_next_id_advances_after_each_workbook_is_created()
    {
        this.LoadWorklog();

        CreateWorkbook("Commodore 64|250469", "First job", "");
        Assert.Equal(2, WorklogManager.PeekNextId());

        var second = CreateWorkbook("Commodore 64|250469", "Second job", "");
        Assert.Equal(2, second.Id);
        Assert.Equal(3, WorklogManager.PeekNextId());
    }

    [Fact]
    public void The_most_recently_created_open_workbook_is_the_active_one_for_a_board()
    {
        // A board can accumulate several workbooks over time - the highest id wins.
        this.LoadWorklog();

        CreateWorkbook("Commodore 64|250469", "Older job", "");
        var newer = CreateWorkbook("Commodore 64|250469", "Newer job", "");

        var active = WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469");

        Assert.Equal(newer.Id, active!.Id);
        Assert.Equal("Newer job", active.Title);
    }

    [Fact]
    public void Workbooks_are_kept_apart_per_board()
    {
        this.LoadWorklog();

        CreateWorkbook("Commodore 64|250469", "C64 job", "");
        CreateWorkbook("Amiga 500|A500", "Amiga job", "");

        Assert.Equal("C64 job", WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469")!.Title);
        Assert.Equal("Amiga job", WorklogManager.GetActiveWorkbookForBoard("Amiga 500|A500")!.Title);
    }

    [Fact]
    public void A_blank_board_key_yields_no_active_workbook_instead_of_throwing()
    {
        this.LoadWorklog();

        CreateWorkbook("Commodore 64|250469", "C64 job", "");

        Assert.Null(WorklogManager.GetActiveWorkbookForBoard(""));
        Assert.Null(WorklogManager.GetActiveWorkbookForBoard("   "));
    }

    [Fact]
    public void The_title_and_note_are_trimmed()
    {
        this.LoadWorklog();

        var record = CreateWorkbook("Commodore 64|250469", "  Mr. Jensens C64  ", "  Some note  ");

        Assert.Equal("Mr. Jensens C64", record.Title);
        Assert.Equal("Some note", record.Note);
    }

    [Fact]
    public void Creating_a_workbook_writes_its_own_index_file_inside_its_own_folder_named_after_its_id()
    {
        string root = this.LoadWorklog();

        CreateWorkbook("Commodore 64|250469", "C64 job", "");

        string ownIndexPath = Path.Combine(root, "1", "index.json");
        Assert.True(File.Exists(ownIndexPath));
        Assert.False(File.Exists(Path.Combine(root, "index.json")), "there must be no central index file");
    }

    [Fact]
    public void A_workbook_survives_a_reload_by_being_read_back_from_its_own_folder()
    {
        string root = this.LoadWorklog();

        CreateWorkbook("Commodore 64|250469", "C64 job", "Some note");

        WorklogManager.LoadFrom(root);

        var active = WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469");
        Assert.NotNull(active);
        Assert.Equal("C64 job", active!.Title);
        Assert.Equal("Some note", active.Note);
        Assert.Equal(2, WorklogManager.PeekNextId());
    }

    [Fact]
    public void Deleting_a_workbook_folder_removes_it_from_the_active_lookup()
    {
        // This is the whole point of the per-workbook-folder model: WorklogManager.DeleteWorkbook
        // below (the Workbooks tab's "Delete workbook" button) is just this - removing the whole
        // folder - with no separate bookkeeping entry anywhere else that could be left dangling.
        string root = this.LoadWorklog();

        CreateWorkbook("Commodore 64|250469", "C64 job", "");
        Assert.NotNull(WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469"));

        Directory.Delete(Path.Combine(root, "1"), recursive: true);

        Assert.Null(WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469"));
    }

    // ---------------------------------------------------------------------------------------
    // DeleteWorkbook / UpdateWorkbook - the Workbooks tab's "Delete workbook" button and the
    // "Edit workbook" dialog (CreateWorkbookWindow.InitializeForEdit).
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void Deleting_a_workbook_removes_its_folder_and_it_stops_appearing_in_the_boards_list()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        bool deleted = WorklogManager.DeleteWorkbook(workbook.Id);

        Assert.True(deleted);
        Assert.False(Directory.Exists(Path.Combine(root, workbook.Id.ToString())));
        Assert.Empty(WorklogManager.GetWorkbooksForBoard("Commodore 64|250469"));
    }

    [Fact]
    public void Deleting_one_workbook_leaves_the_boards_other_workbooks_untouched()
    {
        // The whole point of "delete the next one in the list" working automatically: the board's
        // remaining workbooks must still resolve normally once the deleted one is gone.
        this.LoadWorklog();
        var first = CreateWorkbook("Commodore 64|250469", "First job", "");
        var second = CreateWorkbook("Commodore 64|250469", "Second job", "");

        Assert.True(WorklogManager.DeleteWorkbook(first.Id));

        var remaining = WorklogManager.GetWorkbooksForBoard("Commodore 64|250469");
        Assert.Single(remaining);
        Assert.Equal(second.Id, remaining[0].Id);
    }

    [Fact]
    public void Deleting_the_active_workbook_falls_back_to_the_newest_remaining_one()
    {
        // ResolveActiveWorkbook is what the Workbooks tab and the worklog bar both call after a
        // delete (via Main.RefreshWorklogBar) - this pins down that a saved
        // ActiveWorkbookIdByBoard entry naming the just-deleted workbook is handled the same way
        // an id left stale by any other means already is: fall back to the newest, rather than
        // resolving to nothing.
        this.LoadWorklog();
        var older = CreateWorkbook("Commodore 64|250469", "Older job", "");
        var newer = CreateWorkbook("Commodore 64|250469", "Newer job", "");

        Assert.True(WorklogManager.DeleteWorkbook(newer.Id));

        var workbooks = WorklogManager.GetWorkbooksForBoard("Commodore 64|250469");
        var active = WorklogManager.ResolveActiveWorkbook(workbooks, savedActiveId: newer.Id);

        Assert.NotNull(active);
        Assert.Equal(older.Id, active!.Id);
    }

    [Fact]
    public void Deleting_an_unknown_workbook_returns_false_instead_of_throwing()
    {
        this.LoadWorklog();

        Exception? thrown = Record.Exception(() =>
        {
            bool deleted = WorklogManager.DeleteWorkbook(999);
            Assert.False(deleted);
        });

        Assert.True(thrown is null);
    }

    // ---------------------------------------------------------------------------------------
    // DeleteEntry - the Workbooks tab's per-card "Delete worklog" button. A worklog is a row in
    // entries.json PLUS a "worklog_{id}" folder holding its photo/file bytes, so deleting one has
    // to remove both or the workbook keeps orphaned attachment bytes forever.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void Deleting_a_worklog_removes_only_that_entry_and_leaves_the_others_alone()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        var first = AddEntry(workbook.Id, "Sch", "Bad cap");
        var second = AddEntry(workbook.Id, "Sch", "Cracked trace");
        var third = AddEntry(workbook.Id, "Sch", "Dead PLA");

        Assert.True(WorklogManager.DeleteEntry(workbook.Id, second.Id));

        var remaining = WorklogManager.GetEntries(workbook.Id);
        Assert.Equal(2, remaining.Count);
        Assert.Equal(new[] { first.Id, third.Id }, remaining.Select(e => e.Id).ToArray());
        Assert.Equal(new[] { "Bad cap", "Dead PLA" }, remaining.Select(e => e.Title).ToArray());
    }

    // The surviving entries KEEP their own ids - nothing is renumbered down to close the gap.
    // Those ids are what the board pills, the entry cards and the exported PDF all show ("#4"),
    // and they name each entry's attachment folder on disk, so renumbering would silently relabel
    // worklogs the user has already referred to elsewhere. The gap IS the record of a deletion.
    [Fact]
    public void Deleting_a_worklog_does_not_renumber_the_entries_that_remain()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        AddEntry(workbook.Id, "Sch", "First");
        var second = AddEntry(workbook.Id, "Sch", "Second");
        var third = AddEntry(workbook.Id, "Sch", "Third");

        Assert.True(WorklogManager.DeleteEntry(workbook.Id, second.Id));

        Assert.Equal(new[] { 1, 3 }, WorklogManager.GetEntries(workbook.Id).Select(e => e.Id).ToArray());
        Assert.Equal(3, third.Id);
    }

    // The attachment BYTES go with the entry. Only the row was removed in an earlier shape of
    // this, which left a "worklog_{id}" folder full of photos that nothing on disk named - and
    // worse, one the next entry allocated that same id would then inherit.
    [Fact]
    public void Deleting_a_worklog_also_removes_its_attachments_folder()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        var entry = AddEntry(workbook.Id, "Sch", "Bad cap");

        string attachments = WorklogManager.GetEntryAttachmentsFolder(workbook.Id, entry.Id)!;
        File.WriteAllText(Path.Combine(attachments, "photo.png"), "not really a png");
        Assert.True(Directory.Exists(attachments));

        Assert.True(WorklogManager.DeleteEntry(workbook.Id, entry.Id));

        Assert.False(Directory.Exists(attachments));
    }

    // An entry with no attachments never had a folder (GetEntryAttachmentsFolder creates one only
    // when something is actually attached), and deleting it must not fail over the absence.
    [Fact]
    public void Deleting_a_worklog_with_no_attachments_folder_still_succeeds()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        var entry = AddEntry(workbook.Id, "Sch", "Bad cap");

        Assert.True(WorklogManager.DeleteEntry(workbook.Id, entry.Id));
        Assert.Empty(WorklogManager.GetEntries(workbook.Id));
    }

    // The workbook's own bookkeeping is recomputed, exactly as it is after an add or an edit:
    // deleting the last still-Open entry of an otherwise-finished workbook closes it.
    [Fact]
    public void Deleting_the_last_open_worklog_closes_the_workbook()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        AddEntry(workbook.Id, "Sch", "Done already", state: "Closed");
        var stillOpen = AddEntry(workbook.Id, "Sch", "Outstanding", state: "Open");

        Assert.Equal("Open", WorklogManager.GetWorkbooksForBoard("Commodore 64|250469")[0].Status);

        Assert.True(WorklogManager.DeleteEntry(workbook.Id, stillOpen.Id));

        var reloaded = WorklogManager.GetWorkbooksForBoard("Commodore 64|250469")[0];
        Assert.Equal("Closed", reloaded.Status);
        Assert.Equal(1, reloaded.EntryCount);
    }

    // The other side of the same rule: a workbook with NO entries is Open, per
    // RecomputeWorkbookStatus - so deleting the only (Closed) entry reopens it rather than
    // leaving it Closed on the strength of entries that are gone.
    [Fact]
    public void Deleting_the_only_worklog_leaves_the_workbook_open_with_no_entries()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        var only = AddEntry(workbook.Id, "Sch", "Done", state: "Closed");

        Assert.Equal("Closed", WorklogManager.GetWorkbooksForBoard("Commodore 64|250469")[0].Status);

        Assert.True(WorklogManager.DeleteEntry(workbook.Id, only.Id));

        var reloaded = WorklogManager.GetWorkbooksForBoard("Commodore 64|250469")[0];
        Assert.Equal("Open", reloaded.Status);
        Assert.Equal(0, reloaded.EntryCount);
        Assert.Empty(WorklogManager.GetEntries(workbook.Id));
    }

    [Fact]
    public void Deleting_an_unknown_worklog_returns_false_and_changes_nothing()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        AddEntry(workbook.Id, "Sch", "Bad cap");

        Assert.False(WorklogManager.DeleteEntry(workbook.Id, 999));
        Assert.Single(WorklogManager.GetEntries(workbook.Id));
    }

    [Fact]
    public void Deleting_a_worklog_from_an_unknown_workbook_returns_false_instead_of_throwing()
    {
        this.LoadWorklog();

        Exception? thrown = Record.Exception(() => Assert.False(WorklogManager.DeleteEntry(999, 1)));

        Assert.True(thrown is null);
    }

    [Fact]
    public void Updating_a_workbook_overwrites_its_title_and_note_and_leaves_everything_else_alone()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "Original title", "Original note");
        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        var updated = WorklogManager.UpdateWorkbook(workbook.Id, "  New title  ", "  New note  ");

        Assert.NotNull(updated);

        // The RETURNED record is the one that reached disk, so the caller never has to patch up its
        // own copy - it already carries the trimmed values.
        Assert.Equal("New title", updated!.Title);
        Assert.Equal("New note", updated.Note);

        var stored = WorklogManager.GetWorkbooksForBoard("Commodore 64|250469").Single();
        Assert.Equal("New title", stored.Title);
        Assert.Equal("New note", stored.Note);

        // Untouched: id, board key, status/entryCount (still derived from the real entry above),
        // and start date.
        Assert.Equal(workbook.Id, stored.Id);
        Assert.Equal(workbook.BoardKey, stored.BoardKey);
        Assert.Equal(workbook.StartDate, stored.StartDate);
        Assert.Equal(1, stored.EntryCount);
        Assert.Equal("Open", stored.Status);

        Assert.True(File.Exists(Path.Combine(root, workbook.Id.ToString(), "index.json")));
    }

    [Fact]
    public void Updating_a_workbook_trims_the_title_and_note()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        Assert.NotNull(WorklogManager.UpdateWorkbook(workbook.Id, "  Trimmed title  ", "  Trimmed note  "));

        var stored = WorklogManager.GetWorkbooksForBoard("Commodore 64|250469").Single();
        Assert.Equal("Trimmed title", stored.Title);
        Assert.Equal("Trimmed note", stored.Note);
    }

    [Fact]
    public void Updating_a_workbook_survives_a_reload_by_being_read_back_from_its_own_folder()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "Original title", "");

        WorklogManager.UpdateWorkbook(workbook.Id, "Edited title", "Edited note");
        WorklogManager.LoadFrom(root);

        var stored = WorklogManager.GetWorkbooksForBoard("Commodore 64|250469").Single();
        Assert.Equal("Edited title", stored.Title);
        Assert.Equal("Edited note", stored.Note);
    }

    [Fact]
    public void Updating_an_unknown_workbook_returns_null_instead_of_throwing()
    {
        this.LoadWorklog();

        Exception? thrown = Record.Exception(() =>
        {
            var updated = WorklogManager.UpdateWorkbook(999, "Title", "Note");
            Assert.Null(updated);
        });

        Assert.True(thrown is null);
    }

    // ---------------------------------------------------------------------------------------
    // ID COUNTERS - a deleted id is never handed out again.
    //
    // This replaces the opposite rule, which this file used to pin down: ids were "highest on disk
    // plus one", so deleting the highest workbook let the next one take its number back. That
    // silently made an already-exported PDF ("Workbook_2_...") describe a different repair, and let
    // the new record inherit the deleted one's folder on disk, attachments included.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void Deleting_the_highest_id_workbook_does_not_let_its_id_be_reused()
    {
        this.LoadWorklog();

        var first = CreateWorkbook("Commodore 64|250469", "First job", "");
        Assert.Equal(1, first.Id);

        Assert.True(WorklogManager.DeleteWorkbook(first.Id));

        // Nothing is on disk any more, but the counter remembers #1 was spent.
        Assert.Equal(2, WorklogManager.PeekNextId());
        Assert.Equal(2, CreateWorkbook("Commodore 64|250469", "Second job", "").Id);
    }

    // The example from the request: two workbooks, delete the SECOND, and the next is #3 - not #2.
    // A gap in the numbering is the correct record that #2 existed.
    [Fact]
    public void Deleting_a_workbook_leaves_a_gap_that_the_next_create_skips_over()
    {
        this.LoadWorklog();

        CreateWorkbook("Commodore 64|250469", "First job", "");
        var second = CreateWorkbook("Commodore 64|250469", "Second job", "");
        Assert.Equal(2, second.Id);

        Assert.True(WorklogManager.DeleteWorkbook(second.Id));

        var third = CreateWorkbook("Commodore 64|250469", "Third job", "");
        Assert.Equal(3, third.Id);

        // And the surviving #1 was not disturbed on the way.
        var ids = WorklogManager.GetWorkbooksForBoard("Commodore 64|250469").Select(w => w.Id).OrderBy(id => id);
        Assert.Equal(new[] { 1, 3 }, ids);
    }

    // The counter survives a restart: it is a file, not in-memory state. LoadFrom the SAME root
    // again is what a relaunch does, and the numbering has to carry on rather than restart.
    [Fact]
    public void The_workbook_counter_survives_reloading_the_same_root()
    {
        string root = this.LoadWorklog();

        CreateWorkbook("Commodore 64|250469", "First job", "");
        var second = CreateWorkbook("Commodore 64|250469", "Second job", "");
        Assert.True(WorklogManager.DeleteWorkbook(second.Id));

        WorklogManager.LoadFrom(root);

        Assert.Equal(3, WorklogManager.PeekNextId());
    }

    // NO MIGRATION, by design: nothing seeds a counter from, or rewrites, workbooks written before
    // the counter file existed. What must still hold is that such a workbook is never OVERWRITTEN -
    // a counter starting from zero would otherwise hand out #1 while a "1" folder is sitting there,
    // and CreateWorkbook's Directory.CreateDirectory succeeds silently on an existing folder, so
    // the new workbook's index.json would replace the old one in place and inherit its entries and
    // attachments. The allocator walks past ids already taken on disk instead.
    [Fact]
    public void A_workbook_folder_that_predates_the_counter_is_skipped_rather_than_overwritten()
    {
        string root = this.LoadWorklog();

        // A workbook as an older build left it: a numbered folder with an index.json, and no
        // counter file anywhere.
        string legacyFolder = Path.Combine(root, "1");
        Directory.CreateDirectory(legacyFolder);
        File.WriteAllText(
            Path.Combine(legacyFolder, "index.json"),
            JsonSerializer.Serialize(new WorkbookRecord
            {
                Id = 1,
                BoardKey = "Commodore 64|250469",
                Title = "Written by an older build",
                Status = "Open",
            }));

        var created = CreateWorkbook("Commodore 64|250469", "Brand new job", "");

        Assert.Equal(2, created.Id);

        // The old one is still there, untouched - not replaced by the new record.
        var legacy = WorklogManager.GetWorkbooksForBoard("Commodore 64|250469").Single(w => w.Id == 1);
        Assert.Equal("Written by an older build", legacy.Title);
    }

    // The same floor on the ENTRY side, and for the same reason: a workbook written before
    // LastEntryId existed deserializes it as 0, so without the walk-past its next entry would be
    // #1 on top of the entries it already has.
    [Fact]
    public void Entries_in_a_workbook_that_predates_the_counter_are_not_renumbered_from_one()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        AddEntry(workbook.Id, "Sch", "First");
        AddEntry(workbook.Id, "Sch", "Second");

        // Strip the counter back out of index.json, leaving the workbook exactly as an older build
        // would have written it: entries present, no lastEntryId.
        string indexPath = Path.Combine(root, workbook.Id.ToString(), "index.json");
        var record = JsonSerializer.Deserialize<WorkbookRecord>(File.ReadAllText(indexPath))!;
        record.LastEntryId = 0;
        File.WriteAllText(indexPath, JsonSerializer.Serialize(record));

        Assert.Equal(3, WorklogManager.PeekNextEntryId(workbook.Id));

        var third = AddEntry(workbook.Id, "Sch", "Third");
        Assert.Equal(3, third.Id);
    }

    [Fact]
    public void Deleting_a_worklog_does_not_let_its_id_be_reused()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        AddEntry(workbook.Id, "Sch", "First");
        var second = AddEntry(workbook.Id, "Sch", "Second");
        Assert.Equal(2, second.Id);

        Assert.True(WorklogManager.DeleteEntry(workbook.Id, second.Id));

        // #2 is spent even though nothing on disk carries it any more.
        Assert.Equal(3, WorklogManager.PeekNextEntryId(workbook.Id));
        Assert.Equal(3, AddEntry(workbook.Id, "Sch", "Third").Id);
    }

    // Deleting the ONLY entry still spends its number - the workbook is empty again but does not
    // restart at #1, or the "#1" pill on the board would refer to two different repairs over time.
    [Fact]
    public void Deleting_every_worklog_does_not_restart_entry_numbering()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var only = AddEntry(workbook.Id, "Sch", "Only");
        Assert.True(WorklogManager.DeleteEntry(workbook.Id, only.Id));

        Assert.Empty(WorklogManager.GetEntries(workbook.Id));
        Assert.Equal(2, AddEntry(workbook.Id, "Sch", "Next").Id);
    }

    // Each workbook numbers its own entries: the counter lives in that workbook's index.json, so a
    // second workbook starts at #1 regardless of how many entries the first one has spent.
    [Fact]
    public void Entry_numbering_is_per_workbook_and_not_shared_across_them()
    {
        this.LoadWorklog();
        var first = CreateWorkbook("Commodore 64|250469", "First job", "");
        var second = CreateWorkbook("Commodore 64|250469", "Second job", "");

        AddEntry(first.Id, "Sch", "A");
        AddEntry(first.Id, "Sch", "B");

        Assert.Equal(1, AddEntry(second.Id, "Sch", "C").Id);
    }

    // The entry counter is per-workbook state, so it must survive a reload the same way the
    // workbook counter does.
    [Fact]
    public void The_entry_counter_survives_reloading_the_same_root()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var entry = AddEntry(workbook.Id, "Sch", "First");
        Assert.True(WorklogManager.DeleteEntry(workbook.Id, entry.Id));

        WorklogManager.LoadFrom(root);

        Assert.Equal(2, WorklogManager.PeekNextEntryId(workbook.Id));
    }

    // Deleting a workbook takes its entry counter with it - the counter lives in the folder that is
    // removed. That is correct rather than a leak: entry ids only have to be unique within their
    // own workbook, and the workbook id itself is never reused, so a fresh workbook numbering its
    // entries from #1 cannot collide with anything the deleted one exported.
    [Fact]
    public void A_new_workbook_starts_its_entries_at_one_even_after_another_was_deleted()
    {
        this.LoadWorklog();
        var first = CreateWorkbook("Commodore 64|250469", "First job", "");

        AddEntry(first.Id, "Sch", "A");
        AddEntry(first.Id, "Sch", "B");
        Assert.True(WorklogManager.DeleteWorkbook(first.Id));

        var replacement = CreateWorkbook("Commodore 64|250469", "Second job", "");

        // A new WORKBOOK number, so its "#1" is unambiguous.
        Assert.Equal(2, replacement.Id);
        Assert.Equal(1, AddEntry(replacement.Id, "Sch", "C").Id);
    }

    // ---------------------------------------------------------- an unusable workbook root

    // Regression: CreateWorkbook was the one member with no empty-root guard. Path.Combine("", "1")
    // yields the RELATIVE path "1", so it used to create a workbook folder next to the executable
    // and hand back a record that looked fine in the bar - while GetWorkbookFolder, which does
    // guard, could never find it again, so every entry recorded afterwards was silently discarded.
    [Fact]
    public void Creating_a_workbook_fails_cleanly_when_the_workbook_root_is_unusable()
    {
        WorklogManager.LoadFrom(string.Empty);

        string cwdBefore = Directory.GetCurrentDirectory();

        var record = WorklogManager.CreateWorkbook("Commodore 64|250469", "C64 job", "");

        Assert.Null(record);
        Assert.False(Directory.Exists(Path.Combine(cwdBefore, "1")), "must not create a workbook folder relative to the working directory");
    }

    [Fact]
    public void The_next_id_skips_gaps_left_by_a_deleted_workbook_that_was_not_the_highest()
    {
        // Folders "1" and "3" exist (2 was deleted) - the next id must be 4, the highest plus one,
        // not 2: filling gaps would risk two different workbooks briefly sharing a folder name.
        string root = this.LoadWorklog();

        Directory.CreateDirectory(Path.Combine(root, "1"));
        Directory.CreateDirectory(Path.Combine(root, "3"));

        Assert.Equal(4, WorklogManager.PeekNextId());
    }

    [Fact]
    public void A_subfolder_with_no_index_file_is_skipped_instead_of_throwing()
    {
        // Debris, or a workbook folder created but never finished - either way it must not break
        // the board lookup for every other workbook.
        string root = this.LoadWorklog();

        Directory.CreateDirectory(Path.Combine(root, "1"));
        CreateWorkbook("Commodore 64|250469", "Real job", "");

        var active = WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469");

        Assert.NotNull(active);
        Assert.Equal("Real job", active!.Title);
    }

    [Fact]
    public void A_malformed_index_file_is_skipped_instead_of_throwing()
    {
        string root = this.LoadWorklog();
        Directory.CreateDirectory(Path.Combine(root, "1"));
        File.WriteAllText(Path.Combine(root, "1", "index.json"), "{ this is not json");

        Exception? thrown = Record.Exception(() =>
            WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469"));

        Assert.True(thrown is null);
        Assert.Null(WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469"));
    }

    // Regression: System.Text.Json assigns straight over a property's "= new()" initializer when
    // the JSON carries an explicit null, so these arrive null despite their initializers. Every
    // consumer dereferences them unguarded - the editor's CloneEntry did so six times, on the path
    // that opens a saved entry, before the window was ever shown.
    [Fact]
    public void Explicit_nulls_in_an_entrys_collections_are_read_back_as_empty_lists()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job");
        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        string entriesPath = Path.Combine(root, workbook.Id.ToString(), "entries.json");
        File.WriteAllText(entriesPath, """
        [
          {
            "id": 1,
            "schematicName": "Sch",
            "title": "Bad cap",
            "category": "Issue",
            "state": "Open",
            "componentLabels": null,
            "links": null,
            "comments": null,
            "workDoneItems": null,
            "photos": null,
            "files": null
          }
        ]
        """);

        var entry = WorklogManager.GetEntries(workbook.Id).Single();

        Assert.NotNull(entry.ComponentLabels);
        Assert.NotNull(entry.Links);
        Assert.NotNull(entry.Comments);
        Assert.NotNull(entry.WorkDoneItems);
        Assert.NotNull(entry.Photos);
        Assert.NotNull(entry.Files);
        Assert.Empty(entry.Comments);
    }

    [Fact]
    public void The_written_index_file_is_valid_indented_json()
    {
        string root = this.LoadWorklog();

        CreateWorkbook("Commodore 64|250469", "C64 job", "");

        string json = File.ReadAllText(Path.Combine(root, "1", "index.json"));
        Assert.Contains("\n", json);
        Exception? thrown = Record.Exception(() => JsonDocument.Parse(json));
        Assert.True(thrown is null, "the index file must stay parseable");
    }

    [Fact]
    public void A_missing_workbook_root_folder_is_created_on_load()
    {
        string root = this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N"));
        Assert.False(Directory.Exists(root));

        WorklogManager.LoadFrom(root);

        Assert.True(Directory.Exists(root));
    }

    [Fact]
    public void Adding_an_entry_writes_it_to_the_workbooks_own_entries_file()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var entry = WorklogManager.AddEntry(
            workbook.Id,
            "Schematic 1",
            new Rect(10, 20, 30, 40),
            "  Bad cap  ",
            "  Leaking electrolyte  ",
            "Issue",
            "Open",
            new[] { "C12", "R4" });

        Assert.NotNull(entry);
        Assert.Equal(1, entry!.Id);
        Assert.Equal("Schematic 1", entry.SchematicName);
        Assert.Equal(10, entry.AreaX);
        Assert.Equal(20, entry.AreaY);
        Assert.Equal(30, entry.AreaWidth);
        Assert.Equal(40, entry.AreaHeight);
        Assert.Equal("Bad cap", entry.Title);
        Assert.Equal("Leaking electrolyte", entry.Description);
        Assert.Equal(new[] { "C12", "R4" }, entry.ComponentLabels);

        Assert.True(File.Exists(Path.Combine(root, workbook.Id.ToString(), "entries.json")));

        var storedEntries = WorklogManager.GetEntries(workbook.Id);
        Assert.Single(storedEntries);
        Assert.Equal("Bad cap", storedEntries[0].Title);
    }

    [Fact]
    public void Entry_ids_increment_independently_per_workbook()
    {
        this.LoadWorklog();
        var first = CreateWorkbook("Commodore 64|250469", "First job", "");
        var second = CreateWorkbook("Amiga 500|A500", "Second job", "");

        WorklogManager.AddEntry(first.Id, "Sch", new Rect(0, 0, 1, 1), "A", "", "Note", "Open", Array.Empty<string>());
        var secondEntryInFirst = WorklogManager.AddEntry(first.Id, "Sch", new Rect(0, 0, 1, 1), "B", "", "Note", "Open", Array.Empty<string>());
        var firstEntryInSecond = WorklogManager.AddEntry(second.Id, "Sch", new Rect(0, 0, 1, 1), "C", "", "Note", "Open", Array.Empty<string>());

        Assert.Equal(2, secondEntryInFirst!.Id);
        Assert.Equal(1, firstEntryInSecond!.Id);
    }

    // ---------------------------------------------------------- the "New fault" card's preview badge

    [Fact]
    public void Peek_next_entry_id_previews_1_for_a_workbook_with_no_entries_yet()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        Assert.Equal(1, WorklogManager.PeekNextEntryId(workbook.Id));
    }

    [Fact]
    public void Peek_next_entry_id_matches_what_AddEntry_actually_assigns_next()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "First", "", "Note", "Open", Array.Empty<string>());

        int previewed = WorklogManager.PeekNextEntryId(workbook.Id);
        var secondEntry = WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Second", "", "Note", "Open", Array.Empty<string>());

        Assert.Equal(2, previewed);
        Assert.Equal(previewed, secondEntry!.Id);
    }

    [Fact]
    public void Peek_next_entry_id_is_scoped_to_its_own_workbook_not_the_workbooks_own_id()
    {
        this.LoadWorklog();
        var first = CreateWorkbook("Commodore 64|250469", "First job", "");
        var second = CreateWorkbook("Amiga 500|A500", "Second job", "");
        WorklogManager.AddEntry(first.Id, "Sch", new Rect(0, 0, 1, 1), "A", "", "Note", "Open", Array.Empty<string>());

        // Regression: the "New fault" card's badge used to display the workbook's own id instead of
        // the entry's, so every entry added to workbook #1 previewed as "#1" no matter how many
        // entries it already had.
        Assert.Equal(2, WorklogManager.PeekNextEntryId(first.Id));
        Assert.Equal(1, WorklogManager.PeekNextEntryId(second.Id));
    }

    [Fact]
    public void Peek_next_entry_id_previews_1_for_a_workbook_that_does_not_exist()
    {
        this.LoadWorklog();

        Assert.Equal(1, WorklogManager.PeekNextEntryId(999));
    }

    [Fact]
    public void Adding_an_open_entry_keeps_the_workbook_open()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Desc", "", "Issue", "Open", Array.Empty<string>());

        var active = WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469");
        Assert.NotNull(active);
        Assert.Equal("Open", active!.Status);
        Assert.Equal(1, active.EntryCount);
    }

    [Fact]
    public void A_workbook_auto_closes_once_its_only_entry_is_closed()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Desc", "", "Issue", "Closed", Array.Empty<string>());

        // A Closed workbook is no longer "active" - GetActiveWorkbookForBoard only returns Open ones.
        Assert.Null(WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469"));

        var entries = WorklogManager.GetEntries(workbook.Id);
        Assert.Single(entries);
    }

    // An entry state the app never writes must not silently count as resolved: closing a workbook
    // the user still considers open is the one direction of this rule that loses information.
    [Theory]
    [InlineData("Whatever")]
    [InlineData("half-done")]
    public void An_unrecognised_entry_state_does_not_close_a_workbook(string state)
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Desc", "", "Issue", state, Array.Empty<string>());

        var active = WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469");
        Assert.NotNull(active);
        Assert.Equal("Open", active!.Status);
    }

    [Fact]
    public void A_workbook_stays_open_while_any_entry_is_still_open()
    {
        // Two entries: one closed, one still open - the workbook must not close just because
        // *an* entry was closed, only once *every* entry is.
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Closed one", "", "Issue", "Closed", Array.Empty<string>());
        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Still open", "", "Issue", "Open", Array.Empty<string>());

        var active = WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469");
        Assert.NotNull(active);
        Assert.Equal("Open", active!.Status);
        Assert.Equal(2, active.EntryCount);
    }

    [Fact]
    public void A_workbook_closes_once_its_last_outstanding_entry_is_resolved()
    {
        // Closing the first entry alone must not close the workbook (a second is still open);
        // closing the second one too must close it.
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "First", "", "Issue", "Closed", Array.Empty<string>());
        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Second", "", "Issue", "Closed", Array.Empty<string>());

        Assert.Null(WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469"));
    }

    // ---------------------------------------------------------- the bar's display-only lookup

    [Fact]
    public void The_latest_workbook_lookup_still_finds_a_workbook_that_auto_closed()
    {
        // This is the whole reason GetLatestWorkbookForBoard exists alongside the Open-only
        // GetActiveWorkbookForBoard. Resolving a workbook's last outstanding entry closes it, and
        // when the worklog bar asked only for an *open* workbook, a finished one vanished from the
        // UI completely - it read as data loss even though everything was still on disk.
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Desc", "", "Note", "Closed", Array.Empty<string>());

        Assert.Null(WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469"));

        var latest = WorklogManager.GetLatestWorkbookForBoard("Commodore 64|250469");
        Assert.NotNull(latest);
        Assert.Equal(workbook.Id, latest!.Id);
        Assert.Equal("Closed", latest.Status);
        Assert.Equal(1, latest.EntryCount);
    }

    [Fact]
    public void The_latest_workbook_lookup_returns_the_highest_id_regardless_of_status()
    {
        // Highest id wins even when an older workbook is the one still open - the bar shows the
        // most recent job, not the most recent *open* job.
        this.LoadWorklog();
        var older = CreateWorkbook("Commodore 64|250469", "Older job", "");
        var newer = CreateWorkbook("Commodore 64|250469", "Newer job", "");

        WorklogManager.AddEntry(older.Id, "Sch", new Rect(0, 0, 1, 1), "Still open", "", "Issue", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(newer.Id, "Sch", new Rect(0, 0, 1, 1), "Done", "", "Issue", "Closed", Array.Empty<string>());

        Assert.Equal(older.Id, WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469")!.Id);
        Assert.Equal(newer.Id, WorklogManager.GetLatestWorkbookForBoard("Commodore 64|250469")!.Id);
    }

    [Fact]
    public void Adding_an_open_entry_to_a_closed_workbook_reopens_it()
    {
        // Why the worklog bar needs no separate "Reopen" affordance: "Add worklog" stays available
        // on a closed workbook, and the entry it adds is Open, which RecomputeWorkbookStatus turns
        // straight back into Open.
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Done", "", "Issue", "Closed", Array.Empty<string>());
        Assert.Equal("Closed", WorklogManager.GetLatestWorkbookForBoard("Commodore 64|250469")!.Status);

        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "New fault", "", "Issue", "Open", Array.Empty<string>());

        var reopened = WorklogManager.GetLatestWorkbookForBoard("Commodore 64|250469");
        Assert.Equal("Open", reopened!.Status);
        Assert.Equal(2, reopened.EntryCount);
        Assert.Equal(workbook.Id, WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469")!.Id);
    }

    [Fact]
    public void The_latest_workbook_lookup_is_scoped_to_its_board_and_rejects_a_blank_key()
    {
        this.LoadWorklog();
        CreateWorkbook("Commodore 64|250469", "C64 job", "");

        Assert.Null(WorklogManager.GetLatestWorkbookForBoard("Amiga 500|A500"));
        Assert.Null(WorklogManager.GetLatestWorkbookForBoard(""));
        Assert.Null(WorklogManager.GetLatestWorkbookForBoard("   "));
    }

    // ---------------------------------------------------------------------------------------
    // GetWorkbooksForBoard - the Workbooks tab's list. Unlike the two lookups above it does not
    // reduce a board's history to one workbook, so these pin down that it returns the whole set,
    // in the order the tab renders it.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void A_board_with_no_workbooks_lists_an_empty_set_rather_than_null()
    {
        this.LoadWorklog();

        // The tab does a .Count on this straight away to build its "N workbooks" heading, so a
        // null here would be a crash on the most ordinary case there is - a board nobody has
        // worked on yet.
        var workbooks = WorklogManager.GetWorkbooksForBoard("Commodore 64|250469");

        Assert.NotNull(workbooks);
        Assert.Empty(workbooks);
    }

    [Fact]
    public void Every_workbook_for_a_board_is_listed_newest_first()
    {
        this.LoadWorklog();

        var first = CreateWorkbook("Commodore 64|250469", "Full recap", "");
        var second = CreateWorkbook("Commodore 64|250469", "Dead PLA", "");
        var third = CreateWorkbook("Commodore 64|250469", "No picture", "");

        var workbooks = WorklogManager.GetWorkbooksForBoard("Commodore 64|250469");

        // Descending id, matching the two single-workbook lookups. The id is what the user sees
        // on each card as "#N", so a list that was not in id order would look sorted by nothing.
        Assert.Equal(
            new[] { third.Id, second.Id, first.Id },
            workbooks.Select(w => w.Id).ToArray());
    }

    [Fact]
    public void The_workbook_list_includes_closed_workbooks_not_just_open_ones()
    {
        this.LoadWorklog();

        var open = CreateWorkbook("Commodore 64|250469", "Still going", "");
        var closed = CreateWorkbook("Commodore 64|250469", "Finished", "");

        // Status is derived, not assigned: a workbook closes when every one of its entries is
        // resolved, so this closes it the way the app does rather than writing the field.
        WorklogManager.AddEntry(open.Id, "Sch", new Rect(0, 0, 1, 1), "Still looking", "", "Issue", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(closed.Id, "Sch", new Rect(0, 0, 1, 1), "Sorted", "", "Issue", "Closed", Array.Empty<string>());

        var workbooks = WorklogManager.GetWorkbooksForBoard("Commodore 64|250469");

        // The tab is a history of everything done to a board, so a finished repair must stay in
        // the list - that is the whole point of it. This is the same reasoning that made
        // GetLatestWorkbookForBoard status-blind; an Open-only list would make a workbook vanish
        // the moment it was completed, which reads as data loss.
        Assert.Equal(2, workbooks.Count);
        Assert.Contains(workbooks, w => w.Id == open.Id && w.Status == "Open");
        Assert.Contains(workbooks, w => w.Id == closed.Id && w.Status == "Closed");

        // The card's third line reads "{EntryCount} worklogs", so the count has to come back on
        // the listed records rather than only on the single-workbook lookups.
        Assert.All(workbooks, w => Assert.Equal(1, w.EntryCount));
    }

    [Fact]
    public void The_workbook_list_is_scoped_to_its_board_and_rejects_a_blank_key()
    {
        this.LoadWorklog();
        CreateWorkbook("Commodore 64|250469", "C64 job", "");
        CreateWorkbook("Amiga 500|A500", "Amiga job", "");

        Assert.Single(WorklogManager.GetWorkbooksForBoard("Commodore 64|250469"));
        Assert.Single(WorklogManager.GetWorkbooksForBoard("Amiga 500|A500"));

        // A blank key yields nothing rather than everything on disk. Returning everything would
        // list one board's repairs under another board's name, which is worse than showing none.
        Assert.Empty(WorklogManager.GetWorkbooksForBoard(""));
        Assert.Empty(WorklogManager.GetWorkbooksForBoard("   "));
    }

    // ---------------------------------------------------------------------------------------
    // GetAllWorkbooks - the worklog bar's cross-board picker. Unlike GetWorkbooksForBoard this is
    // the one lookup that is deliberately NOT scoped to a single board.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void No_workbooks_on_disk_lists_an_empty_set_rather_than_null()
    {
        this.LoadWorklog();

        var workbooks = WorklogManager.GetAllWorkbooks();

        Assert.NotNull(workbooks);
        Assert.Empty(workbooks);
    }

    [Fact]
    public void Every_workbook_on_every_board_is_listed_newest_first()
    {
        this.LoadWorklog();

        var first = CreateWorkbook("Commodore 64|250469", "Full recap", "");
        var second = CreateWorkbook("Amiga 500|A500", "Dead PLA", "");
        var third = CreateWorkbook("Commodore 64|250469", "No picture", "");

        var workbooks = WorklogManager.GetAllWorkbooks();

        // Descending id, matching GetWorkbooksForBoard's own order - the picker's items should not
        // suddenly re-sort just because they now span boards.
        Assert.Equal(
            new[] { third.Id, second.Id, first.Id },
            workbooks.Select(w => w.Id).ToArray());
    }

    [Fact]
    public void The_full_list_spans_boards_instead_of_being_scoped_to_one()
    {
        this.LoadWorklog();
        var c64Job = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        var amigaJob = CreateWorkbook("Amiga 500|A500", "Amiga job", "");

        var workbooks = WorklogManager.GetAllWorkbooks();

        // The whole point of this lookup over GetWorkbooksForBoard: both boards' workbooks come
        // back together, each still carrying its own BoardKey so the caller can tell them apart
        // and jump to the right board.
        Assert.Equal(2, workbooks.Count);
        Assert.Contains(workbooks, w => w.Id == c64Job.Id && w.BoardKey == "Commodore 64|250469");
        Assert.Contains(workbooks, w => w.Id == amigaJob.Id && w.BoardKey == "Amiga 500|A500");
    }

    [Fact]
    public void Adding_an_entry_to_a_workbook_with_no_folder_returns_null_instead_of_throwing()
    {
        this.LoadWorklog();

        Exception? thrown = Record.Exception(() =>
        {
            var result = WorklogManager.AddEntry(999, "Sch", new Rect(0, 0, 1, 1), "Desc", "", "Note", "Open", Array.Empty<string>());
            Assert.Null(result);
        });

        Assert.True(thrown is null);
    }

    [Fact]
    public void An_unknown_workbooks_entries_come_back_as_an_empty_list_instead_of_throwing()
    {
        this.LoadWorklog();

        var entries = WorklogManager.GetEntries(999);

        Assert.Empty(entries);
    }

    // Was "starts it with no comments". A new entry now opens its own audit trail with the
    // automatic "Worklog created" line, so the rule this guards - what an entry's comment list
    // holds the moment it is created - is unchanged; only the expected content moved.
    [Fact]
    public void Adding_an_entry_starts_it_with_only_the_created_comment_and_a_matching_title()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var entry = WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        Assert.Equal("Bad cap", entry!.Title);
        Assert.Single(entry.Comments);
        Assert.Equal("Worklog created", entry.Comments[0].Text);
    }

    [Fact]
    public void Updating_an_entry_overwrites_it_in_place_and_recomputes_workbook_status()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        var entry = WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>())!;

        entry.Title = "Bad cap - replaced";
        entry.State = "Closed";
        entry.Links.Add(new WorklogLinkRecord { Id = 1, Headline = "Datasheet", Url = "https://example.com" });

        bool updated = WorklogManager.UpdateEntry(workbook.Id, entry);

        Assert.True(updated);

        var storedEntries = WorklogManager.GetEntries(workbook.Id);
        Assert.Single(storedEntries);
        Assert.Equal("Bad cap - replaced", storedEntries[0].Title);
        Assert.Equal("Closed", storedEntries[0].State);
        Assert.Single(storedEntries[0].Links);

        // Editing State to Closed is how the full editor resolves an entry - the workbook must
        // auto-close exactly as it would from the quick-card flow.
        Assert.Null(WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469"));
    }

    [Fact]
    public void Editing_a_resolved_entry_back_to_open_reopens_an_already_closed_workbook()
    {
        // RecomputeWorkbookStatus's doc comment promises a still-Open entry "keeps (or
        // reopens) the workbook" - this pins down the reopen half specifically, via the full
        // editor's UpdateEntry path rather than AddEntry.
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        var entry = WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Closed", Array.Empty<string>())!;

        Assert.Null(WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469"));

        entry.State = "Open";
        WorklogManager.UpdateEntry(workbook.Id, entry);

        var active = WorklogManager.GetActiveWorkbookForBoard("Commodore 64|250469");
        Assert.NotNull(active);
        Assert.Equal("Open", active!.Status);
    }

    // Regression: the full editor instant-saves after every sub-list change (add a comment, a link,
    // a work-done row). Because UpdateEntry replaces the whole record rather than merging, such a
    // save carries the editor's working copy wholesale - so if that copy still held the title and
    // state the entry was opened with, it silently reverted whatever the user had just retyped.
    // This pins the destructive half of that: a stale-scalar update overwrites good values on disk,
    // which is why WorklogEntryEditorWindow.PersistEntrySilently must sync its direct fields first.
    [Fact]
    public void Updating_an_entry_replaces_the_whole_record_so_stale_fields_overwrite_saved_ones()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        var entry = WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "Leaking", "Issue", "Open", Array.Empty<string>())!;

        // The user retypes the headline and resolves the entry, and it reaches disk.
        entry.Title = "Bad cap - replaced";
        entry.State = "Closed";
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        // A later write built from a copy that never saw those edits takes the record back.
        var staleCopy = WorklogManager.GetEntries(workbook.Id).Single();
        staleCopy.Title = "Bad cap";
        staleCopy.State = "Open";
        staleCopy.Comments.Add(new WorklogCommentRecord { Id = 2, Text = "Added a note", Date = DateTime.Now });
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, staleCopy));

        var stored = WorklogManager.GetEntries(workbook.Id).Single();
        Assert.Equal("Bad cap", stored.Title);
        Assert.Equal("Open", stored.State);

        // Two: the automatic "Worklog created" comment the entry was born with, plus the one the
        // stale copy added. What matters here is that the stale write took the record back
        // wholesale, not the count itself.
        Assert.Equal(2, stored.Comments.Count);
        Assert.Equal("Added a note", stored.Comments[1].Text);

        // And the reverted state feeds straight back into the workbook's Open/Closed status.
        Assert.Equal("Open", WorklogManager.GetLatestWorkbookForBoard("Commodore 64|250469")!.Status);
    }

    // Regression: SaveEntries swallowed every write exception and returned void, so UpdateEntry
    // reported true even when nothing reached disk - the editor then closed looking saved and the
    // user watched their edits revert on the next refresh. A directory sitting where the ".tmp"
    // file must be written fails the write deterministically, with no permissions games.
    [Fact]
    public void A_failed_write_is_reported_rather_than_being_swallowed()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job");
        var entry = WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>())!;

        Directory.CreateDirectory(Path.Combine(root, workbook.Id.ToString(), "entries.json.tmp"));

        entry.Title = "Bad cap - replaced";

        Assert.False(WorklogManager.UpdateEntry(workbook.Id, entry));
        Assert.False(WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Another", "", "Note", "Open", Array.Empty<string>()) is not null);

        // The entry on disk is untouched, which is the point: the caller must not report success.
        Assert.Equal("Bad cap", WorklogManager.GetEntries(workbook.Id).Single().Title);
    }

    [Fact]
    public void Updating_an_unknown_entry_returns_false_instead_of_throwing()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var phantom = new WorklogEntryRecord { Id = 999 };

        Exception? thrown = Record.Exception(() =>
        {
            bool updated = WorklogManager.UpdateEntry(workbook.Id, phantom);
            Assert.False(updated);
        });

        Assert.True(thrown is null);
    }

    [Fact]
    public void Updating_an_entry_in_an_unknown_workbook_returns_false_instead_of_throwing()
    {
        this.LoadWorklog();

        var phantom = new WorklogEntryRecord { Id = 1 };

        Exception? thrown = Record.Exception(() =>
        {
            bool updated = WorklogManager.UpdateEntry(999, phantom);
            Assert.False(updated);
        });

        Assert.True(thrown is null);
    }

    [Fact]
    public void The_entry_attachments_folder_is_created_inside_the_workbooks_own_folder()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        string? attachmentsFolder = WorklogManager.GetEntryAttachmentsFolder(workbook.Id, 1);

        Assert.NotNull(attachmentsFolder);

        // "worklog_{id}", not the older "entry-{id}-files". This folder is visible to the user -
        // in the Workbooks folder on disk AND inside an exported ZIP, which reuses this exact name -
        // and the app says "worklog" everywhere a user can see one. Renamed deliberately, with no
        // migration of existing data: an older folder keeps its name and its attachments simply
        // stop being found.
        Assert.Equal(Path.Combine(root, workbook.Id.ToString(), "worklog_1"), attachmentsFolder);
        Assert.True(Directory.Exists(attachmentsFolder));
    }

    [Fact]
    public void The_entry_attachments_folder_for_an_unknown_workbook_returns_null_instead_of_throwing()
    {
        this.LoadWorklog();

        string? attachmentsFolder = WorklogManager.GetEntryAttachmentsFolder(999, 1);

        Assert.Null(attachmentsFolder);
    }

    // ------------------------------------------------------------- automatic "created" comment

    // A new worklog starts its own history with the fact that it was created, so its Comments list
    // is an audit trail from the first entry rather than starting empty and only recording what
    // happened afterwards.
    [Fact]
    public void A_new_entry_starts_with_a_worklog_created_comment()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var entry = WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        Assert.NotNull(entry);
        Assert.Single(entry!.Comments);
        Assert.Equal("Worklog created", entry.Comments[0].Text);
        Assert.Equal(1, entry.Comments[0].Id);
    }

    // The comment must survive the write, not just exist on the returned object - it is the stored
    // entry the editor will later read back and show.
    [Fact]
    public void The_created_comment_is_written_to_disk()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        var stored = WorklogManager.GetEntries(workbook.Id);

        Assert.Single(stored);
        Assert.Single(stored[0].Comments);
        Assert.Equal("Worklog created", stored[0].Comments[0].Text);
    }

    // Each entry gets its OWN created comment - the list is per-entry, so a second entry must not
    // inherit or share the first one.
    [Fact]
    public void Every_new_entry_gets_its_own_created_comment()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "First", "", "Note", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Second", "", "Note", "Open", Array.Empty<string>());

        var stored = WorklogManager.GetEntries(workbook.Id);

        Assert.Equal(2, stored.Count);
        Assert.All(stored, e =>
        {
            Assert.Single(e.Comments);
            Assert.Equal("Worklog created", e.Comments[0].Text);
        });
    }

    // ------------------------------------------------------------- "Show marked area"

    [Fact]
    public void A_new_entry_shows_its_marked_area_by_default()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var entry = WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        Assert.True(entry!.ShowMarkedArea);
    }

    // The setting is a normal round-tripped field: unticking it must survive the write and the
    // read back, or the area would reappear the next time the board was opened.
    [Fact]
    public void Hiding_the_marked_area_survives_a_save_and_reload()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        var entry = WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>())!;

        entry.ShowMarkedArea = false;
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        Assert.False(WorklogManager.GetEntries(workbook.Id).Single().ShowMarkedArea);
    }

    // THE upgrade case. An entry written by a build that predates this field has no
    // "showMarkedArea" key at all, and must read back as true - a default of false would silently
    // blank every marked area on the board for anyone upgrading.
    [Fact]
    public void An_entry_written_before_the_field_existed_still_shows_its_area()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        // Strip the key back out, reproducing exactly what an older build wrote.
        string entriesPath = Path.Combine(root, workbook.Id.ToString(), "entries.json");
        string json = File.ReadAllText(entriesPath);
        Assert.Contains("showMarkedArea", json);

        using (var document = JsonDocument.Parse(json))
        {
            var stripped = document.RootElement.EnumerateArray()
                .Select(e =>
                {
                    var map = new Dictionary<string, JsonElement>();
                    foreach (var property in e.EnumerateObject())
                    {
                        if (!string.Equals(property.Name, "showMarkedArea", StringComparison.Ordinal))
                        {
                            map[property.Name] = property.Value;
                        }
                    }
                    return map;
                })
                .ToList();

            File.WriteAllText(entriesPath, JsonSerializer.Serialize(stripped));
        }

        Assert.DoesNotContain("showMarkedArea", File.ReadAllText(entriesPath));

        Assert.True(WorklogManager.GetEntries(workbook.Id).Single().ShowMarkedArea);
    }

    // ------------------------------------------------------------- "Mark components completed"

    // A new entry has nothing done yet - a component that was just put in scope is work still to
    // do, so "not started" is the only honest starting point.
    [Fact]
    public void A_new_entry_has_no_completed_components()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var entry = WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Recap", "", "Issue", "Open", new[] { "C1", "C2" });

        Assert.Empty(entry!.CompletedComponentLabels);
    }

    [Fact]
    public void Completed_components_survive_a_save_and_reload()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        var entry = WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Recap", "", "Issue", "Open", new[] { "C1", "C2" })!;

        entry.CompletedComponentLabels = new List<string> { "C1" };
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        Assert.Equal(new[] { "C1" }, WorklogManager.GetEntries(workbook.Id).Single().CompletedComponentLabels);
    }

    // An entry written before this field existed reads back with an empty list rather than null -
    // every consumer enumerates it, so a null would be a crash rather than a missing feature.
    [Fact]
    public void An_entry_written_before_the_completed_field_existed_reads_back_empty()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Recap", "", "Issue", "Open", new[] { "C1" });

        string entriesPath = Path.Combine(root, workbook.Id.ToString(), "entries.json");
        string json = File.ReadAllText(entriesPath);

        using (var document = JsonDocument.Parse(json))
        {
            var stripped = document.RootElement.EnumerateArray()
                .Select(e =>
                {
                    var map = new Dictionary<string, JsonElement>();
                    foreach (var property in e.EnumerateObject())
                    {
                        if (!string.Equals(property.Name, "completedComponentLabels", StringComparison.Ordinal))
                        {
                            map[property.Name] = property.Value;
                        }
                    }
                    return map;
                })
                .ToList();

            File.WriteAllText(entriesPath, JsonSerializer.Serialize(stripped));
        }

        var reloaded = WorklogManager.GetEntries(workbook.Id).Single();

        Assert.NotNull(reloaded.CompletedComponentLabels);
        Assert.Empty(reloaded.CompletedComponentLabels);
    }

    // ------------------------------------------------------------- collapsed list sections

    [Fact]
    public void A_new_entry_has_no_collapsed_sections()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var entry = WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Recap", "", "Issue", "Open", Array.Empty<string>());

        Assert.Empty(entry!.CollapsedSections);
    }

    [Fact]
    public void Collapsed_sections_survive_a_save_and_reload()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        var entry = WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Recap", "", "Issue", "Open", Array.Empty<string>())!;

        entry.CollapsedSections = new List<string> { "EditorCommentsHeader", "EditorPhotosHeader" };
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        Assert.Equal(
            new[] { "EditorCommentsHeader", "EditorPhotosHeader" },
            WorklogManager.GetEntries(workbook.Id).Single().CollapsedSections);
    }

    // An entry written before the field existed reads back with an empty list rather than null -
    // the editor enumerates it on open, so a null would be a crash rather than a missing feature.
    [Fact]
    public void An_entry_written_before_the_collapsed_field_existed_reads_back_empty()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Recap", "", "Issue", "Open", Array.Empty<string>());

        string entriesPath = Path.Combine(root, workbook.Id.ToString(), "entries.json");

        using (var document = JsonDocument.Parse(File.ReadAllText(entriesPath)))
        {
            var stripped = document.RootElement.EnumerateArray()
                .Select(e =>
                {
                    var map = new Dictionary<string, JsonElement>();
                    foreach (var property in e.EnumerateObject())
                    {
                        if (!string.Equals(property.Name, "collapsedSections", StringComparison.Ordinal))
                        {
                            map[property.Name] = property.Value;
                        }
                    }
                    return map;
                })
                .ToList();

            File.WriteAllText(entriesPath, JsonSerializer.Serialize(stripped));
        }

        var reloaded = WorklogManager.GetEntries(workbook.Id).Single();

        Assert.NotNull(reloaded.CollapsedSections);
        Assert.Empty(reloaded.CollapsedSections);
    }

    // Persisting a fold must not carry unsaved direct-field edits with it. The editor writes the
    // fold onto the STORED record rather than its working copy, so a pending Description change or
    // an unticked "Show marked area" - fields the user can still abandon with Cancel - stay off
    // disk. This models that write and pins the rule the editor depends on.
    [Fact]
    public void Writing_a_fold_onto_the_stored_record_leaves_other_fields_untouched()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        var entry = WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Recap", "Original description", "Issue", "Open", Array.Empty<string>())!;

        // A working copy carrying edits the user has NOT saved.
        var workingCopy = WorklogManager.GetEntries(workbook.Id).Single();
        workingCopy.Description = "Edited but not saved";
        workingCopy.ShowMarkedArea = false;

        // The fold is written onto a freshly read record, not onto that working copy.
        var stored = WorklogManager.GetEntries(workbook.Id).Single();
        stored.CollapsedSections = new List<string> { "EditorCommentsHeader" };
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, stored));

        var reloaded = WorklogManager.GetEntries(workbook.Id).Single();

        Assert.Equal(new[] { "EditorCommentsHeader" }, reloaded.CollapsedSections);
        Assert.Equal("Original description", reloaded.Description);
        Assert.True(reloaded.ShowMarkedArea);
        Assert.Equal(entry.Id, reloaded.Id);
    }

    // ------------------------------------------------- "which workbook is active" (pure)

    // The rule Main's worklog bar, "Show worklogs", "Add worklog" AND the Workbooks tab's
    // highlighted card all resolve through. It used to be written out twice, in two different
    // shapes, which is how the highlighted card and the bar came to be able to disagree.
    //
    // Pure, so these need no workbook folders on disk - which is the point of extracting it.
    [Fact]
    public void The_saved_active_workbook_wins_over_the_newest_one()
    {
        // Newest-first, matching GetWorkbooksForBoard's own order.
        var workbooks = new List<WorkbookRecord>
        {
            new() { Id = 3, Title = "Newest" },
            new() { Id = 2, Title = "Middle" },
            new() { Id = 1, Title = "Oldest" },
        };

        Assert.Equal(1, WorklogManager.ResolveActiveWorkbook(workbooks, 1)!.Id);
        Assert.Equal(2, WorklogManager.ResolveActiveWorkbook(workbooks, 2)!.Id);
    }

    // Nothing saved: the newest, which is what every worklog surface defaulted to before workbooks
    // could be activated at all.
    [Fact]
    public void With_no_saved_activation_the_newest_workbook_is_active()
    {
        var workbooks = new List<WorkbookRecord>
        {
            new() { Id = 3, Title = "Newest" },
            new() { Id = 1, Title = "Oldest" },
        };

        Assert.Equal(3, WorklogManager.ResolveActiveWorkbook(workbooks, null)!.Id);
    }

    // A saved id naming a workbook this board no longer has - the folder deleted by hand, or the id
    // left over from another board - must fall back rather than resolve to nothing. Returning null
    // here would make the worklog bar quietly show "no workbook" for a board that plainly has some.
    [Fact]
    public void A_saved_id_that_names_no_workbook_falls_back_to_the_newest()
    {
        var workbooks = new List<WorkbookRecord>
        {
            new() { Id = 3, Title = "Newest" },
            new() { Id = 1, Title = "Oldest" },
        };

        Assert.Equal(3, WorklogManager.ResolveActiveWorkbook(workbooks, 9999)!.Id);
    }

    [Fact]
    public void A_board_with_no_workbooks_has_no_active_workbook()
    {
        Assert.Null(WorklogManager.ResolveActiveWorkbook(new List<WorkbookRecord>(), null));
        Assert.Null(WorklogManager.ResolveActiveWorkbook(new List<WorkbookRecord>(), 1));
    }

    // ------------------------------------------------------- "which states mean finished"

    [Fact]
    public void Closed_is_the_one_resolved_entry_state()
    {
        Assert.True(WorklogManager.IsResolvedState("Closed"));
        Assert.False(WorklogManager.IsResolvedState("Open"));
    }

    // Case- and whitespace-insensitive, unlike the ResolvedEntryStates set it consults. States are
    // read back off disk and can carry a hand edit or an older build's casing; falling through to
    // "open" there draws a RED padlock on a pill whose own label reads "closed", which is
    // indistinguishable from the intended default and so would never be noticed.
    [Theory]
    [InlineData("closed")]
    [InlineData("CLOSED")]
    [InlineData(" Closed ")]
    public void A_differently_cased_closed_state_still_reads_as_resolved(string state)
    {
        Assert.True(WorklogManager.IsResolvedState(state));
    }

    // Anything unrecognised reads as unresolved - the safe direction, since it leaves a workbook
    // open rather than silently closing it.
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    [InlineData("Pending")]
    [InlineData("SomeFutureState")]
    public void An_unrecognised_state_is_not_resolved(string? state)
    {
        Assert.False(WorklogManager.IsResolvedState(state));
    }

    // The workbook axis, not the entry one - RecomputeWorkbookStatus writes this field. Same
    // case-insensitive read, same "anything unrecognised reads as open" fallback every status pill
    // in the app already applies.
    [Theory]
    [InlineData("Open", true)]
    [InlineData("open", true)]
    [InlineData("", true)]
    [InlineData(null, true)]
    [InlineData("Closed", false)]
    [InlineData("closed", false)]
    [InlineData(" Closed ", false)]
    public void A_workbook_status_reads_as_open_unless_it_is_recognisably_closed(string? status, bool expected)
    {
        Assert.Equal(expected, WorklogManager.IsWorkbookStatusOpen(status));
    }

    // The auto-close rule goes through the same IsResolvedState, so an entry stored as "closed"
    // closes its workbook rather than holding it open forever while its own pill reads Closed.
    [Fact]
    public void A_differently_cased_closed_entry_still_closes_its_workbook()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var entry = WorklogManager.AddEntry(
            workbook.Id, "Sch", new Rect(0, 0, 1, 1), "Done", "", "Issue", "Open", Array.Empty<string>())!;

        Assert.Equal("Open", WorklogManager.GetLatestWorkbookForBoard("Commodore 64|250469")!.Status);

        // The casing a hand edit of entries.json, or an older build, could leave behind.
        entry.State = "closed";
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        Assert.Equal("Closed", WorklogManager.GetLatestWorkbookForBoard("Commodore 64|250469")!.Status);
    }

    // ------------------------------------------------- retired entry states, migrated on read

    /// <summary>
    /// Writes one entry to a workbook's entries.json with a raw State string, the way an older build
    /// or a hand edit would leave it. The state has to go in by hand: nothing in the app can produce
    /// a retired value any more, which is exactly why the read-time mapping is the only thing
    /// standing between such a file and being misread.
    /// </summary>
    private static void OverwriteSingleEntryState(string root, int workbookId, string rawState)
    {
        string entriesPath = Path.Combine(root, workbookId.ToString(), "entries.json");

        File.WriteAllText(entriesPath, $$"""
        [
          {
            "id": 1,
            "schematicName": "Sch",
            "title": "Bad cap",
            "category": "Issue",
            "state": "{{rawState}}"
          }
        ]
        """);
    }

    // Entries used to carry Pending/Fixed/RuledOut; they now carry Open/Closed. Without the
    // read-time mapping, an entry saved by an older build is silently misread: ResolvedEntryStates
    // holds only "Closed", so a "Fixed" entry reads as unresolved, the editor renders NEITHER state
    // pill as selected, and the next save writes the retired value straight back.
    //
    // Fixed and RuledOut both map to CLOSED because both counted as resolved under the old rule -
    // mapping either anywhere else would change every affected workbook's status rather than
    // preserve it.
    [Theory]
    [InlineData("Pending", "Open")]
    [InlineData("Fixed", "Closed")]
    [InlineData("RuledOut", "Closed")]
    public void A_retired_entry_state_is_mapped_to_its_replacement_on_read(string stored, string expected)
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job");
        AddEntry(workbook.Id, "Sch", "Bad cap");

        OverwriteSingleEntryState(root, workbook.Id, stored);

        Assert.Equal(expected, WorklogManager.GetEntries(workbook.Id).Single().State);
    }

    // Case-insensitively, for the same reason IsResolvedState is: the value is read off disk, where
    // a hand edit or an older build's casing is exactly what turns up.
    [Theory]
    [InlineData("fixed")]
    [InlineData("FIXED")]
    [InlineData(" Fixed ")]
    [InlineData("ruledout")]
    public void A_differently_cased_retired_state_is_still_mapped(string stored)
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job");
        AddEntry(workbook.Id, "Sch", "Bad cap");

        OverwriteSingleEntryState(root, workbook.Id, stored);

        Assert.Equal("Closed", WorklogManager.GetEntries(workbook.Id).Single().State);
    }

    // The consequence that matters most, and the one no other test here would catch: a workbook the
    // user had FINISHED under the old vocabulary must not reopen itself on upgrade. Without the
    // mapping, "Fixed" reads as unresolved and RecomputeWorkbookStatus reopens the workbook.
    [Fact]
    public void A_workbook_whose_only_entry_was_Fixed_does_not_reopen_itself_on_read()
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job");
        var entry = AddEntry(workbook.Id, "Sch", "Bad cap");

        OverwriteSingleEntryState(root, workbook.Id, "Fixed");

        var migrated = WorklogManager.GetEntries(workbook.Id).Single();

        Assert.True(WorklogManager.IsResolvedState(migrated.State));

        // The write-back path: saving the migrated entry must record "Closed" rather than putting
        // the retired value back on disk, and must leave the workbook closed.
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, migrated));
        Assert.Equal("Closed", WorklogManager.GetEntries(workbook.Id).Single().State);
        Assert.Equal("Closed", WorklogManager.GetLatestWorkbookForBoard("Commodore 64|250469")!.Status);

        // The id is untouched by the migration - it names the entry's attachment folder and every
        // pill the user sees.
        Assert.Equal(entry.Id, migrated.Id);
    }

    // A blank state is an entry that predates the field, or a partially-written record. It reads as
    // Open rather than as an unrecognised value, so the entry has a pill the editor can select.
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void A_blank_stored_state_reads_as_Open(string stored)
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job");
        AddEntry(workbook.Id, "Sch", "Bad cap");

        OverwriteSingleEntryState(root, workbook.Id, stored);

        Assert.Equal("Open", WorklogManager.GetEntries(workbook.Id).Single().State);
    }

    // A current value passes through untouched, trimmed - the mapping must not be a rewrite of
    // everything it sees. An unrecognised one is LEFT ALONE rather than forced to Open, so
    // RecomputeWorkbookStatus treats it as unresolved and the workbook stays open: the safe
    // direction, since the alternative silently closes a workbook.
    [Theory]
    [InlineData("Open", "Open")]
    [InlineData("Closed", "Closed")]
    [InlineData(" Closed ", "Closed")]
    [InlineData("SomeFutureState", "SomeFutureState")]
    public void A_current_or_unknown_state_is_not_rewritten(string stored, string expected)
    {
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job");
        AddEntry(workbook.Id, "Sch", "Bad cap");

        OverwriteSingleEntryState(root, workbook.Id, stored);

        Assert.Equal(expected, WorklogManager.GetEntries(workbook.Id).Single().State);
    }

    // ------------------------------------------------------------- AddEntryRecord

    // AddEntryRecord is what the "Add worklog" flow writes with. Unlike AddEntry, which takes the
    // direct fields and builds a bare record, this takes an ALREADY-BUILT one - because the full
    // editor now opens directly on a new entry and holds the whole thing, sub-lists included, in
    // memory until Save.
    [Fact]
    public void AddEntryRecord_saves_an_entry_with_its_sub_lists_already_populated()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var draft = new WorklogEntryRecord
        {
            Id = 1,
            SchematicName = "Sheet 1",
            AreaX = 10,
            AreaY = 20,
            AreaWidth = 30,
            AreaHeight = 40,
            Title = "Dead VIC",
            Description = "No video at all",
            Category = "Issue",
            State = "Open",
        };
        draft.Comments.Add(new WorklogCommentRecord { Id = 1, Text = "Checked the rail", Date = DateTime.Now });
        draft.Links.Add(new WorklogLinkRecord { Id = 1, Headline = "Pinout", Url = "https://example.com" });

        var saved = WorklogManager.AddEntryRecord(workbook.Id, draft, reservedId: 1);

        Assert.NotNull(saved);
        Assert.Equal(1, saved!.Id);

        // Read back from disk, not from the returned object - the point is that the sub-lists
        // survived the write, which the in-memory record would report either way.
        var stored = Assert.Single(WorklogManager.GetEntries(workbook.Id));
        Assert.Equal("Dead VIC", stored.Title);
        Assert.Equal("No video at all", stored.Description);
        Assert.Equal("Issue", stored.Category);
        Assert.Equal("Sheet 1", stored.SchematicName);
        Assert.Equal(10, stored.AreaX);
        Assert.Equal(40, stored.AreaHeight);
        Assert.Equal("Checked the rail", Assert.Single(stored.Comments).Text);
        Assert.Equal("https://example.com", Assert.Single(stored.Links).Url);
    }

    // THE RACE THIS EXISTS TO CLOSE. The editor reserves an id up front (its attachment folder is
    // named after one), but a peek is not a reservation - another entry can be written in between.
    // AddEntryRecord must therefore re-allocate at write time; trusting the reserved id would
    // produce two entries sharing a number, which entries.json cannot represent and UpdateEntry
    // would resolve to whichever came first.
    [Fact]
    public void AddEntryRecord_reallocates_the_id_when_another_entry_claimed_the_reserved_one()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        // The editor peeks #1 and opens on a draft carrying it...
        int reserved = WorklogManager.PeekNextEntryId(workbook.Id);
        Assert.Equal(1, reserved);

        // ...and while it is open, something else writes an entry, taking #1.
        WorklogManager.AddEntry(
            workbook.Id, "Sheet 1", new Rect(0, 0, 1, 1), "Written meanwhile", "", "Note", "Open", Array.Empty<string>());

        var saved = WorklogManager.AddEntryRecord(
            workbook.Id,
            new WorklogEntryRecord { Id = reserved, SchematicName = "Sheet 1", Title = "The draft" },
            reserved);

        Assert.NotNull(saved);
        Assert.Equal(2, saved!.Id);

        var entries = WorklogManager.GetEntries(workbook.Id);
        Assert.Equal(2, entries.Count);
        Assert.Equal(new[] { 1, 2 }, entries.Select(e => e.Id).OrderBy(id => id));
    }

    // The draft's attachment BYTES are written the moment a photo is added, under the reserved id.
    // When the id changes, that folder has to follow, or the entry's photos point at a folder no
    // entry names - and worse, the next draft to reserve the same number would inherit them.
    [Fact]
    public void AddEntryRecord_moves_the_draft_attachment_folder_when_the_id_changes()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        int reserved = WorklogManager.PeekNextEntryId(workbook.Id);

        // The editor writes a photo into the folder for the id it reserved.
        string reservedFolder = WorklogManager.GetEntryAttachmentsFolder(workbook.Id, reserved)!;
        File.WriteAllText(Path.Combine(reservedFolder, "1_photo.png"), "bytes");

        // Something else takes that id first.
        WorklogManager.AddEntry(
            workbook.Id, "Sheet 1", new Rect(0, 0, 1, 1), "Written meanwhile", "", "Note", "Open", Array.Empty<string>());

        var saved = WorklogManager.AddEntryRecord(
            workbook.Id,
            new WorklogEntryRecord { Id = reserved, SchematicName = "Sheet 1", Title = "The draft" },
            reserved);

        Assert.NotNull(saved);
        Assert.NotEqual(reserved, saved!.Id);

        string finalFolder = WorklogManager.GetEntryAttachmentsFolder(workbook.Id, saved.Id)!;
        Assert.True(
            File.Exists(Path.Combine(finalFolder, "1_photo.png")),
            "the draft's photo did not follow the entry to its allocated id");

        Assert.False(
            Directory.Exists(reservedFolder),
            "the draft's attachment folder was left behind under the reserved id");
    }

    // The common case, and the one where a move would be actively wrong: the reserved id was still
    // free, so the folder is already in the right place and must be left exactly as it is.
    [Fact]
    public void AddEntryRecord_leaves_the_attachment_folder_alone_when_the_id_is_unchanged()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        int reserved = WorklogManager.PeekNextEntryId(workbook.Id);
        string folder = WorklogManager.GetEntryAttachmentsFolder(workbook.Id, reserved)!;
        File.WriteAllText(Path.Combine(folder, "1_photo.png"), "bytes");

        var saved = WorklogManager.AddEntryRecord(
            workbook.Id,
            new WorklogEntryRecord { Id = reserved, SchematicName = "Sheet 1", Title = "The draft" },
            reserved);

        Assert.NotNull(saved);
        Assert.Equal(reserved, saved!.Id);
        Assert.True(File.Exists(Path.Combine(folder, "1_photo.png")));
    }

    // Blank category/state are filled in the same way AddEntry fills them, so a record built by a
    // caller that did not set them cannot reach disk with an empty category the UI has no colour
    // for. CreatedDate likewise.
    [Fact]
    public void AddEntryRecord_defaults_a_blank_category_state_and_date()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var saved = WorklogManager.AddEntryRecord(
            workbook.Id,
            new WorklogEntryRecord { Id = 1, SchematicName = "Sheet 1", Title = "Bare", Category = "", State = "  " },
            reservedId: 1);

        Assert.NotNull(saved);
        Assert.Equal("Note", saved!.Category);
        Assert.Equal("Open", saved.State);
        Assert.NotEqual(default, saved.CreatedDate);
    }

    // GetEntryAttachmentsFolder CREATES the folder, which is right for a caller about to write into
    // it and self-defeating for one about to delete it - resolving to delete would re-create the
    // very folder being removed. The ...Path form exists for those callers and must NOT create.
    [Fact]
    public void GetEntryAttachmentsFolderPath_does_not_create_the_folder()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        string path = WorklogManager.GetEntryAttachmentsFolderPath(workbook.Id, 7)!;

        Assert.False(Directory.Exists(path), "resolving the path created the folder");

        // Same path the creating form resolves - the two must not name different folders.
        Assert.Equal(WorklogManager.GetEntryAttachmentsFolder(workbook.Id, 7), path);
    }

    // The editor's cancelled draft asks this before deleting the folder its RESERVED id names. A
    // peek is not a reservation, so that number can belong to an entry saved meanwhile - whose
    // photos and files live in exactly that folder. Answering wrongly destroys a saved entry's
    // attachments while its entries.json rows survive, pointing at nothing.
    [Fact]
    public void EntryExists_reports_whether_a_saved_entry_uses_the_id()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        Assert.False(WorklogManager.EntryExists(workbook.Id, 1));

        WorklogManager.AddEntry(
            workbook.Id, "Sheet 1", new Rect(0, 0, 1, 1), "Real", "", "Note", "Open", Array.Empty<string>());

        Assert.True(WorklogManager.EntryExists(workbook.Id, 1));
        Assert.False(WorklogManager.EntryExists(workbook.Id, 2));

        // A workbook that does not exist has no entries rather than throwing.
        Assert.False(WorklogManager.EntryExists(9999, 1));
    }

    // A draft whose reserved id is taken meanwhile is moved to the id it is actually allocated,
    // and that destination is now guaranteed to be FREE: since ids are handed out from a counter
    // that also walks past any "worklog_{id}" folder already on disk (see PeekNextEntryIdIn), the
    // allocator can no longer land a new entry on a folder that already has bytes in it.
    //
    // This test used to place a stale folder at reserved + 1 and assert the two were merged. That
    // situation is now unreachable through this path - the allocator skips reserved + 1 precisely
    // because the folder is there - so it asserts the skip instead. MoveEntryAttachmentsFolder's
    // merge behaviour is still real and still covered, directly, by the two tests below it: it
    // remains the safety net for a destination that appears between the allocation and the move.
    [Fact]
    public void AddEntryRecord_is_never_allocated_an_id_whose_attachment_folder_already_exists()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        int reserved = WorklogManager.PeekNextEntryId(workbook.Id);

        // The draft's own photo, under the id it reserved.
        string reservedFolder = WorklogManager.GetEntryAttachmentsFolder(workbook.Id, reserved)!;
        File.WriteAllText(Path.Combine(reservedFolder, "photo_1_draft.png"), "draft bytes");

        // Something else takes that id first, so the draft cannot keep it.
        WorklogManager.AddEntry(
            workbook.Id, "Sheet 1", new Rect(0, 0, 1, 1), "Written meanwhile", "", "Note", "Open", Array.Empty<string>());

        // A stale folder from an earlier cancelled draft sits at the NEXT id along - the one the
        // draft would otherwise have been given.
        string staleFolder = WorklogManager.GetEntryAttachmentsFolder(workbook.Id, reserved + 1)!;
        File.WriteAllText(Path.Combine(staleFolder, "photo_9_orphan.png"), "orphan bytes");

        var saved = WorklogManager.AddEntryRecord(
            workbook.Id,
            new WorklogEntryRecord { Id = reserved, SchematicName = "Sheet 1", Title = "The draft" },
            reserved);

        Assert.NotNull(saved);

        // Skipped PAST the stale folder rather than moved into it: adopting a cancelled draft's
        // photos into someone's new worklog is exactly the outcome the folder check prevents.
        Assert.Equal(reserved + 2, saved!.Id);

        string destinationFolder = Path.Combine(
            WorklogManager.GetEntryAttachmentsFolderPath(workbook.Id, saved.Id)!);

        Assert.True(
            File.Exists(Path.Combine(destinationFolder, "photo_1_draft.png")),
            "the draft's photo did not follow it to the id it was allocated");

        // The orphan is left exactly where it was, still belonging to nothing.
        Assert.True(File.Exists(Path.Combine(staleFolder, "photo_9_orphan.png")));
        Assert.False(File.Exists(Path.Combine(staleFolder, "photo_1_draft.png")));

        Assert.False(
            Directory.Exists(reservedFolder),
            "the emptied reserved folder was left behind");
    }

    // MoveEntryAttachmentsFolder's merge, tested directly rather than through AddEntryRecord.
    //
    // It is reached by reflection because the allocator above can no longer produce a collision
    // through the public path - but the merge is not dead code: it is the safety net for a
    // destination folder that appears between the id being allocated and the bytes being moved
    // (another process, a restored backup, a sync client), and deleting it because today's caller
    // cannot trigger it would remove the handling for the day something can. Same reasoning, and
    // the same technique, as ExternalTargetLauncherTests uses for its private predicates.
    private static void InvokeMoveEntryAttachmentsFolder(string workbookFolder, int fromEntryId, int toEntryId)
    {
        var method = typeof(WorklogManager).GetMethod(
            "MoveEntryAttachmentsFolder",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);

        Assert.NotNull(method);
        method!.Invoke(null, new object[] { workbookFolder, fromEntryId, toEntryId });
    }

    [Fact]
    public void Moving_a_draft_folder_merges_into_an_existing_destination_rather_than_stranding_it()
    {
        // A destination folder that already exists is a real case, not a corruption. Skipping the
        // move there strands the draft's bytes in a folder no entry names, while the entry itself
        // is committed naming files that are not in its own folder - a photo row broken forever,
        // with only a log line nobody reads. So the two folders are MERGED.
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        string workbookFolder = Path.Combine(root, workbook.Id.ToString());

        string from = WorklogManager.GetEntryAttachmentsFolder(workbook.Id, 1)!;
        File.WriteAllText(Path.Combine(from, "photo_1_draft.png"), "draft bytes");

        string to = WorklogManager.GetEntryAttachmentsFolder(workbook.Id, 2)!;
        File.WriteAllText(Path.Combine(to, "photo_9_orphan.png"), "orphan bytes");

        InvokeMoveEntryAttachmentsFolder(workbookFolder, 1, 2);

        // The draft's own bytes made it across - the whole point.
        Assert.True(
            File.Exists(Path.Combine(to, "photo_1_draft.png")),
            "the draft's photo was stranded because the destination folder already existed");

        // And the orphan already there was not destroyed on the way.
        Assert.True(File.Exists(Path.Combine(to, "photo_9_orphan.png")));

        Assert.False(Directory.Exists(from), "the emptied source folder was left behind");
    }

    [Fact]
    public void Moving_a_draft_folder_does_not_overwrite_a_colliding_name()
    {
        // The filenames cannot normally collide across the two folders - they are built from the
        // owning record's own attachment ids - but a name that collides anyway is left alone rather
        // than overwritten: the file already there is the one some record may still name.
        string root = this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");
        string workbookFolder = Path.Combine(root, workbook.Id.ToString());

        string from = WorklogManager.GetEntryAttachmentsFolder(workbook.Id, 1)!;
        File.WriteAllText(Path.Combine(from, "photo_1_shared.png"), "draft bytes");

        string to = WorklogManager.GetEntryAttachmentsFolder(workbook.Id, 2)!;
        File.WriteAllText(Path.Combine(to, "photo_1_shared.png"), "existing bytes");

        InvokeMoveEntryAttachmentsFolder(workbookFolder, 1, 2);

        Assert.Equal("existing bytes", File.ReadAllText(Path.Combine(to, "photo_1_shared.png")));

        // The one that could not be moved stays put rather than being silently dropped, and its
        // folder survives with it.
        Assert.True(File.Exists(Path.Combine(from, "photo_1_shared.png")));
    }

    // AddEntry Trim()s Title and Description; AddEntryRecord must too, or an entry created through
    // the "Add worklog" flow persists padding that the same entry edited later would have lost. The
    // editor's Save gate is IsNullOrWhiteSpace, so "  CPU socket  " gets through it.
    [Fact]
    public void AddEntryRecord_trims_the_title_and_description()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        var saved = WorklogManager.AddEntryRecord(
            workbook.Id,
            new WorklogEntryRecord
            {
                Id = 1,
                SchematicName = "Sheet 1",
                Title = "  CPU socket cold joint  ",
                Description = "\tReflowed pin 12\n"
            },
            reservedId: 1);

        Assert.NotNull(saved);
        Assert.Equal("CPU socket cold joint", saved!.Title);
        Assert.Equal("Reflowed pin 12", saved.Description);

        // And it is the TRIMMED form that reached disk, not just the returned copy.
        var reloaded = WorklogManager.GetEntries(workbook.Id).Single();
        Assert.Equal("CPU socket cold joint", reloaded.Title);
        Assert.Equal("Reflowed pin 12", reloaded.Description);
    }

    // Writing an entry recomputes the workbook's Open/Closed status, exactly as AddEntry does -
    // a Closed-only workbook must close, and the two write paths must not disagree about that.
    [Fact]
    public void AddEntryRecord_recomputes_the_workbooks_status()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        WorklogManager.AddEntryRecord(
            workbook.Id,
            new WorklogEntryRecord { Id = 1, SchematicName = "Sheet 1", Title = "Done", State = "Closed" },
            reservedId: 1);

        var reloaded = WorklogManager.GetLatestWorkbookForBoard("Commodore 64|250469")!;
        Assert.Equal("Closed", reloaded.Status);
        Assert.Equal(1, reloaded.EntryCount);
    }

    // No workbook folder means nothing was persisted, and the caller must be told so rather than
    // being handed a record it would then treat as saved - the same contract AddEntry has.
    [Fact]
    public void AddEntryRecord_returns_null_when_the_workbook_does_not_exist()
    {
        this.LoadWorklog();

        Assert.Null(WorklogManager.AddEntryRecord(
            999,
            new WorklogEntryRecord { Id = 1, SchematicName = "Sheet 1", Title = "Orphan" },
            reservedId: 1));
    }

    [Fact]
    public void AddEntryRecord_returns_null_for_a_null_entry()
    {
        this.LoadWorklog();
        var workbook = CreateWorkbook("Commodore 64|250469", "C64 job", "");

        Assert.Null(WorklogManager.AddEntryRecord(workbook.Id, null!, reservedId: 1));
    }

    // ###########################################################################################
    // "--workbooks-root=" - the same idea as DataManager.ResolveDataRoot's "--data-root=", but for
    // the purely-local Workbooks folder. ResolveExplicitWorkbookRoot is pure (no statics touched, no
    // AppData read), so it is safe to call directly without going through Load().
    // ###########################################################################################
    [Fact]
    public void ResolveExplicitWorkbookRoot_uses_the_workbooks_root_command_line_argument()
    {
        Assert.Equal(
            @"D:\somewhere\Workbooks",
            WorklogManager.ResolveExplicitWorkbookRoot(new[] { @"--workbooks-root=D:\somewhere\Workbooks" }));
    }

    [Fact]
    public void ResolveExplicitWorkbookRoot_strips_surrounding_quotes()
    {
        Assert.Equal(
            @"D:\my data\Workbooks",
            WorklogManager.ResolveExplicitWorkbookRoot(new[] { "--workbooks-root=\"D:\\my data\\Workbooks\"" }));
    }

    [Fact]
    public void ResolveExplicitWorkbookRoot_matches_the_argument_case_insensitively()
    {
        Assert.Equal("X", WorklogManager.ResolveExplicitWorkbookRoot(new[] { "--WORKBOOKS-ROOT=X" }));
    }

    [Fact]
    public void ResolveExplicitWorkbookRoot_takes_the_first_matching_argument()
    {
        Assert.Equal(
            "first",
            WorklogManager.ResolveExplicitWorkbookRoot(new[] { "--workbooks-root=first", "--workbooks-root=second" }));
    }

    [Fact]
    public void ResolveExplicitWorkbookRoot_ignores_unrelated_arguments()
    {
        Assert.Equal(
            "X",
            WorklogManager.ResolveExplicitWorkbookRoot(new[] { "--verbose", "--workbooks-root=X", "--data-root=Y" }));
    }

    [Fact]
    public void ResolveExplicitWorkbookRoot_returns_null_with_no_matching_argument()
    {
        Assert.Null(WorklogManager.ResolveExplicitWorkbookRoot(new[] { "--data-root=X" }));
    }

    [Fact]
    public void ResolveExplicitWorkbookRoot_returns_null_with_no_arguments_at_all()
    {
        Assert.Null(WorklogManager.ResolveExplicitWorkbookRoot(null));
        Assert.Null(WorklogManager.ResolveExplicitWorkbookRoot(Array.Empty<string>()));
    }

}
