using Avalonia;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Covers the shared attach-write path used by BOTH the full editor's Photos section and the
// oscilloscope capture's "Attach image to worklog" flow.
//
// It was extracted precisely because it carries four subtleties that are each invisible when wrong:
// the id must skip ids whose bytes are already in the folder (or an orphan is overwritten), the id
// must be settled before the name (which is built from it), DisplayOrder must be 0-based (or the
// first attachment after a drag-reorder collides), and a failed persist must roll the copied bytes
// back out (or the folder keeps a file entries.json never mentions). Each has a test below.
//
// Touches WorklogManager's static root, so this joins the "Worklog" collection.
[Collection("Worklog")]
public sealed class WorklogAttachmentWriterTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose()
    {
        // Detach from the temp folder so nothing written later can reach the user's real one.
        this.LoadWorklog();
        this.thisWorkspace.Dispose();
    }

    private string LoadWorklog()
    {
        string root = this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N"));
        WorklogManager.LoadFrom(root);
        return root;
    }

    /// <summary>A workbook with one entry, plus a source image file on disk to attach.</summary>
    private (int WorkbookId, int EntryId, string SourcePath) CreateWorkbookWithEntry()
    {
        this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook("Commodore 64|250469", "Test workbook", "");
        Assert.NotNull(workbook);

        var entry = WorklogManager.AddEntry(
            workbook!.Id, "Sch", new Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", new[] { "U8" });
        Assert.NotNull(entry);

        string sourcePath = this.thisWorkspace.Path_("capture.png");
        File.WriteAllBytes(sourcePath, new byte[] { 1, 2, 3 });

        return (workbook.Id, entry!.Id, sourcePath);
    }

    // The happy path, end to end through the id-addressed overload the capture flow uses: the bytes
    // land in the entry's own attachments folder and entries.json records them.
    [Fact]
    public void Attaching_to_an_entry_copies_the_bytes_and_records_the_photo()
    {
        var (workbookId, entryId, sourcePath) = this.CreateWorkbookWithEntry();

        var outcome = WorklogAttachmentWriter.AttachToEntry(
            workbookId, entryId, sourcePath, WorklogAttachmentStorage.PhotoFilePrefix, "Pin 14 at 5V");

        Assert.Equal(WorklogAttachmentWriter.AttachOutcome.Added, outcome);

        var stored = WorklogManager.GetEntries(workbookId).Single(e => e.Id == entryId);
        var photo = Assert.Single(stored.Photos);

        Assert.Equal("Pin 14 at 5V", photo.Comment);

        string folder = WorklogManager.GetEntryAttachmentsFolderPath(workbookId, entryId)!;
        Assert.True(File.Exists(Path.Combine(folder, photo.FileName)));
    }

    // Photos and files are numbered independently within one entry, so attaching a file must not
    // consume the photo sequence's next id.
    [Fact]
    public void Photos_and_files_are_numbered_independently()
    {
        var (workbookId, entryId, sourcePath) = this.CreateWorkbookWithEntry();

        WorklogAttachmentWriter.AttachToEntry(
            workbookId, entryId, sourcePath, WorklogAttachmentStorage.PhotoFilePrefix, "");
        WorklogAttachmentWriter.AttachToEntry(
            workbookId, entryId, sourcePath, WorklogAttachmentStorage.FileFilePrefix, "");

        var stored = WorklogManager.GetEntries(workbookId).Single(e => e.Id == entryId);

        Assert.Equal(1, Assert.Single(stored.Photos).Id);
        Assert.Equal(1, Assert.Single(stored.Files).Id);
    }

    // DisplayOrder is 0-based to match ReorderAttachment's dense renumbering from 0. When this
    // started at 1, the first attachment added after any drag-reorder took the same DisplayOrder as
    // an existing row, and two rows sharing an order sort arbitrarily.
    [Fact]
    public void Display_order_starts_at_zero_and_increments()
    {
        var (workbookId, entryId, sourcePath) = this.CreateWorkbookWithEntry();

        WorklogAttachmentWriter.AttachToEntry(
            workbookId, entryId, sourcePath, WorklogAttachmentStorage.PhotoFilePrefix, "first");
        WorklogAttachmentWriter.AttachToEntry(
            workbookId, entryId, sourcePath, WorklogAttachmentStorage.PhotoFilePrefix, "second");

        var stored = WorklogManager.GetEntries(workbookId).Single(e => e.Id == entryId);

        Assert.Equal(new[] { 0, 1 }, stored.Photos.OrderBy(p => p.Id).Select(p => p.DisplayOrder).ToArray());
    }

    // The orphan case AllocateAttachmentId exists for: a file already sitting in the folder under
    // id 1, with nothing in entries.json referencing it. Plain Max(Id) + 1 would hand out id 1 again
    // and overwrite those bytes. Asserted by the ORPHAN surviving, which is the thing at risk.
    [Fact]
    public void An_orphaned_attachment_file_is_never_overwritten()
    {
        var (workbookId, entryId, sourcePath) = this.CreateWorkbookWithEntry();

        string folder = WorklogManager.GetEntryAttachmentsFolder(workbookId, entryId)!;
        string orphanPath = Path.Combine(folder, "photo_1_orphan.png");
        File.WriteAllBytes(orphanPath, new byte[] { 9, 9, 9 });

        var outcome = WorklogAttachmentWriter.AttachToEntry(
            workbookId, entryId, sourcePath, WorklogAttachmentStorage.PhotoFilePrefix, "");

        Assert.Equal(WorklogAttachmentWriter.AttachOutcome.Added, outcome);

        var photo = Assert.Single(WorklogManager.GetEntries(workbookId).Single(e => e.Id == entryId).Photos);
        Assert.NotEqual(1, photo.Id);

        Assert.Equal(new byte[] { 9, 9, 9 }, File.ReadAllBytes(orphanPath));
    }

    // A failed persist must leave NOTHING behind: no record in the list, and no copied bytes in the
    // folder. Otherwise the folder keeps a file entries.json never mentions, which nothing will ever
    // clean up. Driven through the list-addressed overload, the only one that can be handed a
    // deliberately failing persist.
    [Fact]
    public void A_failed_persist_rolls_the_copied_bytes_back_out()
    {
        var (workbookId, entryId, sourcePath) = this.CreateWorkbookWithEntry();

        string folder = WorklogManager.GetEntryAttachmentsFolder(workbookId, entryId)!;
        var records = new List<WorklogAttachmentRecord>();

        var outcome = WorklogAttachmentWriter.Attach(
            sourcePath,
            folder,
            records,
            WorklogAttachmentStorage.PhotoFilePrefix,
            "",
            () => false,
            out var added);

        Assert.Equal(WorklogAttachmentWriter.AttachOutcome.PersistFailed, outcome);
        Assert.Null(added);
        Assert.Empty(records);

        // The rollback goes through DeleteAttachmentFileAndFolderIfEmpty, which removes the folder
        // itself once its last file is gone - so "no bytes left behind" is satisfied by the folder
        // being empty OR by it no longer existing, and this attach created it.
        Assert.True(!Directory.Exists(folder) || Directory.GetFiles(folder).Length == 0);
    }

    // A missing attachments folder is reported rather than throwing, so the caller can say so in its
    // own UI. Nothing is copied and nothing is recorded.
    [Fact]
    public void A_missing_attachments_folder_is_reported_and_writes_nothing()
    {
        var (_, _, sourcePath) = this.CreateWorkbookWithEntry();
        var records = new List<WorklogAttachmentRecord>();

        var outcome = WorklogAttachmentWriter.Attach(
            sourcePath, null, records, WorklogAttachmentStorage.PhotoFilePrefix, "", () => true, out var added);

        Assert.Equal(WorklogAttachmentWriter.AttachOutcome.NoAttachmentsFolder, outcome);
        Assert.Null(added);
        Assert.Empty(records);
    }

    // An entry the id does not name is its OWN outcome, not the missing-folder one it used to be
    // reported as. The two send whoever reads the log to entirely different places - one to check
    // folder permissions, the other to a worklog that no longer exists - and this path is reachable
    // in normal use: ShowDialog does not block the dispatcher, so the entry can be deleted from
    // another window while the attach dialog is still open on it.
    [Fact]
    public void Attaching_to_an_entry_that_does_not_exist_is_reported_as_the_entry_being_missing()
    {
        var (workbookId, entryId, sourcePath) = this.CreateWorkbookWithEntry();

        var outcome = WorklogAttachmentWriter.AttachToEntry(
            workbookId, entryId + 99, sourcePath, WorklogAttachmentStorage.PhotoFilePrefix, "");

        Assert.Equal(WorklogAttachmentWriter.AttachOutcome.EntryNotFound, outcome);

        // Specifically NOT the folder outcome, which is what it used to return - asserted
        // separately, since the point of the change is the two being distinguishable.
        Assert.NotEqual(WorklogAttachmentWriter.AttachOutcome.NoAttachmentsFolder, outcome);

        // The entry that DOES exist is untouched by the failed attach.
        Assert.Empty(WorklogManager.GetEntries(workbookId).Single(e => e.Id == entryId).Photos);
    }

    // The same, for an entry in a workbook that does not exist at all - GetEntries returns nothing,
    // so there is no entry to find and the outcome is the same one.
    [Fact]
    public void Attaching_into_a_workbook_that_does_not_exist_is_reported_as_the_entry_being_missing()
    {
        var (workbookId, entryId, sourcePath) = this.CreateWorkbookWithEntry();

        var outcome = WorklogAttachmentWriter.AttachToEntry(
            workbookId + 99, entryId, sourcePath, WorklogAttachmentStorage.PhotoFilePrefix, "");

        Assert.Equal(WorklogAttachmentWriter.AttachOutcome.EntryNotFound, outcome);
    }

    // A source file that is not there cannot be copied, and that must be reported rather than
    // recorded - a record naming bytes that never arrived renders as a broken attachment forever.
    [Fact]
    public void A_missing_source_file_fails_the_copy_and_records_nothing()
    {
        var (workbookId, entryId, _) = this.CreateWorkbookWithEntry();

        string folder = WorklogManager.GetEntryAttachmentsFolder(workbookId, entryId)!;
        var records = new List<WorklogAttachmentRecord>();

        var outcome = WorklogAttachmentWriter.Attach(
            this.thisWorkspace.Path_("does-not-exist.png"),
            folder,
            records,
            WorklogAttachmentStorage.PhotoFilePrefix,
            "",
            () => true,
            out var added);

        Assert.Equal(WorklogAttachmentWriter.AttachOutcome.CopyFailed, outcome);
        Assert.Null(added);
        Assert.Empty(records);
    }

    // The id-addressed overload re-reads the entry rather than trusting a caller's copy, because the
    // full editor can be open on that same entry (ShowDialog does not block the dispatcher) and a
    // stale record written back would drop whatever it has since saved. Pinned by attaching AFTER an
    // independent update: the update must survive.
    [Fact]
    public void Attaching_reads_the_entry_fresh_so_a_concurrent_edit_is_not_lost()
    {
        var (workbookId, entryId, sourcePath) = this.CreateWorkbookWithEntry();

        var edited = WorklogManager.GetEntries(workbookId).Single(e => e.Id == entryId);
        edited.Title = "Retitled elsewhere";
        Assert.True(WorklogManager.UpdateEntry(workbookId, edited));

        WorklogAttachmentWriter.AttachToEntry(
            workbookId, entryId, sourcePath, WorklogAttachmentStorage.PhotoFilePrefix, "");

        var stored = WorklogManager.GetEntries(workbookId).Single(e => e.Id == entryId);

        Assert.Equal("Retitled elsewhere", stored.Title);
        Assert.Single(stored.Photos);
    }
}
