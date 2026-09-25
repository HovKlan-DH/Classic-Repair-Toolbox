using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Copies one file into a saved entry's attachments folder and records it - the write half of
    // "add a photo/file to a worklog entry", with no dialog and no UI attached.
    //
    // This exists because there are now TWO ways to attach a photo: the full editor's own Photos
    // section, and the oscilloscope capture's "Attach image to worklog" flow, which files into an
    // entry that is not open in any editor. The sequence below has four separate subtleties that
    // were each fixed once already in the editor's copy:
    //
    //  - the id is allocated through AllocateAttachmentId, which SKIPS ids whose file is already in
    //    the folder, because plain Max(Id) + 1 silently overwrites an orphaned attachment;
    //  - the id is settled BEFORE the name, because the stored name is built from it;
    //  - DisplayOrder is 0-based to match ReorderAttachment's dense renumbering, or the first
    //    attachment added after any drag-reorder collides with an existing row's order;
    //  - a failed persist rolls the copied bytes back out, or the folder keeps a file that
    //    entries.json never mentions.
    //
    // Duplicating all four in the capture flow would mean fixing the next one twice, so the editor
    // now calls this too rather than keeping its own copy.
    //
    // Deliberately takes the records list and a persist callback rather than a workbook id: the
    // editor attaches to an entry it holds in memory and persists via its own PersistEntrySilently
    // (which knows about draft entries), while the capture flow attaches to an entry read fresh off
    // disk. Both shapes fit this; a workbook-id-only API would not fit the editor's.
    // ###########################################################################################
    public static class WorklogAttachmentWriter
    {
        // What an attach attempt did. Anything other than Added leaves the folder and the records
        // list exactly as they were, so a caller can report the failure without cleaning up.
        public enum AttachOutcome
        {
            Added,
            NoAttachmentsFolder,
            CopyFailed,
            PersistFailed,

            // Only AttachToEntry can report this: the entry the id names is not in the workbook.
            // Kept distinct from NoAttachmentsFolder, which it used to be reported as - the two send
            // whoever reads the log to entirely different places, one to check folder permissions
            // and the other to a worklog that simply no longer exists (ShowDialog does not block the
            // dispatcher, so another window can have deleted it since the dialog was opened).
            EntryNotFound
        }

        // ###########################################################################################
        // Copies sourcePath into attachmentsFolder under a generated name, appends a record for it
        // to records, then calls persist. Rolls both back when persist returns false.
        //
        // ownerPrefix is WorklogAttachmentStorage.PhotoFilePrefix or FileFilePrefix - it separates
        // the two id sequences, which are numbered independently within one entry.
        //
        // addedRecord is the record that was appended, so a caller can refresh a row for it; it is
        // null for every outcome other than Added.
        // ###########################################################################################
        public static AttachOutcome Attach(
            string sourcePath,
            string? attachmentsFolder,
            List<WorklogAttachmentRecord> records,
            string ownerPrefix,
            string comment,
            Func<bool> persist,
            out WorklogAttachmentRecord? addedRecord)
        {
            addedRecord = null;

            if (string.IsNullOrWhiteSpace(attachmentsFolder))
            {
                return AttachOutcome.NoAttachmentsFolder;
            }

            // See the header: id before name, and skipping ids whose bytes are already on disk.
            int nextId = WorklogAttachmentStorage.AllocateAttachmentId(
                records,
                ownerPrefix,
                WorklogAttachmentStorage.ListAttachmentFileNames(attachmentsFolder));

            int nextOrder = records.Count == 0 ? 0 : records.Max(r => r.DisplayOrder) + 1;

            string storedFileName = WorklogAttachmentStorage.BuildStoredFileName(
                sourcePath, ownerPrefix, nextId);

            if (!WorklogAttachmentStorage.CopyAttachmentIntoFolder(sourcePath, attachmentsFolder, storedFileName))
            {
                return AttachOutcome.CopyFailed;
            }

            var record = new WorklogAttachmentRecord
            {
                Id = nextId,
                FileName = storedFileName,
                Comment = comment ?? string.Empty,
                DisplayOrder = nextOrder
            };

            records.Add(record);

            if (persist != null && !persist())
            {
                records.RemoveAll(r => r.Id == nextId);
                WorklogAttachmentStorage.DeleteAttachmentFileAndFolderIfEmpty(attachmentsFolder, storedFileName);
                return AttachOutcome.PersistFailed;
            }

            addedRecord = record;
            return AttachOutcome.Added;
        }

        // ###########################################################################################
        // The same attach, against an entry identified by id rather than one held in memory - the
        // shape the oscilloscope capture flow needs, where nothing has the entry open.
        //
        // Reads the entry fresh, appends to it and writes the whole record back through
        // WorklogManager.UpdateEntry. Reading it here rather than taking one from the caller is
        // deliberate: the editor can be open on this same entry (ShowDialog does not block the
        // dispatcher), so a record captured earlier could be stale, and writing that back would drop
        // whatever the editor has since saved.
        // ###########################################################################################
        public static AttachOutcome AttachToEntry(
            int workbookId,
            int entryId,
            string sourcePath,
            string ownerPrefix,
            string comment)
        {
            var entry = WorklogManager.GetEntries(workbookId)
                .FirstOrDefault(candidate => candidate.Id == entryId);

            if (entry == null)
            {
                return AttachOutcome.EntryNotFound;
            }

            string? attachmentsFolder = WorklogManager.GetEntryAttachmentsFolder(workbookId, entryId);

            var records = string.Equals(ownerPrefix, WorklogAttachmentStorage.FileFilePrefix, StringComparison.Ordinal)
                ? entry.Files
                : entry.Photos;

            return Attach(
                sourcePath,
                attachmentsFolder,
                records,
                ownerPrefix,
                comment,
                () => WorklogManager.UpdateEntry(workbookId, entry),
                out _);
        }
    }
}
