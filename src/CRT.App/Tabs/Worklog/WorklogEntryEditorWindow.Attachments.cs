using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // ATTACHMENTS: the Photos and Files sub-lists, which are one implementation parameterised by
    // AttachmentSection (below) - add, edit, delete, thumbnail loading, and pointer-drag reordering
    // (DragContext and the placeholder-based drag machinery). See WorklogEntryEditorWindow.axaml.cs
    // for the file map of the whole partial class.
    // ###########################################################################################
    public partial class WorklogEntryEditorWindow
    {
        // ###########################################################################################
        // ATTACHMENTS - Photos and Files.
        //
        // Both lists have the same storage shape: metadata (file name, comment, order) lives in
        // entries.json with the entry, while the bytes live in the entry's own "worklog_<id>"
        // folder, resolved through WorklogManager.GetEntryAttachmentsFolder. Adding copies the
        // chosen file in there under a name that cannot collide with an existing one - see
        // WorklogAttachmentStorage.
        //
        // ONE implementation serves both, parameterised by the AttachmentSection below, for the
        // same reason the drag-reorder further down is shared: the add/edit/delete paths each
        // encode an ORDERING RULE learned from a real fault, and a second copy is a second place
        // for one of them to be got wrong. The three rules, in the order they must happen:
        //
        //   Add    - copy the bytes, record the metadata, persist, and UNDO THE COPY if the
        //            persist failed. Otherwise the folder holds bytes entries.json never mentions.
        //   Edit   - persist the record BEFORE swapping the file, because the swap deletes what it
        //            replaces. The other order left entries.json naming a file already deleted,
        //            and Cancel then discarded the working copy - the row was permanently broken
        //            with no way back to the original.
        //   Delete - persist BEFORE removing the bytes. Deleting first leaves a still-listed row
        //            pointing at nothing when the save fails.
        //
        // These were written out twice, and had already begun to drift: the file-side copies
        // carried "see the photo path" comments rather than the reasoning itself, which is the
        // admission that one was the original and the other a transcription.
        //
        // The two lists differ on exactly the axes this record names - which records, which rows,
        // which stored-name prefix, which controls, and the noun used in messages. Everything else
        // was identical. Photos additionally carry a decoded thumbnail; that is the ONE branch in
        // the shared code (Thumbnail below), rather than a reason to keep two of everything.
        //
        // Records and Rows are lookups, not captured collections: anything replacing the thisEntry
        // record (rather than mutating its lists in place) would otherwise leave a stale list
        // behind - the same staleness the drag context's own comment describes.
        // ###########################################################################################
        private sealed record AttachmentSection(
            WorklogAttachmentStorage.AttachmentKind Kind,
            string OwnerPrefix,
            Func<List<WorklogAttachmentRecord>> Records,
            Func<ObservableCollection<WorklogAttachmentRow>> Rows,
            Func<ItemsControl> List,
            Func<TextBlock> EmptyText,
            Func<TextBlock> CountText,
            string HeaderKey,
            string Singular,
            string Plural)
        {
            // The noun as it appears mid-sentence in a message ("The photo could not be copied").
            public string Noun => this.Singular;
        }

        private AttachmentSection PhotoAttachments => new(
            WorklogAttachmentStorage.AttachmentKind.Photo,
            WorklogAttachmentStorage.PhotoFilePrefix,
            () => this.thisEntry.Photos,
            () => this.thisPhotoRows,
            () => this.EditorPhotosList,
            () => this.EditorNoPhotosText,
            () => this.EditorPhotosCountText,
            "EditorPhotosHeader",
            "photo",
            "photos");

        private AttachmentSection FileAttachments => new(
            WorklogAttachmentStorage.AttachmentKind.File,
            WorklogAttachmentStorage.FileFilePrefix,
            () => this.thisEntry.Files,
            () => this.thisFileRows,
            () => this.EditorFilesList,
            () => this.EditorNoFilesText,
            () => this.EditorFilesCountText,
            "EditorFilesHeader",
            "file",
            "files");

        // ###########################################################################################
        // Test seam for the two attachment sections.
        //
        // What is worth pinning here is not that a record has the fields it was declared with, but
        // that each section is wired to the RIGHT ONES - a copy/paste slip pointing the Files
        // section at thisEntry.Photos, or at the photo prefix, compiles perfectly and silently
        // makes the two lists share a set of records or a naming scheme. That is exactly the class
        // of fault unifying the two implementations could introduce, so it is the thing to assert.
        //
        // Returns the section's wiring already resolved to values, so the test needs no access to
        // the record type itself.
        // ###########################################################################################
        internal (string Prefix, string HeaderKey, string Singular, string Plural, bool IsPhotoKind,
            int RecordCount, int RowCount, ItemsControl List, TextBlock EmptyText, TextBlock CountText)
            DescribeAttachmentSectionForTests(bool photos)
        {
            var section = photos ? this.PhotoAttachments : this.FileAttachments;

            return (
                section.OwnerPrefix,
                section.HeaderKey,
                section.Singular,
                section.Plural,
                section.Kind == WorklogAttachmentStorage.AttachmentKind.Photo,
                section.Records().Count,
                section.Rows().Count,
                section.List(),
                section.EmptyText(),
                section.CountText());
        }

        // Drives the shared row rebuild for one list, so a test can prove the SAME method serves
        // both sections - including the thumbnail branch, which is the one place they differ.
        // ###########################################################################################
        // The marked area on the window's WORKING COPY - which is what the tick handler edits, and
        // what Save writes back. Initialize deliberately clones the caller's record (see CloneEntry)
        // so Cancel cannot mutate it, so a test reading the record it passed in would see nothing
        // change no matter what the editor did.
        // ###########################################################################################
        internal Rect WorkingEntryAreaForTests =>
            new(this.thisEntry.AreaX, this.thisEntry.AreaY, this.thisEntry.AreaWidth, this.thisEntry.AreaHeight);

        internal void RefreshAttachmentRowsForTests(bool photos) =>
            this.RefreshAttachmentRows(photos ? this.PhotoAttachments : this.FileAttachments);

        // Appends a record directly to one section's list, bypassing the file copy an add would do.
        // The rebuild is what is under test here, not the storage.
        internal void AddAttachmentRecordForTests(bool photos, int id, string fileName, string comment)
        {
            var section = photos ? this.PhotoAttachments : this.FileAttachments;
            section.Records().Add(new WorklogAttachmentRecord
            {
                Id = id,
                FileName = fileName,
                Comment = comment,
                DisplayOrder = section.Records().Count
            });
        }

        // ###########################################################################################
        // Rebuilds one attachment list's rows from its records.
        //
        // The thumbnail disposal is the photo-only half, and it matters: each thumbnail is a decoded
        // Bitmap holding an unmanaged surface, and this runs on every add/edit/delete/reorder.
        // Without disposing the old ones each refresh orphaned a full set until a finalizer
        // eventually ran. They are collected BEFORE Clear() but disposed AFTER it - an Image is
        // still bound to the bitmap until the row leaves the collection, and disposing one out from
        // under a live binding risks a render against a freed surface.
        //
        // Files never load a thumbnail, so for them the collected list is empty and this costs a
        // no-op LINQ pass.
        // ###########################################################################################
        private void RefreshAttachmentRows(AttachmentSection section)
        {
            var rows = section.Rows();

            var discardedThumbnails = rows.Select(row => row.Thumbnail).Where(bitmap => bitmap != null).ToList();

            rows.Clear();

            foreach (var bitmap in discardedThumbnails)
            {
                bitmap!.Dispose();
            }

            bool wantsThumbnail = section.Kind == WorklogAttachmentStorage.AttachmentKind.Photo;

            foreach (var record in section.Records().OrderBy(r => r.DisplayOrder))
            {
                rows.Add(new WorklogAttachmentRow
                {
                    Id = record.Id,
                    FileName = record.FileName,
                    DisplayFileName = WorklogAttachmentStorage.GetDisplayFileName(record.FileName, section.OwnerPrefix, record.Id),
                    Comment = record.Comment,
                    Thumbnail = wantsThumbnail ? this.TryLoadPhotoThumbnail(record.FileName) : null
                });
            }

            section.EmptyText().IsVisible = rows.Count == 0 && this.IsListSectionExpanded(section.HeaderKey);
            section.CountText().Text = FormatItemCount(rows.Count, section.Singular, section.Plural);
        }

        private void RefreshPhotoRows() => this.RefreshAttachmentRows(this.PhotoAttachments);

        private void RefreshFileRows() => this.RefreshAttachmentRows(this.FileAttachments);

        // ###########################################################################################
        // Resolves the on-disk path of one of this entry's attachments, or null when the workbook
        // folder cannot be resolved or the file is not there.
        // ###########################################################################################
        private string? ResolveAttachmentPath(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return null;
            }

            string? attachmentsFolder = WorklogManager.GetEntryAttachmentsFolder(this.thisWorkbookId, this.thisEntry.Id);
            if (attachmentsFolder == null)
            {
                return null;
            }

            // Fully qualified: this file also uses Avalonia.Controls.Shapes, which has its own Path.
            string path = System.IO.Path.Combine(attachmentsFolder, fileName);
            return File.Exists(path) ? path : null;
        }

        // ###########################################################################################
        // Decodes a row thumbnail, scaled down on load rather than at full resolution - a phone
        // photo is several thousand pixels wide and the row shows it at 64, so decoding the full
        // image would spend memory the list never uses. Failure is not fatal: the row renders with
        // a "missing" marker instead, since a photo file can be deleted or corrupted outside the app.
        // ###########################################################################################
        private Bitmap? TryLoadPhotoThumbnail(string fileName)
        {
            string? path = this.ResolveAttachmentPath(fileName);
            if (path == null)
            {
                return null;
            }

            try
            {
                using var stream = File.OpenRead(path);
                return Bitmap.DecodeToWidth(stream, 256);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load worklog photo thumbnail [{fileName}]: {ex.Message}");
                return null;
            }
        }

        // ###########################################################################################
        // Adds one attachment: collect the file and comment, copy the bytes into the entry's
        // attachments folder, then record the metadata. The record is only added once the copy has
        // succeeded - a row pointing at a file that never landed would show as permanently broken -
        // and the copy is undone if the metadata then fails to persist. See the ATTACHMENTS header
        // for why that ordering is not negotiable.
        // ###########################################################################################
        private async void OnAddPhotoClick(object? sender, RoutedEventArgs e)
        {
            // async void cannot be awaited, so anything thrown after the first await reaches the
            // global handler instead of this window. GetEntryAttachmentsFolder calls
            // Directory.CreateDirectory, which throws on a read-only or disconnected folder - a
            // reportable condition, not a crash.
            await this.AddAttachmentGuardedAsync(this.PhotoAttachments);
        }

        private async void OnAddFileClick(object? sender, RoutedEventArgs e)
        {
            await this.AddAttachmentGuardedAsync(this.FileAttachments);
        }

        private async Task AddAttachmentGuardedAsync(AttachmentSection section)
        {
            try
            {
                await this.AddAttachmentAsync(section);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to add worklog {section.Noun}: {ex.Message}");
                this.ShowSaveFailed($"The {section.Noun} could not be added - see the log for details.");
            }
        }

        private async Task AddAttachmentAsync(AttachmentSection section)
        {
            var dialog = new WorklogAddPhotoWindow();
            if (section.Kind == WorklogAttachmentStorage.AttachmentKind.File)
            {
                dialog.InitializeForFileKind();
            }

            var result = await dialog.ShowDialog<WorklogAddPhotoWindow.PhotoResult?>(this);
            if (result == null || string.IsNullOrWhiteSpace(result.SourcePath))
            {
                return;
            }

            string? attachmentsFolder = WorklogManager.GetEntryAttachmentsFolder(this.thisWorkbookId, this.thisEntry.Id);
            if (attachmentsFolder == null)
            {
                this.ShowSaveFailed($"Could not resolve where to store the {section.Noun}.");
                return;
            }

            var records = section.Records();

            // The id/name/order allocation and the roll-back-on-failed-persist all live in
            // WorklogAttachmentWriter, which the oscilloscope capture's attach flow calls too - see
            // its header for the four subtleties that used to be written out here. The rows are
            // refreshed after the write either way, since a rollback removes the record again.
            var outcome = WorklogAttachmentWriter.Attach(
                result.SourcePath,
                attachmentsFolder,
                records,
                section.OwnerPrefix,
                result.Comment,
                this.PersistEntrySilently,
                out _);

            if (outcome == WorklogAttachmentWriter.AttachOutcome.CopyFailed)
            {
                this.ShowSaveFailed($"The {section.Noun} could not be copied into the worklog.");
                return;
            }

            this.EnsureListSectionExpanded(section.HeaderKey);
            this.RefreshAttachmentRows(section);
        }

        // ###########################################################################################
        // Clicking a photo row's thumbnail opens the full-size viewer. Separate from the row's Edit
        // button on purpose: viewing is the common action and editing is the deliberate one, so the
        // large target views and the small explicit one edits.
        // ###########################################################################################
        private void OnPhotoThumbnailPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is not Control { Tag: int id })
            {
                return;
            }

            // Left button only, so a right-click does not open the viewer.
            if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            {
                return;
            }

            var photo = this.thisEntry.Photos.FirstOrDefault(p => p.Id == id);
            if (photo == null)
            {
                return;
            }

            // Stops the click also reaching the row, which would open the editor behind the viewer.
            e.Handled = true;

            // The display name, not the stored one - the user never chose the "photo3_" prefix.
            var viewer = new WorklogPhotoViewerWindow();
            viewer.Initialize(
                WorklogAttachmentStorage.GetDisplayFileName(photo.FileName, WorklogAttachmentStorage.PhotoFilePrefix, photo.Id),
                photo.Comment,
                this.ResolveAttachmentPath(photo.FileName));
            viewer.ShowDialog(this);
        }

        // ###########################################################################################
        // Editing an attachment reopens the same modal pre-filled, matching the comment and
        // work-done rows. A replacement file is copied in alongside the old one and the record
        // repointed; the previous file is deliberately left on disk rather than deleted, because an
        // entry that has not been saved yet can still be cancelled, and deleting here would take
        // the original with it. See the note on Delete below - the same reasoning applies.
        // ###########################################################################################
        private async void OnEditPhotoClick(object? sender, RoutedEventArgs e)
        {
            // See OnAddPhotoClick for why the body is wrapped.
            await this.EditAttachmentGuardedAsync(this.PhotoAttachments, sender);
        }

        private async void OnEditFileClick(object? sender, RoutedEventArgs e)
        {
            await this.EditAttachmentGuardedAsync(this.FileAttachments, sender);
        }

        private async Task EditAttachmentGuardedAsync(AttachmentSection section, object? sender)
        {
            try
            {
                await this.EditAttachmentAsync(section, sender);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to edit worklog {section.Noun}: {ex.Message}");
                this.ShowSaveFailed($"The {section.Noun} could not be updated - see the log for details.");
            }
        }

        private async Task EditAttachmentAsync(AttachmentSection section, object? sender)
        {
            if (sender is not Button { Tag: int id })
            {
                return;
            }

            var record = section.Records().FirstOrDefault(r => r.Id == id);
            if (record == null)
            {
                return;
            }

            bool isPhoto = section.Kind == WorklogAttachmentStorage.AttachmentKind.Photo;

            var dialog = new WorklogAddPhotoWindow();
            dialog.InitializeForEdit(
                WorklogAttachmentStorage.GetDisplayFileName(record.FileName, section.OwnerPrefix, record.Id),
                record.Comment,
                // Only a photo has an image to preview; a document has none.
                isPhoto ? this.ResolveAttachmentPath(record.FileName) : null,
                section.Kind);

            var result = await dialog.ShowDialog<WorklogAddPhotoWindow.PhotoResult?>(this);
            if (result == null)
            {
                return;
            }

            string previousFileName = record.FileName;
            string previousComment = record.Comment;

            string? attachmentsFolder = null;
            string newStoredFileName = string.Empty;

            if (!string.IsNullOrWhiteSpace(result.SourcePath))
            {
                attachmentsFolder = WorklogManager.GetEntryAttachmentsFolder(this.thisWorkbookId, this.thisEntry.Id);
                if (attachmentsFolder == null)
                {
                    this.ShowSaveFailed($"Could not resolve where to store the {section.Noun}.");
                    return;
                }

                newStoredFileName = WorklogAttachmentStorage.BuildStoredFileName(
                    result.SourcePath, section.OwnerPrefix, record.Id);

                record.FileName = newStoredFileName;
            }

            record.Comment = result.Comment;

            // The record is saved BEFORE the file is swapped, because the swap deletes the file it
            // replaces. Doing it the other way round meant a failed save left entries.json naming a
            // file that had already been deleted - and Cancel then discarded the working copy, so
            // the row was permanently broken with no way back to the original.
            if (!this.PersistEntrySilently())
            {
                record.FileName = previousFileName;
                record.Comment = previousComment;
                this.RefreshAttachmentRows(section);
                return;
            }

            if (attachmentsFolder != null)
            {
                // Copies the new file in and removes the one it replaces, leaving exactly one file
                // behind whether or not the stored name changed - see TryReplaceAttachmentFile.
                this.RollBackAttachmentFileNameIfSwapFailed(
                    record,
                    previousFileName,
                    newStoredFileName,
                    result.SourcePath!,
                    attachmentsFolder,
                    $"The {section.Noun} could not be copied into the worklog.");
            }

            this.RefreshAttachmentRows(section);
        }

        // ###########################################################################################
        // Removes an attachment, metadata and bytes both. Deleting the file is safe because the
        // stored name carries the record's own id (see BuildStoredFileName), so it can only ever
        // belong to the record being removed - the app copied it in and nothing else points at it.
        //
        // The file goes only after the metadata change has been persisted: if the save fails the
        // row is still listed, and deleting first would leave it pointing at nothing.
        //
        // DeleteAttachmentFileAndFolderIfEmpty also removes the entry's shared attachments folder
        // once nothing - photo or file - is left in it, so deleting an entry's last attachment
        // leaves no empty folder behind on disk.
        // ###########################################################################################
        private void OnDeletePhotoClick(object? sender, RoutedEventArgs e)
        {
            this.DeleteAttachment(this.PhotoAttachments, sender);
        }

        private void OnDeleteFileClick(object? sender, RoutedEventArgs e)
        {
            this.DeleteAttachment(this.FileAttachments, sender);
        }

        private void DeleteAttachment(AttachmentSection section, object? sender)
        {
            if (sender is not Button { Tag: int id })
            {
                return;
            }

            var records = section.Records();

            var record = records.FirstOrDefault(r => r.Id == id);
            if (record == null)
            {
                return;
            }

            string fileName = record.FileName;

            records.RemoveAll(r => r.Id == id);
            this.RefreshAttachmentRows(section);

            if (this.PersistEntrySilently())
            {
                WorklogAttachmentStorage.DeleteAttachmentFileAndFolderIfEmpty(
                    WorklogManager.GetEntryAttachmentsFolder(this.thisWorkbookId, this.thisEntry.Id),
                    fileName);
            }
        }

        // ###########################################################################################
        // Drag-to-reorder for the Photos and Files rows, replacing up/down buttons in both.
        //
        // Only the row's empty space starts a drag: the thumbnail, the file link and the icon
        // buttons handle their own pointer events and mark them handled, so pressing those never
        // begins a drag. That is also why the row shows the north/south cursor only over that empty
        // space - the cursor is set on the panel that carries the drag, not on the whole row.
        //
        // A press alone does not start the drag; it only arms it. The drag begins once the pointer
        // has actually moved a few pixels, so a plain click on a row cannot reorder anything by
        // accident.
        //
        // One implementation serves both lists, parameterised by the DragContext below. The logic
        // here is subtle in three places (the frozen boundary snapshot, the re-entrancy guard, the
        // placeholder-index-as-target rule), and a second copy would be a second place for those to
        // be got wrong.
        // ###########################################################################################
        // Built from the AttachmentSection for the list being dragged, so the drag and the
        // add/edit/delete paths cannot disagree about which records, rows or control a list has.
        //
        // Records stays a lookup rather than a captured list, and the reason is the whole point of
        // this record: the context survives the entire gesture, so anything replacing the thisEntry
        // record mid-drag (rather than mutating its lists in place) would leave the release handler
        // reordering a detached list and persisting nothing the user could see. Reading the list
        // through thisEntry at use time cannot go stale that way.
        private sealed record DragContext(
            ItemsControl List,
            ObservableCollection<WorklogAttachmentRow> Rows,
            Func<List<WorklogAttachmentRecord>> Records,
            Action Refresh);

        // Built once per gesture rather than on every press - these were properties allocating a
        // fresh record plus a bound delegate on every pointer press, including presses that
        // immediately bailed.
        private DragContext CreateDragContext(AttachmentSection section) => new(
            section.List(),
            section.Rows(),
            section.Records,
            () => this.RefreshAttachmentRows(section));

        private DragContext? thisActiveDragContext;

        private int thisDraggedPhotoId = -1;

        private Point thisPhotoDragStartPoint;

        private bool thisIsDraggingPhoto;

        // Far enough that a click with a shaky hand is not a drag, small enough to feel immediate.
        private const double PhotoDragThreshold = 4.0;

        private void OnPhotoRowDragHandlePointerPressed(object? sender, PointerPressedEventArgs e)
        {
            this.BeginRowDrag(sender, e, this.CreateDragContext(this.PhotoAttachments));
        }

        private void OnFileRowDragHandlePointerPressed(object? sender, PointerPressedEventArgs e)
        {
            this.BeginRowDrag(sender, e, this.CreateDragContext(this.FileAttachments));
        }

        private void BeginRowDrag(object? sender, PointerPressedEventArgs e, DragContext context)
        {
            if (sender is not Control { Tag: int id })
            {
                return;
            }

            if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            {
                return;
            }

            this.thisActiveDragContext = context;
            this.thisDraggedPhotoId = id;
            this.thisPhotoDragStartPoint = e.GetPosition(context.List);
            this.thisIsDraggingPhoto = false;
        }

        private void OnPhotoRowDragHandlePointerMoved(object? sender, PointerEventArgs e)
        {
            if (this.thisDraggedPhotoId < 0 || this.thisActiveDragContext == null)
            {
                return;
            }

            var context = this.thisActiveDragContext;

            if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            {
                // The button was released somewhere that did not reach the release handler (outside
                // the window, say); without this the next move would resume a drag the user ended.
                this.ResetPhotoDragState();
                return;
            }

            var current = e.GetPosition(context.List);

            if (!this.thisIsDraggingPhoto &&
                Math.Abs(current.Y - this.thisPhotoDragStartPoint.Y) < PhotoDragThreshold &&
                Math.Abs(current.X - this.thisPhotoDragStartPoint.X) < PhotoDragThreshold)
            {
                return;
            }

            if (!this.thisIsDraggingPhoto)
            {
                // Only a placeholder that was actually established starts the drag. If the row has
                // gone (a refresh landed between press and first move), the boundaries are empty
                // and every later step is a no-op that still LOOKS like a drag: no visual feedback
                // while dragging, and on release a full Refresh that re-decodes every thumbnail
                // from disk for a gesture that changed nothing. Better to end it here.
                if (!this.BeginPhotoDragPlaceholder())
                {
                    this.ResetPhotoDragState();
                    return;
                }
            }

            this.thisIsDraggingPhoto = true;

            // Move the dragged row to wherever the pointer now is, so the gap follows the pointer
            // and the surrounding rows shift into the order the drop will produce. The row is drawn
            // as an outlined slot while it is the placeholder, so what the user sees is the space
            // it will occupy rather than the row itself trailing the cursor.
            //
            // Guarded against re-entry: Move() reorders the collection, which makes Avalonia recycle
            // the row containers, which changes the element under the cursor and raises further
            // pointer events synchronously. Those re-enter this handler and move again, and the list
            // flickers between orders for as long as the pointer is held there. The flag makes the
            // nested calls no-ops so one physical mouse move produces exactly one reorder.
            if (this.thisIsApplyingPhotoPlaceholderMove)
            {
                return;
            }

            this.thisIsApplyingPhotoPlaceholderMove = true;
            try
            {
                this.MovePhotoPlaceholderTo(context, this.ResolvePhotoDropIndex(context, current));
            }
            finally
            {
                this.thisIsApplyingPhotoPlaceholderMove = false;
            }
        }

        private bool thisIsApplyingPhotoPlaceholderMove;

        private void OnPhotoRowDragHandlePointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            if (this.thisDraggedPhotoId < 0 || !this.thisIsDraggingPhoto || this.thisActiveDragContext == null)
            {
                this.ResetPhotoDragState();
                return;
            }

            var context = this.thisActiveDragContext;
            int draggedId = this.thisDraggedPhotoId;

            // The placeholder is already sitting at the drop position, so its index in the row list
            // IS the target - no need to re-measure against the pointer, which would disagree with
            // what the user was just shown if the pointer sat between two rows.
            int targetIndex = IndexOfPhotoRow(context, draggedId);

            this.ResetPhotoDragState();

            if (targetIndex < 0)
            {
                context.Refresh();
                return;
            }

            WorklogAttachmentStorage.ReorderAttachment(context.Records(), draggedId, targetIndex);
            context.Refresh();
            this.PersistEntrySilently();
        }

        // ###########################################################################################
        // Turns the dragged row into the placeholder, sized to the height it currently occupies so
        // the gap does not jump when its content is swapped for the empty outline.
        // ###########################################################################################
        // Returns false when no placeholder could be established, so the caller can abandon the
        // gesture instead of running a drag with no boundaries and no visible row.
        private bool BeginPhotoDragPlaceholder()
        {
            var context = this.thisActiveDragContext;
            if (context == null)
            {
                return false;
            }

            // Cleared up front so an early return below cannot leave the PREVIOUS drag's boundaries
            // in place - ResolvePhotoDropIndex would then measure this drag against the other list's
            // geometry, and with photo rows several times taller than file rows the drop lands at a
            // wildly wrong index and is persisted immediately.
            this.thisPhotoRowDragBoundaries.Clear();

            int index = IndexOfPhotoRow(context, this.thisDraggedPhotoId);
            if (index < 0)
            {
                return false;
            }

            var row = context.Rows[index];

            var container = context.List.ContainerFromIndex(index);
            if (container != null && container.Bounds.Height > 0)
            {
                row.PlaceholderHeight = container.Bounds.Height;
            }

            this.CapturePhotoRowBoundaries(context);

            row.IsDropPlaceholder = true;
            return true;
        }

        // ###########################################################################################
        // The Y positions of the row boundaries as they are at the moment the drag starts, used to
        // decide which slot the pointer is over for the rest of the gesture.
        //
        // A snapshot rather than live measurement, because measuring live feeds the swap back into
        // its own input: moving the placeholder re-lays out the list, which moves the very rows the
        // next measurement reads, which can select a different slot, which moves them back - the
        // rows oscillate every frame. That feedback is unavoidable once rows differ in height (they
        // do now that each image is sized by its own aspect ratio), because a swap shifts the
        // layout by the difference between two row heights rather than leaving it unchanged.
        //
        // Against a frozen frame the pointer position alone decides the slot, so the same pointer
        // position always gives the same answer and there is nothing to oscillate.
        // ###########################################################################################
        private readonly List<double> thisPhotoRowDragBoundaries = new();

        private void CapturePhotoRowBoundaries(DragContext context)
        {
            this.thisPhotoRowDragBoundaries.Clear();

            // One entry per row, always - index i in this list means row i. Skipping a row whose
            // container is not realized would shorten the list and shift every later boundary's
            // meaning by one, so an unmeasurable row gets an interpolated midpoint instead and the
            // two lists stay aligned. ResolvePhotoDropIndex relies on that 1:1 correspondence to
            // return an index into the row collection.
            double runningY = 0;

            for (int i = 0; i < context.Rows.Count; i++)
            {
                var container = context.List.ContainerFromIndex(i);
                Point? topLeft = container?.TranslatePoint(new Point(0, 0), context.List);

                double height = container != null && container.Bounds.Height > 0
                    ? container.Bounds.Height
                    : context.Rows[i].PlaceholderHeight;

                double top = topLeft?.Y ?? runningY;

                // The midpoint of each row as laid out before anything moved. The pointer being
                // past a midpoint means the drop belongs after that row.
                this.thisPhotoRowDragBoundaries.Add(top + (height / 2.0));

                runningY = top + height;
            }
        }

        // ###########################################################################################
        // Moves the placeholder row to the given index, leaving the collection untouched when it is
        // already there - a Move on every pointer frame would rebuild containers continuously and
        // make the list flicker.
        // ###########################################################################################
        private void MovePhotoPlaceholderTo(DragContext context, int targetIndex)
        {
            if (targetIndex < 0)
            {
                return;
            }

            int currentIndex = IndexOfPhotoRow(context, this.thisDraggedPhotoId);
            if (currentIndex < 0)
            {
                return;
            }

            targetIndex = Math.Clamp(targetIndex, 0, context.Rows.Count - 1);
            if (targetIndex == currentIndex)
            {
                return;
            }

            context.Rows.Move(currentIndex, targetIndex);
        }

        private static int IndexOfPhotoRow(DragContext context, int id)
        {
            for (int i = 0; i < context.Rows.Count; i++)
            {
                if (context.Rows[i].Id == id)
                {
                    return i;
                }
            }

            return -1;
        }

        // ###########################################################################################
        // Ends the drag and returns every row to its normal appearance. Clearing the flag on all
        // rows rather than just the dragged one means an interrupted drag (the window closing, a
        // refresh landing mid-drag) cannot stitch a row permanently as a placeholder.
        // ###########################################################################################
        private void ResetPhotoDragState()
        {
            // Both lists are cleared, not just the active one: the flag must never survive a drag,
            // and an interrupted gesture can leave the context null while a row still carries it.
            foreach (var row in this.thisPhotoRows)
            {
                row.IsDropPlaceholder = false;
            }

            foreach (var row in this.thisFileRows)
            {
                row.IsDropPlaceholder = false;
            }

            // Cleared so the next drag cannot resolve against the previous drag's layout.
            this.thisPhotoRowDragBoundaries.Clear();

            this.thisActiveDragContext = null;
            this.thisDraggedPhotoId = -1;
            this.thisIsDraggingPhoto = false;
        }

        // ###########################################################################################
        // Which slot the pointer is over, measured against the boundaries captured when the drag
        // started (see CapturePhotoRowBoundaries for why a live measurement oscillates).
        //
        // Above the first boundary gives 0 and past the last gives the final index, so a drag flung
        // past either end lands at that end instead of being discarded.
        // ###########################################################################################
        private int ResolvePhotoDropIndex(DragContext context, Point pointerInList)
        {
            if (context.Rows.Count == 0 || this.thisPhotoRowDragBoundaries.Count == 0)
            {
                return -1;
            }

            for (int i = 0; i < this.thisPhotoRowDragBoundaries.Count; i++)
            {
                if (pointerInList.Y < this.thisPhotoRowDragBoundaries[i])
                {
                    // Also covers a pointer dragged above the list entirely.
                    return i;
                }
            }

            // Past the last midpoint - the drop belongs at the end.
            return this.thisPhotoRowDragBoundaries.Count - 1;
        }

        // ###########################################################################################
        // Swaps an attachment's bytes for a newly picked file and, if that fails, puts the record's
        // stored name back so entries.json never names bytes that were never written.
        //
        // Shared by the photo and file edit paths, which had the same flaw independently: the
        // roll-back save's result was ignored, so if THAT save failed too the record was left
        // naming the new file while only the old bytes existed. The row then resolved to null
        // forever - "That file could no longer be found" - with no way back to the original.
        //
        // When the roll-back cannot be persisted the record is still restored in memory, so what
        // the user sees matches the bytes on disk, and the message says the entry needs saving.
        // That is the best available outcome: the alternative is leaving a name on screen that
        // nothing on disk backs.
        // ###########################################################################################
        private void RollBackAttachmentFileNameIfSwapFailed(
            WorklogAttachmentRecord record,
            string previousFileName,
            string newStoredFileName,
            string sourcePath,
            string attachmentsFolder,
            string failureMessage)
        {
            if (WorklogAttachmentStorage.TryReplaceAttachmentFile(
                    sourcePath,
                    attachmentsFolder,
                    previousFileName,
                    newStoredFileName,
                    out _))
            {
                return;
            }

            // The record already names the new file, so put it back before re-saving.
            record.FileName = previousFileName;

            if (this.PersistEntrySilently())
            {
                this.ShowSaveFailed(failureMessage);
                return;
            }

            this.ShowSaveFailed(failureMessage + " The entry could not be saved either - use Save to retry.");
        }

        // ###########################################################################################
        // Opens the clicked file through ExternalTargetLauncher, scoped to the entry's attachments
        // folder rather than the data root: worklog workbooks live in AppData, outside the data root
        // the launcher defaults to, so without the override every attachment would be refused as
        // "outside allowed scope". The containment check still applies - just against this folder.
        // ###########################################################################################
        private void OnFileRowPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is not Control { Tag: int id })
            {
                return;
            }

            // Marked handled IMMEDIATELY, before any other guard can return.
            //
            // This link sits inside the row's drag Panel, whose PointerPressed arms a reorder. The
            // Panel is skipped only when this handler marks the event handled - so every early
            // return below that happens BEFORE this line lets the press fall through and arm a drag
            // the user never started. The next few pixels of movement then cross the drag threshold,
            // turn the row into a placeholder, and on release commit a reorder and save it.
            //
            // That was the intermittent "clicking a file link does nothing" fault: a right-click, or
            // a click on a row whose record had just been replaced, silently armed a phantom drag
            // instead of opening anything. It is why the guards below now run after this line, and
            // why this must stay first.
            e.Handled = true;

            // Left button only. This launches an external application through the OS shell, so a
            // right-click reaching for a context menu must not open the document - unlike the photo
            // thumbnail, whose press opens a modal the user can simply close.
            if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            {
                return;
            }

            var file = this.thisEntry.Files.FirstOrDefault(f => f.Id == id);
            if (file == null)
            {
                return;
            }

            string? path = this.ResolveAttachmentPath(file.FileName);
            if (path == null)
            {
                this.ShowSaveFailed("That file could no longer be found.");
                return;
            }

            string? attachmentsFolder = WorklogManager.GetEntryAttachmentsFolder(this.thisWorkbookId, this.thisEntry.Id);
            bool opened = attachmentsFolder != null && ExternalTargetLauncher.TryOpen(path, attachmentsFolder);
            if (!opened)
            {
                this.ShowSaveFailed("That file could not be opened.");
            }
        }
    }
}
