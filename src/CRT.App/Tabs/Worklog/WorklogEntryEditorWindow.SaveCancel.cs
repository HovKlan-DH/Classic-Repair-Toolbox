using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;

namespace CRT
{
    // ###########################################################################################
    // The discard/cancel/save flow: whether the entered work would be lost, the confirmation
    // dialog, cancelling (discarding a draft's attachments), and Save itself. See
    // WorklogEntryEditorWindow.axaml.cs for the file map of the whole partial class.
    // ###########################################################################################
    public partial class WorklogEntryEditorWindow
    {
        // ###########################################################################################
        // Cancel/Escape discards pending edits to the direct fields (Title/Description/category/
        // state), but still reports WasSaved when a Links/Comments/Work-done/Photos/Files change
        // already made it to disk via PersistEntrySilently, so the caller knows to refresh.
        //
        // "Open" is the limit of what Cancel can undo: an instant-save commits the direct fields
        // along with the sub-list change (see PersistEntrySilently), so once one has run, the
        // direct-field values at that moment are already on disk and Cancel cannot take them back.
        // ###########################################################################################
        // ###########################################################################################
        // Whether abandoning this window right now would throw away work the user has entered.
        //
        // TRUE ONLY FOR A DRAFT, and only one that actually holds something. A saved worklog writes
        // every sub-list change to disk the moment it is made (PersistEntrySilently), so the most
        // Cancel can cost there is a half-typed Title or Description - which is exactly what Cancel
        // is FOR, and prompting about it on every Escape would train the user to dismiss the prompt.
        //
        // A draft is the opposite: nothing has been written, so everything goes - the typed fields,
        // every comment and work-done row, and the attachment bytes with them. Because those
        // sub-lists behave as "instantly saved" on a saved worklog, they FEEL saved here too, which
        // is precisely how the reported data loss happened.
        //
        // An empty draft (opened and immediately abandoned) returns false: there is nothing to
        // protect, and a prompt there is pure friction on the most common way to back out.
        // ###########################################################################################
        // ###########################################################################################
        // Test seam for the discard guard's DECISION - whether closing right now would lose work.
        //
        // The decision is the whole of the correctness here: whether the confirmation appears at all,
        // and it must be true for a draft holding anything and false for every saved worklog. The
        // dialog it then opens cannot be driven headlessly (ShowDialog resolves with no user), so
        // that half is verified by running the app - see the tests in WorklogEditorDiscardGuardTests.
        // ###########################################################################################
        internal bool WouldDiscardEnteredWorkForTests() => this.WouldDiscardEnteredWork();

        private bool WouldDiscardEnteredWork()
        {
            if (!this.thisIsDraftEntry || this.thisWasSuccessfullySaved)
            {
                return false;
            }

            // Read from the CONTROLS, not from thisEntry: the direct fields are only synced into the
            // record on a save or an instant-save, so a user who typed a title and immediately hit
            // Escape has it nowhere else. This is the case the guard most needs to catch.
            if (!string.IsNullOrWhiteSpace(this.EditorTitleTextBox?.Text) ||
                !string.IsNullOrWhiteSpace(this.EditorDescriptionTextBox?.Text))
            {
                return true;
            }

            // Compared against what the draft STARTED with, not against zero. A new worklog is
            // seeded with the automatic "Worklog created" audit comment (InitializeForNewEntry), so
            // a zero test reports work to lose the instant the editor opens and would prompt on
            // every cancel - including the untouched empty form, which is the most common way to
            // back out. Snapshotting is used rather than special-casing that one comment because it
            // stays correct if anything else is ever seeded into a draft.
            return this.thisEntry.Comments.Count > this.thisDraftInitialCommentCount
                || this.thisEntry.WorkDoneItems.Count > this.thisDraftInitialWorkDoneCount
                || this.thisEntry.Links.Count > this.thisDraftInitialLinkCount
                || this.thisEntry.Photos.Count > this.thisDraftInitialPhotoCount
                || this.thisEntry.Files.Count > this.thisDraftInitialFileCount;
        }

        // ###########################################################################################
        // Asks before discarding, and reports whether the caller may go ahead.
        //
        // Returns true immediately when there is nothing to lose, so the ordinary Escape-out-of-an
        // -empty-form path is unchanged and costs no dialog.
        //
        // The dialog itself takes no arguments: it asks one question, and WouldDiscardEnteredWork
        // above has already decided that the question is worth asking. An earlier version listed
        // what would be lost ("the title X, the description and 1 work done row"); that was cut back
        // on request, so nothing here has to be described to it.
        // ###########################################################################################
        private async System.Threading.Tasks.Task<bool> ConfirmDiscardIfNeededAsync()
        {
            if (!this.WouldDiscardEnteredWork())
            {
                return true;
            }

            return await new DiscardWorklogChangesWindow().ShowDialog<bool>(this);
        }

        // ###########################################################################################
        // Cancel/Escape. Asks first when abandoning would lose entered work - see
        // ConfirmDiscardIfNeededAsync - and otherwise behaves exactly as before.
        //
        // async void because it is an event handler: there is nothing to await it. The guard sets
        // thisIsClosingConfirmed before closing, so the Closing handler below knows the question has
        // already been asked and does not ask it twice.
        // ###########################################################################################
        private async void OnCancelClick(object? sender, RoutedEventArgs e)
        {
            if (!await this.ConfirmDiscardIfNeededAsync())
            {
                return;
            }

            this.thisIsClosingConfirmed = true;

            // A draft has written nothing to entries.json, so there is nothing for the caller to
            // refresh and nothing for Cancel to have failed to undo - whatever
            // thisHasPersistedChange picked up along the way describes a save that never happened.
            if (this.thisIsDraftEntry)
            {
                this.DiscardDraftAttachments();
                this.WasSaved = false;
                this.Close(false);
                return;
            }

            this.WasSaved = this.thisHasPersistedChange;
            this.Close(this.WasSaved);
        }

        // ###########################################################################################
        // Removes the attachment bytes a cancelled draft wrote.
        //
        // Photo and file bytes are copied to disk the moment they are added - they have to be, a
        // photo cannot be shown from nowhere - so a draft that added one and was then cancelled
        // would leave that folder behind naming an entry that does not exist. Worse than untidy:
        // WorklogManager.AddEntryRecord moves the reserved folder into place for the NEXT draft that
        // reserves the same id, so the abandoned photos would reappear on an unrelated entry.
        //
        // Only ever called for a draft, and only for the folder named after its own reserved id.
        // Failure is logged rather than surfaced: the user asked to cancel, and there is nothing
        // useful they could do about it.
        //
        // TWO things this must not do, both of which the obvious version did:
        //
        //  - It must not resolve the path through WorklogManager.GetEntryAttachmentsFolder, which
        //    CREATES the folder. Resolving to delete would re-create the very folder this is about
        //    to remove, leaving an empty one behind where there had been none. Hence the
        //    ...FolderPath form, which only builds the path.
        //  - It must not delete a folder that a SAVED entry now owns. The reserved id is a peek, not
        //    a reservation: an entry saved elsewhere while this draft was open can legitimately hold
        //    that number, and its photo and file bytes live in exactly this folder. Deleting it
        //    would destroy a saved entry's attachments while its entries.json rows survived, each
        //    pointing at nothing. So the id is checked against what is actually on disk first.
        // ###########################################################################################
        private void DiscardDraftAttachments()
        {
            if (this.thisEntry.Photos.Count == 0 && this.thisEntry.Files.Count == 0)
            {
                return;
            }

            if (WorklogManager.EntryExists(this.thisWorkbookId, this.thisDraftReservedEntryId))
            {
                // Another entry claimed the reserved number while this draft was open. Leaving the
                // bytes behind is untidy; deleting them destroys that entry's attachments.
                Logger.Warning(
                    $"Not discarding cancelled draft attachments for workbook [#{this.thisWorkbookId}] entry [#{this.thisDraftReservedEntryId}]: " +
                    "a saved entry now uses that id");
                return;
            }

            string? attachmentsFolder = WorklogManager.GetEntryAttachmentsFolderPath(
                this.thisWorkbookId,
                this.thisDraftReservedEntryId);

            if (attachmentsFolder == null)
            {
                return;
            }

            try
            {
                if (System.IO.Directory.Exists(attachmentsFolder))
                {
                    System.IO.Directory.Delete(attachmentsFolder, recursive: true);
                    Logger.Info($"Discarded attachments of cancelled draft worklog entry [#{this.thisDraftReservedEntryId}]");
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to discard cancelled draft attachments [{attachmentsFolder}]: [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Commits the working copy back via WorklogManager.UpdateEntry, which also recomputes the
        // workbook's Open/Closed status - editing State to Closed here is exactly how the user
        // resolves an entry from the full editor.
        // ###########################################################################################
        private void OnSaveClick(object? sender, RoutedEventArgs e)
        {
            this.SyncDirectFieldsToEntry();

            if (this.thisIsDraftEntry)
            {
                // First time this entry has reached disk. AddEntryRecord allocates the real id (the
                // draft's was only reserved) and moves the attachment folder if the two differ.
                var saved = WorklogManager.AddEntryRecord(
                    this.thisWorkbookId,
                    this.thisEntry,
                    this.thisDraftReservedEntryId);

                if (saved == null)
                {
                    this.ShowSaveFailed(DefaultSaveFailedMessage);
                    return;
                }

                // No longer a draft: the record exists, so a reopen or any later instant-save must
                // go down the ordinary UpdateEntry path rather than adding a second copy.
                //
                // thisWasSuccessfullySaved is the belt to that braces - the Closing handler's
                // attachment cleanup checks both, so this entry's freshly-moved attachment folder is
                // safe even if the draft flag is ever left set. See that handler.
                this.thisIsDraftEntry = false;
                this.thisWasSuccessfullySaved = true;
                this.thisEntry = saved;
                this.SavedNewEntry = saved;
                this.EditorIdText.Text = $"#{saved.Id}";

                this.WasSaved = true;

                // A successful save has nothing left to lose, so the Closing handler must not ask
                // "discard unsaved work?" on the way out. Set EXPLICITLY rather than relying on
                // WouldDiscardEnteredWork happening to short-circuit on the draft flag: this is the
                // flag whose whole job is "the close question is already answered", and a later
                // change to what that predicate covers must not be able to reintroduce a discard
                // prompt after a save. See thisIsClosingConfirmed.
                this.thisIsClosingConfirmed = true;

                this.Close(this.WasSaved);
                return;
            }

            if (!WorklogManager.UpdateEntry(this.thisWorkbookId, this.thisEntry))
            {
                // Nothing reached disk. Closing here would report success and the user would watch
                // their edits revert on the next refresh, so keep the window open with what they
                // typed still in it and say so. The log carries the underlying reason.
                this.ShowSaveFailed(DefaultSaveFailedMessage);
                return;
            }

            this.WasSaved = true;

            // As above: the work is on disk, so the close question is answered.
            this.thisIsClosingConfirmed = true;

            this.Close(this.WasSaved);
        }
    }
}
