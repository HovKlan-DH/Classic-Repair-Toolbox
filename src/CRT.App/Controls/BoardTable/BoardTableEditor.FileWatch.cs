using Avalonia;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;

namespace CRT
{
    // ###########################################################################################
    // BoardTableEditor - WATCHING THE DRAFT FILE WHILE THE TABLE IS ON SCREEN (owner request,
    // 2026-09-24: "it would be nice to see a warning that edits have been detected in Excel, and
    // that the content on screen may be invalid").
    //
    // Excel is the one writer of a draft that nothing in the application can hold back (the
    // Contribute tab's editor and the label editor wait for the table - see
    // TabDrafts.HasUnsavedTableEditsFor). So while the table is on screen the file is checked every
    // couple of seconds (DraftTableSession.CheckFile - one stat call when nothing changed):
    //
    //   - CHANGED, and the table has nothing unsaved: it reloads, keeping its place, and says so.
    //   - CHANGED, and the table HAS unsaved edits: a warning bar says the table is out of date and
    //     its edits can no longer be saved into the draft, with a Reload button, and "Save changes"
    //     is greyed out - that save would be refused anyway.
    //   - OPEN in Excel (or LibreOffice), going by the lock file it leaves beside the workbook: a
    //     notice asks to edit it in one place at a time - shown the whole time, from the moment the
    //     table opens, since that advice matters most BEFORE any editing starts (owner's
    //     choice, 2026-09-24, after briefly showing it only once the table had unsaved edits).
    //     "Save changes" is OFF while the workbook is really held open there (owner request) -
    //     that save would fail. Not on the lock file alone: a crashed Excel leaves it behind, and
    //     the table must not be blocked for good (see DraftTableSession.IsHeldOpen).
    //
    // A reload because of an outside change raises ReloadedFromOutside, which the Drafts tab answers
    // exactly as it answers Saved: the draft's "N rows changed" and the board on screen (component
    // list, schematics) were read from the file before Excel changed it.
    //
    // *** NEVER A RELOAD UNDER THE CONTRIBUTOR'S HANDS. *** A cell being typed in has not reached
    // the model yet, so "nothing unsaved" would reload the typing away; a row being dragged is
    // mid-gesture. Either waits for the next check, which by then finds unsaved edits (or none).
    //
    // Watching runs only while the editor is on screen and holds a table: a tab switch detaches it,
    // and coming back checks at once.
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        private readonly DispatcherTimer thisFileCheckTimer = new() { Interval = TimeSpan.FromSeconds(2) };

        // The draft changed on disk and the table has unsaved edits: the warning bar is up and Save
        // is off.
        private bool thisDraftChangedOnDisk;

        // Between a cell starting its editor and the edit ending (committed or cancelled).
        private bool thisIsEditingCell;

        // Excel (or another program) holds the workbook open, so a save cannot land: Save is off.
        private bool thisDraftHeldElsewhere;

        private bool thisIsOnScreen;

        internal bool IsWatchingDraftFileForTests => this.thisFileCheckTimer.IsEnabled;

        // Raised after the table re-read its draft because something outside it (Excel, say) had
        // changed the file - see the header.
        public event EventHandler? ReloadedFromOutside;

        private void WireDraftFileWatching()
        {
            this.thisFileCheckTimer.Tick += (_, _) => this.CheckDraftFile();

            // After OnBeginningEdit, so an edit it refused (a red ghost) never counts as one.
            this.TableGrid.BeginningEdit += (_, e) => this.thisIsEditingCell = !e.Cancel;
            this.TableGrid.CellEditEnded += (_, _) => this.thisIsEditingCell = false;
        }

        // ###########################################################################################
        // Looks at the draft file now and brings the table and its notices in line. Called by the
        // timer, when the table comes back on screen, and by the Drafts tab after it imported files
        // into the same draft.
        // ###########################################################################################
        public void CheckDraftFile()
        {
            if (this.thisSession is not { } session)
            {
                this.ShowOpenElsewhere(default);
                this.SetChangedOnDisk(false);
                return;
            }

            // The table's own save is writing the file right now (SaveAsync) - it would read as a
            // change from outside, and the save re-reads the file when it is done anyway.
            if (this.thisSaveInFlight)
            {
                return;
            }

            DraftFileStatus status = session.CheckFile();

            // Gone altogether is not a change to warn about: the Drafts tab closes such a table.
            if (status.IsMissing || !status.ChangedSinceRead)
            {
                this.SetChangedOnDisk(false);
            }
            else if (this.thisIsEditingCell || this.thisRowDrag is not null)
            {
                // Under the contributor's hands - the next check decides.
            }
            else if (!this.HasUnsavedChanges)
            {
                this.ReloadKeepingPlace();
                this.ShowStatus("Updated - the draft was changed outside this table.");
                this.ReloadedFromOutside?.Invoke(this, EventArgs.Empty);
            }
            else
            {
                this.SetChangedOnDisk(true);
            }

            this.ShowOpenElsewhere(status);
        }

        // ###########################################################################################
        // *** THE CHECKS AGAIN, WITH NOTHING EDITED (code review, 2026-10-04). *** A picture the
        // table marked "missing" may have been put in the draft folder meanwhile - which the watch
        // above cannot see, since the workbook did not change - and the mark stayed while the draft's
        // row and Submit found nothing wrong. The host calls this when the table comes back into
        // view (the Drafts tab shown, CRT's window in front again): the file lookup looks again for
        // every file it did not find (DiskFileLookup), so only those cost anything. Not under a cell
        // being typed in or a row being dragged - the next time decides.
        // ###########################################################################################
        public void RecheckFiles()
        {
            if (this.thisDocument is null || this.thisIsEditingCell || this.thisRowDrag is not null)
            {
                return;
            }

            this.thisDocument.RefreshProblems();
            this.OnDocumentChanged(this, EventArgs.Empty);
        }

        // The notice, and Save, as the file is open elsewhere or not.
        private void ShowOpenElsewhere(DraftFileStatus status)
        {
            this.OpenElsewhereBar.IsVisible = status.IsOpenElsewhere;

            bool held = status.IsOpenElsewhere && status.IsHeldOpenElsewhere;
            if (held != this.thisDraftHeldElsewhere)
            {
                this.thisDraftHeldElsewhere = held;
                this.UpdateToolbar();
            }
        }

        // Straight away when a table is read, rather than on the next check two seconds later: a
        // draft already open in Excel is exactly when the notice is most use.
        private void ShowWhetherOpenElsewhere(DraftTableSession session) => this.ShowOpenElsewhere(session.CheckFile());

        // ###########################################################################################
        // Asked NOW (not from the last check): whether a save would be refused because the draft is
        // held open elsewhere - so the Drafts tab's close prompt does not offer a Save that would
        // fail and loop, exactly the reason DraftChangedOnDisk offers none.
        // ###########################################################################################
        internal bool IsDraftHeldOpenElsewhere() =>
            this.thisSession?.CheckFile() is { IsOpenElsewhere: true, IsHeldOpenElsewhere: true };

        private void SetChangedOnDisk(bool changed)
        {
            this.thisDraftChangedOnDisk = changed;
            this.ChangedOnDiskBar.IsVisible = changed;
            this.UpdateToolbar();
        }

        // On while there is a table and it is on screen.
        private void UpdateFileWatching() =>
            this.thisFileCheckTimer.IsEnabled = this.thisSession is not null && this.thisIsOnScreen;

        protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnAttachedToVisualTree(e);

            this.thisIsOnScreen = true;
            this.UpdateFileWatching();

            // Another tab (or Excel) may have written the draft while this one was away.
            this.CheckDraftFile();
        }

        protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnDetachedFromVisualTree(e);

            this.thisIsOnScreen = false;
            this.UpdateFileWatching();

            // A card left open over a table that has left the screen would float over whatever
            // replaced it.
            this.HideFilePreview();
        }
    }
}
