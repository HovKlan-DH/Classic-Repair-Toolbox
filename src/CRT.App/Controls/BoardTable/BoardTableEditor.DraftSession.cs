using Avalonia;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // BoardTableEditor - THE DRAFTS TAB'S TABLE ON THE DRAFT HELD IN MEMORY (CRT 4.0 step 1, the
    // plan of 2026-10-24: "Every change is saved within a second ... Undo and redo work across every
    // editor for the whole session ... What goes away: every Save button in drafts ... and the table's
    // file watch with its Excel bar"). Built 2026-10-25.
    //
    // In the Contribute tab the table no longer reads and writes its draft's file itself. It shows the
    // draft's DraftSession - the one every editor of that draft is to share (DraftManager.Sessions) -
    // and:
    //
    //   - EVERY CHANGE GOES INTO THE SESSION once it has settled - the cell committed, the drag
    //     dropped (PushToSession): the table's rows applied to the session's content as ONE step,
    //     which says where it was (BoardTableSessionEdit). A change that leaves every row as it was
    //     (an empty row inserted) is no step. The table is not built again for its own change - what
    //     is on screen already is the change.
    //   - THE SESSION SAVES ITSELF a second after the last change, off the UI thread, while nobody
    //     waits (the save timer - DraftSession.IsSaveDue; not under the BusyOverlay: nobody pressed
    //     for it). Where "Save changes" was, the table says Saved, Saving... or "Could not save -
    //     <reason>" with Try again (DraftSaveWording). There is no Save and no Reload.
    //   - UNDO AND REDO ARE THE SESSION'S, and the table is built again from the session afterwards,
    //     its place kept - on the change itself when it was one of the table's.
    //   - A DRAFT FILE CHANGED OUTSIDE THE TABLE is read again when CRT itself wrote it (the host calls
    //     ReadDraftFileAgain - after the component editor, the label editor, "Schematic images"...),
    //     when the table comes back on screen, and is caught by any save after it (ChangedOnDisk):
    //     then the bar offers the version in the file - one undoable step - or the table's, saved over
    //     it. Nothing is thrown away unasked, and nothing polls the file any more.
    //   - NEW COMPONENTS stay where they were typed while the table is open, and are saved into their
    //     category in label order (BoardTableSheet.BuildRows), as the explicit save did - said once.
    //   - WHAT GOES WITH A CHANGE is said once it is made - what was drawn on a schematic image taken
    //     off the board - with Ctrl+Z to bring it back, instead of a question before a save there no
    //     longer is (BoardTableSessionWording.WentWith).
    //
    // The Maintainer tab's tables are documents, not drafts, and keep their explicit Save.
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        // The draft behind the table in the Contribute tab - null in document mode.
        private DraftSession? thisDraftSession;

        // Checks four times a second whether a save is due - the save itself waits a second after
        // the last change (DraftSession.SaveDelay).
        private readonly DispatcherTimer thisSaveTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };

        // The save running now, on the pool - one at a time.
        private Task<DraftWorkbookEditOutcome>? thisSaveInFlight;

        private bool thisPushPosted;

        // Said once per table: new components are saved into their category (NewComponentsPlaced).
        private bool thisToldAboutPlacing;

        // Whether a cell is in its editor - between the grid starting its editor and the edit ending.
        private bool thisIsEditingCell;

        // The checks' lookup of a row's files, kept across the table being built again from the
        // session - it remembers which files it found (DiskFileLookup).
        private IBoardFileLookup? thisFiles;

        // False keeps the timer from saving - for a test that saves when it says so (SaveNow).
        internal bool AutoSave { get; set; } = true;

        // Raised after the table read its draft again because something outside it changed the file
        // - the Contribute tab answers it as it answers Saved: the draft's row and the board on screen
        // were read before that change.
        public event EventHandler? ReloadedFromOutside;

        internal DraftSession? DraftSessionForTests => this.thisDraftSession;

        internal BoardTableDocument? DocumentForTests => this.thisDocument;

        // ###########################################################################################
        // Opens a draft as a table, on `session` - the draft held in memory, shared with every other
        // editor of it (the host's DraftSessions). `published` is what differences are coloured
        // against - null for a board with nothing published.
        //
        // `dataRoot` is the downloaded data, where the checks look for the files a row names after
        // the draft's own folder (owner request, 2026-10-02) - the host's to hand over, since this
        // control knows nothing of DataManager. Empty: files are not looked for. `baselineLabel` names
        // the published side in a changed cell's tooltip (BoardTableDocument.SourceBaselineLabel).
        // ###########################################################################################
        public bool Load(DraftSession session, BoardData? published, string dataRoot = "", string? baselineLabel = null)
        {
            ArgumentNullException.ThrowIfNull(session);

            // Another draft starts with an empty search (agreed, 2026-10-02); the same one keeps it.
            if (!string.Equals(session.ExcelDataFile, this.thisExcelDataFile, StringComparison.OrdinalIgnoreCase))
            {
                this.thisSearchText = string.Empty;
                this.SearchBox.Text = string.Empty;
            }

            this.thisExcelDataFile = session.ExcelDataFile;
            this.thisPublished = published;
            this.thisBaselineLabel = baselineLabel;
            this.thisToldAboutPlacing = false;

            this.thisFiles = string.IsNullOrWhiteSpace(dataRoot)
                ? null
                : new DiskFileLookup(dataRoot, DraftFolderLayout.GetBoardFolder(session.DraftsRoot, session.ExcelDataFile), session.ExcelDataFile);

            BoardTableDocument document = BoardTableDocument.Create(published, session.Content.Board, baselineLabel, this.thisFiles);

            this.Attach(document, session, keepSheet: null);
            this.ShowStatus(document.HasBaseline
                ? string.Empty
                : "Nothing of this board is published yet, so no row is marked as added, changed or deleted - every row is your own.");

            return true;
        }

        // ###########################################################################################
        // The same, opening the draft's session itself - for a host that shares none, and for the
        // tests. False when there is no readable draft.
        // ###########################################################################################
        public bool Load(string draftsRoot, string excelDataFile, BoardData? published, string dataRoot = "", string? baselineLabel = null) =>
            DraftSession.Open(draftsRoot, excelDataFile) is DraftSession session &&
            this.Load(session, published, dataRoot, baselineLabel);

        // ---- the session joining and leaving the table ----------------------------------------

        private void UseDraftSession(DraftSession? session)
        {
            if (ReferenceEquals(session, this.thisDraftSession))
            {
                return;
            }

            if (this.thisDraftSession is { } previous)
            {
                previous.ChangeComing -= this.OnDraftSessionChangeComing;
                previous.Changed -= this.OnDraftSessionChanged;
                previous.SaveStateChanged -= this.OnDraftSaveStateChanged;
            }

            this.thisDraftSession = session;

            if (session is not null)
            {
                session.ChangeComing += this.OnDraftSessionChangeComing;
                session.Changed += this.OnDraftSessionChanged;
                session.SaveStateChanged += this.OnDraftSaveStateChanged;
            }

            this.thisSaveTimer.IsEnabled = session is not null;
            this.ShowSaveState();
        }

        private void WireDraftSession()
        {
            this.thisSaveTimer.Tick += (_, _) => this.SaveIfDue();

            // After OnBeginningEdit, so an edit it refused (a red ghost) never counts as one.
            this.TableGrid.BeginningEdit += (_, e) => this.thisIsEditingCell = !e.Cancel;
            this.TableGrid.CellEditEnded += (_, _) =>
            {
                this.thisIsEditingCell = false;
                this.CellEditEnded?.Invoke(this, EventArgs.Empty);
                this.SchedulePush();
            };
        }

        // ---- the table's changes into the session ----------------------------------------------

        // ###########################################################################################
        // Hands the table's change to the session soon - after the grid has finished committing the
        // edit and the sheets have refreshed (both posted at the same priority, so they run first).
        // ###########################################################################################
        private void SchedulePush()
        {
            if (this.thisDraftSession is null || this.thisPushPosted || this.thisDocument?.HasUnsavedChanges != true)
            {
                return;
            }

            this.thisPushPosted = true;

            Dispatcher.UIThread.Post(() =>
            {
                this.thisPushPosted = false;
                this.PushToSession();
            }, DispatcherPriority.Background);
        }

        // ###########################################################################################
        // The table's rows applied to the session, as one step - nothing when nothing changed since
        // the last push, or while a cell is being typed in or a row dragged (the end of either pushes
        // again).
        // ###########################################################################################
        private void PushToSession()
        {
            if (this.thisDraftSession is not { } session ||
                this.thisDocument is not { HasUnsavedChanges: true } document ||
                this.thisIsEditingCell ||
                this.thisRowDrag is not null)
            {
                return;
            }

            Func<BoardContent, BoardContent?> edit = document.TakeSessionEdit();
            bool placing = document.Sheets.Any(sheet => sheet.HasRowsToPlaceOnSave);

            // Where the change was, for an undo to put the cursor back on it.
            BoardTableSessionEdit origin = this.WhereTheLatestChangeWas();
            BoardContent before = session.Content;

            bool changed = session.Apply(BoardTableSessionWording.Step(origin.Sheet), edit, origin);

            // In the session now, or nothing to put there: the table holds nothing of its own.
            document.MarkSaved();

            if (changed)
            {
                string? wentWith = BoardTableSessionWording.WentWith(DraftEditRules.LostWith(before, session.Content));

                if (wentWith is not null)
                {
                    this.ShowStatus(wentWith);
                }
                else if (placing && !this.thisToldAboutPlacing)
                {
                    this.thisToldAboutPlacing = true;
                    this.ShowStatus(BoardTableSessionWording.NewComponentsPlaced);
                }
            }

            this.UpdateToolbar();
        }

        private BoardTableSessionEdit WhereTheLatestChangeWas()
        {
            BoardTableHistoryResult? latest = this.thisDocument?.History.Latest;
            BoardTableSheet? sheet = latest?.Sheet ?? this.thisCurrentSheet;
            BoardTableRow? row = latest is null ? this.CurrentRow : latest.Row;

            return new BoardTableSessionEdit(
                this,
                sheet?.Name ?? string.Empty,
                sheet is null || row is null ? -1 : sheet.Rows.IndexOf(row),
                latest?.Column ?? this.CurrentCell?.ColumnIndex ?? -1);
        }

        // ---- the session's changes onto the table ----------------------------------------------

        // ###########################################################################################
        // The session changed. The table's own edit is already on screen; anything else - an undo, a
        // redo, the file read again, another editor's edit - builds the table again from the
        // session, keeping its place, or on the change undone or redone when it was the table's.
        // ###########################################################################################
        private void OnDraftSessionChanged(object? sender, DraftSessionChange change)
        {
            // A write run off the UI thread (DraftSession.WriteFileDirectly) raises it there.
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.Post(() => this.OnDraftSessionChanged(sender, change));
                return;
            }

            if (!ReferenceEquals(sender, this.thisDraftSession))
            {
                return;
            }

            if (change.Kind == DraftSessionChangeKind.Edit && change.Origin is BoardTableSessionEdit own && ReferenceEquals(own.Editor, this))
            {
                return;
            }

            BoardTableSessionEdit? focus = change.Kind is DraftSessionChangeKind.Undo or DraftSessionChangeKind.Redo &&
                                           change.Origin is BoardTableSessionEdit edit && ReferenceEquals(edit.Editor, this)
                ? edit
                : null;

            this.ShowSessionContent(focus);

            switch (change.Kind)
            {
                case DraftSessionChangeKind.Undo:
                    this.ShowStatus(BoardTableSessionWording.Undone(change.Description));
                    break;

                case DraftSessionChangeKind.Redo:
                    this.ShowStatus(BoardTableSessionWording.Redone(change.Description));
                    break;
            }
        }

        // ###########################################################################################
        // ANOTHER CHANGE IS COMING - another editor's, an undo from elsewhere, the file read again: what
        // the table holds goes to the session first, a cell still in its editor committed, so the
        // coming change lands on top of it and the table, built again afterwards, loses nothing
        // (DraftSession.ChangeComing). Its own pushes it ignores. Only on the UI thread, where every
        // change but a direct write is made - and a direct write is from a window holding CRT's main
        // window, with nothing typed in the table meanwhile.
        // ###########################################################################################
        private void OnDraftSessionChangeComing(object? sender, DraftSessionChange change)
        {
            if (!ReferenceEquals(sender, this.thisDraftSession) ||
                !Dispatcher.UIThread.CheckAccess() ||
                (change.Origin is BoardTableSessionEdit own && ReferenceEquals(own.Editor, this)))
            {
                return;
            }

            this.TableGrid.CommitEdit();
            this.PushToSession();
        }

        // ###########################################################################################
        // The table built again from the session's content - the same sheet, the cursor on the same
        // row (found by its values, else by its place), or at `focus`.
        // ###########################################################################################
        private void ShowSessionContent(BoardTableSessionEdit? focus)
        {
            if (this.thisDraftSession is not { } session)
            {
                return;
            }

            this.TableGrid.CommitEdit();

            string? sheetName = focus is { Sheet.Length: > 0 } ? focus.Sheet : this.thisCurrentSheet?.Name;
            int rowIndex = focus is not null
                ? focus.Row
                : this.thisCurrentSheet is null || this.CurrentRow is null ? -1 : this.thisCurrentSheet.Rows.IndexOf(this.CurrentRow);
            int column = focus is { Column: >= 0 } ? focus.Column : this.CurrentCell?.ColumnIndex ?? 0;
            List<string>? currentValues = focus is null ? this.CurrentRow?.Cells.Select(cell => cell.Text).ToList() : null;

            BoardTableDocument document = BoardTableDocument.Create(this.thisPublished, session.Content.Board, this.thisBaselineLabel, this.thisFiles);
            this.Attach(document, session, sheetName);

            if (rowIndex >= 0 && this.thisCurrentSheet is { Rows.Count: > 0 } sheet)
            {
                BoardTableRow? same = currentValues is null
                    ? null
                    : sheet.Rows.FirstOrDefault(candidate =>
                        !candidate.IsDeleted && candidate.Cells.Select(cell => cell.Text).SequenceEqual(currentValues));

                this.SelectCell(same ?? sheet.Rows[Math.Min(rowIndex, sheet.Rows.Count - 1)], Math.Clamp(column, 0, sheet.Columns.Count - 1));
            }
        }

        // ###########################################################################################
        // Ctrl+Z / Ctrl+Y on a draft: the session's step, whichever editor made it - what is on screen
        // handed over first, so it is the latest step.
        // ###########################################################################################
        private void StepDraftSession(bool undo)
        {
            if (this.thisDraftSession is not { } session)
            {
                return;
            }

            this.TableGrid.CommitEdit();
            this.PushToSession();

            if (undo ? session.Undo() : session.Redo())
            {
                this.UpdateToolbar();
                this.FocusGrid();
            }
        }

        // ###########################################################################################
        // CRT itself wrote the draft's file - the component editor, the label editor, "Schematic
        // images" - or the table came back on screen: the file read again when it changed, as one
        // undoable step. Nothing while a change of the table's waits to be saved - that save refuses
        // and asks which to keep - or under a cell being typed in or a row being dragged.
        // ###########################################################################################
        public void ReadDraftFileAgain()
        {
            if (this.thisDraftSession is not { } session || this.thisIsEditingCell || this.thisRowDrag is not null)
            {
                return;
            }

            this.PushToSession();

            if (session.ReadFileIfChanged(DraftSaveWording.ReadAgainStep))
            {
                this.ShowStatus(DraftSaveWording.UpdatedFromOutside);
                this.ReloadedFromOutside?.Invoke(this, EventArgs.Empty);
            }
        }

        // ---- saving -----------------------------------------------------------------------------

        // What the timer does: a change not handed over yet goes to the session, and a save due runs.
        private void SaveIfDue()
        {
            if (this.thisDraftSession is not { } session || this.thisSaveInFlight is not null)
            {
                return;
            }

            if (this.thisDocument?.HasUnsavedChanges == true)
            {
                this.PushToSession();
            }

            if (this.AutoSave && session.IsSaveDueNow)
            {
                _ = this.SaveDraftSessionAsync();
            }
        }

        // Runs the timer's check now - for a test, whose session clock it moves itself.
        internal void SaveIfDueForTests() => this.SaveIfDue();

        // ###########################################################################################
        // Saves the session off the UI thread, after any save already running. Never throws: a save
        // that fails is Failed, said on the table with its reason.
        // ###########################################################################################
        private async Task<DraftWorkbookEditOutcome> SaveDraftSessionAsync(bool overwriteOutsideChange = false)
        {
            if (this.thisDraftSession is not { } session)
            {
                return DraftWorkbookEditOutcome.NoDraft;
            }

            while (this.thisSaveInFlight is { } running)
            {
                await running;
            }

            bool hadChanges = session.HasUnsavedChanges;
            Task<DraftWorkbookEditOutcome> save = session.SaveAsync(overwriteOutsideChange);
            this.thisSaveInFlight = save;

            DraftWorkbookEditOutcome outcome;

            try
            {
                outcome = await save;
            }
            finally
            {
                if (ReferenceEquals(this.thisSaveInFlight, save))
                {
                    this.thisSaveInFlight = null;
                }
            }

            this.AfterDraftSave(session, outcome, hadChanges);

            return outcome;
        }

        private void AfterDraftSave(DraftSession session, DraftWorkbookEditOutcome outcome, bool hadChanges)
        {
            if (!ReferenceEquals(session, this.thisDraftSession))
            {
                return;
            }

            this.ShowSaveState();
            this.UpdateToolbar();

            if (outcome == DraftWorkbookEditOutcome.Saved && hadChanges)
            {
                this.Saved?.Invoke(this, EventArgs.Empty);
            }
        }

        // ###########################################################################################
        // EVERYTHING ON SCREEN SAVED NOW - before the host acts on the draft's file (submitting it,
        // managing its files), closes the table or quits CRT: a cell still in its editor committed,
        // the change handed to the session, and the session saved, waiting for a save already
        // running. True when nothing is left unsaved.
        // ###########################################################################################
        public async Task<bool> SaveNowAsync()
        {
            if (this.thisDraftSession is not { } session)
            {
                return !this.HasUnsavedChanges;
            }

            this.TableGrid.CommitEdit();
            this.PushToSession();

            if (session.HasUnsavedChanges || this.thisSaveInFlight is not null)
            {
                await this.SaveDraftSessionAsync();
            }

            return !this.HasUnsavedChanges;
        }

        // ###########################################################################################
        // The same, on the calling thread - for the tests, which drive a save without waiting for the
        // timer. A save already running on the pool is waited for (DraftSession saves one at a time).
        // ###########################################################################################
        public DraftWorkbookEditOutcome SaveNow()
        {
            if (this.thisDraftSession is not { } session)
            {
                return DraftWorkbookEditOutcome.NoDraft;
            }

            this.TableGrid.CommitEdit();
            this.PushToSession();

            bool hadChanges = session.HasUnsavedChanges;
            DraftWorkbookEditOutcome outcome = session.Save();
            this.AfterDraftSave(session, outcome, hadChanges);

            return outcome;
        }

        // ###########################################################################################
        // Throws away what the table holds that is not saved, by reading the file again - one
        // undoable step, so even this can be taken back while CRT runs. For the host, after the
        // contributor chose Discard (a save that could not be made). False when the file cannot be
        // read; the host then lets the session go.
        // ###########################################################################################
        public bool DiscardUnsavedChanges()
        {
            if (this.thisDraftSession is not { } session)
            {
                return false;
            }

            this.TableGrid.CancelEdit();
            this.thisDocument?.MarkSaved();

            return session.ReloadFromFile(DraftSaveWording.UsedSavedVersionStep);
        }

        // Why the last save failed, or null - what the host's question says.
        public DraftWorkbookEditOutcome? SaveFailure => this.thisDraftSession?.SaveFailure;

        // A save's state changed - on the pool, where saves run: shown on the UI thread.
        private void OnDraftSaveStateChanged(object? sender, EventArgs e)
        {
            if (Dispatcher.UIThread.CheckAccess())
            {
                this.ShowSaveState();
                return;
            }

            Dispatcher.UIThread.Post(this.ShowSaveState);
        }

        // ###########################################################################################
        // Saved, Saving... or "Could not save - <reason>" where "Save changes" is in document mode,
        // with Try again after a failure - or, when the file changed outside the table, the bar with
        // its two choices instead.
        // ###########################################################################################
        private void ShowSaveState()
        {
            bool session = this.thisDraftSession is not null;

            this.SaveStatePanel.IsVisible = session;

            if (this.thisDraftSession is not { } draft)
            {
                this.ChangedOnDiskBar.IsVisible = false;
                this.NoteUnsavedState();
                return;
            }

            DraftSaveState state = this.thisDocument?.HasUnsavedChanges == true && draft.SaveState == DraftSaveState.Saved
                ? DraftSaveState.Pending
                : draft.SaveState;

            this.SaveStateText.Text = DraftSaveWording.StateLine(state, draft.SaveFailure);

            bool changedOutside = state == DraftSaveState.Failed && draft.SaveFailure == DraftWorkbookEditOutcome.ChangedOnDisk;

            this.TryAgainButton.IsVisible = state == DraftSaveState.Failed && !changedOutside;
            this.ChangedOnDiskBar.IsVisible = changedOutside;

            this.NoteUnsavedState();
        }

        // The save running on the pool now - for a test to wait for.
        internal Task<DraftWorkbookEditOutcome>? SaveInFlightForTests => this.thisSaveInFlight;

        // "Try again": the failed content saved again now.
        internal async Task<DraftWorkbookEditOutcome> TryAgainAsync()
        {
            this.thisDraftSession?.Retry();
            return await this.SaveDraftSessionAsync();
        }

        // "Keep the table's version": saved over the file that changed outside the table.
        internal Task<DraftWorkbookEditOutcome> KeepTableVersionAsync() =>
            this.SaveDraftSessionAsync(overwriteOutsideChange: true);

        private async void OnTryAgainClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.TryAgainAsync();
                this.FocusGrid();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler (AsyncVoidHandlersGuardTests).
                CrtLog.Critical($"OnTryAgainClick failed - [{ex}]");
            }
        }

        private async void OnKeepTableVersionClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.KeepTableVersionAsync();
                this.FocusGrid();
            }
            catch (Exception ex)
            {
                CrtLog.Critical($"OnKeepTableVersionClick failed - [{ex}]");
            }
        }

        private void OnUseSavedVersionClick(object? sender, RoutedEventArgs e)
        {
            if (this.DiscardUnsavedChanges())
            {
                this.ShowStatus(DraftSaveWording.UsedSavedVersion);
            }

            this.FocusGrid();
        }

        // ---- on screen and off ------------------------------------------------------------------

        // ###########################################################################################
        // *** THE CHECKS AGAIN, WITH NOTHING EDITED (code review, 2026-10-04). *** A picture the
        // table marked "missing" may have been put in the draft folder meanwhile - the board file did
        // not change - and the mark stayed while the draft's row and Submit found nothing wrong. The
        // host calls this when the table comes back into view (the Contribute tab shown, CRT's window in
        // front again): the file lookup looks again for every file it did not find
        // (DiskFileLookup), so only those cost anything. Not under a cell being typed in or a row
        // being dragged - the next time decides.
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

        protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnAttachedToVisualTree(e);

            // Another part of CRT may have written the draft while the table was away.
            this.ReadDraftFileAgain();
        }

        protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnDetachedFromVisualTree(e);

            // A card left open over a table that has left the screen would float over whatever
            // replaced it.
            this.HideFilePreview();
        }
    }

    // ###########################################################################################
    // The table's mark on a step it handed the draft's session: which table, and where the change
    // was - its sheet, the row's place on it, and the column (-1 for a whole row). An undo or a redo
    // of that step puts the cursor there.
    // ###########################################################################################
    internal sealed record BoardTableSessionEdit(BoardTableEditor Editor, string Sheet, int Row, int Column);
}
