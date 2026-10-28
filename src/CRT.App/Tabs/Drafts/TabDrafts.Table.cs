using Avalonia.Controls;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // TABLE MODE on the Contribute tab - "Table view" (owner request, 2026-09-24; on the right of the
    // workspace since CRT 4.0 step 2).
    //
    // Owns: which draft's table is open, showing it or the start panel, and what happens to the
    // table's changes when it closes or the draft's file is acted on. The table itself is
    // BoardTableEditor; this part only decides when it opens and closes.
    //
    // *** THE TABLE SAVES ITSELF (CRT 4.0 step 1, 2026-10-25). *** It works on its draft's session
    // (DraftManager.Sessions), which saves a second after every change - so closing the table,
    // submitting, managing files and quitting first save whatever is still waiting (SaveTableNowAsync),
    // and ask only when that save could not be made: Discard (the file's version, as an undoable
    // step) or Cancel, with the reason. A submission never goes without what is on screen, and an
    // import never lands under changes it would make unsaveable.
    //
    // *** EVERY OTHER EDITOR CHANGES THE SAME DRAFT IN MEMORY (DraftEdits, 2026-10-25). *** The
    // component editor, the label editor, the calibration, "Schematic images" and "Bring my draft up
    // to date" go through the session the table works on, so the table shows each change as it is
    // made. A file written any other way is read again when CRT's window is activated, and after a
    // reload of the board on screen (ReadTableDraftAgain) - one undoable step.
    // ###########################################################################################
    public partial class TabDrafts
    {
        internal bool IsTableOpen => this.thisTableEntry is not null;

        internal bool HasUnsavedTableEdits => this.IsTableOpen && this.TableEditor.HasUnsavedChanges;

        // ###########################################################################################
        // The same, with a cell still in its editor committed first (code review, 2026-10-13). The
        // grid does not commit when focus moves to a control outside it - Submit, "Close table",
        // quitting CRT - so a value typed and not yet confirmed was neither counted as unsaved nor
        // saved: quitting asked nothing and the value was gone. For the moments a decision is made;
        // HasUnsavedTableEdits itself does not commit, since background checks read it and must not
        // close an editor under the contributor's typing.
        // ###########################################################################################
        internal bool CommitCellsAndHasUnsavedTableEdits()
        {
            if (this.IsTableOpen)
            {
                this.TableEditor.CommitCellEdit();
            }

            // The component pane's boxes too (CRT 4.0 step 2): what they hold goes into its draft, which
            // then has a change not saved yet.
            if (this.IsComponentOpen)
            {
                this.ComponentPane.CommitHeld();
            }

            return this.HasUnsavedTableEdits ||
                   (this.IsComponentOpen && this.ComponentPane.HasUnsavedChanges) ||
                   TabDrafts.OtherUnsavedDrafts().Count > 0;
        }

        // ###########################################################################################
        // Every draft held in memory with changes not saved - the table's and the pane's, and one
        // neither shows any more: a save the pane started as it closed that failed, say. Without
        // those, quitting asked nothing and the change was gone.
        // ###########################################################################################
        private static List<DraftSession> OtherUnsavedDrafts() =>
            DraftManager.Sessions.Open
                .Where(session => session.HasUnsavedChanges && TabDrafts.IsOfCurrentDraftsFolder(session))
                .ToList();

        // ###########################################################################################
        // *** EVERY DRAFT HELD IN MEMORY IS SAVED WHEN DUE (code review, 2026-10-26). *** The table and
        // the pane save their own draft by their own timer - their Saved events are what the draft's
        // row and the board on screen follow - and a change in a draft neither shows was saved only if
        // whoever let go of it started a save. This one timer saves those: every session with a save
        // due (DraftSessions.Due) that neither holds, of the current drafts folder. A failed save waits
        // for a change or Try again, as everywhere (DraftSession.IsSaveDue). Saves run one after another,
        // off the UI thread (DraftSession.SaveAsync).
        // ###########################################################################################
        private readonly DispatcherTimer thisDraftSaveTimer = new() { Interval = TimeSpan.FromMilliseconds(500) };

        private Task thisDueDraftsSave = Task.CompletedTask;

        private void StartDraftSaveTimer()
        {
            this.thisDraftSaveTimer.Tick += (_, _) =>
            {
                if (this.thisDueDraftsSave.IsCompleted)
                {
                    this.thisDueDraftsSave = this.SaveDueDraftsNoEditorHoldsAsync(DateTimeOffset.UtcNow);
                }
            };

            this.thisDraftSaveTimer.Start();
        }

        internal async Task SaveDueDraftsNoEditorHoldsAsync(DateTimeOffset now)
        {
            DraftSession? table = this.thisTableEntry is { } tableEntry
                ? DraftManager.Sessions.Find(DraftManager.DraftsRoot, tableEntry.ExcelDataFile)
                : null;
            DraftSession? pane = this.ComponentPane.Session;

            IEnumerable<DraftSession> heldByNobody = DraftManager.Sessions.Due(now)
                .Where(session => !ReferenceEquals(session, table) && !ReferenceEquals(session, pane) && TabDrafts.IsOfCurrentDraftsFolder(session));

            foreach (DraftSession session in heldByNobody)
            {
                await session.SaveAsync();
            }
        }

        // Whether a session is of the drafts folder CRT uses now - DraftManager.Sessions outlives a
        // test's drafts folder, and a session of another must never be touched from here.
        private static bool IsOfCurrentDraftsFolder(DraftSession session) =>
            string.Equals(
                Path.GetFullPath(session.DraftsRoot),
                Path.GetFullPath(DraftManager.DraftsRoot.Length == 0 ? "." : DraftManager.DraftsRoot),
                StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // Whether the table is open on THIS board's draft and holds changes not saved yet - asked
        // before the draft is retired once published (Main.RetirePublishedDraftsAsync's isInUse). The
        // other editors no longer ask (CRT 4.0 step 1, 2026-10-25): they change the same draft held in
        // memory, so their saves land on top of the table's changes.
        // ###########################################################################################
        internal bool HasUnsavedTableEditsFor(string? excelDataFile) =>
            (this.HasUnsavedTableEdits &&
             this.thisTableEntry is not null &&
             string.Equals(this.thisTableEntry.ExcelDataFile, excelDataFile?.Trim(), StringComparison.OrdinalIgnoreCase)) ||
            (this.thisComponentEntry is not null &&
             string.Equals(this.thisComponentEntry.ExcelDataFile, excelDataFile?.Trim(), StringComparison.OrdinalIgnoreCase) &&
             this.ComponentPane.HasUnsavedChanges);

        private bool IsTableOpenFor(HardwareBoardEntry entry) =>
            this.thisTableEntry is not null &&
            string.Equals(this.thisTableEntry.ExcelDataFile, entry.ExcelDataFile, StringComparison.OrdinalIgnoreCase);

        // The row's "Edit in table format" / "Close table" button.
        internal async Task ToggleTableAsync(HardwareBoardEntry entry)
        {
            if (this.IsTableOpenFor(entry))
            {
                await this.CloseTableAsync();
            }
            else
            {
                await this.OpenTableAsync(entry);
            }
        }

        // Opens this draft's table, or keeps it open when it already is - for the Contribute tab's
        // "Table view" and the board parts on the left of the tab (TabDrafts.Workspace.cs).
        internal async Task ShowTableAsync(HardwareBoardEntry entry)
        {
            if (this.IsTableOpenFor(entry))
            {
                this.FocusTableIfOpen();
                return;
            }

            await this.OpenTableAsync(entry);
        }

        // ###########################################################################################
        // Opens one draft's table. The published board is what the colours compare against - none
        // for a board that exists only as a draft, where nothing is coloured.
        // ###########################################################################################
        internal async Task OpenTableAsync(HardwareBoardEntry entry)
        {
            if (this.IsTableOpen && !await this.CloseTableAsync())
            {
                return;
            }

            // A picture's name the Pictures part cannot take keeps it open (OpenPaneAsync has why).
            if (this.IsComponentOpen && this.ComponentPane.KeepsRefusedPictureName())
            {
                return;
            }

            // The right side shows one thing: the component pane goes, what it holds handed to the
            // draft the table is about to show - and saved before the table reads the draft's file,
            // under CRT's "please wait" (code review, 2026-10-26): the contributor pressed for the
            // table, and a large board's save takes a second or two.
            this.CloseComponent();

            Task paneSave = this.ComponentPane.LastSave;

            if (!paneSave.IsCompleted)
            {
                await BusyOverlay.RunLocalAsync(this, CrtWaitWording.SavingDraft, () => paneSave);
            }

            DraftStatus? status = DraftStatusReader.Resolve(
                DataManager.DataRoot,
                DraftManager.DraftsRoot,
                entry.ExcelDataFile,
                entry.IsPublished);

            // Reading the published board to colour against is a noticeable wait on a large board,
            // so it runs under the "please wait" overlay (2026-09-28).
            BoardData? published = this.PublishedBoardOverrideForTests is not null
                ? this.PublishedBoardOverrideForTests(entry)
                : await BusyOverlay.RunLocalAsync(this, CrtWaitWording.OpeningTable, () => TabDrafts.LoadPublishedBoardAsync(entry, status));

            // The published side is named after the source it was downloaded from (owner request,
            // 2026-10-05): a value BETA holds and stable does not yet is otherwise taken for a mistake.
            bool betaSource = this.BetaSourceOverrideForTests ?? UserSettings.DownloadDataFromTestSource;

            // Resting on a file cell shows the file - the published copy and the draft's.
            this.TableEditor.FileSource = new DraftTableFileSource(
                DataManager.DataRoot,
                DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, entry.ExcelDataFile),
                betaSource,
                entry.ExcelDataFile);

            // ###########################################################################################
            // *** THE DRAFT HELD IN MEMORY, SHARED (CRT 4.0 step 1). *** The one session of this draft
            // (DraftManager.Sessions) - kept for the whole run, so what the table did is still one
            // Ctrl+Z away when it is opened again. Something may have written the file while the
            // table was closed: read again first, when nothing of the session's waits to be saved.
            // ###########################################################################################
            DraftSession? session = DraftManager.Sessions.Get(DraftManager.DraftsRoot, entry.ExcelDataFile);
            session?.ReadFileIfChanged(DraftSaveWording.ReadAgainStep);

            // The downloaded data too, so the table's checks find a row's files where a submit
            // finds them (owner request, 2026-10-02).
            if (session is null ||
                !this.TableEditor.Load(
                    session,
                    published,
                    DataManager.DataRoot,
                    BoardTableDocument.SourceBaselineLabel(betaSource)))
            {
                Logger.Warning($"Could not open the table for the draft of [{entry.ExcelDataFile}] - its board file could not be read");
                return;
            }

            this.thisTableEntry = entry;
            this.RefreshDrafts();
            this.FocusTableIfOpen();
        }

        // ###########################################################################################
        // Puts keyboard focus in the open table, so typing lands in a cell rather than in the
        // always-on component filter on the left (see Main.ShouldReturnFocusToComponentSearch, which
        // keeps it there once it is). Called when a table opens and when the tab is shown again.
        //
        // Posted, not immediate: a table just opened, or a tab just switched to, has not been laid
        // out yet, and a control that is not yet visible cannot take focus.
        // ###########################################################################################
        internal void FocusTableIfOpen()
        {
            if (this.IsTableOpen)
            {
                Dispatcher.UIThread.Post(this.TableEditor.FocusGrid, DispatcherPriority.Background);
            }
        }

        // ###########################################################################################
        // Closes the table - what it holds saved first, and a question only when that save could not
        // be made. False when the contributor chose to stay (Cancel): the table stays open with the
        // reason on screen.
        //
        // *** ITS OWN BOARD'S DRAFT ONLY (code review, 2026-10-26). *** It settled every draft held in
        // memory (ConfirmLeavingTableAsync, quitting's) - so another board's change that could not be
        // saved was asked about as the table's, at every close, and Discard read that other draft's
        // file back without it ever being shown. That one waits for its own board, or quitting.
        // ###########################################################################################
        internal async Task<bool> CloseTableAsync()
        {
            if (this.thisTableEntry is not { } entry)
            {
                return true;
            }

            if (!await this.SettleTableEditsAsync(entry))
            {
                return false;
            }

            this.CloseTableWithoutAsking();

            // A draft made only to open the table goes again, nothing in it (TabDrafts.Workspace.cs).
            this.LetGoOfDraftMadeOnlyToLook(entry);
            return true;
        }

        // ###########################################################################################
        // Closes the table, without asking, if it is open on this board's draft - for a draft that
        // is about to be deleted by something other than this tab (retirement after publishing).
        // The caller has already made sure nothing unsaved is lost: see HasUnsavedTableEditsFor.
        // ###########################################################################################
        internal void CloseTableIfOpenFor(string? excelDataFile)
        {
            if (this.thisTableEntry is not null &&
                string.Equals(this.thisTableEntry.ExcelDataFile, excelDataFile?.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                this.CloseTableWithoutAsking();
            }

            // The component pane lets go of it too - nothing may be written into a draft that goes.
            if (this.thisComponentEntry is not null &&
                string.Equals(this.thisComponentEntry.ExcelDataFile, excelDataFile?.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                this.ForgetComponent();
                this.ApplyTableMode();
            }
        }

        private void CloseTableWithoutAsking()
        {
            this.TableEditor.Clear();
            this.thisTableEntry = null;
            this.RefreshDrafts();
            this.CatchUpBoardOnScreen();
        }

        // ###########################################################################################
        // Everything the table holds saved now - under CRT's "please wait", since the contributor
        // pressed for it and the largest board takes a second or two. True when nothing is left
        // unsaved (or no table is open).
        // ###########################################################################################
        internal async Task<bool> SaveTableNowAsync()
        {
            if (!this.IsTableOpen)
            {
                return true;
            }

            this.TableEditor.CommitCellEdit();

            if (!this.TableEditor.HasUnsavedChanges)
            {
                return true;
            }

            return await BusyOverlay.RunLocalAsync(this, CrtWaitWording.SavingDraft, this.TableEditor.SaveNowAsync);
        }

        // ###########################################################################################
        // What the table AND the component pane hold, saved now (CRT 4.0 step 2) - for quitting and
        // installing an update (Main.ConfirmLeavingTablesAsync). True when nothing is left unsaved.
        // ###########################################################################################
        internal async Task<bool> SaveDraftEditorsNowAsync()
        {
            // Both, whatever the first says - each holds changes of its own.
            bool tableSaved = await this.SaveTableNowAsync();
            bool paneSaved = await this.SaveComponentNowAsync();

            // Then any other draft held in memory with changes not saved (OtherUnsavedDrafts).
            bool othersSaved = true;

            foreach (DraftSession session in TabDrafts.OtherUnsavedDrafts())
            {
                DraftWorkbookEditOutcome outcome = await BusyOverlay.RunLocalAsync(this, CrtWaitWording.SavingDraft, () => session.SaveAsync());
                othersSaved &= outcome == DraftWorkbookEditOutcome.Saved;
            }

            return tableSaved && paneSaved && othersSaved;
        }

        // ###########################################################################################
        // True when it is fine to leave the table and the component pane: everything saved - now, if
        // a change was still waiting - or, when that save could not be made, the contributor chose to
        // discard what could not be saved. EVERY draft held in memory - for quitting and installing an
        // update (Main.ConfirmLeavingTablesAsync); closing the table settles its own board's only
        // (CloseTableAsync). `owner` is for Main, which asks while this tab may not be the one on screen.
        // ###########################################################################################
        internal async Task<bool> ConfirmLeavingTableAsync(Window? owner = null)
        {
            if (await this.SaveDraftEditorsNowAsync())
            {
                return true;
            }

            if (await this.AskAboutUnsavedTableEditsAsync(owner) != UnsavedTableEditsChoice.Discard)
            {
                return false;
            }

            this.DiscardUnsavedDraftChanges();
            return true;
        }

        // ###########################################################################################
        // Before an action on the draft FILE (importing files, submitting): what the table or the
        // component pane holds is saved, or - when that cannot be done - thrown away at the
        // contributor's word, either staying open. True when the action may go ahead.
        //
        // *** AND THE DRAFT HELD IN MEMORY ITSELF (code review, 2026-10-26). *** A change neither the
        // table nor the pane shows waits there too - the pane's close-save refused, the header's Undo
        // with nothing open - and Submit sent the board file without it. Only this board's draft is
        // saved, asked about and thrown away; another board's waits for its own action or quitting.
        // ###########################################################################################
        internal async Task<bool> SettleTableEditsAsync(HardwareBoardEntry entry)
        {
            bool tableSaved = !this.IsTableOpenFor(entry) || await this.SaveTableNowAsync();
            bool paneSaved = !this.IsComponentOpenFor(entry) || await this.SaveComponentNowAsync();

            DraftSession? session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile);
            bool draftSaved = session is null || !session.HasUnsavedChanges ||
                              await BusyOverlay.RunLocalAsync(this, CrtWaitWording.SavingDraft, () => session.SaveAsync()) == DraftWorkbookEditOutcome.Saved;

            if (tableSaved && paneSaved && draftSaved)
            {
                return true;
            }

            if (await this.AskAboutUnsavedTableEditsAsync(owner: null, onlyOf: session) != UnsavedTableEditsChoice.Discard)
            {
                return false;
            }

            this.DiscardUnsavedDraftChanges(onlyOf: session);
            return true;
        }

        // What could not be saved, thrown away - the table's, and the component pane's draft read again.
        // `onlyOf`: of the other drafts held in memory, only that one - an action on one board's file.
        private void DiscardUnsavedDraftChanges(DraftSession? onlyOf = null)
        {
            if (this.IsTableOpen && this.TableEditor.HasUnsavedChanges)
            {
                this.DiscardTableChanges();
            }

            if (this.IsComponentOpen && this.ComponentPane.HasUnsavedChanges && !this.ComponentPane.DiscardUnsavedChanges() &&
                this.thisComponentEntry is { } entry)
            {
                DraftManager.Sessions.Close(DraftManager.DraftsRoot, entry.ExcelDataFile);
                this.ForgetComponent();
                this.ApplyTableMode();
            }

            // Any other draft held in memory that could not be saved: its file read again, or - when
            // that cannot be read - let go of, so nothing saves it back later.
            foreach (DraftSession session in TabDrafts.OtherUnsavedDrafts().Where(session => onlyOf is null || ReferenceEquals(session, onlyOf)))
            {
                if (!session.ReloadFromFile(DraftSaveWording.ReadAgainStep))
                {
                    DraftManager.Sessions.Close(session.DraftsRoot, session.ExcelDataFile);
                }
            }
        }

        // ###########################################################################################
        // The table's unsaved changes thrown away: the draft's file read again - one undoable step
        // (BoardTableEditor.DiscardUnsavedChanges). A file that cannot be read lets the draft's
        // session go instead, so nothing saves it back later.
        // ###########################################################################################
        private void DiscardTableChanges()
        {
            if (!this.TableEditor.DiscardUnsavedChanges() && this.thisTableEntry is { } entry)
            {
                DraftManager.Sessions.Close(DraftManager.DraftsRoot, entry.ExcelDataFile);
            }
        }

        // ###########################################################################################
        // Asked only once a save could not be made, so there is no Save to offer: Discard or Cancel,
        // and why it was not saved - changed outside the table, or not writable. `onlyOf`: see
        // DiscardUnsavedDraftChanges.
        // ###########################################################################################
        private async Task<UnsavedTableEditsChoice> AskAboutUnsavedTableEditsAsync(Window? owner, DraftSession? onlyOf = null)
        {
            DraftWorkbookEditOutcome? failure = this.IsTableOpen && this.TableEditor.HasUnsavedChanges
                ? this.TableEditor.SaveFailure
                : this.IsComponentOpen && this.ComponentPane.HasUnsavedChanges
                    ? this.ComponentPane.SaveFailure
                    : TabDrafts.OtherUnsavedDrafts()
                        .Where(session => onlyOf is null || ReferenceEquals(session, onlyOf))
                        .Select(session => session.SaveFailure)
                        .FirstOrDefault(outcome => outcome is not null);

            UnsavedTableEditsPrompt prompt = failure == DraftWorkbookEditOutcome.ChangedOnDisk
                ? UnsavedTableEditsPrompt.DraftChangedOnDisk
                : UnsavedTableEditsPrompt.DraftNotWritten;

            if (this.UnsavedTableEditsAnswerForTests is not null)
            {
                return this.UnsavedTableEditsAnswerForTests(prompt);
            }

            Window? ownerWindow = owner ?? TopLevel.GetTopLevel(this) as Window;
            if (ownerWindow is null)
            {
                return UnsavedTableEditsChoice.Cancel;
            }

            var window = new UnsavedTableEditsWindow();
            window.Initialize(prompt);

            return await window.ShowDialog<UnsavedTableEditsChoice?>(ownerWindow) ?? UnsavedTableEditsChoice.Cancel;
        }

        // ###########################################################################################
        // The draft's file may have been written past the draft held in memory - the editors no longer
        // do (DraftEdits), but whatever else reloads the board on screen might have - so an open table
        // reads it again when it changed (BoardTableEditor.ReadDraftFileAgain: one undoable step, and
        // nothing while a change of the table's waits). Main calls this from ReloadCurrentBoardFromDisk.
        // ###########################################################################################
        internal void ReadTableDraftAgain()
        {
            if (this.IsTableOpen)
            {
                this.TableEditor.ReadDraftFileAgain();
            }
        }

        // ###########################################################################################
        // The table on the right, or the start panel (CRT 4.0 step 2) - the header shows the draft the
        // table is open on (TabDrafts.Workspace.cs). The table is never inside a ScrollViewer: it
        // would lay out every row.
        //
        // Runs at the end of every RefreshDrafts. If the open draft is no longer listed (discarded
        // from elsewhere, or published and retired), the table closes: there is nothing left for
        // it to save into.
        // ###########################################################################################
        private void ApplyTableMode()
        {
            bool listed = this.thisTableEntry is not null && this.Drafts.Any(item =>
                string.Equals(item.ExcelDataFile, this.thisTableEntry.ExcelDataFile, StringComparison.OrdinalIgnoreCase));

            if (this.thisTableEntry is not null && !listed)
            {
                this.TableEditor.Clear();
                this.thisTableEntry = null;
            }

            // The component pane's draft too - gone, so nothing is written into it.
            if (this.thisComponentEntry is { } paneEntry &&
                !this.Drafts.Any(item => string.Equals(item.ExcelDataFile, paneEntry.ExcelDataFile, StringComparison.OrdinalIgnoreCase)))
            {
                this.ForgetComponent();
            }

            this.TableEditor.IsVisible = this.IsTableOpen;
            this.ComponentPane.IsVisible = !this.IsTableOpen && this.IsComponentOpen;
            this.ShowStartPanel(!this.IsTableOpen && !this.IsComponentOpen);
        }

        // ###########################################################################################
        // The table's draft saved itself (or was read again because something outside the table
        // changed it): the draft's row here follows at once. The board on screen - its component list
        // and schematics - follows once the table is closed or the tab is left (CatchUpBoardOnScreen):
        // a draft saves a second after every change, and reading the whole board again each time
        // would make the table lag while it is typed into. The board was read from the file before;
        // until the explicit save went, it waited for that.
        // ###########################################################################################
        private void OnTableSaved(object? sender, EventArgs e)
        {
            this.thisBoardOnScreenBehind = this.TableEditor.ExcelDataFile;

            if (this.thisMainWindow is not null)
            {
                this.thisMainWindow.ApplyContributeTabState();
            }
            else
            {
                this.RefreshDrafts();
            }
        }

        // The draft whose saves the board on screen has not caught up with yet - see OnTableSaved.
        private string? thisBoardOnScreenBehind;

        // The board on screen read again when the table's saves have changed its draft.
        internal void CatchUpBoardOnScreen()
        {
            if (this.thisBoardOnScreenBehind is not { } excelDataFile)
            {
                return;
            }

            this.thisBoardOnScreenBehind = null;

            if (this.ReloadBoardOverrideForTests is { } reload)
            {
                reload(excelDataFile);
            }
            else
            {
                this.thisMainWindow?.ReloadCurrentBoardAfterDraftFileImport(excelDataFile);
            }
        }

        // The published board to colour against, or null when there is none (a draft-only board,
        // or a published file that is missing - the table then simply shows no colours rather
        // than calling every row an addition).
        private static async Task<BoardData?> LoadPublishedBoardAsync(HardwareBoardEntry entry, DraftStatus? status)
        {
            if (entry.IsDraftOnly || status?.IsNewBoard == true)
            {
                return null;
            }

            string publishedPath = DraftBoardSource.PublishedPathOf(DataManager.DataRoot, entry.ExcelDataFile);
            if (publishedPath.Length == 0 || !File.Exists(publishedPath))
            {
                return null;
            }

            return await BoardDataReader.LoadAsync(publishedPath, publishedPath);
        }
    }
}
