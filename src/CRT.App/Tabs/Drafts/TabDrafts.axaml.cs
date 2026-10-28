using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using Handlers.Online;
using System;
using System.Globalization;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;

namespace CRT
{
    // ###########################################################################################
    // THE CONTRIBUTE TAB - the class keeps the Drafts tab's name (CRT 4.0 step 2, 2026-10-25 - the
    // plan: "One tab instead of Contribute and Drafts"). It is about the board shown in CRT's
    // drop-down lists: its draft's header (or what it is, with none), its parts and components on
    // the left, and the table view, a component or a start panel on the right (TabDrafts.Workspace.cs). "My
    // drafts" lists every board with a local, unpublished draft (`Drafts`, which Main also reads for
    // the drop-downs' chips and the tab's title) and shows the one picked. Always on screen.
    //
    // *** TABLE MODE (2026-09-24). *** "Table view" opens the board's draft as editable sheets
    // (BoardTableEditor) on the right; the header shows the draft the table is open on.
    //
    // The table saves itself (CRT 4.0 step 1, 2026-10-25). Anything that acts on the draft's file
    // first saves what is still waiting, and asks only when that save cannot be made
    // (ConfirmLeavingTableAsync, SettleTableEditsAsync): closing the table, managing its files,
    // submitting it (a submission must include what is on screen), and quitting the application
    // (Main.OnWindowClosing). Discarding the draft throws the table away with it.
    //
    // FILE MAP:
    //   TabDrafts.axaml.cs  - the drafts' rows and their actions (discard, files, drift, submit), the badge
    //   TabDrafts.Workspace.cs - the board the tab is about: its header, the left column (its parts,
    //                         its components, "Add component", "Table view") and "My drafts"
    //   TabDrafts.Table.cs  - table mode: opening/closing the table on its draft's session, saving
    //                         it (and the component pane) before the draft's file is acted on, and the
    //                         question when that fails
    //   TabDrafts.Component.cs - the component pane on the right (ComponentPane): opening it on a
    //                         component or on a new one's form, closing, what follows its changes,
    //                         and its "Show on board" and "Preview popup"
    //   TabDrafts.Header.cs - the draft's header on its row: save state, Undo/Redo, Try again,
    //                         "Review changes" with "Undo changes to <component>", and the summary
    //                         the Submit and Discard draft dialogs show
    //   TabDrafts.OutsideChanges.cs - reading the list again when the tab is shown or CRT's window
    //                         comes back to the front, so a draft edited in Excel shows its new numbers
    //   TabDrafts.UpdateRequired.cs - "CRT has to be updated" over the whole tab, when the server
    //                         turns this version of CRT away from submissions
    //   TabDrafts.SubmitOthers.cs - the other drafts the Submit dialog offers to send with one, and
    //                         making each ready once Submit is pressed
    //   TabDrafts.Readiness.cs - the start panel's "Getting this board ready" checklist of a new
    //                         board, and "What's missing" on an existing one
    //   TabDrafts.MarkAutomatically.cs - a picture's "Mark all automatically": every component's
    //                         rectangle from its KiCad footprint, on a picture aligned with KiCad
    //   TabDrafts.Drop.cs   - pictures or the KiCad project folder dropped on the tab, added to the
    //                         board shown through the window that manages them
    // ###########################################################################################
    public partial class TabDrafts : UserControl
    {
        private Main? thisMainWindow;

        // The board whose table is open, or null when the tab is showing its ordinary list.
        private HardwareBoardEntry? thisTableEntry;

        // Lets a headless test answer the unsaved-edits prompt, which would otherwise block on
        // ShowDialog. null (the default) is the shipped path: the real UnsavedTableEditsWindow.
        internal Func<UnsavedTableEditsPrompt, UnsavedTableEditsChoice>? UnsavedTableEditsAnswerForTests { get; set; }

        // Lets a headless test open a table without a real published board on disk. null is the
        // shipped path: the published workbook, read through BoardDataReader's cache.
        internal Func<HardwareBoardEntry, BoardData?>? PublishedBoardOverrideForTests { get; set; }

        // Lets a headless test say which source the data came from without touching UserSettings.
        // null is the shipped path: "Download data from the BETA source" as set.
        internal bool? BetaSourceOverrideForTests { get; set; }

        // The same override-then-real-static pattern TabWorkbooks.BoardKeyOverrideForTests uses,
        // and for the same reason: RefreshDrafts normally reads DataManager.HardwareBoards, a
        // static list a headless test cannot seed without loading a real main Excel workbook.
        // null (the default) leaves RefreshDrafts reading the real static exactly as the shipped
        // app does.
        internal List<HardwareBoardEntry>? HardwareBoardsOverrideForTests { get; set; }

        // Answers the "draft discarded" notice instead of the server (2026-09-28) - a test sends
        // nothing over the network (test rule 6).
        internal Func<long, string, System.Threading.CancellationToken, Task<int?>>? DraftDiscardSendOverrideForTests { get; set; }

        // Stands in for the delete itself - a test can make it stop part-way as a locked workbook does.
        internal Func<string, bool>? DiscardOverrideForTests { get; set; }

        public ObservableCollection<DraftListItem> Drafts { get; } = new();

        // The boards behind the rows above, as of the last RefreshDrafts. Main builds the Hardware
        // and Board drop-downs' "Draft" chips from exactly this list (see Main.DraftBadges.cs), so a
        // chip appears for precisely the boards this tab lists and for no others.
        private readonly List<HardwareBoardEntry> thisDraftedBoards = new();

        internal IReadOnlyList<HardwareBoardEntry> DraftedBoards => this.thisDraftedBoards;

        public TabDrafts()
        {
            this.InitializeComponent();

            this.MyDraftsComboBox.ItemsSource = this.Drafts;

            this.TableEditor.Saved += this.OnTableSaved;

            // The table reading its draft again, because something outside it changed the file,
            // needs the same refreshes as a save.
            this.TableEditor.ReloadedFromOutside += this.OnTableSaved;

            this.WireComponentPane();

            // Pictures or the KiCad project dropped anywhere on the tab (TabDrafts.Drop.cs).
            this.WireDrop();
        }

        // ###########################################################################################
        // Called once from Main's constructor, the same pattern every other tab's Initialize uses.
        // ###########################################################################################
        internal void Initialize(Main mainWindow)
        {
            this.thisMainWindow = mainWindow;

            // Every draft no editor holds, saved when due (TabDrafts.Table.cs) - for the run, from here:
            // a tab a test builds has no Main and no timer.
            this.StartDraftSaveTimer();

            // The component pane's "Capture from my oscilloscope" (TabDrafts.Component.cs).
            this.UseScopeCapture(
                mainWindow.TabOscilloscopeControl.CaptureOscilloscopePngAsync,
                mainWindow.TabOscilloscopeControl.HasActiveEstablishedOscilloscopeSessionForTitleState());

            // The Credits part's "Add me" - by the name and contact the settings remember (2026-10-26).
            this.ComponentPane.CreditIdentity = () => (UserSettings.CreditName, UserSettings.CreditContact);
            this.ComponentPane.RememberCreditIdentity = UserSettings.RememberCreditIdentity;
        }

        // ###########################################################################################
        // Rebuilds the drafted-boards list from disk and shows/hides the tab to match - called from
        // Main whenever the set of drafts could have changed: at startup, after a board load (a
        // draft's own row count could have changed since the tab was last shown), and after a
        // discard on this tab itself.
        //
        // internal virtual would be the more MVVM-flavoured seam, but this codebase has none (see
        // CLAUDE.md) - a plain method the caller (Main) drives directly, same as TabResources.LoadData.
        // ###########################################################################################
        internal void RefreshDrafts()
        {
            this.RefreshSubmissionBadge();

            // The rows read now - put on the list by KeepUnchangedRows, which keeps every row
            // that would draw the same.
            var fresh = new List<DraftListItem>();
            this.thisDraftedBoards.Clear();

            var hardwareBoards = this.HardwareBoardsOverrideForTests ?? DataManager.HardwareBoards;
            var draftedBoards = DraftManager.EnumerateDraftedBoards(hardwareBoards);

            // Read once for every row - each row's badge is its board's latest submission.
            IReadOnlyList<SubmissionReceipt> receipts = this.ReceiptsOverrideForTests ?? SubmissionReceiptStore.All;

            foreach (var entry in draftedBoards.OrderBy(e => e.ShortHardwareBoardLabel, StringComparer.OrdinalIgnoreCase))
            {
                DraftStatus? status = DraftStatusReader.Resolve(
                    DataManager.DataRoot,
                    DraftManager.DraftsRoot,
                    entry.ExcelDataFile,
                    entry.IsPublished);

                if (status == null)
                {
                    // Vanished between EnumerateDraftedBoards' own read and this one (discarded
                    // from another surface, or the folder was deleted by hand) - skip rather than
                    // show a row with nothing behind it.
                    continue;
                }

                // *** THE COUNT IS DERIVED, AND A FRESH ONE COSTS TWO BOARD PARSES (Phase 6). ***
                // It used to be a cheap sum of the draft's own delta rows. There is no delta list
                // any more - the contributor can edit the workbook in Excel, so the only honest
                // answer is to compare the two workbooks.
                //
                // *** CountChangesCached, NOT CountChanges (code review, 2026-09-25). *** This runs
                // after every save of any kind, for EVERY drafted board, on the UI thread - and
                // re-parsing boards nobody had touched froze the window for seconds per save. The
                // cached count is keyed by each file's size and write time, so a board edited
                // anywhere (Excel included) is still counted afresh.
                this.thisDraftedBoards.Add(entry);

                SubmissionReceipt? lastSubmission = SubmissionReceiptPresenter.LatestForBoard(
                    receipts,
                    BoardDescriptorRules.BoardIdFromExcelDataFile(entry.ExcelDataFile),
                    status.CreatedUtc);

                fresh.Add(new DraftListItem(
                    entry,
                    status,
                    DraftStatusReader.CountChangesCached(status),

                    // The table's own checks, on the row (owner report, 2026-10-02) - remembered
                    // until the workbook or its sidecar changes, so this costs nothing for a draft
                    // nobody touched, and catches an edit made in Excel.
                    DraftStatusReader.CountProblemsCached(
                        status,
                        DataManager.DataRoot,
                        DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, entry.ExcelDataFile)),
                    ResolveDriftState(entry, status),
                    this.DiscardAsync,
                    this.ViewDriftAsync,
                    this.SubmitAsync,
                    isTableOpen: this.IsTableOpenFor(entry),
                    lastSubmission,
                    TabDrafts.IsAlreadySent(entry, status, lastSubmission),

                    // What is not a row - files and KiCad calibrations (code review, 2026-10-13): a
                    // draft whose only work is a calibration or a picture is something to send.
                    otherChanges: DraftStatusReader.CountOtherChangesCached(status, DraftManager.DraftsRoot),

                    // The draft's header - its save state and Undo/Redo, kept for the run (TabDrafts.Header.cs).
                    header: this.HeaderFor(entry),
                    headerActions: this.HeaderActions));
            }

            this.ForgetHeadersExcept(this.thisDraftedBoards);

            TabDrafts.KeepUnchangedRows(this.Drafts, fresh);

            this.ApplyTableMode();

            // The header, the left column and "My drafts" follow (TabDrafts.Workspace.cs).
            this.ShowWorkspace();
        }

        // ###########################################################################################
        // *** THE LIST IS READ AGAIN IN PLACE, NOT REBUILT (owner report, 2026-10-10: "I do need
        // to click it twice before it does react"). ***
        //
        // RefreshDrafts runs whenever CRT's window comes to the front (TabDrafts.OutsideChanges.cs)
        // - and a click INTO the window is what brings it there, so it ran between that click's
        // press and its release. Replacing every row threw away the button being pressed, and the
        // release landed on a new one not even laid out yet: the click did nothing, and only the
        // second one worked. A row that draws the same (DraftListItem.ShowsTheSameAs) is now kept,
        // buttons and all (InPlaceList.KeepUnchanged); only a row that changed is replaced, and the
        // list grows or shrinks at its end.
        // ###########################################################################################
        internal static void KeepUnchangedRows(ObservableCollection<DraftListItem> rows, IReadOnlyList<DraftListItem> fresh) =>
            InPlaceList.KeepUnchanged(rows, fresh, (row, other) => row.ShowsTheSameAs(other));

        // ###########################################################################################
        // Confirms via DiscardDraftWindow (the same shape as DeleteWorkbookWindow - a permanent,
        // local action deserves the same "are you sure" this app gives a permanent delete anywhere
        // else), then discards and refreshes both this tab and its own visibility - the same way a
        // Workbooks card's delete refreshes the Workbooks tab afterward.
        // ###########################################################################################
        private async Task DiscardAsync(HardwareBoardEntry entry)
        {
            if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
            {
                return;
            }

            // This board's submissions still with the maintainers - read BEFORE the draft goes, since
            // "sent from this draft" is judged against the draft's own creation time.
            IReadOnlyList<SubmissionReceipt> unfinished = TabDrafts.UnfinishedSubmissionsOf(entry);

            // What goes, listed - the draft's summary in "Review changes"' words, a dozen lines at most.
            DraftChangeSummary lost = await BusyOverlay.RunLocalAsync(this, CrtWaitWording.ReadingChanges, () => this.SummaryForAsync(entry));

            var confirmWindow = new DiscardDraftWindow();
            confirmWindow.Initialize(entry.ToString(), unfinished, lost.AsLines(max: 12));

            bool? confirmed = await confirmWindow.ShowDialog<bool?>(ownerWindow);
            if (confirmed != true)
            {
                return;
            }

            this.DiscardConfirmed(entry, unfinished);
        }

        // Which of this board's submissions a discard is reported for (DraftDiscardContract).
        private static IReadOnlyList<SubmissionReceipt> UnfinishedSubmissionsOf(HardwareBoardEntry entry)
        {
            DraftStatus? status = DraftStatusReader.Resolve(
                DataManager.DataRoot,
                DraftManager.DraftsRoot,
                entry.ExcelDataFile,
                entry.IsPublished);

            return DraftDiscardContract.WhichToReport(
                SubmissionReceiptStore.All,
                BoardDescriptorRules.BoardIdFromExcelDataFile(entry.ExcelDataFile),
                status?.CreatedUtc);
        }

        // ###########################################################################################
        // Everything a confirmed Discard does. Split from the confirmation so a headless test can
        // drive the real path - the dialog above cannot be answered from a test.
        // ###########################################################################################
        //
        // `unfinished` (owner request, 2026-09-28): this board's submissions still with the
        // maintainers, who are told the draft was discarded - see DraftDiscardReporter.
        internal void DiscardConfirmed(HardwareBoardEntry entry, IReadOnlyList<SubmissionReceipt>? unfinished = null)
        {
            // The whole draft is going, so the table's unsaved edits go with it - the discard
            // confirmation above already said everything local is lost. Closed first so the table
            // is not left pointing at a folder that no longer exists.
            if (this.IsTableOpenFor(entry))
            {
                this.CloseTableWithoutAsking();
            }

            if (this.IsComponentOpenFor(entry))
            {
                this.ForgetComponent();
            }

            this.ShowDiscardFailure(null);

            // *** A FAILED DELETE IS REPORTED, AND THE REFRESH BELOW STILL RUNS. *** The usual cause
            // is the draft workbook being open in Excel, which locks it on Windows - part of the
            // folder may already be gone by then, so the list and the drop-downs must still catch
            // up with what is really on disk.
            bool discarded = this.DiscardOverrideForTests?.Invoke(entry.ExcelDataFile)
                ?? DraftManager.DiscardDraft(entry.ExcelDataFile);

            if (!discarded)
            {
                this.ShowDiscardFailure(entry.ToString());
            }

            // A discarded board that was DRAFT-ONLY (session 2c, task 9) has just stopped existing
            // entirely - not merely lost its edits. Without this it would linger in the
            // hardware/board drop-downs until the next restart, pointing at a folder that is gone.
            // Harmless for an ordinary draft over a synced board, which this leaves in place.
            DataManager.RefreshDraftOnlyBoards();

            // ###########################################################################################
            // *** THE MAINTAINERS ARE TOLD (owner request, 2026-09-28). *** Recorded on the receipts
            // first, so a notice that cannot be sent now (no network) is sent at the next launch;
            // then sent in the background - the discard itself never waits on the server.
            // ###########################################################################################
            // *** ONLY A DISCARD THAT FINISHED (code review, 2026-10-13). *** One that stopped
            // part-way (the workbook open in Excel) leaves the draft on the tab to be discarded again
            // - or kept - and telling the maintainers now would report a discard that may never happen.
            // The next Discard that finishes reports it: the list of unfinished submissions is read
            // afresh each time.
            if (discarded && unfinished is { Count: > 0 })
            {
                SubmissionReceiptStore.MarkDraftDiscarded(unfinished.Select(receipt => receipt.SubmissionId), DateTimeOffset.UtcNow);
                _ = DraftDiscardReporter.ReportPendingAsync(this.DraftDiscardSendOverrideForTests ?? new SubmissionClient().ReportDraftDiscardedAsync);
            }

            this.RefreshDrafts();
            this.thisMainWindow?.ApplyContributeTabState();
            this.thisMainWindow?.RefreshHardwareAndBoardSelectionsAfterDraftChange();
        }

        // ###########################################################################################
        // Shows (or, with null, hides) the line saying a discard did not finish. Worded for the
        // likely cause and the fix, since "could not delete" alone leaves the contributor with
        // nothing to do about it.
        // ###########################################################################################
        private void ShowDiscardFailure(string? boardDisplayName)
        {
            this.DiscardFailedText.IsVisible = boardDisplayName is not null;
            this.DiscardFailedText.Text = boardDisplayName is null
                ? string.Empty
                : $"Part of the draft for {boardDisplayName} could not be deleted. This usually means one of its " +
                  "files is open in another program. Close it there and press Discard draft again.";
        }

        // Lets a headless test read the line without reaching into the named control.
        internal string? DiscardFailureTextForTests =>
            this.DiscardFailedText.IsVisible ? this.DiscardFailedText.Text : null;

        // ###########################################################################################
        // Whether the official data has moved under one drafted board, resolved CHEAPLY - this
        // runs for every drafted board on every refresh, and RefreshDrafts is called at startup,
        // after every discard and after every file import.
        //
        // So it never loads a board: it reads the cached revision for a board already parsed (always
        // true of the selected one), and otherwise reads ONLY the revision date off the workbook
        // rather than its ten sheets (BoardDataReader.ReadRevisionDateOnly). A draft-only board is
        // skipped entirely - it has no official file, and drift is meaningless for it.
        //
        // The full per-row report is deliberately NOT built here. That needs the whole board and is
        // only worth it when the contributor actually asks to see it - see ViewDriftAsync.
        // ###########################################################################################
        private static DraftDriftState ResolveDriftState(HardwareBoardEntry entry, DraftStatus status)
        {
            if (entry.IsDraftOnly || status.IsNewBoard)
            {
                return DraftDriftState.Unknown;
            }

            // *** KEYED BY PATH, NOT BY THE BOARD IDENTITY (Phase 6). *** The board cache is keyed
            // by the file that was actually read, because a drafted board has two workbooks and
            // "view as officially published" switches between them. Asking with the ExcelDataFile
            // would always miss, quietly falling through to the re-read below - correct, but a
            // whole Excel open per board on a tab that lists them all.
            string excelPath = DraftBoardSource.PublishedPathOf(
                DataManager.DataRoot,
                entry.ExcelDataFile);

            // The one rule the row, the main window's notice and the window share - by CONTENT for
            // a draft that kept what it was copied from, by revision otherwise (DraftDriftCheck).
            return DraftDriftCheck.Resolve(status, excelPath);
        }

        // ###########################################################################################
        // Shows how one board's drafted rows line up against the official data as it stands now.
        //
        // This is the ONE place a full board load is justified for drift - the per-row report needs
        // the official BoardData, and the contributor has explicitly asked for it by clicking. The
        // same load-on-demand-from-a-button shape ManageFilesAsync already uses.
        // ###########################################################################################
        private async Task ViewDriftAsync(HardwareBoardEntry entry)
        {
            if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
            {
                return;
            }

            DraftStatus? status = DraftStatusReader.Resolve(
                DataManager.DataRoot,
                DraftManager.DraftsRoot,
                entry.ExcelDataFile,
                entry.IsPublished);

            if (status == null)
            {
                return;
            }

            // ###########################################################################################
            // *** THE REPORT KNOWS THE DRAFT'S BASE (code review, 2026-10-13). *** A draft seeded since
            // CRT keeps what it was copied from shows the contributor's changes and the published
            // update apart, the rows both changed, and offers to bring the draft up to date; an older
            // draft is compared with the published board as it is now, and the window says so
            // (DraftChangeReport.For - the same report the main window's notice opens).
            // ###########################################################################################
            string publishedPath = DraftBoardSource.PublishedPathOf(
                DataManager.DataRoot,
                entry.ExcelDataFile);

            string draftsRoot = DraftManager.DraftsRoot;

            // Reading two workbooks of a large board is a noticeable wait (2026-09-28).
            DraftChangeReport report = await BusyOverlay.RunLocalAsync(
                this,
                CrtWaitWording.ComparingWithOfficial,
                () => Task.Run(() => DraftChangeReport.For(status, draftsRoot, publishedPath)));

            var window = new DraftDriftWindow();
            window.Initialize(entry.ToString(), entry.ExcelDataFile, report);
            await window.ShowDialog(ownerWindow);

            // Brought up to date: the board on screen is this draft's, read again.
            if (window.BroughtUpToDate)
            {
                this.thisMainWindow?.ReloadCurrentBoardAfterDraftFileImport(entry.ExcelDataFile);
            }

            // Dismissing the warning or bringing the draft up to date both change the row.
            this.RefreshDrafts();
        }

        // ###########################################################################################
        // Opens the "KiCad data" window for one drafted board (session 2c, task 9) - its pictures are
        // the Pictures part's since step 3, the window's "Schematic images" half gone.
        //
        // The board labels and the "Important signals" rows handed over - what the KiCad report
        // judges the KiCad data against - are the DRAFT's, read from its workbook on disk
        // (BoardForFilesWindowAsync).
        //
        // `toImport`: a KiCad project folder dropped on the tab (TabDrafts.Drop.cs) - imported as the
        // window opens, as its own drop target imports it.
        // ###########################################################################################
        private async Task ManageFilesAsync(HardwareBoardEntry entry, IReadOnlyList<string>? toImport = null)
        {
            if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
            {
                return;
            }

            // The table's changes are saved first: every change the window makes is saved with what
            // the draft holds, so one the table could not save would refuse each of them too - asked
            // about here, once, rather than behind every click in the window.
            if (!await this.SettleTableEditsAsync(entry))
            {
                return;
            }

            // Under CRT's "please wait" (owner request, 2026-10-10): reading the board took a few
            // seconds with nothing on screen, and then the window opened. The window says what it
            // reads itself, under its own (BoardFilesWindow.Opening).
            BoardData? boardData = await BusyOverlay.RunLocalAsync(
                this,
                CrtWaitWording.OpeningBoardFiles,
                () => TabDrafts.BoardForFilesWindowAsync(entry));

            var boardLabels = boardData?.Components
                .Select(component => component.BoardLabel)
                .Where(label => !string.IsNullOrWhiteSpace(label))
                .ToList() ?? new List<string>();

            var window = new BoardFilesWindow();
            // The "Important signals" rows from the same draft, for the KiCad report's net names
            // (2026-10-10).
            window.Initialize(entry.ToString(), entry.ExcelDataFile, boardLabels, boardData?.KiCadImportantSignals);

            if (toImport is { Count: > 0 })
            {
                window.ImportOnceOpen(toImport);
            }

            await window.ShowDialog(ownerWindow);

            this.CatchUpAfterBoardFiles(entry, window.ChangedTheDraft);
        }

        // ###########################################################################################
        // After "KiCad data" closes - and ONLY when it changed the draft (owner report, 2026-10-10:
        // closing the window it was then refreshed the detached schematic thumbnails although nothing
        // had changed; the board on screen was reloaded every time). Then the row's count follows,
        // and the board on screen is read again if it is this one - new KiCad files are new traces
        // (the pictures' own changes are the Pictures part's, saved as it makes them). An open table needs
        // nothing: it showed each change as the window made it (the draft held in memory, CRT 4.0
        // step 1).
        // ###########################################################################################
        internal void CatchUpAfterBoardFiles(HardwareBoardEntry entry, bool changedTheDraft)
        {
            if (!changedTheDraft)
            {
                return;
            }

            this.RefreshDrafts();

            if (this.ReloadBoardOverrideForTests is { } reload)
            {
                reload(entry.ExcelDataFile);
            }
            else
            {
                this.thisMainWindow?.ReloadCurrentBoardAfterDraftFileImport(entry.ExcelDataFile);
            }
        }

        // Stands in for Main's reload of the board on screen in a test, which has no Main.
        internal Action<string>? ReloadBoardOverrideForTests { get; set; }

        // ###########################################################################################
        // Sends one drafted board for review (Phase 4) - the point of the whole Contribute tab.
        //
        // WHAT IS SUBMITTED IS THE DRAFT WORKBOOK AS IT IS ON DISK - the board's complete intended
        // state (BoardToSendAsync). The server compares it with the board as published NOW, so a
        // draft started before somebody else's work was published must be brought up to date first
        // (DraftRebaser) or that work reads as removed; the drift line on the row says so.
        //
        // NO ACCOUNT IS INVOLVED. The dialog asks for an email address so a maintainer can contact
        // the contributor and so the DECISION reaches them: SubmissionNotifier mails the address when
        // a submission is published to BETA, published to stable, sent back for changes or rejected
        // (corrected 2026-10-13 - this said no outcome mail was sent, which stopped being true when
        // the notifier was added). "My submissions" shows the same states.
        //
        // THE DRAFT IS LEFT ALONE, deliberately - nothing here discards it. An accepted submission
        // is not published data yet, and a contributor whose draft vanished on submit would lose
        // the ability to keep working while review happens. It goes away when the work is actually
        // published and the synced data starts carrying it.
        // ###########################################################################################
        private async Task SubmitAsync(HardwareBoardEntry entry)
        {
            if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
            {
                return;
            }

            DraftStatus? status = DraftStatusReader.Resolve(
                DataManager.DataRoot,
                DraftManager.DraftsRoot,
                entry.ExcelDataFile,
                entry.IsPublished);

            if (status == null)
            {
                return;
            }

            // What is sent is the draft FILE, so edits still sitting unsaved in the table would
            // silently be left out of the submission.
            if (!await this.SettleTableEditsAsync(entry))
            {
                return;
            }

            // ###########################################################################################
            // *** THE SAME DRAFT IS NOT SENT TWICE (owner request, 2026-10-03). *** The row greys
            // Submit out, but it can be a moment behind - an edit in Excel shows only when CRT's
            // window is next activated - so the draft is looked at again here, now that the
            // table's edits are settled. What it gives is also what the receipt remembers.
            // ###########################################################################################
            string draftFingerprint = TabDrafts.FingerprintOf(entry, status);

            SubmissionReceipt? lastSubmission = SubmissionReceiptPresenter.LatestForBoard(
                this.ReceiptsOverrideForTests ?? SubmissionReceiptStore.All,
                BoardDescriptorRules.BoardIdFromExcelDataFile(entry.ExcelDataFile),
                status.CreatedUtc);

            if (SubmissionReceiptPresenter.IsAlreadySent(lastSubmission, draftFingerprint))
            {
                this.RefreshDrafts();
                return;
            }

            // The board as it should READ after publishing. Since Phase 6 that is simply the
            // draft's own workbook rather than an official board with drafted rows merged over it.
            // The submission contract is unchanged: SubmissionManifestBuilder has always wanted the
            // complete intended state.
            var mergedData = await TabDrafts.BoardToSendAsync(status);
            if (mergedData == null)
            {
                return;
            }

            // ###########################################################################################
            // *** ERRORS ARE FIXED BEFORE ANYTHING IS SENT (owner request, 2026-10-02: "All error
            // should be fixed before submission can be done"). *** The table's own checks, over the
            // board a submit would send, with its files looked for where a submit looks
            // (BoardDataChecks; every error there is one the server would refuse). With any, nothing
            // is sent: the table opens on exactly those rows, saying why.
            // ###########################################################################################
            if (await this.StopForErrorsAsync(entry, mergedData))
            {
                return;
            }

            SubmissionIdentity identity = TabDrafts.IdentityFor(entry, status);

            // What the draft changes, in words - "Review changes"' summary, shown before anything is sent.
            DraftChangeSummary changes = await BusyOverlay.RunLocalAsync(this, CrtWaitWording.ReadingChanges, () => this.SummaryForAsync(entry));

            // The credits that work deserves, against the board as published (CreditSuggestions).
            List<CreditSuggestion> credits = CreditSuggestions.For(await this.PublishedBoardForCreditsAsync(entry, status), mergedData);

            var window = this.SubmitWindowFactoryForTests?.Invoke() ?? new SubmitDraftWindow();

            window.Initialize(
                entry.ToString(),
                mergedData,
                identity,
                DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, entry.ExcelDataFile),
                DataManager.DataRoot,
                TabDrafts.CalibrationsFor(entry, mergedData),

                // A signed-in maintainer sends with the account (2026-10-01) - see UseMaintainerAccount.
                this.thisMaintainerAccount,
                draftFingerprint,
                changes,
                credits,
                added => this.AddCreditsToDraftAsync(entry, status, added));

            // The other drafts, to send with this one under the same message (TabDrafts.SubmitOthers.cs).
            window.OfferOtherDrafts(this.OtherDraftOffers(entry), this.PrepareOtherDraftAsync);

            await window.ShowDialog(ownerWindow);

            // The draft is untouched either way, but the row is rebuilt regardless: a submission
            // that failed on a missing file is very often followed by the user fixing it, and a
            // stale row would then still describe the state that failed.
            this.RefreshDrafts();
        }

        // ###########################################################################################
        // Who a draft's submission is about - built from its board identity and, for a new board, its
        // registration. Shared by the board Submit was pressed on and every other draft sent with it.
        // ###########################################################################################
        private static SubmissionIdentity IdentityFor(HardwareBoardEntry entry, DraftStatus status)
        {
            // ###########################################################################################
            // THE REGISTRATION IS THE SOURCE OF TRUTH FOR A DRAFT-ONLY BOARD, not the board entry.
            //
            // A board entry is built ONCE, when the hardware/board list is assembled, and it is only
            // rebuilt when something calls RefreshDraftOnlyBoards. So an entry can be stale - and
            // for a draft-only board it can carry blank names, because EnumerateDraftOnlyBoards
            // reads them out of the draft's own registration and a draft whose registration was
            // missing at that moment produces an entry with nothing in it.
            //
            // That is exactly what happened on 2026-09-21: the label editor had destroyed the
            // registration (see DraftWriterNewBoardPreservationTests), the entry was built from
            // the damaged draft, and the submission went out with empty Hardware and Board. The
            // server answered "does not name the hardware" / "does not name the board" - correct,
            // and impossible to act on from the client's own screen, which showed both names.
            //
            // Reading the draft we have just loaded closes that window: it is the freshest copy on
            // disk, and for a draft-only board it is where these values actually live.
            //
            // ###########################################################################################
            // *** THE FALLBACK IS THE PATH SEGMENT, NOT THE BOARD ENTRY'S DISPLAY NAME (fixed
            // 2026-09-23). ***
            //
            // BoardId below is built from the ExcelDataFile path, and SubmissionValidator rebuilds
            // the id from Manufacturer/Hardware/Board and compares the two. entry.HardwareName and
            // entry.BoardName are the MASTER WORKBOOK's display names - for the C64 250407 they are
            // "Commodore 64" and "250407 (long board)", while the path segments are "C64" and
            // "250407" - so using them made the rebuilt id "Commodore/Commodore 64/250407 (long
            // board)" and every submission for a published board was refused with 400
            // "identity.board_id_mismatch".
            //
            // A draft-only board is unaffected either way: NewBoardIdentity builds its folders
            // FROM the registered names, so for those the two agree by construction. That is why
            // the registration still wins here - it remains the right answer for the case this
            // fallback chain was originally written for.
            // ###########################################################################################
            string hardwareName = status.NewBoard?.HardwareName is { Length: > 0 } registeredHardware
                ? registeredHardware
                : NewBoardIdentity.ExtractHardware(entry.ExcelDataFile);

            string boardName = status.NewBoard?.BoardName is { Length: > 0 } registeredBoard
                ? registeredBoard
                : NewBoardIdentity.ExtractBoard(entry.ExcelDataFile);

            return new SubmissionIdentity
            {
                // "Manufacturer/Hardware/Board", NOT the ExcelDataFile - the .xlsx on the end names
                // a file, and for a draft-only board one that does not exist. This is the
                // `boards` table's primary key (see 0001_initial.sql), so the shape has to be
                // exactly what the server expects.
                BoardId = BoardDescriptorRules.BoardIdFromExcelDataFile(entry.ExcelDataFile),
                Manufacturer = NewBoardIdentity.ExtractManufacturer(entry.ExcelDataFile),
                Hardware = hardwareName,
                Board = boardName,

                // What the draft was started against, NOT what the official file says right now.
                // The server needs to know which revision these edits were made on top of; reading
                // the current one here would claim the contributor had seen changes they never saw.
                BaseRevision = status.BaseRevision,

                // A new board's notes from "Create board" (owner request, 2026-10-05): the
                // maintainer's placement starts with them, and from it they reach the main Excel
                // data file's notes column. A draft of a published board has none.
                HardwareNotes = status.NewBoard?.NotesForSubmission() ?? string.Empty,
                ApplicationVersion = AppConfig.AppDisplayVersionString,
                CreatedUtc = DateTimeOffset.UtcNow
            };
        }

        // ###########################################################################################
        // The draft's KiCad calibrations, read from the draft's own board file rather than from the
        // board sent - BoardData has no calibration section, which is exactly why a contributor's
        // calibration work never travelled before this was wired up.
        // ###########################################################################################
        private static IReadOnlyList<KiCadCalibrationEntry> CalibrationsFor(HardwareBoardEntry entry, BoardData board) =>
            DraftBoardSource.CollectCalibrations(
                DraftBoardSource.ResolveWritablePath(DataManager.DataRoot, DraftManager.DraftsRoot, entry.ExcelDataFile),
                board);

        // The same, read for Submit's credit suggestions - under CRT's "please wait", since Submit was
        // pressed for it (code review, 2026-10-26: a large board not read yet was read with nothing on
        // screen, between the "reading your changes" wait and the dialog).
        internal Task<BoardData?> PublishedBoardForCreditsAsync(HardwareBoardEntry entry, DraftStatus status) =>
            BusyOverlay.RunLocalAsync(this, CrtWaitWording.ReadingPublishedBoard, () => this.PublishedBoardForAsync(entry, status));

        // The board as published, or null for a new one - what the credits are suggested against.
        private async Task<BoardData?> PublishedBoardForAsync(HardwareBoardEntry entry, DraftStatus status) =>
            this.PublishedBoardOverrideForTests is not null
                ? this.PublishedBoardOverrideForTests(entry)
                : await TabDrafts.LoadPublishedBoardAsync(entry, status);

        // ###########################################################################################
        // *** WHAT A SUBMIT SENDS IS THE DRAFT WORKBOOK ON DISK, read now - the file its fingerprint
        // was just taken from (owner report, 2026-10-11). *** It was DataManager.LoadBoardDataAsync,
        // which answered from its cache: an image renamed in "Schematic images" for a board not
        // selected in the drop-down lists went to the queue under its OLD name, with its highlights
        // under it and its KiCad calibration left out (collected under the board's names) - and the
        // receipt, fingerprinted from the file, then greyed Submit out as "already sent". It also
        // answers the PUBLISHED board while "view boards as officially published" is ticked.
        // TabDraftsSubmitBoardTests holds both.
        // ###########################################################################################
        internal static Task<BoardData?> BoardToSendAsync(DraftStatus status) =>
            Task.Run(() => BoardDataReader.ReadWorkbookUncached(status.WorkbookPath));

        // ###########################################################################################
        // The board "KiCad data" is handed - its labels and its "Important
        // signals" rows, which the KiCad report judges the draft's KiCad data against: the draft
        // WORKBOOK as it is on disk, like BoardToSendAsync and for its reasons (code review,
        // 2026-10-11). It was DataManager.LoadBoardDataAsync, which answers the PUBLISHED board under
        // "view boards as officially published" - so the report's "will NOT light up" list and its
        // "already used" marks described the published rows, not the ones just drafted.
        // ###########################################################################################
        internal static Task<BoardData?> BoardForFilesWindowAsync(HardwareBoardEntry entry) =>
            Task.Run(() => DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, entry.ExcelDataFile));

        // ###########################################################################################
        // The credits ticked in the Submit dialog into the draft, as ONE step ("Credits") saved at
        // once - each in its category's group (BoardPartEdits) - then the board a submit sends and its
        // fingerprint read again, so what goes out and what the receipt remembers carry them. Null
        // when the draft could not take them; the draft's header says why.
        // ###########################################################################################
        internal async Task<(BoardData Board, string Fingerprint)?> AddCreditsToDraftAsync(HardwareBoardEntry entry, DraftStatus status, IReadOnlyList<CreditEntry> credits)
        {
            DraftWorkbookEditOutcome outcome = await DraftEdits.ApplyAndSaveAsync(
                entry.ExcelDataFile,
                DraftStepWording.Credits,
                board => BoardPartEdits.WithCredits(board, [.. board.Credits, .. credits]));

            if (outcome != DraftWorkbookEditOutcome.Saved || await TabDrafts.BoardToSendAsync(status) is not { } board)
            {
                Logger.Warning($"Could not put the credits into the draft of [{entry.ExcelDataFile}] before submitting - [{outcome}]");
                return null;
            }

            return (board, TabDrafts.FingerprintOf(entry, status));
        }

        // What the draft holds, as DraftFingerprint gives it - see that class.
        private static string FingerprintOf(HardwareBoardEntry entry, DraftStatus status) =>
            DraftFingerprint.Compute(
                status.WorkbookPath,
                DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, entry.ExcelDataFile));

        // Whether the row's Submit is greyed out because the draft was already sent as it is. The
        // draft is only read when its latest submission remembers what it sent, so a draft never
        // sent costs nothing here.
        private static bool IsAlreadySent(HardwareBoardEntry entry, DraftStatus status, SubmissionReceipt? lastSubmission) =>
            lastSubmission is { DraftFingerprint.Length: > 0 } &&
            SubmissionReceiptPresenter.IsAlreadySent(lastSubmission, TabDrafts.FingerprintOf(entry, status));

        // Lets a headless test drive the submit flow without standing up a real dialog that would
        // try to reach the server. null (the default) is the shipped path - a plain
        // new SubmitDraftWindow() - exactly as every other window on this tab is constructed.
        internal Func<SubmitDraftWindow>? SubmitWindowFactoryForTests { get; set; }

        // ###########################################################################################
        // True when `board` - what a submit of this draft would send - has errors, which are then
        // shown in the table instead of anything being sent. Apart from SubmitAsync so a test can
        // drive it without the dialog's window.
        // ###########################################################################################
        internal async Task<bool> StopForErrorsAsync(HardwareBoardEntry entry, BoardData board)
        {
            int errors = TabDrafts.CountErrors(entry, board);

            if (errors == 0)
            {
                return false;
            }

            await this.ShowErrorsBeforeSubmitAsync(entry, errors);
            return true;
        }

        // The errors in the board a submit of this draft would send - the table's own checks.
        private static int CountErrors(HardwareBoardEntry entry, BoardData board) =>
            BoardDataChecks.Check(
                    BoardCheckRows.From(board),
                    new DiskFileLookup(DataManager.DataRoot, DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, entry.ExcelDataFile), entry.ExcelDataFile),
                    BoardCheckScope.Everything)
                .Count(problem => problem.Level == BoardProblemLevel.Error);

        // ###########################################################################################
        // Submit pressed on a draft with errors: its table, showing only the rows with errors (the
        // colour key's "Errors" picked, which takes the table to the first sheet with any), and a
        // sentence saying what to do. Already open, the table is kept as it is, unsaved edits and all.
        // ###########################################################################################
        private async Task ShowErrorsBeforeSubmitAsync(HardwareBoardEntry entry, int errors)
        {
            if (!this.IsTableOpenFor(entry))
            {
                await this.OpenTableAsync(entry);
            }

            if (!this.IsTableOpenFor(entry))
            {
                return;
            }

            // Not when every error is on a schematic, in no row: the filter would empty the table.
            bool rowsWithErrors = this.TableEditor.CommitAndGetDocument()?.HasRowsShownBy(BoardTableRowKinds.Errors) == true;

            if (rowsWithErrors)
            {
                this.TableEditor.Filter = BoardTableRowKinds.Errors;
            }

            this.TableEditor.ShowMessage(BoardTableProblemWording.SubmitBlocked(errors, rowsWithErrors));
            this.FocusTableIfOpen();
        }

        // ###########################################################################################
        // The maintainer signed in on the Maintainer tab, or null (owner request, 2026-10-01: "When I
        // am a maintainer, and I have logged in, then I want to use that email address everywhere").
        // Handed over by Main (ShareMaintainerSignIn) - this tab never reaches into the Maintainer
        // tab - and passed to the Submit dialog, which shows the account's address and sends the
        // submission with the account (ContactAddress).
        // ###########################################################################################
        private ReviewSession? thisMaintainerAccount;

        internal void UseMaintainerAccount(ReviewSession? account) => this.thisMaintainerAccount = account;

        internal ReviewSession? MaintainerAccountForTests => this.thisMaintainerAccount;

        // ###########################################################################################
        // Creates a whole new hardware/board of the user's own (session 2c, task 9) - the same
        // action the Contribute tab offers, repeated here because this is where someone already
        // working on drafts looks for it.
        // ###########################################################################################
        // ###########################################################################################
        // Opens the "my submissions" view (Phase 4, task 6).
        //
        // Offered unconditionally, including with nothing sent yet: its empty state says what will
        // appear there, which teaches the feature, whereas a button that materialises only after
        // the first submission is one nobody discovers until they no longer need to be told.
        // ###########################################################################################
        private async void OnMySubmissionsClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
                {
                    return;
                }

                var window = new MySubmissionsWindow();
                window.Initialize();

                await window.ShowDialog(ownerWindow);

                // The window can mark comments as read and can refresh state from the server, so the
                // badge is re-read once it closes rather than left showing what was true before it
                // opened.
                //
                // *** VIA THE VISIBILITY PATH, NOT JUST THE BADGE. *** Unread feedback is what keeps
                // this tab on screen for somebody with no local drafts, so reading the last comment
                // must also let the tab go away again - otherwise it lingers empty until restart.
                //
                // *** AND THROUGH RETIREMENT FIRST (code review, 2026-09-25). *** The window's Refresh
                // can move a submission to "published", and nothing retired its draft until the next
                // launch. RetirePublishedDraftsAsync ends in ApplyContributeTabState, which calls
                // RefreshDrafts (and so the badge) on its way through.
                if (this.thisMainWindow is not null)
                {
                    await this.thisMainWindow.RetirePublishedDraftsAsync();
                }
                else
                {
                    this.RefreshSubmissionBadge();
                }
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT
                // (code review, 2026-10-13 - AsyncVoidHandlersGuardTests).
                Logger.Critical($"OnMySubmissionsClick failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // Shows or hides the "you have unread maintainer feedback" badge on the My submissions button.
        //
        // *** IT COUNTS UNREAD COMMENTS, NOT SUBMISSIONS IN ANY PARTICULAR STATE. *** Tying it to
        // "changes requested" would miss a maintainer who left a note while approving, and would
        // keep shouting at somebody who has already read the note and is working on it. What the
        // contributor needs to know is "somebody said something you have not seen".
        //
        // Reads the store directly rather than taking a count as an argument: this is called from
        // several places (a refresh, the window closing) and a parameter would be one more thing
        // for a future caller to compute differently.
        // ###########################################################################################
        private void RefreshSubmissionBadge()
        {
            var badge = this.FindControl<Border>("MySubmissionsBadge");
            var text = this.FindControl<TextBlock>("MySubmissionsBadgeText");

            if (badge is null || text is null)
            {
                return;
            }

            int unread = this.UnreadCommentCountOverrideForTests
                ?? SubmissionReceiptStore.UnreadCommentCount();

            badge.IsVisible = unread > 0;
            text.Text = unread.ToString(CultureInfo.InvariantCulture);
        }

        // Lets a headless test drive the badge without standing up the real receipts store, which
        // is a static singleton reading the user's own AppData folder - the same seam reasoning as
        // HardwareBoardsOverrideForTests above.
        internal int? UnreadCommentCountOverrideForTests { get; set; }

        // The same, for the rows' submission badges. null is the shipped path: the real store.
        internal IReadOnlyList<SubmissionReceipt>? ReceiptsOverrideForTests { get; set; }

        private void OnAddNewBoardClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisMainWindow == null)
            {
                return;
            }

            _ = this.thisMainWindow.OpenNewBoardWindowAsync();
        }
    }

    // ###########################################################################################
    // One draft of the Contribute tab - its header, and an entry of "My drafts" - a board with a
    // local draft, plus how many rows it has
    // changed in total (across every BoardData section, not just Components) so a user can tell a
    // typo fix from a whole new board's worth of edits at a glance.
    // ###########################################################################################
    public sealed class DraftListItem
    {
        public string DisplayName { get; }

        // The board this row acts on - its buttons are bound to it (see ShowsTheSameAs).
        private readonly HardwareBoardEntry thisEntry;

        // The same board, for "My drafts", which shows the board picked (TabDrafts.Workspace.cs).
        internal HardwareBoardEntry Entry => this.thisEntry;

        // ###########################################################################################
        // Whether this row draws exactly what `other` would, for the same board - TabDrafts
        // keeps an unchanged row rather than replacing it (TabDrafts.KeepUnchangedRows).
        //
        // Every value the row shows is compared (ShownValues); the commands are not. The board
        // ENTRY must be the very same object, because the kept row's buttons act on its entry.
        // ###########################################################################################
        internal bool ShowsTheSameAs(DraftListItem other) =>
            ReferenceEquals(this.thisEntry, other.thisEntry) && ShownValues.Same(this, other);

        // The badge beside the name: this board's latest submission, in "My submissions"' words
        // and colour. Hidden when it was never sent.
        // ###########################################################################################
        // *** THE DRAFT'S ERRORS AND WARNINGS, ON ITS ROW (owner report, 2026-10-02: "It must
        // check for errors when creating the draft, and if the board changes "offline", outside of
        // app"). *** The table's own checks (DraftStatusReader.CountProblemsCached), so a draft with
        // something to fix says so the moment it exists or is changed in Excel - not only once its
        // table is opened. Red for errors (Submit opens the table on them instead of sending),
        // amber for warnings, nothing at all for a clean draft.
        // ###########################################################################################
        public bool HasErrors { get; }
        public string ErrorsText { get; } = string.Empty;
        public bool HasWarnings { get; }
        public string WarningsText { get; } = string.Empty;

        public bool HasSubmission { get; }
        public string SubmissionStateText { get; } = string.Empty;
        public string SubmissionTooltip { get; } = string.Empty;
        public Avalonia.Media.IBrush? SubmissionAccentBrush { get; }

        // Which board this row is - how table mode finds the open draft's row again after a
        // refresh has rebuilt every item.
        public string ExcelDataFile { get; }

        // Whether the table view is open on this draft - its save state is then the table's to say.
        // A row is rebuilt rather than changed when the table opens or closes, the same way every
        // other value on it is. The table, the pictures and the KiCad data are opened from the left
        // of the Contribute tab since CRT 4.0 step 2, not from the row.
        public bool IsTableOpen { get; }
        public string RowSummary { get; }
        public ICommand DiscardCommand { get; }
        public ICommand ViewDriftCommand { get; }
        public ICommand SubmitCommand { get; }

        // ###########################################################################################
        // THE DRAFT'S HEADER (CRT 4.0 step 1's last slice, 2026-10-25): its save state, Undo and Redo,
        // Try again after a save that could not be made, and "Review changes". The live part is
        // `Header`, one per draft for the run (TabDrafts.Header.cs) - compared by reference, so a row
        // rebuilt while the draft saves is still the same row.
        // ###########################################################################################
        public DraftHeaderState Header { get; }
        public ICommand UndoCommand { get; }
        public ICommand RedoCommand { get; }
        public ICommand TryAgainCommand { get; }
        public ICommand ReviewChangesCommand { get; }

        // Whether this draft can be sent at all, and why not when it cannot. Disabled rather than
        // hidden - see the markup's own comment.
        public bool CanSubmit { get; }
        public string SubmitTooltip { get; }

        // Why the Submit dialog cannot send this draft along with another's, in its short words - or
        // null when it can (slice 5). Errors count here: no table opens on another board's.
        internal string? NotOfferedWithOthers { get; }

        // Whether to show the drift line and the "What changed" button (the "Updated officially"
        // chip beside the name was removed, owner request 2026-09-28). Only a real
        // difference counts - an unrecorded base revision is not evidence of drift, so it stays
        // silent (see DraftRevisionComparer).
        public bool HasDrift { get; }

        // One plain line saying what happened, worded to match what the data actually supports:
        // "updated" only when both revisions parsed and the official one is genuinely later,
        // "changed" otherwise. Never promises an ordering that could not be established.
        public string DriftSummary { get; }

        private static string Capitalised(string text) =>
            text.Length == 0 ? text : char.ToUpperInvariant(text[0]) + text[1..];

        // A line under the row saying why Submit is greyed out, or that the draft is finished with
        // (code review, 2026-10-13) - the reason used to live in the button's tooltip alone.
        public string NoteText { get; } = string.Empty;
        public bool HasNote => this.NoteText.Length > 0;

        public DraftListItem(
            HardwareBoardEntry entry,
            DraftStatus status,
            int changeCount,
            BoardProblemCounts problems,
            DraftDriftState driftState,
            Func<HardwareBoardEntry, Task> discard,
            Func<HardwareBoardEntry, Task> viewDrift,
            Func<HardwareBoardEntry, Task> submit,
            bool isTableOpen,
            SubmissionReceipt? lastSubmission,
            bool alreadySent = false,
            DraftOtherChanges otherChanges = default,
            DraftHeaderState? header = null,
            DraftRowHeaderActions? headerActions = null)
        {
            this.thisEntry = entry;

            this.Header = header ?? new DraftHeaderState();

            DraftRowHeaderActions actions = headerActions ?? DraftRowHeaderActions.None;
            this.UndoCommand = new ActionCommand(() => Dispatcher.UIThread.InvokeAsync(() => actions.Undo(entry)));
            this.RedoCommand = new ActionCommand(() => Dispatcher.UIThread.InvokeAsync(() => actions.Redo(entry)));
            this.TryAgainCommand = new ActionCommand(() => Dispatcher.UIThread.InvokeAsync(() => actions.TryAgain(entry)));
            this.ReviewChangesCommand = new ActionCommand(() => Dispatcher.UIThread.InvokeAsync(() => actions.ReviewChanges(entry)));
            this.DisplayName = entry.ToString();
            this.ExcelDataFile = entry.ExcelDataFile;

            // ###########################################################################################
            // *** THE SUBMISSION BADGE (owner request, 2026-09-27): "When I have submitted my
            // submission to the server, then I need to see that somehow". ***
            //
            // The state words and their colour are "My submissions"' own (DescribeReceiptState,
            // ClassifyReceipt, AccentFor - a submission the server no longer knows reads "No longer on
            // the server"), so the badge moves with the review - "Submitted - awaiting ...",
            // then "Published to the BETA source", or "Changes requested" in orange - and can never say
            // anything that window does not. Its tooltip says which submission and when.
            // ###########################################################################################
            this.HasSubmission = lastSubmission is not null;

            if (lastSubmission is not null)
            {
                this.SubmissionStateText = SubmissionReceiptPresenter.DescribeReceiptState(lastSubmission);
                this.SubmissionTooltip = SubmissionReceiptPresenter.DescribeLastSubmission(lastSubmission);
                this.SubmissionAccentBrush = SubmissionListItem.AccentFor(
                    SubmissionReceiptPresenter.ClassifyReceipt(lastSubmission));
            }

            this.HasErrors = problems.Errors > 0;
            this.ErrorsText = problems.Errors == 1 ? "1 error" : $"{problems.Errors} errors";
            this.HasWarnings = problems.Warnings > 0;
            this.WarningsText = problems.Warnings == 1 ? "1 warning" : $"{problems.Warnings} warnings";

            this.IsTableOpen = isTableOpen;

            int rowCount = changeCount;

            // A board that exists only as a draft (session 2c, task 9) is described as what it is
            // rather than by a row count. "0 rows changed" would be actively misleading on a
            // freshly created board: nothing was CHANGED because the whole board is new, and the
            // number says nothing about the thing the user actually made.
            string rowText = rowCount == 1 ? "1 row" : $"{rowCount} rows";

            // Files and KiCad calibrations count too (code review, 2026-10-13) - "0 rows changed" on
            // a draft whose work is a re-calibrated overlay said there was nothing in it.
            this.RowSummary = status.IsNewBoard
                ? (rowCount == 0 ? "New board, nothing added yet" : $"New board, {rowText} so far")
                : (rowCount, otherChanges.Any) switch
                {
                    (0, true) => DraftListItem.Capitalised(otherChanges.Describe()),
                    (_, true) => $"{rowText} changed, {otherChanges.Describe()}",
                    _ => $"{rowText} changed",
                };

            // ###########################################################################################
            // *** PUBLISHED, CHANGED BY A MAINTAINER, AND NOTHING NEW SINCE (code review, 2026-10-13). ***
            // Retirement deletes a draft only when it equals the published board, so a submission a
            // maintainer corrected before publishing left its draft behind for good - and that
            // draft, sent again, would undo the correction. Said here, where the contributor looks,
            // instead of in the Wiki alone; the drift line is not shown, since the "update" is the
            // contributor's own work as published.
            // ###########################################################################################
            bool finishedWith =
                alreadySent &&
                lastSubmission is { AmendedByMaintainer: true } &&
                string.Equals(lastSubmission.LastKnownState, "published", StringComparison.OrdinalIgnoreCase);

            this.HasDrift = !finishedWith && DraftDriftWording.IsWarning(driftState);

            this.DriftSummary = this.HasDrift
                ? DraftDriftWording.Line(driftState, status.HasBase)
                : string.Empty;

            // An EMPTY draft has nothing to review. That is a real state, not a theoretical one: a
            // board created through "Add a new board" exists as a registration before a single
            // row or image is added to it, and this tab lists it from that moment. Sending it would
            // put an empty board in front of a maintainer, so the button says why instead.
            //
            // A drafted board with rows is submittable, drift or no drift - the server diffs
            // against the base revision itself, and refusing to send while the official data has
            // moved would strand a contributor behind a change they did not make.
            //
            // *** BUT NOT WHAT WAS JUST SENT (owner request, 2026-10-03: "It should not be possible
            // to submit the same data again"). *** While the draft holds exactly what its latest
            // submission sent, Submit is greyed out and says so - SubmissionReceiptPresenter.IsAlreadySent.
            bool hasSomething = rowCount > 0 || otherChanges.Any;

            this.CanSubmit = hasSomething && !alreadySent;
            this.NotOfferedWithOthers = SubmitDraftsWording.NotOffered(hasSomething, alreadySent, problems.Errors);

            this.SubmitTooltip = !hasSomething
                ? (status.IsNewBoard
                    ? "There is nothing to send yet. Add some data to this board first."
                    : "There is nothing to send yet - this draft holds no change from the published data.")
                : alreadySent && lastSubmission is not null
                    ? SubmissionReceiptPresenter.DescribeAlreadySent(lastSubmission)
                : this.HasErrors
                    ? $"This draft has {this.ErrorsText} to fix first - Submit opens its table on them instead of sending."
                    : "Send this board for review. You do not need an account - only an email address, " +
                      "so you can be told whether it was accepted.";

            this.NoteText = finishedWith
                ? "Your work is published, with some details changed by a maintainer, so this draft will not tidy " +
                  "itself away. Look at the board, then discard this draft - sending it again would undo the maintainer's changes."
                : this.CanSubmit ? string.Empty : this.SubmitTooltip;

            this.DiscardCommand = new ActionCommand(() => Dispatcher.UIThread.InvokeAsync(() => discard(entry)));
            this.ViewDriftCommand = new ActionCommand(() => Dispatcher.UIThread.InvokeAsync(() => viewDrift(entry)));
            this.SubmitCommand = new ActionCommand(() => Dispatcher.UIThread.InvokeAsync(() => submit(entry)));
        }
    }

    // What the header's buttons do for a draft - TabDrafts.Header.cs. None does nothing.
    public sealed record DraftRowHeaderActions(
        Func<HardwareBoardEntry, Task> Undo,
        Func<HardwareBoardEntry, Task> Redo,
        Func<HardwareBoardEntry, Task> TryAgain,
        Func<HardwareBoardEntry, Task> ReviewChanges)
    {
        public static DraftRowHeaderActions None { get; } =
            new(_ => Task.CompletedTask, _ => Task.CompletedTask, _ => Task.CompletedTask, _ => Task.CompletedTask);
    }
}
