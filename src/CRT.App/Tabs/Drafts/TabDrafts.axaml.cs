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
    // THE DRAFTS TAB (NewContributeStrategy.md Phase 2, session 2b) - lists every hardware/board
    // board with a local, unpublished draft, and lets the user discard one. Hidden entirely when
    // there are none - see Main.ApplyDraftsTabVisibility, the same conditional-tab pattern
    // WorkbooksTabItem/OscilloscopeTabItem already use.
    //
    // Deliberately simple: no per-row detail beyond a row count, no editing. Session 2c
    // ("Authoring into drafts") is what actually lets a contributor CREATE a draft through the
    // app; until then this tab can only show and discard whatever a draft.json already contains
    // (hand-written, or from a future submission "request changes" round-trip in a later phase).
    //
    // *** TABLE MODE (2026-09-24). *** "Edit in table format" on a row opens that draft's workbook
    // as editable sheets (BoardTableEditor) directly below it, and hides every other row while it
    // is open. `Drafts` itself is NOT filtered for that - Main reads its count to decide whether
    // this tab is shown at all - so the list is hidden and the open draft's row is shown on its
    // own, in OpenDraftRow (ApplyTableMode; the markup explains why it cannot be the list itself).
    //
    // Anything that would lose or bypass the table's unsaved edits asks first
    // (ConfirmLeavingTableAsync): closing the table, discarding the draft, managing its files,
    // submitting it (a submission must include what is on screen), and quitting the application
    // (Main.OnWindowClosing).
    //
    // FILE MAP:
    //   TabDrafts.axaml.cs  - the list, its row actions (discard, files, drift, submit), the badge
    //   TabDrafts.Table.cs  - table mode: opening/closing the table, and the unsaved-edits prompts
    //   TabDrafts.OutsideChanges.cs - reading the list again when the tab is shown or CRT's window
    //                         comes back to the front, so a draft edited in Excel shows its new numbers
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

        public ObservableCollection<DraftListItem> Drafts { get; } = new();

        // The boards behind the rows above, as of the last RefreshDrafts. Main builds the Hardware
        // and Board drop-downs' "Draft" chips from exactly this list (see Main.DraftBadges.cs), so a
        // chip appears for precisely the boards this tab lists and for no others.
        private readonly List<HardwareBoardEntry> thisDraftedBoards = new();

        internal IReadOnlyList<HardwareBoardEntry> DraftedBoards => this.thisDraftedBoards;

        public TabDrafts()
        {
            this.InitializeComponent();

            this.DraftsItemsControl.ItemsSource = this.Drafts;

            this.TableEditor.Saved += this.OnTableSaved;

            // A reload because Excel (say) changed the draft needs the same refreshes as a save.
            this.TableEditor.ReloadedFromOutside += this.OnTableSaved;
        }

        // ###########################################################################################
        // Called once from Main's constructor, the same pattern every other tab's Initialize uses.
        // ###########################################################################################
        internal void Initialize(Main mainWindow)
        {
            this.thisMainWindow = mainWindow;
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

            this.Drafts.Clear();
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

                this.Drafts.Add(new DraftListItem(
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
                    this.ManageFilesAsync,
                    this.ViewDriftAsync,
                    this.SubmitAsync,
                    isTableOpen: this.IsTableOpenFor(entry),
                    this.ToggleTableAsync,
                    lastSubmission,
                    TabDrafts.IsAlreadySent(entry, status, lastSubmission)));
            }

            this.ApplyTableMode();

            bool hasAnyDrafts = this.Drafts.Count > 0;

            // In table mode the list gives way to the open draft's own row - see ApplyTableMode.
            this.DraftsScrollViewer.IsVisible = hasAnyDrafts && !this.IsTableOpen;
            this.EmptyStateText.IsVisible = !hasAnyDrafts;

            // *** THE EMPTY STATE HAS TO EXPLAIN WHY THE TAB IS EVEN HERE. *** With no drafts this
            // tab is normally hidden, so the one case it appears empty is when unread maintainer
            // feedback is holding it open (Main.ApplyDraftsTabVisibility). "No local drafts yet"
            // would then be a true sentence that answers the wrong question, and the contributor
            // would have no idea a maintainer had written to them.
            if (!hasAnyDrafts)
            {
                int unread = this.UnreadCommentCountOverrideForTests
                    ?? SubmissionReceiptStore.UnreadCommentCount();

                this.EmptyStateText.Text = unread > 0
                    ? "A maintainer has replied to something you sent. Open \"My submissions\" above to read it."
                    : "No local drafts yet. Edits made through the Contribute tab that have not been submitted will appear here.";
            }
        }

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

            var confirmWindow = new DiscardDraftWindow();
            confirmWindow.Initialize(entry.ToString(), unfinished);

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

            this.ShowDiscardFailure(null);

            // *** A FAILED DELETE IS REPORTED, AND THE REFRESH BELOW STILL RUNS. *** The usual cause
            // is the draft workbook being open in Excel, which locks it on Windows - part of the
            // folder may already be gone by then, so the list and the drop-downs must still catch
            // up with what is really on disk.
            if (!DraftManager.DiscardDraft(entry.ExcelDataFile))
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
            if (unfinished is { Count: > 0 })
            {
                SubmissionReceiptStore.MarkDraftDiscarded(unfinished.Select(receipt => receipt.SubmissionId), DateTimeOffset.UtcNow);
                _ = DraftDiscardReporter.ReportPendingAsync(this.DraftDiscardSendOverrideForTests ?? new SubmissionClient().ReportDraftDiscardedAsync);
            }

            this.RefreshDrafts();
            this.thisMainWindow?.ApplyDraftsTabVisibility();
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
                  "files is open in another program, such as its workbook in Excel. Close it there and press " +
                  "Discard again.";
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

            string? officialRevision = BoardDataReader.TryGetCachedRevisionDate(excelPath);

            if (officialRevision == null)
            {
                officialRevision = BoardDataReader.ReadRevisionDateOnly(excelPath);
            }

            // Compared directly rather than through DraftDriftDetector.CompareRevisions, which takes
            // a BoardDraft - the new-board guard it applies is already handled above, off the
            // marker, so all that is left is the revision comparison itself.
            return DraftRevisionComparer.Compare(status.BaseRevision, officialRevision);
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
            // *** THE PER-ROW REPORT IS NOW A COMPARISON, AND IT ANSWERS A NARROWER QUESTION
            // (Phase 6, 2026-09-23). ***
            //
            // It used to classify each DRAFTED row against the official data as it stands now, and
            // could therefore say two things this cannot: "the row you edited has since
            // DISAPPEARED officially" and "a row you added now exists officially too". Both were
            // consequences of the merge - a Modified row whose key vanished was silently dropped,
            // an Added row whose key collided lost to the official one.
            //
            // There is no merge any more, so neither consequence exists and neither can be
            // reported. What the contributor gets instead is what they actually changed, derived
            // by comparing their workbook against the published one - and the revision comparison
            // above still tells them the published board has moved.
            //
            // Recovering the sharper warning would mean comparing against the revision the draft
            // was SEEDED from, which the draft folder does not keep a copy of. That is a real
            // feature, not a line of code, and is deliberately not attempted here.
            // ###########################################################################################
            string publishedPath = DraftBoardSource.PublishedPathOf(
                DataManager.DataRoot,
                entry.ExcelDataFile);

            // Reading the published board of a large board is a noticeable wait (2026-09-28).
            BoardData? official = await BusyOverlay.RunLocalAsync(
                this,
                CrtWaitWording.ComparingWithOfficial,
                () => BoardDataReader.LoadAsync(publishedPath, publishedPath));

            var report = new DraftChangeReport
            {
                State = ResolveDriftState(entry, status),
                BaseRevision = status.BaseRevision,
                OfficialRevision = official?.RevisionDate ?? string.Empty,
                Rows = DraftStatusReader.DescribeChanges(status),
            };

            var window = new DraftDriftWindow();
            window.Initialize(entry.ToString(), entry.ExcelDataFile, report);

            await window.ShowDialog(ownerWindow);

            // A rebase inside the window clears the warning, so the row has to be rebuilt.
            this.RefreshDrafts();
        }

        // ###########################################################################################
        // Opens the board-image OR the KiCad import window for one drafted board (session 2c,
        // task 9) - one button each since 2026-09-24, both landing in BoardFilesWindow.
        //
        // The board labels handed over come from the MERGED BoardData for that board - official
        // plus draft - so the KiCad report compares against exactly what the board renders, whether
        // this is a brand-new board (every label drafted) or an overlay on a published one.
        // ###########################################################################################
        private async Task ManageFilesAsync(HardwareBoardEntry entry, BoardFilesSection section)
        {
            if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
            {
                return;
            }

            // Importing an image or KiCad data writes the draft's workbook, which the open table
            // could then never save over - so its edits are settled first.
            if (!await this.SettleTableEditsAsync(entry))
            {
                return;
            }

            var boardData = await DataManager.LoadBoardDataAsync(entry);

            var boardLabels = boardData?.Components
                .Select(component => component.BoardLabel)
                .Where(label => !string.IsNullOrWhiteSpace(label))
                .ToList() ?? new List<string>();

            var window = new BoardFilesWindow();
            window.Initialize(entry.ToString(), entry.ExcelDataFile, boardLabels, section);

            await window.ShowDialog(ownerWindow);

            // Importing a board image adds a Schematics row, so the row count shown here changes -
            // and the board itself has to be reloaded for a newly imported image or KiCad file to
            // actually appear on the Schematics tab.
            this.RefreshDrafts();
            this.thisMainWindow?.ReloadCurrentBoardAfterDraftFileImport(entry.ExcelDataFile);

            // An import adds rows (a Board schematics row per image), so an open table catches up.
            if (this.IsTableOpenFor(entry))
            {
                this.TableEditor.CheckDraftFile();
            }
        }

        // ###########################################################################################
        // Sends one drafted board for review (Phase 4) - the point of the whole Drafts tab.
        //
        // WHAT IS SUBMITTED IS THE MERGED BOARD, not the draft. The server diffs the submission
        // against the base revision itself, so what it needs is the board as it should READ after
        // publishing - official rows with the drafted ones applied over them. Sending only the
        // drafted rows would mean the server had to reconstruct that merge from a draft format it
        // has no reason to know about.
        //
        // NO ACCOUNT IS INVOLVED. The dialog asks for an email address purely so a MAINTAINER can
        // contact the contributor if they need to; there is no sign-in, and pressing this button
        // is the whole interaction. A maintainer account is created by the project owner afterwards,
        // for a NEW board's author, and has nothing to do with sending.
        //
        // *** NOTHING EMAILS THE CONTRIBUTOR THE OUTCOME (corrected 2026-09-23). *** That address
        // is stored and shown to the maintainer, and no submission flow sends to it - the outcome is
        // reported in "My submissions" and nowhere else. See SubmissionReceipt's header.
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
            // draft's own workbook - LoadBoardDataAsync resolves it - rather than an official
            // board with drafted rows merged over it. The submission contract is unchanged:
            // SubmissionManifestBuilder has always wanted the complete intended state.
            var mergedData = await DataManager.LoadBoardDataAsync(entry);
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

            var identity = new SubmissionIdentity
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

            var window = this.SubmitWindowFactoryForTests?.Invoke() ?? new SubmitDraftWindow();

            window.Initialize(
                entry.ToString(),
                mergedData,
                identity,
                DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, entry.ExcelDataFile),
                DataManager.DataRoot,

                // The draft's KiCad calibrations, read from the draft's own JSON SIDECAR rather
                // than from mergedData - BoardData has no calibration section, which is exactly why
                // a contributor's calibration work never travelled before this was wired up.
                DraftBoardSource.CollectCalibrations(
                    DraftBoardSource.ResolveWritablePath(
                        DataManager.DataRoot,
                        DraftManager.DraftsRoot,
                        entry.ExcelDataFile),
                    mergedData),

                // A signed-in maintainer sends with the account (2026-10-01) - see UseMaintainerAccount.
                this.thisMaintainerAccount,
                draftFingerprint);

            await window.ShowDialog(ownerWindow);

            // The draft is untouched either way, but the row is rebuilt regardless: a submission
            // that failed on a missing file is very often followed by the user fixing it, and a
            // stale row would then still describe the state that failed.
            this.RefreshDrafts();
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
                    new DiskFileLookup(DataManager.DataRoot, DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, entry.ExcelDataFile)),
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
            // launch. RetirePublishedDraftsAsync ends in ApplyDraftsTabVisibility, which calls
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
    // One row of the Drafts tab's list - a board with a local draft, plus how many rows it has
    // changed in total (across every BoardData section, not just Components) so a user can tell a
    // typo fix from a whole new board's worth of edits at a glance.
    // ###########################################################################################
    public sealed class DraftListItem
    {
        public string DisplayName { get; }

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

        // "Edit in table format" / "Close table" (2026-09-24). A row is rebuilt rather than
        // changed when the table opens or closes, the same way every other value on it is.
        public bool IsTableOpen { get; }
        public string TableButtonText { get; }
        public string TableButtonTooltip { get; }
        public ICommand EditTableCommand { get; }
        public string RowSummary { get; }
        public ICommand DiscardCommand { get; }

        // One per button - "Schematic images" and "KiCad data" each open BoardFilesWindow on its
        // own section (owner request, 2026-09-24).
        public ICommand ManageSchematicImagesCommand { get; }
        public ICommand ManageKiCadDataCommand { get; }
        public ICommand ViewDriftCommand { get; }
        public ICommand SubmitCommand { get; }

        // Whether this draft can be sent at all, and why not when it cannot. Disabled rather than
        // hidden - see the markup's own comment.
        public bool CanSubmit { get; }
        public string SubmitTooltip { get; }

        // Whether to show the drift line and the "What changed" button (the "Updated officially"
        // chip beside the name was removed, owner request 2026-09-28). Only a real
        // difference counts - an unrecorded base revision is not evidence of drift, so it stays
        // silent (see DraftRevisionComparer).
        public bool HasDrift { get; }

        // One plain line saying what happened, worded to match what the data actually supports:
        // "updated" only when both revisions parsed and the official one is genuinely later,
        // "changed" otherwise. Never promises an ordering that could not be established.
        public string DriftSummary { get; }

        public DraftListItem(
            HardwareBoardEntry entry,
            DraftStatus status,
            int changeCount,
            BoardProblemCounts problems,
            DraftDriftState driftState,
            Func<HardwareBoardEntry, Task> discard,
            Func<HardwareBoardEntry, BoardFilesSection, Task> manageFiles,
            Func<HardwareBoardEntry, Task> viewDrift,
            Func<HardwareBoardEntry, Task> submit,
            bool isTableOpen,
            Func<HardwareBoardEntry, Task> toggleTable,
            SubmissionReceipt? lastSubmission,
            bool alreadySent = false)
        {
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
            this.TableButtonText = isTableOpen ? "Close table" : "Edit in table format";
            this.TableButtonTooltip = isTableOpen
                ? "Close the table. You are asked first if it has edits that are not saved."
                : "Edit this draft's data as sheets and rows, like its Excel file - with every " +
                  "difference from the published data coloured in";
            this.EditTableCommand = new ActionCommand(() => Dispatcher.UIThread.InvokeAsync(() => toggleTable(entry)));

            int rowCount = changeCount;

            // A board that exists only as a draft (session 2c, task 9) is described as what it is
            // rather than by a row count. "0 rows changed" would be actively misleading on a
            // freshly created board: nothing was CHANGED because the whole board is new, and the
            // number says nothing about the thing the user actually made.
            string rowText = rowCount == 1 ? "1 row" : $"{rowCount} rows";

            this.RowSummary = status.IsNewBoard
                ? (rowCount == 0 ? "New board, nothing added yet" : $"New board, {rowText} so far")
                : $"{rowText} changed";

            this.HasDrift = driftState is DraftDriftState.OfficialIsNewer or DraftDriftState.Changed;

            this.DriftSummary = driftState switch
            {
                DraftDriftState.OfficialIsNewer =>
                    "The official data for this board has been updated since you started. Your edits are still applied.",
                DraftDriftState.Changed =>
                    "The official data for this board has changed since you started. Your edits are still applied.",
                _ => string.Empty,
            };

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
            this.CanSubmit = rowCount > 0 && !alreadySent;

            this.SubmitTooltip = rowCount == 0
                ? "There is nothing to send yet. Add some data to this board first."
                : alreadySent && lastSubmission is not null
                    ? SubmissionReceiptPresenter.DescribeAlreadySent(lastSubmission)
                : this.HasErrors
                    ? $"This draft has {this.ErrorsText} to fix first - Submit opens its table on them instead of sending."
                    : "Send this board for review. You do not need an account - only an email address, " +
                      "so you can be told whether it was accepted.";

            this.DiscardCommand = new ActionCommand(() => Dispatcher.UIThread.InvokeAsync(() => discard(entry)));
            this.ManageSchematicImagesCommand = new ActionCommand(() =>
                Dispatcher.UIThread.InvokeAsync(() => manageFiles(entry, BoardFilesSection.SchematicImages)));
            this.ManageKiCadDataCommand = new ActionCommand(() =>
                Dispatcher.UIThread.InvokeAsync(() => manageFiles(entry, BoardFilesSection.KiCadData)));
            this.ViewDriftCommand = new ActionCommand(() => Dispatcher.UIThread.InvokeAsync(() => viewDrift(entry)));
            this.SubmitCommand = new ActionCommand(() => Dispatcher.UIThread.InvokeAsync(() => submit(entry)));
        }
    }
}
