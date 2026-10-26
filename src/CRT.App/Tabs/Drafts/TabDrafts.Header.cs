using Avalonia.Controls;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE DRAFT'S HEADER, ON ITS ROW (CRT 4.0 step 1's last slice, 2026-10-25 - the plan:
    //
    //     | ZX Spectrum 16K/48K - Issue 4B     Draft: 3 changes - Saved   [Undo] [Redo] |
    //     | [Review changes] [Submit]                                  [Discard draft]  |
    //
    // "Undo and redo work across every editor for the whole session ... Each way of throwing work
    // away says what it covers: Undo - one step. Undo changes to IC1 - one component, back to its
    // published values. Discard draft - everything, after listing what will be lost. Review changes
    // shows a readable summary, and the Submit dialog shows the same").
    //
    // Owns: one DraftHeaderState per listed draft, kept for the run (thisHeaders); the row's Undo,
    // Redo and Try again; "Review changes" with its "Undo changes to <component>"; and the summary the
    // Submit and Discard draft dialogs show (SummaryForAsync).
    //
    // *** A STEP TAKEN FROM THE ROW IS SAVED AT ONCE. *** With the table open, the table takes it -
    // its Ctrl+Z, its status line, its timer. With the table closed nothing else would save it, so
    // the row does, off the UI thread, and the board on screen catches up.
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        private readonly Dictionary<string, DraftHeaderState> thisHeaders = new(StringComparer.OrdinalIgnoreCase);

        private DraftRowHeaderActions? thisHeaderActions;

        private DraftRowHeaderActions HeaderActions =>
            this.thisHeaderActions ??= new DraftRowHeaderActions(this.UndoAsync, this.RedoAsync, this.TryAgainAsync, this.ReviewChangesAsync);

        // The row's header of this draft, following the draft's session as it is now - none until an
        // editor opens one; the next refresh follows it.
        private DraftHeaderState HeaderFor(HardwareBoardEntry entry)
        {
            if (!this.thisHeaders.TryGetValue(entry.ExcelDataFile, out DraftHeaderState? header))
            {
                header = new DraftHeaderState();
                this.thisHeaders[entry.ExcelDataFile] = header;
            }

            header.Follow(DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile));
            return header;
        }

        // The headers of drafts no longer listed let go of their sessions.
        private void ForgetHeadersExcept(IEnumerable<HardwareBoardEntry> listed)
        {
            var keep = new HashSet<string>(listed.Select(entry => entry.ExcelDataFile), StringComparer.OrdinalIgnoreCase);

            foreach (string gone in this.thisHeaders.Keys.Where(key => !keep.Contains(key)).ToList())
            {
                this.thisHeaders[gone].Forget();
                this.thisHeaders.Remove(gone);
            }
        }

        internal DraftHeaderState? HeaderForTests(string excelDataFile) =>
            this.thisHeaders.TryGetValue(excelDataFile, out DraftHeaderState? header) ? header : null;

        // ---- Undo, Redo, Try again ----------------------------------------------------------------

        internal Task UndoAsync(HardwareBoardEntry entry) => this.StepAsync(entry, undo: true);

        internal Task RedoAsync(HardwareBoardEntry entry) => this.StepAsync(entry, undo: false);

        private async Task StepAsync(HardwareBoardEntry entry, bool undo)
        {
            // The open table takes the step itself: it says what was undone, and shows where.
            if (this.IsTableOpenFor(entry))
            {
                if (undo)
                {
                    this.TableEditor.Undo();
                }
                else
                {
                    this.TableEditor.Redo();
                }

                return;
            }

            if (DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile) is not DraftSession session ||
                !(undo ? session.Undo() : session.Redo()))
            {
                return;
            }

            await session.SaveAsync();
            this.AfterDraftChanged(entry);
        }

        // ###########################################################################################
        // "Try again" after a save that could not be made - the open table's own, or the row's. A file
        // changed somewhere else is not tried again (it would be refused again, for ever): the table
        // opens on it, and its bar offers both versions - nothing thrown away unasked.
        // ###########################################################################################
        internal async Task TryAgainAsync(HardwareBoardEntry entry)
        {
            if (this.IsTableOpenFor(entry))
            {
                await this.TableEditor.TryAgainAsync();
                return;
            }

            if (DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile) is not DraftSession session)
            {
                return;
            }

            if (session.SaveFailure == DraftWorkbookEditOutcome.ChangedOnDisk)
            {
                await this.OpenTableAsync(entry);
                return;
            }

            session.Retry();
            await session.SaveAsync();
            this.AfterDraftChanged(entry);
        }

        // What a change made from this tab needs: the row's numbers, and the board on screen.
        private void AfterDraftChanged(HardwareBoardEntry entry)
        {
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

        // ---- Review changes -----------------------------------------------------------------------

        // ###########################################################################################
        // "Review changes": the draft's summary (DraftChangeSummary), each component with its own
        // "Undo changes to <component>". The table's held change goes in first, so the summary is the
        // draft as it is on screen.
        // ###########################################################################################
        internal async Task ReviewChangesAsync(HardwareBoardEntry entry)
        {
            if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
            {
                return;
            }

            DraftChangeSummary summary = await BusyOverlay.RunLocalAsync(this, CrtWaitWording.ReadingChanges, () => this.SummaryForAsync(entry));

            var window = this.ReviewWindowFactoryForTests?.Invoke() ?? new ReviewChangesWindow();
            window.Initialize(entry.ToString(), summary, label => this.UndoComponentChangesAsync(entry, label));

            await window.ShowDialog(ownerWindow);
        }

        // A test's window, which it then drives - the real one would block on ShowDialog.
        internal Func<ReviewChangesWindow>? ReviewWindowFactoryForTests { get; set; }

        // ###########################################################################################
        // What the draft changes, read off the UI thread - the draft as every editor shows it, against
        // the published board (DraftChangeSummaryReader). What Review changes, Submit and Discard
        // draft show.
        // ###########################################################################################
        internal Task<DraftChangeSummary> SummaryForAsync(HardwareBoardEntry entry)
        {
            DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile)?.CollectHeldChanges(DraftSaveWording.ReadAgainStep);

            string dataRoot = DataManager.DataRoot;
            string draftsRoot = DraftManager.DraftsRoot;
            bool draftOnly = TabDrafts.IsDraftOnly(entry);
            BoardData? publishedOverride = this.PublishedBoardOverrideForTests?.Invoke(entry);

            return Task.Run(() => DraftChangeSummaryReader.Read(dataRoot, draftsRoot, entry.ExcelDataFile, draftOnly, publishedOverride));
        }

        // ###########################################################################################
        // "Undo changes to <component>" from Review changes: the component back to its published
        // values as one step (DraftEdits.UndoComponentChangesAsync), then the summary read again - or
        // why it could not be done.
        // ###########################################################################################
        internal async Task<(DraftChangeSummary Summary, string? Failure)> UndoComponentChangesAsync(HardwareBoardEntry entry, string label)
        {
            string dataRoot = DataManager.DataRoot;
            bool draftOnly = TabDrafts.IsDraftOnly(entry);

            BoardData? published = this.PublishedBoardOverrideForTests is { } publishedFor
                ? publishedFor(entry)
                : await Task.Run(() => DraftChangeSummaryReader.PublishedBoard(dataRoot, entry.ExcelDataFile, draftOnly));

            DraftWorkbookEditOutcome outcome = await DraftEdits.UndoComponentChangesAsync(
                entry.ExcelDataFile,
                label,
                published,
                draftOnly ? string.Empty : DraftBoardSource.PublishedPathOf(dataRoot, entry.ExcelDataFile));

            if (outcome == DraftWorkbookEditOutcome.Saved)
            {
                this.AfterDraftChanged(entry);
            }

            return (await this.SummaryForAsync(entry), outcome == DraftWorkbookEditOutcome.Saved ? null : DraftSaveWording.EditNotSaved(outcome));
        }

        // A board that exists only as a draft has no published board to compare with.
        private static bool IsDraftOnly(HardwareBoardEntry entry) =>
            entry.IsDraftOnly ||
            DraftStatusReader.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, entry.ExcelDataFile, entry.IsPublished)?.IsNewBoard == true;
    }
}
