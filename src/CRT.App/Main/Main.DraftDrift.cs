using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.IO;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // The drift warning on the BOARD itself (NewContributeStrategy.md Phase 2, session 2d) - the
    // secondary surface, next to the Drafts tab's own chip. See Main.axaml.cs for the file map of
    // the whole partial class.
    //
    // It puts the warning in front of someone who is looking at the board rather than at the
    // Drafts tab, which is where a contributor actually spends their time. The Drafts tab remains
    // the primary home: it is the one surface that is ABOUT drafts, and drift is a per-board fact.
    //
    // This reuses the existing SyncBanner rather than adding a fourth banner mechanism, and follows
    // _isShowingDataSyncDisabledBanner's precedent for shared ownership of that one text slot:
    // a sync flow always wins, because a sync in progress is both more urgent and transient. The
    // drift banner re-evaluates once the sync has settled.
    // ###########################################################################################
    public partial class Main
    {
        // The drift state of the board currently on screen, resolved on every board load. Kept so
        // the banner and the "What changed" button agree without re-reading anything.
        private DraftDriftState _currentBoardDriftState = DraftDriftState.Unknown;

        private bool _isShowingDraftDriftBanner;

        // ###########################################################################################
        // Works out whether the board now on screen has drifted, and shows or hides the banner to
        // match. Called after every board load, and again after a sync settles.
        //
        // Costs nothing: DataManager published the draft's base revision and the official
        // RevisionDate as part of the load that just happened, so this reads two values it already
        // has rather than touching the disk.
        // ###########################################################################################
        private void EvaluateDraftDriftForCurrentBoard()
        {
            this._currentBoardDriftState = ResolveCurrentBoardDriftState();

            bool shouldWarn = this._currentBoardDriftState
                is DraftDriftState.OfficialIsNewer or DraftDriftState.Changed;

            // Suppressed while "view boards as officially published" is on. The drift is still TRUE
            // and _currentBoardDriftState still records it - but warning about a draft, on a screen
            // that is claiming to show the published version, contradicts what the toggle promises.
            // Suppressed at the banner rather than by blinding the data, so nothing else has to
            // know about the toggle.
            if (shouldWarn && !UserSettings.ViewOfficialPublishedOnly)
            {
                this.ShowDraftDriftBanner();
            }
            else
            {
                this.HideDraftDriftBanner();
            }
        }

        // ###########################################################################################
        // The drift state of the board on screen, or Unknown when there is nothing to compare.
        //
        // A draft-only board is Unknown by construction: it has no official counterpart, so its
        // deliberately blank base revision must never be read as drift.
        // ###########################################################################################
        private DraftDriftState ResolveCurrentBoardDriftState()
        {
            if (this._currentBoardData == null || DataManager.LastLoadedDraftIsNewBoard)
            {
                return DraftDriftState.Unknown;
            }

            return DraftRevisionComparer.Compare(
                DataManager.LastLoadedDraftBaseRevision,
                this._currentBoardData.RevisionDate);
        }

        // ###########################################################################################
        // Shows the drift banner, unless a sync flow already owns the banner - a sync in progress
        // is more urgent and is transient, and stealing its text slot would hide it.
        // ###########################################################################################
        private void ShowDraftDriftBanner()
        {
            if (this._isShowingDataSyncDisabledBanner)
            {
                return;
            }

            // Worded to match what the data supports: "updated" only when both revisions parsed and
            // the official one is genuinely later. See DraftRevisionComparer.
            string what = this._currentBoardDriftState == DraftDriftState.OfficialIsNewer
                ? "has been updated"
                : "has changed";

            // The second sentence is load-bearing: without it this reads as "your work may be
            // lost", which is false - Drafts/ and Data/ are separate trees and the overlay still
            // applies.
            this.SyncBannerText.Text =
                $"The official data for this board {what} since you started your draft. Your edits are still applied.";

            this.SyncBannerRefreshButton.IsVisible = false;
            this.SyncBannerDriftButton.IsVisible = true;
            this.SyncBanner.IsVisible = true;
            this._isShowingDraftDriftBanner = true;
        }

        // ###########################################################################################
        // Hides the drift banner without disturbing any other banner flow - the same narrow shape
        // HideDataSyncDisabledBanner uses, and for the same reason: five different flows write to
        // this one text slot.
        // ###########################################################################################
        private void HideDraftDriftBanner()
        {
            if (!this._isShowingDraftDriftBanner)
            {
                return;
            }

            this.SyncBanner.IsVisible = false;
            this.SyncBannerDriftButton.IsVisible = false;
            this._isShowingDraftDriftBanner = false;
        }

        // ###########################################################################################
        // Opens the same report the Drafts tab's own "What changed" button opens. Loads the PURE
        // official board - the report asks whether each drafted row has an official counterpart,
        // and a merged board would already contain the very rows being asked about.
        // ###########################################################################################
        private async void OnSyncBannerDriftClick(object? sender, RoutedEventArgs e)
        {
            var entry = this.GetCurrentBoardEntry();
            if (entry == null)
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

            // Built by COMPARISON since Phase 6 - see DraftChangeReport for what that can and
            // cannot say, and TabDrafts.ViewDriftAsync for the same construction.
            string publishedPath = DraftBoardSource.PublishedPathOf(
                DataManager.DataRoot,
                entry.ExcelDataFile);

            BoardData? official = await BusyOverlay.RunLocalAsync(
                this,
                CrtWaitWording.ComparingWithOfficial,
                () => BoardDataReader.LoadAsync(publishedPath, publishedPath));

            var report = new DraftChangeReport
            {
                State = this._currentBoardDriftState,
                BaseRevision = status.BaseRevision,
                OfficialRevision = official?.RevisionDate ?? string.Empty,
                Rows = DraftStatusReader.DescribeChanges(status),
            };

            var window = new DraftDriftWindow();
            window.Initialize(entry.ToString(), entry.ExcelDataFile, report);

            await window.ShowDialog(this);

            // Dismissing the warning inside the window moves the draft's base revision, so the
            // banner has to be re-evaluated against what is now on disk.
            DraftStatus? refreshed = DraftStatusReader.Resolve(
                DataManager.DataRoot,
                DraftManager.DraftsRoot,
                entry.ExcelDataFile,
                entry.IsPublished);

            this._currentBoardDriftState = DraftRevisionComparer.Compare(
                refreshed?.BaseRevision ?? string.Empty,
                this._currentBoardData?.RevisionDate ?? string.Empty);

            if (this._currentBoardDriftState is DraftDriftState.InSync or DraftDriftState.Unknown)
            {
                this.HideDraftDriftBanner();
            }

            this.TabDrafts.RefreshDrafts();
        }
    }
}
