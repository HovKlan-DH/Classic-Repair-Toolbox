using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.Theming;
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
    // This reuses the SyncBanner card, and follows _isShowingDataSyncDisabledBanner's precedent
    // for shared ownership of that one text slot. A data check used to share it too, and win; it
    // has a card of its own since the notices float (2026-10-11), so the two no longer meet.
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
        // Cheap: the draft's marker, and the published workbook's revision and content hash, both
        // remembered by the file's length and write time (DraftDriftCheck).
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

            var entry = this.GetCurrentBoardEntry();

            if (entry == null)
            {
                return DraftDriftState.Unknown;
            }

            // The one rule the Drafts row and the window use too (code review, 2026-10-13): by
            // CONTENT for a draft that kept what it was copied from, so a second publish on the same
            // day is seen; by revision otherwise.
            return DraftDriftCheck.Resolve(
                this.CurrentDraftStatus(entry),
                DraftBoardSource.PublishedPathOf(DataManager.DataRoot, entry.ExcelDataFile));
        }

        private DraftStatus? CurrentDraftStatus(HardwareBoardEntry entry) =>
            DraftStatusReader.Resolve(
                DataManager.DataRoot,
                DraftManager.DraftsRoot,
                entry.ExcelDataFile,
                entry.IsPublished);

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

            // *** NOT "YOUR EDITS ARE STILL APPLIED" ANY MORE (code review, 2026-10-13). *** That was
            // untrue: a draft is a complete copy and does not get the update by itself. The line
            // says so, and offers to bring the draft up to date when it can (DraftDriftWording).
            var entry = this.GetCurrentBoardEntry();
            bool canBringUpToDate = entry != null && this.CurrentDraftStatus(entry) is { HasBase: true };

            NoticeEmphasis.Show(
                this.SyncBannerText,
                DraftDriftWording.Line(this._currentBoardDriftState, canBringUpToDate));

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
            try
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

                // The same report the Drafts tab's "What changed" opens (DraftChangeReport.For).
                string publishedPath = DraftBoardSource.PublishedPathOf(
                    DataManager.DataRoot,
                    entry.ExcelDataFile);

                string draftsRoot = DraftManager.DraftsRoot;

                DraftChangeReport report = await BusyOverlay.RunLocalAsync(
                    this,
                    CrtWaitWording.ComparingWithOfficial,
                    () => Task.Run(() => DraftChangeReport.For(status, draftsRoot, publishedPath)));

                var window = new DraftDriftWindow();
                window.Initialize(entry.ToString(), entry.ExcelDataFile, report);
                await window.ShowDialog(this);

                // Brought up to date: the board on screen is this draft's, read again.
                if (window.BroughtUpToDate)
                {
                    this.ReloadCurrentBoardAfterDraftFileImport(entry.ExcelDataFile);
                }

                // Dismissing the warning or bringing the draft up to date changes the drift.
                this._currentBoardDriftState = this.ResolveCurrentBoardDriftState();

                if (this._currentBoardDriftState is DraftDriftState.InSync or DraftDriftState.Unknown)
                {
                    this.HideDraftDriftBanner();
                }

                this.TabDrafts.RefreshDrafts();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT
                // (code review, 2026-10-13 - AsyncVoidHandlersGuardTests).
                Logger.Critical($"OnSyncBannerDriftClick failed - [{ex}]");
            }
        }
    }
}
