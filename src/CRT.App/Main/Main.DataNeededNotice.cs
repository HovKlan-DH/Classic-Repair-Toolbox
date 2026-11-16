using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // *** A BOARD CHECKED AGAIN WHOSE DATA IS NOT HERE (owner request, 2026-11-14: "If then clicked/checked, and it does not
    // have the data, then it should inform about this in the bottom-right corner with an info-panel, so the user should take a
    // data refresh"). *** A data check leaves out the hardware and boards unchecked under "Visible hardware, boards and
    // schematics" (DataSyncScope). Checked again on the "Configuration" tab, each board under the key is looked at off the UI
    // thread (DataManager.BoardsWithoutData: its board file not there or behind, or a stored file it names not there), and the
    // card names those - one by name, several by count (DataNeededWording). "Check for new data" runs the data check, which
    // takes them in now they are checked.
    //
    // The card stays until closed, until its boards are all unchecked again, or until a data check runs (HideDataNeededNotice,
    // from Main.CheckForDataUpdatesNowAsync). See Main.axaml.cs for the file map.
    // ###########################################################################################
    public partial class Main
    {
        // The boards the card names: checked, their data not here - in the order they were found.
        private List<HardwareBoardEntry> thisBoardsNeedingData = [];

        // What the card's last look found, for a test - the look runs on the pool and ends on the UI thread.
        internal Task LookForDataNeededForTests { get; private set; } = Task.CompletedTask;

        // From ApplyCatalogueVisibility, for a hardware or a board key - checked or unchecked.
        private void FollowCatalogueChangeInDataNeeded(string changedKey)
        {
            this.LookForDataNeededForTests = this.LookForDataNeededAsync(changedKey);
        }

        // ###########################################################################################
        // How many times the card was let go of - a data check starting, or the card closed. A look that ended after that
        // says nothing (code review, 2026-11-16): the check began after the box was ticked, so it brings the board's data, and
        // the card it put back stayed after the data had come.
        // ###########################################################################################
        private int thisDataNeededLetGo;

        private async Task LookForDataNeededAsync(string changedKey)
        {
            int letGoBefore = this.thisDataNeededLetGo;

            try
            {
                IReadOnlySet<string> uncheckedKeys = UserSettings.CatalogueUncheckedKeysSnapshot;

                List<HardwareBoardEntry> checkedNow = DataManager.HardwareBoards
                    .Where(entry => Main.KeyNamesBoardOrItsHardware(changedKey, entry))
                    .Where(entry => CatalogueVisibility.IsBoardVisible(uncheckedKeys, entry.HardwareName, entry.BoardName))
                    .ToList();

                List<HardwareBoardEntry> missing = checkedNow.Count == 0
                    ? []
                    : await Task.Run(() => DataManager.BoardsWithoutData(checkedNow));

                if (letGoBefore != this.thisDataNeededLetGo)
                    return;

                // As the checkboxes stand once the look is done - another may have been clicked meanwhile.
                uncheckedKeys = UserSettings.CatalogueUncheckedKeysSnapshot;

                this.thisBoardsNeedingData = this.thisBoardsNeedingData
                    .Concat(missing)
                    .Where(entry => CatalogueVisibility.IsBoardVisible(uncheckedKeys, entry.HardwareName, entry.BoardName))
                    .DistinctBy(entry => entry.ExcelDataFile, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                this.ShowDataNeededNotice();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not tell whether the boards checked have their data - [{ex.Message}]");
            }
        }

        // Whether a hardware key or a board key of the "Configuration" tab's tree names `entry` - the board's own key, or its hardware's.
        private static bool KeyNamesBoardOrItsHardware(string key, HardwareBoardEntry entry) =>
            CatalogueVisibility.KeyComparer.Equals(key, CatalogueVisibility.BuildKey(entry.HardwareName)) ||
            CatalogueVisibility.KeyComparer.Equals(key, CatalogueVisibility.BuildKey(entry.HardwareName, entry.BoardName));

        private void ShowDataNeededNotice()
        {
            this.DataNeededBanner.IsVisible = this.thisBoardsNeedingData.Count > 0;

            NoticeEmphasis.Show(this.DataNeededBannerText, this.thisBoardsNeedingData.Count > 0
                ? DataNeededWording.Of(this.thisBoardsNeedingData)
                : string.Empty);
        }

        // A data check is running - the very one the card asks for, or another: what it names comes with it.
        internal void HideDataNeededNotice()
        {
            this.thisDataNeededLetGo++;
            this.thisBoardsNeedingData = [];
            this.ShowDataNeededNotice();
        }

        private void OnDataNeededBannerDismiss(object? sender, RoutedEventArgs e) => this.HideDataNeededNotice();

        // ###########################################################################################
        // *** A CHECK RUNNING AS THE BUTTON IS PRESSED IS FOLLOWED BY ONE MORE (code review, 2026-11-16). *** The check running
        // worked out what it takes as it began (DataSyncScope), before these boards were checked - so it brings none of them. The
        // card went and nothing more happened. Now one more check runs as soon as the running one ends (FollowDataCheckAskedFor,
        // from StopDataSyncStatusIconSpin), and that one takes them in.
        // ###########################################################################################
        private bool thisDataCheckAskedForMeanwhile;

        // The check the card's button runs - a test's stand-in, as a real check goes to the network.
        internal Func<Task>? DataCheckOverrideForTests { get; set; }

        private void OnDataNeededBannerCheckClick(object? sender, RoutedEventArgs e)
        {
            this.HideDataNeededNotice();

            if (this.IsDataSyncStatusIconSpinning)
            {
                this.thisDataCheckAskedForMeanwhile = true;
                return;
            }

            this.RunDataCheckFromCard();
        }

        // From StopDataSyncStatusIconSpin as the last check running ends.
        private void FollowDataCheckAskedFor()
        {
            if (!this.thisDataCheckAskedForMeanwhile)
                return;

            this.thisDataCheckAskedForMeanwhile = false;

            // Posted: the check that just ended is still finishing its own card.
            Dispatcher.UIThread.Post(() =>
            {
                if (!this.IsDataSyncStatusIconSpinning)
                    this.RunDataCheckFromCard();
            });
        }

        // "async void" in all but name - see OnBoardSelectionChanged; the icon's own click runs the same.
        private async void RunDataCheckFromCard()
        {
            try
            {
                await (this.DataCheckOverrideForTests?.Invoke() ?? this.CheckForDataUpdatesNowAsync());
            }
            catch (Exception ex)
            {
                Logger.Critical($"Checking for data updates failed: [{ex.Message}]");
            }
        }
    }
}
