using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE BANNER UNDER THE TABS ABOUT WHICH SOURCE TO DOWNLOAD FROM. See Main.axaml.cs for the file
    // map of the whole partial class. It carries one of two notices, never both - each holds only
    // for one setting of "Download data from the BETA source":
    //
    //   - "YOUR SUBMISSION IS NOW IN THE BETA SOURCE - TICK BETA TO TRY IT" (owner request,
    //     2026-10-03), while this machine downloads from the STABLE source and a submission of
    //     theirs has been accepted into BETA;
    //   - "YOUR SUBMISSION IS NOW IN THE STABLE SOURCE - SWITCH BACK FROM BETA" (owner request,
    //     2026-09-27), while this machine downloads from BETA and a submission has been published
    //     to production.
    //
    // WHICH submissions, and the words, are SubmissionReceiptPresenter's (pure, unit tested). The
    // banner is re-derived from the receipts each time rather than raised once on a state change:
    // a notice raised as the app closed would otherwise be lost for good. Closing it is remembered
    // per submission and per notice.
    // ###########################################################################################
    public partial class Main
    {
        // The submissions the banner is showing, so closing it remembers exactly those - and which
        // of the two notices it is.
        private IReadOnlyList<long> thisSourceSwitchNoticeIds = [];
        private bool thisSourceNoticeIsBetaTry;

        // ###########################################################################################
        // Shows or hides the banner to match the receipts and the source settings. Called after
        // every submission status check and whenever "Download data from BETA source" or "Check for
        // new or updated data at application launch" changes. Never throws: it runs on the launch
        // path, where an exception would fault a discarded Task.
        // ###########################################################################################
        internal void RefreshSourceSwitchNotice()
        {
            try
            {
                IReadOnlyList<SubmissionReceipt> receipts = SubmissionReceiptStore.All;
                bool onBeta = UserSettings.DownloadDataFromTestSource;

                IReadOnlyList<SubmissionReceipt> switchBack = SubmissionReceiptPresenter.NeedingSourceSwitchNotice(receipts, onBeta);
                IReadOnlyList<SubmissionReceipt> tryBeta = SubmissionReceiptPresenter.NeedingBetaTryNotice(receipts, onBeta);

                this.thisSourceNoticeIsBetaTry = switchBack.Count == 0 && tryBeta.Count > 0;
                IReadOnlyList<SubmissionReceipt> shown = this.thisSourceNoticeIsBetaTry ? tryBeta : switchBack;

                this.thisSourceSwitchNoticeIds = shown.Select(receipt => receipt.SubmissionId).ToList();

                this.SourceSwitchBanner.IsVisible = shown.Count > 0;
                this.SourceSwitchBannerText.Text = shown.Count == 0
                    ? string.Empty
                    : this.thisSourceNoticeIsBetaTry
                        ? SubmissionReceiptPresenter.DescribeBetaTryNotice(shown, UserSettings.CheckDataOnLaunch)
                        : SubmissionReceiptPresenter.DescribeSourceSwitchNotice(shown);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not refresh the BETA source notice: [{ex.Message}]");
            }
        }

        // Takes the contributor to where the setting is. Deliberately does not flip it here: either
        // way it starts a download, which the Configuration tab already explains and asks about.
        private void OnSourceSwitchBannerConfigurationClick(object? sender, RoutedEventArgs e)
        {
            if (this.MainTabControl != null && this.ConfigurationTabItem != null)
            {
                this.MainTabControl.SelectedItem = this.ConfigurationTabItem;
            }
        }

        private void OnSourceSwitchBannerDismiss(object? sender, RoutedEventArgs e)
        {
            if (this.thisSourceNoticeIsBetaTry)
            {
                SubmissionReceiptStore.DismissBetaNotice(this.thisSourceSwitchNoticeIds);
            }
            else
            {
                SubmissionReceiptStore.DismissSourceNotice(this.thisSourceSwitchNoticeIds);
            }

            this.RefreshSourceSwitchNotice();
        }
    }
}
