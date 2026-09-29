using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // "YOUR SUBMISSION IS NOW IN THE ONLINE SOURCE - SWITCH BACK FROM BETA" (owner request,
    // 2026-09-27). See Main.axaml.cs for the file map of the whole partial class.
    //
    // A contributor checks a submission in the BETA data by downloading from the BETA source. Once
    // a maintainer promotes it to production, they should go back to the normal online source - and
    // nothing told them so. This banner does, for every published submission of theirs they have
    // not dismissed it for, while this machine downloads from BETA.
    //
    // WHICH submissions, and the words, are SubmissionReceiptPresenter's (pure, unit tested). The
    // banner is re-derived from the receipts each time rather than raised once on a state change:
    // a published submission is never asked about again, so a notice raised as the app closed
    // would otherwise be lost for good. Dismissing it is remembered per submission.
    // ###########################################################################################
    public partial class Main
    {
        // The submissions the banner is showing, so dismissing it remembers exactly those.
        private IReadOnlyList<long> thisSourceSwitchNoticeIds = [];

        // ###########################################################################################
        // Shows or hides the banner to match the receipts and the source setting. Called after every
        // submission status check and whenever "Download data from BETA source" changes. Never
        // throws: it runs on the launch path, where an exception would fault a discarded Task.
        // ###########################################################################################
        internal void RefreshSourceSwitchNotice()
        {
            try
            {
                IReadOnlyList<SubmissionReceipt> needing = SubmissionReceiptPresenter.NeedingSourceSwitchNotice(
                    SubmissionReceiptStore.All,
                    UserSettings.DownloadDataFromTestSource);

                this.thisSourceSwitchNoticeIds = needing.Select(receipt => receipt.SubmissionId).ToList();

                this.SourceSwitchBanner.IsVisible = needing.Count > 0;
                this.SourceSwitchBannerText.Text = needing.Count > 0
                    ? SubmissionReceiptPresenter.DescribeSourceSwitchNotice(needing)
                    : string.Empty;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not refresh the switch-back-from-BETA notice: [{ex.Message}]");
            }
        }

        // Takes the contributor to where the setting is. Deliberately does not flip it here: turning
        // BETA off starts a download of the normal data, which the Configuration tab already
        // explains and asks about.
        private void OnSourceSwitchBannerConfigurationClick(object? sender, RoutedEventArgs e)
        {
            if (this.MainTabControl != null && this.ConfigurationTabItem != null)
            {
                this.MainTabControl.SelectedItem = this.ConfigurationTabItem;
            }
        }

        private void OnSourceSwitchBannerDismiss(object? sender, RoutedEventArgs e)
        {
            SubmissionReceiptStore.DismissSourceNotice(this.thisSourceSwitchNoticeIds);
            this.RefreshSourceSwitchNotice();
        }
    }
}
