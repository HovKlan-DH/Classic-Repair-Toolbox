using System;
using Avalonia.Controls;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Online;
using Handlers.Theming;

namespace CRT
{
    // ###########################################################################################
    // THE DRAFTS TAB'S BADGE (owner request, 2026-09-30: "Like the "Maintainer" tab now has a badge,
    // then I think the "Draft" also should have same functionality, as the contributor likewise can
    // receive an update").
    //
    // WHAT THIS PART OWNS: the number beside "Drafts" in the row of tabs, and asking the server
    // again every minute while CRT runs, so the number arrives without a restart.
    //
    // *** THE SAME NUMBER AS THE "MY SUBMISSIONS" BUTTON INSIDE THE TAB. *** Both are
    // SubmissionReceiptStore.UnreadCommentCount - submissions with a decision or a comment the
    // contributor has not seen - so the tab says what opening it will show, and reading the news in
    // "My submissions" clears both. It is drawn from ApplyDraftsTabVisibility, the one place every
    // change to drafts or receipts already ends up (after a status check, after "My submissions"
    // closes, after a draft is retired or discarded) - which is also what brings the tab itself
    // back for news about a submission whose draft is gone.
    //
    // *** THE CHECK IS THE LAUNCH CHECK, AGAIN (SubmissionStatusRefresh). *** Only still-open
    // submissions are asked about, failures are silent, and the launch check's own follow-up runs
    // after it: retire what is now published, then ApplyDraftsTabVisibility. How often, and when not
    // at all, is SubmissionStatusRefresh.PeriodicInterval / ChecksPeriodically.
    //
    // *** onFinished, NOT onChanged - THE LAUNCH CHECK'S RULE (code review, 2026-10-01). *** This
    // retired drafts only when a receipt MOVED on this tick. A receipt that turns "published" before
    // the board it was published into has synced finds nothing to retire - and never moves again,
    // so no later tick retried it, and the draft stayed until a restart. Retirement depends on a
    // state, so it follows every check (Main.StartAsync's note says the same of the launch check).
    // ###########################################################################################
    public partial class Main
    {
        private DispatcherTimer? thisSubmissionCheckTimer;
        private bool thisSubmissionCheckRunning;

        // The status check's "update CRT" has been said in the banner - once a run (see below).
        private bool thisSubmissionsOutdatedBannerShown;

        // ###########################################################################################
        // *** THE CHECKS STOP WHEN THE SERVER TURNS SUBMISSIONS AWAY - ONE RECORD OF THAT (code
        // review, 2026-10-10). *** It answers every minute the same way (code review, 2026-10-04),
        // and that is what covers the Drafts tab (AppUpdateRequirement's Drafts reason). The checks
        // kept a flag of their own, set only by their own refusal, so a 426 met by Submit, "My
        // submissions" or the discard notice covered the tab while the checks went on asking - each
        // refused, each asking /api/health again. Now they stop on the requirement, whatever set it.
        // ###########################################################################################
        private bool SubmissionChecksStopped => this.thisAppUpdateRequirement.IsRequiredFor(AppUpdateArea.Drafts);

        // ###########################################################################################
        // The status check was answered "update CRT": the server's words go in the banner that already
        // says an application update is needed - beside its main Excel reason when that is shown too,
        // never instead of it - and into the requirement, which covers the Drafts tab and stops the
        // minute checks (the same as the refusal's signal does, so the two never disagree). The
        // banner says it once: a contributor who closed it is not shown it again.
        // ###########################################################################################
        internal void ShowSubmissionChecksOutdated(string serversWords)
        {
            if (this.thisAppUpdateRequirement.ApplyRefusal(AppUpdateArea.Drafts, serversWords, Main.OwnVersion))
                this.ApplyAppUpdateRequired();

            if (this.thisSubmissionsOutdatedBannerShown)
                return;

            this.thisSubmissionsOutdatedBannerShown = true;
            this.ShowSubmissionsRequireAppUpdateBanner(SubmissionStatusRefresh.DescribeOutdated(serversWords));
        }

        // Whether the minute checks have stopped for an "update CRT" answer - for tests.
        internal bool SubmissionChecksOutdatedForTests => this.SubmissionChecksStopped;

        // The count, or no badge at all for nothing unread.
        private void ShowDraftsTabBadge(int unread)
        {
            string? text = TabBadge.Text(unread);

            this.DraftsTabBadgeText.Text = text ?? string.Empty;
            this.DraftsTabBadge.IsVisible = text is not null;
        }

        // The badge as drawn, or null when none is - for tests.
        internal string? DraftsTabBadgeForTests =>
            this.DraftsTabBadge.IsVisible ? this.DraftsTabBadgeText.Text : null;

        // Started once the window is up and the launch check has been started (StartAsync).
        private void StartSubmissionChecks()
        {
            if (this.thisSubmissionCheckTimer is not null)
                return;

            this.thisSubmissionCheckTimer = new DispatcherTimer { Interval = SubmissionStatusRefresh.PeriodicInterval };
            this.thisSubmissionCheckTimer.Tick += async (_, _) => await this.CheckSubmissionsAsync();
            this.thisSubmissionCheckTimer.Start();
        }

        private async System.Threading.Tasks.Task CheckSubmissionsAsync()
        {
            if (this.thisSubmissionCheckRunning ||
                this.SubmissionChecksStopped ||
                !SubmissionStatusRefresh.ChecksPeriodically(
                    SubmissionReceiptStore.All, DateTimeOffset.UtcNow, this.WindowState == WindowState.Minimized))
            {
                return;
            }

            this.thisSubmissionCheckRunning = true;

            try
            {
                // The launch check's own follow-up, after every check - see the header. The Drafts
                // tab is drawn again only when a receipt moved or a draft was retired (code review,
                // 2026-10-04: it was rebuilt every minute for nothing).
                await SubmissionStatusRefresh.RefreshQuietlyAsync(
                    new SubmissionClient().GetStatusAsync,
                    DateTimeOffset.UtcNow,
                    onFinished: changedCount => Dispatcher.UIThread.Post(
                        () => _ = this.RetirePublishedDraftsAsync(refreshDraftsTab: changedCount > 0)),
                    onOutdated: message => Dispatcher.UIThread.Post(() => this.ShowSubmissionChecksOutdated(message)));
            }
            finally
            {
                this.thisSubmissionCheckRunning = false;
            }
        }
    }
}
