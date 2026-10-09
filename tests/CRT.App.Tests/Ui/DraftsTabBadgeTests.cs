using Avalonia.Controls;
using Handlers.DataHandling;
using Handlers.Online;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Drafts tab's badge (owner request, 2026-09-30: "Like the "Maintainer" tab now has a badge,
// then I think the "Draft" also should have same functionality, as the contributor likewise can
// receive an update") - Main.SubmissionChecks.cs.
//
// On a real Main, built and never shown, with the receipts store, the settings and the workbook
// folder all pointed at a temp folder - so no test reads or writes the user's own files. In the
// "HeadlessUi" collection for the dispatcher; it resets the receipts store, as MySubmissionsWindowTests
// does, since one class can join only one collection.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class DraftsTabBadgeTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public DraftsTabBadgeTests()
    {
        SubmissionReceiptStore.LoadFrom(Path.Combine(this.thisWorkspace.Root, "submissions.json"));
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile(Guid.NewGuid().ToString("N") + ".json", "{}"));
    }

    public void Dispose()
    {
        SubmissionReceiptStore.LoadFrom(string.Empty);
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile(Guid.NewGuid().ToString("N") + ".json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private static SubmissionReceipt Receipt(long id, string state) => new()
    {
        SubmissionId = id,
        UploadToken = "tok" + id,
        BoardId = "Commodore/C64/250407/Data C64 250407.xlsx",
        Summary = "Fixed U8 pinout",
        SentUtc = new DateTimeOffset(2026, 9, 21, 12, 0, 0, TimeSpan.Zero),
        LastKnownState = state
    };

    // ###########################################################################################
    // The number beside "Drafts" is the number of submissions with news not seen yet - the same as
    // the "My submissions" button inside the tab - and reading the news takes it away. A submission
    // still waiting for a maintainer is not news.
    // ###########################################################################################
    [Fact]
    public void The_Drafts_tab_counts_the_submissions_with_unread_news_and_reading_them_clears_it()
    {
        UiTest.Run(() =>
        {
            SubmissionReceiptStore.Record(Receipt(41, "changes_requested"));
            SubmissionReceiptStore.Record(Receipt(42, "merged"));
            SubmissionReceiptStore.Record(Receipt(43, "pending"));

            var main = new CRT.Main();
            main.ApplyDraftsTabVisibility();

            Assert.Equal("2", main.DraftsTabBadgeForTests);
            Assert.Equal(SubmissionReceiptStore.UnreadCommentCount().ToString(), main.DraftsTabBadgeForTests);

            // News keeps the tab on screen even with no drafts - otherwise the badge had nowhere to be.
            Assert.True(main.DraftsTabItem.IsVisible);

            SubmissionReceiptStore.AcknowledgeComment(41);
            SubmissionReceiptStore.AcknowledgeComment(42);
            main.ApplyDraftsTabVisibility();

            Assert.Null(main.DraftsTabBadgeForTests);
        });
    }

    // ###########################################################################################
    // The minute check's follow-up redraws the Drafts tab only when a receipt moved or a draft was
    // retired (code review, 2026-10-04): every draft row and both drop-down templates were rebuilt
    // once a minute for a receipt that had said the same thing for weeks. The launch check and "My
    // submissions" closing still redraw every time.
    // ###########################################################################################
    [Fact]
    public async Task The_minute_checks_follow_up_leaves_the_Drafts_tab_alone_when_nothing_moved()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();
            SubmissionReceiptStore.Record(Receipt(41, "changes_requested"));

            // Nothing moved on this tick, nothing to retire: not drawn again.
            await main.RetirePublishedDraftsAsync(refreshDraftsTab: false);
            Assert.Null(main.DraftsTabBadgeForTests);

            // A receipt moved (or the launch check): drawn, badge and all.
            await main.RetirePublishedDraftsAsync(refreshDraftsTab: true);
            Assert.Equal("1", main.DraftsTabBadgeForTests);
        });
    }

    // ###########################################################################################
    // *** "UPDATE CRT" FROM THE STATUS CHECK REACHES THE CONTRIBUTOR (code review, 2026-10-04). ***
    // The banner that already says a newer application is needed shows the server's words, and the
    // minute checks stop for this run - every one would get the same answer. Shown once: closed, it
    // is not raised again a minute later.
    // ###########################################################################################
    [Fact]
    public void An_update_CRT_answer_shows_the_servers_words_once_and_stops_the_minute_checks()
    {
        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            const string words = "Please update CRT to version [3.2.0] or newer.";

            main.ShowSubmissionChecksOutdated(words);

            Assert.Equal(SubmissionStatusRefresh.DescribeOutdated(words), main.RequiresAppUpdateBannerTextForTests);
            Assert.True(main.SubmissionChecksOutdatedForTests);

            main.MainExcelRequiresAppUpdateBanner.IsVisible = false;
            main.ShowSubmissionChecksOutdated(words);

            Assert.Null(main.RequiresAppUpdateBannerTextForTests);
        });
    }

    // ###########################################################################################
    // *** BOTH REASONS FOR A NEWER CRT SHARE THE BANNER, NEITHER REPLACES THE OTHER (code review,
    // 2026-10-04). *** The status check's "update CRT" and the newer main Excel data used to write
    // the same text block, so whichever came second wiped the first - and the server's words, shown
    // once a run, were then gone for good. Dismissing forgets both; a reason raised again after
    // that comes back alone.
    // ###########################################################################################
    [Fact]
    public void The_update_CRT_answer_and_the_newer_main_Excel_data_are_both_said_in_either_order()
    {
        UiTest.Run(() =>
        {
            const string words = "Please update CRT to version [3.2.0] or newer.";
            string submissions = SubmissionStatusRefresh.DescribeOutdated(words);

            var serverFirst = new CRT.Main();
            serverFirst.ShowSubmissionChecksOutdated(words);
            serverFirst.ShowMainExcelRequiresAppUpdateBannerForTests();

            string both = serverFirst.RequiresAppUpdateBannerTextForTests!;
            Assert.Contains(submissions, both, StringComparison.Ordinal);
            Assert.Contains("Newer main Excel data file is available", both, StringComparison.Ordinal);

            var excelFirst = new CRT.Main();
            excelFirst.ShowMainExcelRequiresAppUpdateBannerForTests();
            excelFirst.ShowSubmissionChecksOutdated(words);

            Assert.Equal(both, excelFirst.RequiresAppUpdateBannerTextForTests);

            // Dismissed, then the main Excel reason raised again by the next sync: that one alone.
            excelFirst.GetControl<Button>("MainExcelRequiresAppUpdateBannerDismissButton")
                .RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
            Assert.Null(excelFirst.RequiresAppUpdateBannerTextForTests);

            excelFirst.ShowMainExcelRequiresAppUpdateBannerForTests();

            Assert.DoesNotContain(submissions, excelFirst.RequiresAppUpdateBannerTextForTests!, StringComparison.Ordinal);
            Assert.StartsWith("Newer main Excel data file is available", excelFirst.RequiresAppUpdateBannerTextForTests!, StringComparison.Ordinal);
        });
    }

    // Nothing sent at all: no badge (and, with no drafts either, no tab).
    [Fact]
    public void Nothing_sent_is_no_badge()
    {
        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            main.ApplyDraftsTabVisibility();

            Assert.Null(main.DraftsTabBadgeForTests);
            Assert.False(main.DraftsTabItem.IsVisible);
        });
    }

    // ###########################################################################################
    // The check while CRT runs asks only when there is something to ask about and a badge to see:
    // never while the window is minimised, never for a contributor whose submissions are all
    // decided - and so never for somebody who has not contributed at all.
    // ###########################################################################################
    [Fact]
    public void The_check_while_CRT_runs_asks_only_about_open_submissions_with_the_window_up()
    {
        var now = new DateTimeOffset(2026, 9, 30, 12, 0, 0, TimeSpan.Zero);

        SubmissionReceipt[] open = [Receipt(41, "pending"), Receipt(42, "rejected")];
        SubmissionReceipt[] decided = [Receipt(42, "rejected"), Receipt(44, "published")];

        Assert.True(SubmissionStatusRefresh.ChecksPeriodically(open, now, windowMinimised: false));
        Assert.False(SubmissionStatusRefresh.ChecksPeriodically(open, now, windowMinimised: true));
        Assert.False(SubmissionStatusRefresh.ChecksPeriodically(decided, now, windowMinimised: false));
        Assert.False(SubmissionStatusRefresh.ChecksPeriodically([], now, windowMinimised: false));

        // ONE minute since 2026-10-01 (owner decision), down from five. The check is what retires a
        // published draft, and the project owner accepted the traffic: "this is not a high volume
        // thing ... not that many users globally". ChecksPeriodically above is what keeps that cheap.
        Assert.Equal(TimeSpan.FromMinutes(1), SubmissionStatusRefresh.PeriodicInterval);
    }

    // One rule for every tab badge: nothing is no badge, and past 99 it stops counting.
    [Fact]
    public void A_tab_badge_is_hidden_at_zero_and_stops_at_99_plus()
    {
        Assert.Null(TabBadge.Text(0));
        Assert.Null(TabBadge.Text(-3));
        Assert.Equal("1", TabBadge.Text(1));
        Assert.Equal("99", TabBadge.Text(99));
        Assert.Equal("99+", TabBadge.Text(100));
    }
}
