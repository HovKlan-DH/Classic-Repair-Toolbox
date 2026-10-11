using System.Reflection;
using Avalonia.Controls;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// *** "THE CONTRIBUTOR DISCARDED THEIR DRAFT" IS SEEN WHEREVER A MAINTAINER DECIDES (owner request,
// 2026-09-28: "it should be clearly visible on the system and for the maintainer(s) - both in the
// BETA to PROD queue, but also in the normal queue"). ***
//
// The review queue's row, the opened submission (above its table), the "Beta > Prod" list and its
// plan, and the Boards screen's submission list - each drawn here from an answer carrying it, and
// checked to say nothing where it is absent. No server: every screen is handed its answer.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class DraftDiscardShownTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static readonly DateTimeOffset Discarded = new(2026, 9, 28, 9, 15, 0, TimeSpan.Zero);

    private static ReviewQueueRow Row(long id, DateTimeOffset? discarded) =>
        new(id, "Commodore/C128/310378", "pending", $"Change {id}", "dennis@example.com",
            DateTimeOffset.UtcNow.AddHours(-1), IsNewBoard: false, AwaitsYou: true, DraftDiscardedUtc: discarded);

    [Fact]
    public void The_queue_row_says_the_contributor_discarded_their_draft_and_only_that_row()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            main.ApplyQueueResponse(new ReviewQueueResponse(
                CanPublish: true,
                Submissions: [Row(41, Discarded), Row(42, null)],
                IsAdministrator: true));

            List<string> texts = main.QueueTextsForTests().ToList();

            Assert.Single(texts, text => text == DraftDiscardWording.Mark(Discarded));
            Assert.Equal(texts.IndexOf("Change 41") + 2, texts.IndexOf(DraftDiscardWording.Mark(Discarded)));
        });
    }

    // Above the submission's views, before anything else - the first thing to settle before approving.
    [Fact]
    public void An_opened_submission_warns_above_its_views_to_ask_the_contributor_first()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            ReviewQueueRow row = Row(41, Discarded);

            typeof(TabMaintainer).GetMethod("ShowSubmission", Any)!.Invoke(main, [row]);
            main.ShowDetail(new ReviewSubmissionDetail(
                row,
                true,
                new ReviewChangeSummaryView(false, []),
                [],
                new ReviewSubmissionAssets([]),
                [],
                new Dictionary<string, string>(),
                new Dictionary<string, string>(),
                [],
                null));

            TextBlock first = main.FindControl<StackPanel>("BeforeApprovingPanel")!.Children.OfType<TextBlock>().First();

            Assert.Equal(DraftDiscardWording.SubmissionWarning(Discarded, "dennis@example.com"), first.Text);
        });
    }

    [Fact]
    public async Task The_beta_list_marks_a_board_carrying_a_discarded_draft()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();

            await main.ApplyBetaListAsync(new ProductionListResponse(true,
            [
                new ProductionBoardRow("Commodore/C128/310378", "Commodore", "C128", "310378", null, "hash", null, null, CarriesDiscardedDraft: true),
                new ProductionBoardRow("Commodore/C64/250407", "Commodore", "C64", "250407", null, "hash", null, null)
            ]), background: false);

            Assert.Single(main.BetaTextsForTests(), text => text == DraftDiscardWording.ListMark);
        });
    }

    // ###########################################################################################
    // The owner's own advice - push it back and ask - in THE AMBER NOTICE PANEL under the board's
    // name (owner request, 2026-10-11: "it should show this as a normal highlight, which has the
    // yellow'ish panel and shown in same place as normal"). It was a red line among the contributors,
    // inside the panel that scrolls; the queue says the same thing in BeforeApprovingNotice, at the
    // top of its side, so BETA's side says it the same way and in the same place.
    // ###########################################################################################
    [Fact]
    public void The_beta_plan_says_who_discarded_their_draft_in_the_notice_panel_above_everything_else()
    {
        UiTest.Run(() =>
        {
            var view = new BetaView();
            var carried = new CarriedSubmission(41, "dennis@example.com", "Change 41", Discarded.AddDays(-1), DraftDiscardedUtc: Discarded);
            var kept = new CarriedSubmission(42, "anna@example.com", "Change 42", Discarded.AddDays(-1));

            view.ShowPlanForTests(Plan(carried, kept), BetaRow());

            var notice = view.FindControl<Border>("BeforePublishingNotice")!;
            List<string> warnings = view.FindControl<StackPanel>("BeforePublishingPanel")!.Children.OfType<TextBlock>().Select(TabMaintainer.TextOf).ToList();

            Assert.True(notice.IsVisible);
            Assert.Contains("Notice", notice.Classes);
            Assert.Equal([DraftDiscardWording.BetaWarning(carried)], warnings);
            Assert.All(view.FindControl<StackPanel>("BeforePublishingPanel")!.Children, line => Assert.Contains("NoticeText", line.Classes));
            Assert.Contains("pushing it back to the queue", DraftDiscardWording.BetaWarning(carried), StringComparison.Ordinal);

            // Which contribution, and its contributor in bold as on the carried line (code review,
            // 2026-10-11) - away from that line, the warning has to say it itself.
            TextBlock warning = view.FindControl<StackPanel>("BeforePublishingPanel")!.Children.OfType<TextBlock>().Single();
            Assert.StartsWith("#41 (Change 41) - ", TabMaintainer.TextOf(warning), StringComparison.Ordinal);
            Assert.Equal(["dennis@example.com"], Handlers.Theming.BoldRuns.BoldPiecesOf(warning));

            // Said once: no longer a red line among the contributors as well.
            List<string> carrying = view.FindControl<StackPanel>("CarryingPanel")!.Children.OfType<TextBlock>().Select(TabMaintainer.TextOf).ToList();
            Assert.DoesNotContain(DraftDiscardWording.BetaWarning(carried), carrying);

            // Directly under the board's name - before who must approve, and outside the panel
            // that scrolls, so it is on screen whatever the plan holds.
            List<Control> top = ((DockPanel)view.Content!).Children.ToList();
            Assert.Equal(top.IndexOf(view.FindControl<TextBlock>("SelectedBoardText")!) + 1, top.IndexOf(notice));
        });
    }

    // Nothing to settle: no panel - and none carried over to the next board, or to an empty view.
    [Fact]
    public void The_beta_notice_panel_is_gone_when_nothing_carried_was_discarded()
    {
        UiTest.Run(() =>
        {
            var view = new BetaView();
            var notice = view.FindControl<Border>("BeforePublishingNotice")!;

            view.ShowPlanForTests(Plan(new CarriedSubmission(41, "dennis@example.com", "Change 41", Discarded, DraftDiscardedUtc: Discarded)), BetaRow());
            Assert.True(notice.IsVisible);

            view.ShowPlanForTests(Plan(new CarriedSubmission(42, "anna@example.com", "Change 42", Discarded)), BetaRow());
            Assert.False(notice.IsVisible);
            Assert.Empty(view.FindControl<StackPanel>("BeforePublishingPanel")!.Children);

            view.ShowPlanForTests(Plan(new CarriedSubmission(41, "dennis@example.com", "Change 41", Discarded, DraftDiscardedUtc: Discarded)), BetaRow());
            view.ShowPlanForTests(null, null);
            Assert.False(notice.IsVisible);
        });
    }

    private static ProductionPlanView Plan(params CarriedSubmission[] carrying) =>
        new(
            "Commodore/C128/310378",
            "2026-September-27",
            new string('c', 64),
            TouchesSharedFiles: false,
            CanPublish: true,
            Refusal: null,
            UnchangedCount: 10,
            Files: [],
            Problems: [],
            Approval: null,
            Removals: null,
            Carrying: carrying);

    private static ProductionBoardRow BetaRow() =>
        new("Commodore/C128/310378", "Commodore", "C128", "310378", null, "hash", null, null);

    [Fact]
    public void The_boards_screen_marks_the_submission_whose_draft_was_discarded()
    {
        UiTest.Run(() =>
        {
            var view = new BoardDetailView();

            view.ShowDetailForTests(new BoardDetailAnswer(
                new BoardOverviewEntry("Commodore/C128/310378", "Commodore", "C128", "310378", true, false, true, true, null, null, null, 1),
                [],
                [],
                [
                    new BoardSubmissionEntry(41, "dennis@example.com", "Change 41", "merged", Discarded.AddDays(-2), Discarded.AddDays(-1), null, DraftDiscardedUtc: Discarded),
                    new BoardSubmissionEntry(40, "dennis@example.com", "Change 40", "published", Discarded.AddDays(-9), Discarded.AddDays(-8), null)
                ]));

            Assert.Single(view.TextsForTests(), text => text == DraftDiscardWording.Mark(Discarded));
        });
    }
}
