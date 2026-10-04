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
// plan, and the Systems screen's submission list - each drawn here from an answer carrying it, and
// checked to say nothing where it is absent. No server: every screen is handed its answer.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class DraftDiscardShownTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static readonly DateTimeOffset Discarded = new(2026, 9, 28, 9, 15, 0, TimeSpan.Zero);

    private static ReviewQueueRow Row(long id, DateTimeOffset? discarded) =>
        new(id, "Commodore/C128/310378", "pending", $"Change {id}", "dennis@example.com",
            DateTimeOffset.UtcNow.AddHours(-1), IsNewSystem: false, AwaitsYou: true, DraftDiscardedUtc: discarded);

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
    public async Task The_beta_list_marks_a_system_carrying_a_discarded_draft()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();

            await main.ApplyBetaListAsync(new ProductionListResponse(true,
            [
                new ProductionSystemRow("Commodore/C128/310378", "Commodore", "C128", "310378", null, "hash", null, null, CarriesDiscardedDraft: true),
                new ProductionSystemRow("Commodore/C64/250407", "Commodore", "C64", "250407", null, "hash", null, null)
            ]), background: false);

            Assert.Single(main.BetaTextsForTests(), text => text == DraftDiscardWording.ListMark);
        });
    }

    // The owner's own advice, under the contributor's line: push it back and ask them.
    [Fact]
    public void The_beta_plan_says_who_discarded_their_draft_and_to_consider_pushing_it_back()
    {
        UiTest.Run(() =>
        {
            var view = new BetaView();
            var carried = new CarriedSubmission(41, "dennis@example.com", "Change 41", Discarded.AddDays(-1), DraftDiscardedUtc: Discarded);

            view.ShowPlanForTests(
                new ProductionPlanView(
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
                    Carrying: [carried]),
                new ProductionSystemRow("Commodore/C128/310378", "Commodore", "C128", "310378", null, "hash", null, null));

            List<string> lines = view.FindControl<StackPanel>("CarryingPanel")!.Children.OfType<TextBlock>().Select(block => block.Text ?? string.Empty).ToList();

            Assert.Contains(DraftDiscardWording.BetaWarning(carried), lines);
            Assert.Contains("pushing it back to the queue", DraftDiscardWording.BetaWarning(carried), StringComparison.Ordinal);
        });
    }

    [Fact]
    public void The_systems_screen_marks_the_submission_whose_draft_was_discarded()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(
                new SystemOverviewEntry("Commodore/C128/310378", "Commodore", "C128", "310378", true, false, true, true, null, null, null, 1),
                [],
                [],
                [
                    new SystemSubmissionEntry(41, "dennis@example.com", "Change 41", "merged", Discarded.AddDays(-2), Discarded.AddDays(-1), null, DraftDiscardedUtc: Discarded),
                    new SystemSubmissionEntry(40, "dennis@example.com", "Change 40", "published", Discarded.AddDays(-9), Discarded.AddDays(-8), null)
                ]));

            Assert.Single(view.TextsForTests(), text => text == DraftDiscardWording.Mark(Discarded));
        });
    }
}
