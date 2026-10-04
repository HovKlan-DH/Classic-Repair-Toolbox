using Avalonia.Controls;
using Avalonia.LogicalTree;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The "Systems" screen's right-hand panel, SystemView (owner request, 2026-09-27: "Right-side
// will show data per system - e.g. who has contributed to it and who is set as maintainer etc.").
//
// Drawn from an answer without a server; with no client the panel asks nobody. The words are
// SystemsDisplay's and tested there - these pin what the panel puts where, and that every empty
// section SAYS it is empty.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SystemViewTests
{
    private static readonly DateTimeOffset Noon = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

    private static SystemOverviewEntry System() =>
        new("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, true, true, "2026-September-25", null, null, 1);

    [Fact]
    public void With_nothing_chosen_the_panel_says_what_to_do()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            Assert.Equal(["Select a system"], view.TextsForTests());
        });
    }

    // ###########################################################################################
    // Every view's words, in the order TextsForTests reads them: which system and what is off about
    // it, who maintains it, its statistics (none counted here, and SAID so), who has contributed and
    // how it went, then its history - a card per submission, with what the contributor was told.
    // ###########################################################################################
    [Fact]
    public void A_system_reads_as_its_state_then_maintainers_contributors_and_history()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(
                SystemViewTests.System(),
                [new PoolMaintainerEntry(7, "Anna", "anna@example.com")],
                [new SystemContributorEntry("hest@mailscan.dk", null, 2, 0, 0, 0, SystemViewTests.Noon)],
                [new SystemSubmissionEntry(41, "hest@mailscan.dk", "Corrected U8.", "returned", SystemViewTests.Noon, null, "U7 is the wrong revision.")]));

            Assert.Equal(
                [
                    "Commodore / C64 / 250407",
                    "BETA ahead of the stable source - 1 maintainer",
                    "Maintainer",
                    "Anna (anna@example.com)",
                    SystemsDisplay.NoViewCountsLine,
                    "Contributor",
                    "hest@mailscan.dk",
                    "2 accepted - last sent 2026-September-25",
                    "History - 1 submission, newest first",
                    "September 2026",
                    "25 Sep",
                    "#41 - Corrected U8.",
                    "Taken back out of BETA - waiting for review again",
                    "2026-September-25 - sent by hest@mailscan.dk",
                    "Told the contributor: U7 is the wrong revision."
                ],
                view.TextsForTests());

            // The revisions are the stage line's since 2026-10-04 - with the newest submission, and
            // where BETA waits.
            Assert.Equal(
                [
                    $"Submitted: #41 Taken back out of BETA - waiting for review again - sent {SubmissionReceiptPresenter.FormatDate(SystemViewTests.Noon)}",
                    $"BETA: Revision 2026-September-25 - ahead of stable - waiting under {MaintainerScreenWording.BetaQueueQuoted}",
                    "Stable: In the stable source"
                ],
                view.StagesForTests());

            Assert.Equal("Commodore/C64/250407", view.ShownSystem!.SystemId);
        });
    }

    // ###########################################################################################
    // Board views (owner request, 2026-09-27) are the STATISTICS view's since 2026-10-03 - and no
    // longer on the system's status line ("remove the 'no views in 30 days' from the data").
    // ###########################################################################################
    [Fact]
    public void A_systems_board_views_are_its_Statistics_view_and_not_on_its_status_line()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(
                SystemViewTests.System() with { ViewsLast30Days = 48 },
                [new PoolMaintainerEntry(7, "Anna", "anna@example.com")],
                [],
                [],
                Views: new BoardViewStatistics(12, 48, 310, 3, [new("DK", "Denmark", 120)])));

            Assert.Equal(
                [
                    "Views in CRT",
                    "[12] in the last 7 days - [48] in 30 days - [310] in 12 months",
                    "Most views in the last 12 months: Denmark (120).",
                    "Not counted above: [3] from CRTs downloading BETA data in the last 30 days.",
                    "A view is this board on screen in CRT for at least 10 seconds, counted every time."
                ],
                view.SectionTextsForTests("ViewsSection"));

            Assert.Contains("BETA ahead of the stable source - 1 maintainer", view.TextsForTests());
            Assert.DoesNotContain(view.TextsForTests(), text => text.Contains("views in 30 days", StringComparison.Ordinal));
        });
    }

    // No counts from the server: the view says so - never "no views", which would be a claim about
    // the board the server never made.
    [Fact]
    public void Without_counts_from_the_server_the_Statistics_view_says_none_were_sent()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(SystemViewTests.System(), [], [], []));

            Assert.Equal([SystemsDisplay.NoViewCountsLine], view.SectionTextsForTests("ViewsSection"));
            Assert.DoesNotContain(view.TextsForTests(), text => text.StartsWith("No views", StringComparison.Ordinal));
        });
    }

    // Nobody assigned, nobody contributed, nothing submitted: each section says so rather than
    // leaving a blank that reads as not loaded.
    [Fact]
    public void Every_empty_section_says_it_is_empty()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(SystemViewTests.System() with { MaintainerCount = 0 }, [], [], []));

            IReadOnlyList<string> texts = view.TextsForTests();

            Assert.Contains("Nobody maintains this system - its submissions go to the administrator.", texts);
            Assert.Contains("Nobody has contributed to this system through CRT yet.", texts);
            Assert.Contains("Nothing has happened to this system yet", texts);
        });
    }

    // Signed out: nothing of the previous account's system stays.
    [Fact]
    public void Clearing_the_panel_leaves_nothing_of_the_system()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(SystemViewTests.System(), [], [], []));
            view.Clear();

            Assert.Null(view.ShownSystem);
            Assert.Equal(["Select a system"], view.TextsForTests());
        });
    }

    // ###########################################################################################
    // A system not in CRT's drop-down lists gets the placement panel in its Maintainer view (owner
    // request, 2026-09-27), and a line above the views saying where it is; a listed one names the
    // entry CRT shows it under instead.
    // ###########################################################################################
    private static SystemListingAnswer Listing() =>
        new(
            true,
            [new SystemListingRow("Commodore/C64/250407", "Commodore 64", "250407 (long board)", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx")],
            [
                new UnlistedSystemEntry(
                    "Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", false, true, null,
                    new SystemPlacement("Commodore 128", "310378 Open128", string.Empty, null)),
            ]);

    private static SystemDetailAnswer Detail(SystemOverviewEntry system) => new(system, [], [], []);

    [Fact]
    public void A_system_not_in_the_lists_gets_the_placement_panel_and_a_listed_one_names_its_entry()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();
            view.UseListing(SystemViewTests.Listing());

            view.ShowDetailForTests(SystemViewTests.Detail(
                new("Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", false, null, false, true, null, null, null, 0)));

            Assert.True(view.PlacementForTests.IsVisible);
            Assert.Equal("Commodore/C128/310378 Open128", view.PlacementForTests.ShownSystemId);
            Assert.DoesNotContain(view.TextsForTests(), text => text.StartsWith("In CRT's drop-down lists", StringComparison.Ordinal));
            Assert.Contains("Needs a place in the drop-down lists - give it one under Maintainer", view.TextsForTests());

            view.ShowDetailForTests(SystemViewTests.Detail(SystemViewTests.System()));

            Assert.False(view.PlacementForTests.IsVisible);
            Assert.Contains("In CRT's drop-down lists as Commodore 64 / 250407 (long board)", view.TextsForTests());
            Assert.DoesNotContain(view.TextsForTests(), text => text.StartsWith("Needs a place", StringComparison.Ordinal));
        });
    }

    // Signed out: the listing goes with the session - nothing of it is left for the next account.
    [Fact]
    public void Clearing_the_panel_forgets_the_listing()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();
            view.UseListing(SystemViewTests.Listing());

            view.Clear();
            view.ShowDetailForTests(SystemViewTests.Detail(
                new("Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", false, null, false, true, null, null, null, 0)));

            Assert.False(view.PlacementForTests.IsVisible);
        });
    }

    // The chosen system's own status line carries the same bold as its row in the list.
    [Fact]
    public void The_chosen_systems_status_line_has_what_is_still_to_be_done_in_bold()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(SystemViewTests.Detail(
                new("Commodore/C128/310378 Open128", "Commodore", "C128", "310378 Open128", true, false, true, true, null, null, null, 0)));

            TextBlock state = view.FindControl<TextBlock>("SystemStateText")!;
            List<Avalonia.Controls.Documents.Run> runs = state.Inlines!.OfType<Avalonia.Controls.Documents.Run>().ToList();

            Assert.Equal(
                ["not in the stable source yet", "nobody assigned"],
                runs.Where(run => run.FontWeight == Avalonia.Media.FontWeight.Bold).Select(run => run.Text));
            Assert.Contains("In BETA - not in the stable source yet - nobody assigned", view.TextsForTests());
        });
    }

    // ###########################################################################################
    // THE HISTORY VIEW (owner request, 2026-10-04: "the full history of what has happened with this
    // board, in an 'easy to overview' way"): by month, newest first, the day in its own column, one
    // card per submission - its steps oldest first, where it stands now, and WHAT IT CHANGED - and a
    // line for anything else. The grouping and words are SystemHistoryDisplay's, tested there; this
    // pins what the view puts where.
    // ###########################################################################################
    [Fact]
    public void The_history_view_shows_a_card_per_submission_with_what_it_changed_and_other_events_as_lines()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            var changes = new SubmissionChanges(
                false,
                [new SectionChanges("Components", 0, 1, 0, 0, [], [new ChangedRowFact("U8", ["Part-number"])], [], [])],
                FileChanges.None);

            view.ShowDetailForTests(new SystemDetailAnswer(
                SystemViewTests.System(),
                [],
                [],
                [new SystemSubmissionEntry(9, "hest@mailscan.dk", "Corrected U8.", "merged", SystemViewTests.Noon.AddDays(-1), SystemViewTests.Noon, null, Changes: changes)],
                null,
                [
                    new SystemHistoryEntry(SystemViewTests.Noon.AddDays(2), SystemHistoryEvents.PublishedToProduction, "Dennis", null, "revision 2026-September-27"),
                    new SystemHistoryEntry(SystemViewTests.Noon, SystemHistoryEvents.Decided, "Anna", 9, "merged"),
                    new SystemHistoryEntry(SystemViewTests.Noon.AddDays(-1), SystemHistoryEvents.Sent, "hest@mailscan.dk", 9, "Corrected U8."),
                ]));

            Assert.Equal(
                [
                    "History - 1 submission and 1 other event, newest first",
                    "September 2026",
                    "27 Sep",
                    "Published to the stable source",
                    "by Dennis - revision 2026-September-27",
                    "25 Sep",
                    "#9 - Corrected U8.",
                    "Published to the BETA source",
                    "2026-September-24 - sent by hest@mailscan.dk",
                    "2026-September-25 - Published to the BETA source by Anna",
                    SystemHistoryDisplay.ChangesHeading,
                    "Components",
                    "[1] changed: U8 (Part-number)"
                ],
                view.SectionTextsForTests("HistoryItemsSection"));

            // A submission is ONE card - its sending and its decision are not lines of their own.
            Assert.Single(view.FindControl<StackPanel>("HistoryItemsSection")!.GetLogicalDescendants().OfType<Border>(), border => border.Classes.Contains("HistoryCard"));
        });
    }

    // ###########################################################################################
    // The Maintainer view is the maintainers and nothing more (owner request, 2026-10-04) - adding,
    // inviting and removing them is Account > Maintainers (MaintainerPoolViewTests).
    // ###########################################################################################
    [Fact]
    public void The_maintainer_view_lists_every_maintainer_and_offers_no_controls()
    {
        UiTest.Run(() =>
        {
            var view = new SystemView();

            view.ShowDetailForTests(new SystemDetailAnswer(
                SystemViewTests.System(),
                [new PoolMaintainerEntry(7, "Anna", "anna@example.com"), new PoolMaintainerEntry(8, "Bo", "bo@example.com")],
                [],
                [],
                [new MaintainerInvitationEntry(3, "new@example.com", SystemViewTests.Noon, SystemViewTests.Noon.AddDays(14))]));

            Assert.Equal(
                ["Maintainers (2)", "Anna (anna@example.com)", "Bo (bo@example.com)"],
                view.SectionTextsForTests("MaintainersSection"));

            Assert.DoesNotContain(view.TextsForTests(), text => text.StartsWith("Invited", StringComparison.Ordinal));
            Assert.Null(view.FindControl<ComboBox>("AddAccountCombo"));
            Assert.Null(view.FindControl<Button>("InviteButton"));
        });
    }
}
