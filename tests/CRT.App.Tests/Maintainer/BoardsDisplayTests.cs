using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers BoardsDisplay - how the "Boards" screen reads (owner request, 2026-09-27).
//
// The one rule that is more than formatting: a submission's state is said in CRT's OWN words
// (SubmissionReceiptPresenter.DescribeState), so the maintainer and the contributor describe one
// submission identically. The test asserts the words themselves, so a change to CRT's wording that
// leaves this screen behind - or the reverse - fails here.
// ###########################################################################################
public sealed class BoardsDisplayTests
{
    // Midday UTC, so FormatDate's local-time conversion cannot move the date on any machine.
    private static readonly DateTimeOffset Noon = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

    private static BoardOverviewEntry Board(
        bool inBeta = true,
        bool? inProduction = true,
        bool awaiting = false,
        bool accepting = true,
        string? beta = null,
        string? production = null,
        DateTimeOffset? published = null,
        int maintainers = 1) =>
        new("Commodore/C64/250407", "Commodore", "C64", "250407", inBeta, inProduction, awaiting, accepting, beta, production, published, maintainers);

    // ###########################################################################################
    // *** EVERY LIST OF BOARDS IN THE BOARDS SCREEN'S ORDER (owner request, 2026-10-09: "the list of
    // boards should be identical to the order in the 'Systems' list"). *** Account's Maintainers and
    // "Delete a board" follow the Boards screen top to bottom; a board that list does not hold comes
    // after, by name - as does every board before the list has been read.
    // ###########################################################################################
    [Fact]
    public void A_list_of_boards_follows_the_Boards_screen_and_puts_what_it_lacks_last_by_name()
    {
        string[] boards = ["Commodore/C64/250407", "Amstrad/CPC 664/MC0005A", "Commodore/VIC-20/250403", "Zx/Spectrum/4B", "Acme/X/1"];
        string[] boardsScreen = ["Commodore/VIC-20/250403", "commodore/c64/250407", "Amstrad/CPC 664/MC0005A"];

        IReadOnlyList<string> ordered = BoardsDisplay.InBoardsListOrder(boards, id => id, id => id, boardsScreen);

        Assert.Equal(["Commodore/VIC-20/250403", "Commodore/C64/250407", "Amstrad/CPC 664/MC0005A", "Acme/X/1", "Zx/Spectrum/4B"], ordered);
    }

    [Fact]
    public void Before_the_Boards_screen_has_been_read_a_list_of_boards_is_by_name()
    {
        string[] boards = ["Commodore/C64/250407", "Amstrad/CPC 664/MC0005A"];

        Assert.Equal(["Amstrad/CPC 664/MC0005A", "Commodore/C64/250407"], BoardsDisplay.InBoardsListOrder(boards, id => id, id => id, null));
        Assert.Equal(["Amstrad/CPC 664/MC0005A", "Commodore/C64/250407"], BoardsDisplay.InBoardsListOrder(boards, id => id, id => id, []));
    }

    // ###########################################################################################
    // *** HOW A BOARD'S SUBMISSIONS WENT (owner request, 2026-10-09: "19 submissions in total; 2
    // rejected, 1 in BETA, 17 in stable"). *** The three always, the waiting ones only when there
    // are any; nothing from a server that does not count them.
    // ###########################################################################################
    [Theory]
    [InlineData(19, 0, 2, 1, 16, "19 submissions in total; 2 rejected, 1 in BETA, 16 in stable")]
    [InlineData(20, 1, 2, 1, 16, "20 submissions in total; 1 waiting, 2 rejected, 1 in BETA, 16 in stable")]
    [InlineData(1, 0, 0, 1, 0, "1 submission in total; 0 rejected, 1 in BETA, 0 in stable")]
    [InlineData(0, 0, 0, 0, 0, "No submissions yet")]
    public void A_boards_submissions_are_said_in_one_line(int total, int waiting, int rejected, int inBeta, int inStable, string expected)
    {
        Assert.Equal(expected, BoardsDisplay.SubmissionCountsLine(new BoardSubmissionCounts(total, waiting, rejected, inBeta, inStable)));
    }

    [Fact]
    public void A_server_that_does_not_count_submissions_gets_no_line()
    {
        Assert.Null(BoardsDisplay.SubmissionCountsLine(null));
    }

    // ###########################################################################################
    // *** WHO A POOL CHANGE IS ABOUT, AND WHO MADE IT, IN BOLD (owner report, 2026-10-09:
    // "dennis@... made a maintainer / by dennis@... - this is I do not understand?"). *** What
    // happened first, then the person it happened to; the grey line says who did it.
    // ###########################################################################################
    [Fact]
    public void A_maintainer_added_names_who_was_added_and_who_added_them_in_bold()
    {
        var entry = new BoardHistoryEntry(Noon, BoardHistoryEvents.MaintainerAdded, "Dennis", null, "Anna");

        Assert.Equal(
            [("Maintainer added: ", false), ("Anna", true)],
            BoardsDisplay.HistoryWhatRuns(entry).Select(run => (run.Text, run.IsCount)));
        Assert.Equal(
            [("by ", false), ("Dennis", true)],
            BoardsDisplay.HistoryFooterRuns(entry).Select(run => (run.Text, run.IsCount)));
    }

    // A contributor and a maintainer on their views: the name bold, the address beside it plain.
    [Fact]
    public void A_contributor_and_a_maintainer_are_named_in_bold()
    {
        var contributor = new BoardContributorEntry("dora@example.com", "Dora", 1, 0, 0, 0, Noon);
        var maintainer = new PoolMaintainerEntry(7, "", "bo@example.com");

        Assert.Equal(
            [("Dora", true), (" (dora@example.com)", false)],
            BoardsDisplay.ContributorRuns(contributor).Select(run => (run.Text, run.IsCount)));
        Assert.Equal([("bo@example.com", true)], BoardsDisplay.MaintainerRuns(maintainer).Select(run => (run.Text, run.IsCount)));
    }

    [Fact]
    public void A_board_is_named_by_its_three_parts()
    {
        Assert.Equal("Commodore / C64 / 250407", BoardsDisplay.Name(BoardsDisplayTests.Board()));
        Assert.Equal("Acme/X/1", BoardsDisplay.Name(new BoardOverviewEntry("Acme/X/1", "", "", "", true, null, false, true, null, null, null, 0)));
    }

    // ###########################################################################################
    // WHERE ITS DATA IS, one phrase per case - and NOTHING for the ordinary case, in BETA and the
    // stable source (owner request, 2026-10-03: "only show where there is something odd/off").
    // ###########################################################################################
    [Theory]
    [InlineData(true, true, false, "")]
    [InlineData(true, true, true, "BETA ahead of the stable source")]
    [InlineData(true, false, true, "In BETA - not in the stable source yet")]
    [InlineData(true, false, false, "In BETA - not in the stable source yet")]
    [InlineData(false, false, false, "Not published yet")]
    [InlineData(false, true, false, "In the stable source - not in BETA")]
    public void Where_a_boards_data_is_is_said_in_one_phrase(bool inBeta, bool inProduction, bool awaiting, string expected)
    {
        Assert.Equal(expected, BoardsDisplay.Where(BoardsDisplayTests.Board(inBeta, inProduction, awaiting)));
    }

    // Without a stable tree to look in, a published board is the ordinary case too: nothing said,
    // and nothing claimed about the stable source.
    [Fact]
    public void Without_a_production_tree_nothing_is_claimed_about_production()
    {
        Assert.Equal(string.Empty, BoardsDisplay.Where(BoardsDisplayTests.Board(inProduction: null)));
        Assert.Equal("In BETA - not in the stable source yet", BoardsDisplay.Where(BoardsDisplayTests.Board(inProduction: null, awaiting: true)));
    }

    // The list's grey line: where, only when it is off; how many maintain it; closed only when it is.
    [Fact]
    public void A_boards_list_line_says_where_only_when_it_is_off_who_and_whether_it_is_closed()
    {
        Assert.Equal("2 maintainers", BoardsDisplay.ListLine(BoardsDisplayTests.Board(maintainers: 2)));
        Assert.Equal("1 maintainer - closed to contributions", BoardsDisplay.ListLine(BoardsDisplayTests.Board(accepting: false)));
        Assert.Equal(
            "Not published yet - nobody assigned - closed to contributions",
            BoardsDisplay.ListLine(BoardsDisplayTests.Board(inBeta: false, inProduction: false, accepting: false, maintainers: 0)));
    }

    // Nobody assigned is said, with where its submissions go - it is what this screen is opened to find.
    [Fact]
    public void A_board_nobody_maintains_says_where_its_submissions_go()
    {
        Assert.Equal("Nobody maintains this board - its submissions go to the administrator.", BoardsDisplay.MaintainersHeading(0));
        Assert.Equal("Maintainer", BoardsDisplay.MaintainersHeading(1));
        Assert.Equal("Maintainers (3)", BoardsDisplay.MaintainersHeading(3));
        Assert.Equal("Anna (anna@example.com)", BoardsDisplay.MaintainerLine(new PoolMaintainerEntry(7, "Anna", "anna@example.com")));
    }

    // ###########################################################################################
    // A contributor's record: only the counts that are not zero, then when they last sent one. A
    // contributor with no account is their address alone.
    // ###########################################################################################
    [Fact]
    public void A_contributors_record_names_only_what_happened()
    {
        var contributor = new BoardContributorEntry("hest@mailscan.dk", null, Accepted: 3, Waiting: 0, ChangesRequested: 1, Rejected: 0, BoardsDisplayTests.Noon);

        Assert.Equal("hest@mailscan.dk", BoardsDisplay.ContributorName(contributor));
        Assert.Equal("3 accepted, 1 sent back for changes - last sent 2026-September-25", BoardsDisplay.ContributorRecord(contributor));

        var named = new BoardContributorEntry("anna@example.com", "Anna", 0, 1, 0, 2, null);

        Assert.Equal("Anna (anna@example.com)", BoardsDisplay.ContributorName(named));
        Assert.Equal("1 waiting, 2 rejected", BoardsDisplay.ContributorRecord(named));

        Assert.Equal(string.Empty, BoardsDisplay.ContributorRecord(new BoardContributorEntry("x@example.com", null, 0, 0, 0, 0, null)));

        // A board the account does not maintain comes without addresses (2026-10-05): a contributor
        // with no account is said as one - "(no address)" would claim they gave none - and a named
        // one, like a maintainer, is the name alone.
        var anonymous = new BoardContributorEntry(null, null, 1, 0, 0, 0, null);

        Assert.Equal("A contributor without an account", BoardsDisplay.ContributorName(anonymous, addressesHidden: true));
        Assert.Equal("(no address)", BoardsDisplay.ContributorName(anonymous));
        Assert.Equal("Anna", BoardsDisplay.ContributorName(named with { Email = null }, addressesHidden: true));
        Assert.Equal("Anna", BoardsDisplay.MaintainerLine(new PoolMaintainerEntry(7, "Anna", string.Empty)));
    }

    [Fact]
    public void The_contributor_heading_says_when_there_are_none()
    {
        Assert.Equal("Nobody has contributed to this board through CRT yet.", BoardsDisplay.ContributorsHeading(0));
        Assert.Equal("Contributors (2)", BoardsDisplay.ContributorsHeading(2));
    }

    // ###########################################################################################
    // *** THE STATE IN CRT's OWN WORDS. *** "merged" is "Published to the BETA source" and "returned"
    // is "Taken back out of BETA - waiting for review again" in CRT's "My submissions"; the
    // maintainer reads exactly the same. These strings are asserted, not re-derived, so either
    // side changing alone fails.
    // ###########################################################################################
    [Theory]
    [InlineData("merged", "Published to the BETA source")]
    [InlineData("published", "Published to the stable source")]
    [InlineData("returned", "Taken back out of BETA - waiting for review again")]
    [InlineData("pending", "Submitted - awaiting feedback from a maintainer")]
    [InlineData("withdrawn", "Replaced by a newer submission")]
    public void A_submission_is_described_in_the_words_CRT_uses(string state, string words)
    {
        var submission = new BoardSubmissionEntry(41, "hest@mailscan.dk", "Corrected U8.", state, BoardsDisplayTests.Noon, null, null);

        // The History view's card states it where it stands now (BoardHistoryDisplay, 2026-10-04).
        Assert.Equal(words, BoardHistoryDisplay.Card(submission, []).State);
    }

    // What the contributor was told, labelled by who read it; nothing when nothing was said.
    [Fact]
    public void What_the_contributor_was_told_is_shown_and_a_missing_description_says_so()
    {
        var told = new BoardSubmissionEntry(38, null, "  ", "returned", BoardsDisplayTests.Noon, null, " U7 is the wrong revision. ");

        Assert.Equal("Told the contributor: U7 is the wrong revision.", BoardsDisplay.SubmissionComment(told));
        Assert.Equal("(no description given)", BoardsDisplay.SubmissionTitle(told));

        Assert.Null(BoardsDisplay.SubmissionComment(told with { DecisionComment = null }));
    }

    // ###########################################################################################
    // *** WHAT IS STILL TO BE DONE IS MARKED, AND SHOWN IN BOLD (owner request, 2026-09-27: "'nobody
    // assigned' is important, as that should be done - the same for 'not in production yet'"). ***
    // The pieces join to exactly the line as it always read.
    // ###########################################################################################
    [Fact]
    public void Nobody_assigned_and_not_in_production_yet_are_marked_as_still_to_be_done()
    {
        BoardOverviewEntry board = BoardsDisplayTests.Board(inProduction: false, awaiting: true, maintainers: 0);

        IReadOnlyList<StatusPart> parts = BoardsDisplay.ListLineParts(board);

        Assert.Equal(["not in the stable source yet", "nobody assigned"], parts.Where(part => part.IsToDo).Select(part => part.Text));
        Assert.Equal("In BETA - not in the stable source yet - nobody assigned", BoardsDisplay.ListLine(board));
        Assert.Equal(BoardsDisplay.ListLine(board), string.Concat(parts.Select(part => part.Text)));
    }

    [Theory]
    [InlineData(true, true, true, "BETA ahead of the stable source")]      // waits for "Publish to production"
    [InlineData(false, true, false, "not in BETA")]                 // production holds what BETA lacks
    public void Other_outstanding_states_are_marked_too(bool inBeta, bool inProduction, bool awaiting, string marked)
    {
        IReadOnlyList<StatusPart> parts = BoardsDisplay.ListLineParts(
            BoardsDisplayTests.Board(inBeta: inBeta, inProduction: inProduction, awaiting: awaiting));

        Assert.Equal([marked], parts.Where(part => part.IsToDo).Select(part => part.Text));
    }

    // Nothing to do, nothing marked - including a turned-down new board ("Not published yet") and a
    // board closed to contributions on purpose.
    [Fact]
    public void A_board_with_nothing_outstanding_has_nothing_marked()
    {
        Assert.DoesNotContain(BoardsDisplay.ListLineParts(BoardsDisplayTests.Board(maintainers: 2)), part => part.IsToDo);
        Assert.DoesNotContain(
            BoardsDisplay.ListLineParts(BoardsDisplayTests.Board(inBeta: false, inProduction: false, accepting: false, maintainers: 1)),
            part => part.IsToDo);
    }

    // ###########################################################################################
    // *** WHAT IS OFF ABOUT THE DROP-DOWN LISTS IS FLAGGED (owner request, 2026-10-04: "I do not
    // expect there should be cases where something can only be listed in stable? If so, it must be
    // flagged in the left-sided menu 'Systems' list"). *** Marked as still to be done, so the line
    // shows it in bold - and the mirror case too: a board the stable source holds that its list
    // leaves out, which stable CRT users cannot reach.
    // ###########################################################################################
    [Fact]
    public void A_board_listed_only_in_the_stable_list_is_flagged_as_to_do()
    {
        BoardOverviewEntry board = BoardsDisplayTests.Board() with { ListedInBeta = false, ListedInStable = true };

        Assert.Equal("listed only in the stable source's drop-down list - 1 maintainer", BoardsDisplay.ListLine(board));
        Assert.Equal(["listed only in the stable source's drop-down list"], BoardsDisplay.ListLineParts(board).Where(part => part.IsToDo).Select(part => part.Text));
    }

    [Fact]
    public void A_board_in_the_stable_source_that_its_list_leaves_out_is_flagged_as_to_do()
    {
        BoardOverviewEntry board = BoardsDisplayTests.Board(awaiting: true) with { ListedInBeta = true, ListedInStable = false };

        Assert.Equal("BETA ahead of the stable source - missing from the stable source's drop-down list - 1 maintainer", BoardsDisplay.ListLine(board));
        Assert.Equal(
            ["BETA ahead of the stable source", "missing from the stable source's drop-down list"],
            BoardsDisplay.ListLineParts(board).Where(part => part.IsToDo).Select(part => part.Text));
    }

    // ###########################################################################################
    // Nothing is flagged that is as it should be or not known: both lists naming it; a NEW board in
    // BETA only, not promoted yet, which is listed in BETA alone by design; a list that could not be
    // read; and a board BETA does not hold at all, whose "not in BETA" already says it.
    // ###########################################################################################
    [Fact]
    public void Nothing_about_the_lists_is_said_when_they_are_as_they_should_be_or_unknown()
    {
        Assert.Equal("1 maintainer", BoardsDisplay.ListLine(BoardsDisplayTests.Board() with { ListedInBeta = true, ListedInStable = true }));
        Assert.Equal(
            "In BETA - not in the stable source yet - 1 maintainer",
            BoardsDisplay.ListLine(BoardsDisplayTests.Board(inProduction: false, awaiting: true) with { ListedInBeta = true, ListedInStable = false }));
        Assert.Equal("1 maintainer", BoardsDisplay.ListLine(BoardsDisplayTests.Board() with { ListedInBeta = null, ListedInStable = true }));
        Assert.Equal("1 maintainer", BoardsDisplay.ListLine(BoardsDisplayTests.Board() with { ListedInBeta = true, ListedInStable = null }));
        Assert.Equal(
            "In the stable source - not in BETA - 1 maintainer",
            BoardsDisplay.ListLine(BoardsDisplayTests.Board(inBeta: false) with { ListedInBeta = false, ListedInStable = true }));
    }

    // An invitation nobody has accepted, on the administrator's screen: who, and how long its code works.
    [Fact]
    public void An_open_invitation_says_who_when_and_until_when()
    {
        var invitation = new MaintainerInvitationEntry(3, "new@example.com", BoardsDisplayTests.Noon, BoardsDisplayTests.Noon.AddDays(14));

        Assert.Equal("Invited: new@example.com", BoardsDisplay.InvitationLine(invitation));
        Assert.Equal(
            "Sent 2026-September-25 - the code works until 2026-October-9, and they become a maintainer when they use it",
            BoardsDisplay.InvitationFooter(invitation));
    }

    // ###########################################################################################
    // THE BOARD'S HISTORY (owner request, 2026-09-27: "I would like to see the date, newest first,
    // to understand what has happened to a system"): what happened, and a grey line of by whom and
    // what more - the date is the History view's own column since 2026-10-04. Decisions in CRT's
    // own state words.
    // ###########################################################################################
    [Theory]
    [InlineData(BoardHistoryEvents.Sent, 9L, "Corrected U8.", "hest@mailscan.dk", "#9 sent", "from hest@mailscan.dk - Corrected U8.")]
    [InlineData(BoardHistoryEvents.Decided, 9L, "merged", "Anna", "#9 - Published to the BETA source", "by Anna")]
    [InlineData(BoardHistoryEvents.Decided, 9L, "changes_requested", null, "#9 - Changes requested", "")]
    [InlineData(BoardHistoryEvents.PublishedToProduction, null, "revision 2026-September-25; 3 file(s) copied", "admin@example.com", "Published to the stable source", "by admin@example.com - revision 2026-September-25; 3 file(s) copied")]
    [InlineData(BoardHistoryEvents.PushedBack, null, "1 submission(s) returned to the queue; Not ready.", "admin@example.com", "Pushed back from BETA to the queue", "by admin@example.com - 1 submission(s) returned to the queue; Not ready.")]
    [InlineData(BoardHistoryEvents.RejectedFromBeta, null, "1 submission(s) rejected; Not for this board.", "admin@example.com", "Rejected in BETA and taken out of it", "by admin@example.com - 1 submission(s) rejected; Not for this board.")]
    [InlineData(BoardHistoryEvents.MaintainerAdded, null, "anna@example.com", "admin@example.com", "Maintainer added: anna@example.com", "by admin@example.com")]
    [InlineData(BoardHistoryEvents.MaintainerRemoved, null, "anna@example.com", "admin@example.com", "Maintainer removed: anna@example.com", "by admin@example.com")]
    [InlineData(BoardHistoryEvents.Invited, null, "new@example.com", "admin@example.com", "Invited to be a maintainer: new@example.com", "by admin@example.com")]
    [InlineData(BoardHistoryEvents.InvitationAccepted, null, "new@example.com", "new@example.com", "new@example.com accepted the invitation and became a maintainer", "")]
    [InlineData(BoardHistoryEvents.Placed, null, "Commodore 64 / 250407", "anna@example.com", "Placed in the drop-down lists", "by anna@example.com - Commodore 64 / 250407")]
    [InlineData(BoardHistoryEvents.Amended, 9L, "amendment 2, 3 file(s)", "anna@example.com", "#9 changed by a maintainer", "by anna@example.com - amendment 2, 3 file(s)")]
    [InlineData(BoardHistoryEvents.Deleted, null, "2 file(s) removed from BETA, 2 from the stable data; 3 submission(s) deleted, 0 of them open", "admin@example.com", "Deleted from the BETA and stable data and the database", "by admin@example.com - 2 file(s) removed from BETA, 2 from the stable data; 3 submission(s) deleted, 0 of them open")]
    [InlineData("something.new", null, null, null, "something.new", "")]
    public void A_history_entry_says_what_happened_and_by_whom(string kind, long? submission, string? detail, string? who, string line, string footer)
    {
        var entry = new BoardHistoryEntry(BoardsDisplayTests.Noon, kind, who, submission, detail);

        Assert.Equal(line, BoardsDisplay.HistoryWhat(entry));
        Assert.Equal(footer, BoardsDisplay.HistoryFooter(entry));
    }

    [Fact]
    public void A_decision_carries_what_the_contributor_was_told()
    {
        var entry = new BoardHistoryEntry(BoardsDisplayTests.Noon, BoardHistoryEvents.Decided, "Anna", 9, "rejected", "Wrong board.");

        Assert.Equal("Told the contributor: Wrong board.", BoardsDisplay.HistoryNote(entry));
        Assert.Null(BoardsDisplay.HistoryNote(entry with { Note = null }));
    }

    // -----------------------------------------------------------------------------------
    // Board views (owner request, 2026-09-27)
    // -----------------------------------------------------------------------------------

    // The list line says nothing about views any more (owner request, 2026-10-03: "remove the 'no
    // views in 30 days' from the data in the left-side list") - they are the Statistics view's.
    [Theory]
    [InlineData(null)]
    [InlineData(0)]
    [InlineData(48)]
    public void The_list_line_says_nothing_about_views(int? views)
    {
        BoardOverviewEntry board = BoardsDisplayTests.Board() with { ViewsLast30Days = views };

        Assert.Equal("1 maintainer", BoardsDisplay.ListLine(board));
    }

    // ###########################################################################################
    // The section: the three counts with their numbers bold, the countries, the BETA-source views
    // counted apart - each line only when there is something to say.
    // ###########################################################################################
    [Fact]
    public void The_views_section_says_the_counts_the_countries_and_the_beta_views_apart()
    {
        var views = new BoardViewStatistics(12, 48, 310, 3, [new("DK", "Denmark", 120), new("DE", "Germany", 60)]);

        Assert.Equal("Views in CRT", BoardsDisplay.ViewsHeading(views));

        IReadOnlyList<ReviewNoteRun> counts = BoardsDisplay.ViewCountRuns(views);
        Assert.Equal("[12] in the last 7 days - [48] in 30 days - [310] in 12 months", string.Concat(counts.Select(run => run.Text)));
        Assert.Equal(["12", "48", "310"], counts.Where(run => run.IsCount).Select(run => run.Text));

        Assert.Equal("Most views in the last 12 months: Denmark (120), Germany (60).", BoardsDisplay.ViewCountriesLine(views));
        Assert.Equal(
            "Not counted above: [3] from CRTs downloading BETA data in the last 30 days.",
            string.Concat(BoardsDisplay.ViewBetaRuns(views)!.Select(run => run.Text)));

        Assert.Equal("A view is this board on screen in CRT for at least 10 seconds, counted every time.", BoardsDisplay.ViewExplanation);
    }

    [Fact]
    public void A_board_nobody_viewed_says_so_in_its_heading_and_nothing_else()
    {
        var none = new BoardViewStatistics(0, 0, 0, 0, []);

        Assert.True(BoardsDisplay.HasNoViews(none));
        Assert.Equal("No views of this board in CRT counted in the last 12 months.", BoardsDisplay.ViewsHeading(none));
        Assert.Null(BoardsDisplay.ViewCountriesLine(none));
        Assert.Null(BoardsDisplay.ViewBetaRuns(none));

        // Only BETA-source views is still something to show.
        Assert.False(BoardsDisplay.HasNoViews(none with { FromBetaLast30Days = 2 }));
    }
}
