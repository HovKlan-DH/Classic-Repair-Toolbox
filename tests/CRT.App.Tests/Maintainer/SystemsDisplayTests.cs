using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers SystemsDisplay - how the "Systems" screen reads (owner request, 2026-09-27).
//
// The one rule that is more than formatting: a submission's state is said in CRT's OWN words
// (SubmissionReceiptPresenter.DescribeState), so the maintainer and the contributor describe one
// submission identically. The test asserts the words themselves, so a change to CRT's wording that
// leaves this screen behind - or the reverse - fails here.
// ###########################################################################################
public sealed class SystemsDisplayTests
{
    // Midday UTC, so FormatDate's local-time conversion cannot move the date on any machine.
    private static readonly DateTimeOffset Noon = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

    private static SystemOverviewEntry System(
        bool inBeta = true,
        bool? inProduction = true,
        bool awaiting = false,
        bool accepting = true,
        string? beta = null,
        string? production = null,
        DateTimeOffset? published = null,
        int maintainers = 1) =>
        new("Commodore/C64/250407", "Commodore", "C64", "250407", inBeta, inProduction, awaiting, accepting, beta, production, published, maintainers);

    [Fact]
    public void A_system_is_named_by_its_three_parts()
    {
        Assert.Equal("Commodore / C64 / 250407", SystemsDisplay.Name(SystemsDisplayTests.System()));
        Assert.Equal("Acme/X/1", SystemsDisplay.Name(new SystemOverviewEntry("Acme/X/1", "", "", "", true, null, false, true, null, null, null, 0)));
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
    public void Where_a_systems_data_is_is_said_in_one_phrase(bool inBeta, bool inProduction, bool awaiting, string expected)
    {
        Assert.Equal(expected, SystemsDisplay.Where(SystemsDisplayTests.System(inBeta, inProduction, awaiting)));
    }

    // Without a stable tree to look in, a published system is the ordinary case too: nothing said,
    // and nothing claimed about the stable source.
    [Fact]
    public void Without_a_production_tree_nothing_is_claimed_about_production()
    {
        Assert.Equal(string.Empty, SystemsDisplay.Where(SystemsDisplayTests.System(inProduction: null)));
        Assert.Equal("In BETA - not in the stable source yet", SystemsDisplay.Where(SystemsDisplayTests.System(inProduction: null, awaiting: true)));
    }

    // The list's grey line: where, only when it is off; how many maintain it; closed only when it is.
    [Fact]
    public void A_systems_list_line_says_where_only_when_it_is_off_who_and_whether_it_is_closed()
    {
        Assert.Equal("2 maintainers", SystemsDisplay.ListLine(SystemsDisplayTests.System(maintainers: 2)));
        Assert.Equal("1 maintainer - closed to contributions", SystemsDisplay.ListLine(SystemsDisplayTests.System(accepting: false)));
        Assert.Equal(
            "Not published yet - nobody assigned - closed to contributions",
            SystemsDisplay.ListLine(SystemsDisplayTests.System(inBeta: false, inProduction: false, accepting: false, maintainers: 0)));
    }

    // Nobody assigned is said, with where its submissions go - it is what this screen is opened to find.
    [Fact]
    public void A_system_nobody_maintains_says_where_its_submissions_go()
    {
        Assert.Equal("Nobody maintains this system - its submissions go to the administrator.", SystemsDisplay.MaintainersHeading(0));
        Assert.Equal("Maintainer", SystemsDisplay.MaintainersHeading(1));
        Assert.Equal("Maintainers (3)", SystemsDisplay.MaintainersHeading(3));
        Assert.Equal("Anna (anna@example.com)", SystemsDisplay.MaintainerLine(new PoolMaintainerEntry(7, "Anna", "anna@example.com")));
    }

    // ###########################################################################################
    // A contributor's record: only the counts that are not zero, then when they last sent one. A
    // contributor with no account is their address alone.
    // ###########################################################################################
    [Fact]
    public void A_contributors_record_names_only_what_happened()
    {
        var contributor = new SystemContributorEntry("hest@mailscan.dk", null, Accepted: 3, Waiting: 0, ChangesRequested: 1, Rejected: 0, SystemsDisplayTests.Noon);

        Assert.Equal("hest@mailscan.dk", SystemsDisplay.ContributorName(contributor));
        Assert.Equal("3 accepted, 1 sent back for changes - last sent 2026-September-25", SystemsDisplay.ContributorRecord(contributor));

        var named = new SystemContributorEntry("anna@example.com", "Anna", 0, 1, 0, 2, null);

        Assert.Equal("Anna (anna@example.com)", SystemsDisplay.ContributorName(named));
        Assert.Equal("1 waiting, 2 rejected", SystemsDisplay.ContributorRecord(named));

        Assert.Equal(string.Empty, SystemsDisplay.ContributorRecord(new SystemContributorEntry("x@example.com", null, 0, 0, 0, 0, null)));

        // A system the account does not maintain comes without addresses (2026-10-05): a contributor
        // with no account is said as one - "(no address)" would claim they gave none - and a named
        // one, like a maintainer, is the name alone.
        var anonymous = new SystemContributorEntry(null, null, 1, 0, 0, 0, null);

        Assert.Equal("A contributor without an account", SystemsDisplay.ContributorName(anonymous, addressesHidden: true));
        Assert.Equal("(no address)", SystemsDisplay.ContributorName(anonymous));
        Assert.Equal("Anna", SystemsDisplay.ContributorName(named with { Email = null }, addressesHidden: true));
        Assert.Equal("Anna", SystemsDisplay.MaintainerLine(new PoolMaintainerEntry(7, "Anna", string.Empty)));
    }

    [Fact]
    public void The_contributor_heading_says_when_there_are_none()
    {
        Assert.Equal("Nobody has contributed to this system through CRT yet.", SystemsDisplay.ContributorsHeading(0));
        Assert.Equal("Contributors (2)", SystemsDisplay.ContributorsHeading(2));
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
        var submission = new SystemSubmissionEntry(41, "hest@mailscan.dk", "Corrected U8.", state, SystemsDisplayTests.Noon, null, null);

        // The History view's card states it where it stands now (SystemHistoryDisplay, 2026-10-04).
        Assert.Equal(words, SystemHistoryDisplay.Card(submission, []).State);
    }

    // What the contributor was told, labelled by who read it; nothing when nothing was said.
    [Fact]
    public void What_the_contributor_was_told_is_shown_and_a_missing_description_says_so()
    {
        var told = new SystemSubmissionEntry(38, null, "  ", "returned", SystemsDisplayTests.Noon, null, " U7 is the wrong revision. ");

        Assert.Equal("Told the contributor: U7 is the wrong revision.", SystemsDisplay.SubmissionComment(told));
        Assert.Equal("(no description given)", SystemsDisplay.SubmissionTitle(told));

        Assert.Null(SystemsDisplay.SubmissionComment(told with { DecisionComment = null }));
    }

    // ###########################################################################################
    // *** WHAT IS STILL TO BE DONE IS MARKED, AND SHOWN IN BOLD (owner request, 2026-09-27: "'nobody
    // assigned' is important, as that should be done - the same for 'not in production yet'"). ***
    // The pieces join to exactly the line as it always read.
    // ###########################################################################################
    [Fact]
    public void Nobody_assigned_and_not_in_production_yet_are_marked_as_still_to_be_done()
    {
        SystemOverviewEntry system = SystemsDisplayTests.System(inProduction: false, awaiting: true, maintainers: 0);

        IReadOnlyList<StatusPart> parts = SystemsDisplay.ListLineParts(system);

        Assert.Equal(["not in the stable source yet", "nobody assigned"], parts.Where(part => part.IsToDo).Select(part => part.Text));
        Assert.Equal("In BETA - not in the stable source yet - nobody assigned", SystemsDisplay.ListLine(system));
        Assert.Equal(SystemsDisplay.ListLine(system), string.Concat(parts.Select(part => part.Text)));
    }

    [Theory]
    [InlineData(true, true, true, "BETA ahead of the stable source")]      // waits for "Publish to production"
    [InlineData(false, true, false, "not in BETA")]                 // production holds what BETA lacks
    public void Other_outstanding_states_are_marked_too(bool inBeta, bool inProduction, bool awaiting, string marked)
    {
        IReadOnlyList<StatusPart> parts = SystemsDisplay.ListLineParts(
            SystemsDisplayTests.System(inBeta: inBeta, inProduction: inProduction, awaiting: awaiting));

        Assert.Equal([marked], parts.Where(part => part.IsToDo).Select(part => part.Text));
    }

    // Nothing to do, nothing marked - including a turned-down new system ("Not published yet") and a
    // system closed to contributions on purpose.
    [Fact]
    public void A_system_with_nothing_outstanding_has_nothing_marked()
    {
        Assert.DoesNotContain(SystemsDisplay.ListLineParts(SystemsDisplayTests.System(maintainers: 2)), part => part.IsToDo);
        Assert.DoesNotContain(
            SystemsDisplay.ListLineParts(SystemsDisplayTests.System(inBeta: false, inProduction: false, accepting: false, maintainers: 1)),
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
    public void A_system_listed_only_in_the_stable_list_is_flagged_as_to_do()
    {
        SystemOverviewEntry system = SystemsDisplayTests.System() with { ListedInBeta = false, ListedInStable = true };

        Assert.Equal("listed only in the stable source's drop-down list - 1 maintainer", SystemsDisplay.ListLine(system));
        Assert.Equal(["listed only in the stable source's drop-down list"], SystemsDisplay.ListLineParts(system).Where(part => part.IsToDo).Select(part => part.Text));
    }

    [Fact]
    public void A_board_in_the_stable_source_that_its_list_leaves_out_is_flagged_as_to_do()
    {
        SystemOverviewEntry system = SystemsDisplayTests.System(awaiting: true) with { ListedInBeta = true, ListedInStable = false };

        Assert.Equal("BETA ahead of the stable source - missing from the stable source's drop-down list - 1 maintainer", SystemsDisplay.ListLine(system));
        Assert.Equal(
            ["BETA ahead of the stable source", "missing from the stable source's drop-down list"],
            SystemsDisplay.ListLineParts(system).Where(part => part.IsToDo).Select(part => part.Text));
    }

    // ###########################################################################################
    // Nothing is flagged that is as it should be or not known: both lists naming it; a NEW system in
    // BETA only, not promoted yet, which is listed in BETA alone by design; a list that could not be
    // read; and a board BETA does not hold at all, whose "not in BETA" already says it.
    // ###########################################################################################
    [Fact]
    public void Nothing_about_the_lists_is_said_when_they_are_as_they_should_be_or_unknown()
    {
        Assert.Equal("1 maintainer", SystemsDisplay.ListLine(SystemsDisplayTests.System() with { ListedInBeta = true, ListedInStable = true }));
        Assert.Equal(
            "In BETA - not in the stable source yet - 1 maintainer",
            SystemsDisplay.ListLine(SystemsDisplayTests.System(inProduction: false, awaiting: true) with { ListedInBeta = true, ListedInStable = false }));
        Assert.Equal("1 maintainer", SystemsDisplay.ListLine(SystemsDisplayTests.System() with { ListedInBeta = null, ListedInStable = true }));
        Assert.Equal("1 maintainer", SystemsDisplay.ListLine(SystemsDisplayTests.System() with { ListedInBeta = true, ListedInStable = null }));
        Assert.Equal(
            "In the stable source - not in BETA - 1 maintainer",
            SystemsDisplay.ListLine(SystemsDisplayTests.System(inBeta: false) with { ListedInBeta = false, ListedInStable = true }));
    }

    // An invitation nobody has accepted, on the administrator's screen: who, and how long its code works.
    [Fact]
    public void An_open_invitation_says_who_when_and_until_when()
    {
        var invitation = new MaintainerInvitationEntry(3, "new@example.com", SystemsDisplayTests.Noon, SystemsDisplayTests.Noon.AddDays(14));

        Assert.Equal("Invited: new@example.com", SystemsDisplay.InvitationLine(invitation));
        Assert.Equal(
            "Sent 2026-September-25 - the code works until 2026-October-9, and they become a maintainer when they use it",
            SystemsDisplay.InvitationFooter(invitation));
    }

    // ###########################################################################################
    // THE SYSTEM'S HISTORY (owner request, 2026-09-27: "I would like to see the date, newest first,
    // to understand what has happened to a system"): what happened, and a grey line of by whom and
    // what more - the date is the History view's own column since 2026-10-04. Decisions in CRT's
    // own state words.
    // ###########################################################################################
    [Theory]
    [InlineData(SystemHistoryEvents.Sent, 9L, "Corrected U8.", "hest@mailscan.dk", "#9 sent", "from hest@mailscan.dk - Corrected U8.")]
    [InlineData(SystemHistoryEvents.Decided, 9L, "merged", "Anna", "#9 - Published to the BETA source", "by Anna")]
    [InlineData(SystemHistoryEvents.Decided, 9L, "changes_requested", null, "#9 - Changes requested", "")]
    [InlineData(SystemHistoryEvents.PublishedToProduction, null, "revision 2026-September-25; 3 file(s) copied", "admin@example.com", "Published to the stable source", "by admin@example.com - revision 2026-September-25; 3 file(s) copied")]
    [InlineData(SystemHistoryEvents.PushedBack, null, "1 submission(s) returned to the queue; Not ready.", "admin@example.com", "Pushed back from BETA to the queue", "by admin@example.com - 1 submission(s) returned to the queue; Not ready.")]
    [InlineData(SystemHistoryEvents.RejectedFromBeta, null, "1 submission(s) rejected; Not for this board.", "admin@example.com", "Rejected in BETA and taken out of it", "by admin@example.com - 1 submission(s) rejected; Not for this board.")]
    [InlineData(SystemHistoryEvents.MaintainerAdded, null, "anna@example.com", "admin@example.com", "anna@example.com made a maintainer", "by admin@example.com")]
    [InlineData(SystemHistoryEvents.MaintainerRemoved, null, "anna@example.com", "admin@example.com", "anna@example.com removed as a maintainer", "by admin@example.com")]
    [InlineData(SystemHistoryEvents.Invited, null, "new@example.com", "admin@example.com", "new@example.com invited to be a maintainer", "by admin@example.com")]
    [InlineData(SystemHistoryEvents.InvitationAccepted, null, "new@example.com", "new@example.com", "new@example.com accepted the invitation and became a maintainer", "")]
    [InlineData(SystemHistoryEvents.Placed, null, "Commodore 64 / 250407", "anna@example.com", "Placed in the drop-down lists", "by anna@example.com - Commodore 64 / 250407")]
    [InlineData(SystemHistoryEvents.Amended, 9L, "amendment 2, 3 file(s)", "anna@example.com", "#9 changed by a maintainer", "by anna@example.com - amendment 2, 3 file(s)")]
    [InlineData(SystemHistoryEvents.Deleted, null, "2 file(s) removed from BETA, 2 from the stable data; 3 submission(s) deleted, 0 of them open", "admin@example.com", "Deleted from the BETA and stable data and the database", "by admin@example.com - 2 file(s) removed from BETA, 2 from the stable data; 3 submission(s) deleted, 0 of them open")]
    [InlineData("something.new", null, null, null, "something.new", "")]
    public void A_history_entry_says_what_happened_and_by_whom(string kind, long? submission, string? detail, string? who, string line, string footer)
    {
        var entry = new SystemHistoryEntry(SystemsDisplayTests.Noon, kind, who, submission, detail);

        Assert.Equal(line, SystemsDisplay.HistoryWhat(entry));
        Assert.Equal(footer, SystemsDisplay.HistoryFooter(entry));
    }

    [Fact]
    public void A_decision_carries_what_the_contributor_was_told()
    {
        var entry = new SystemHistoryEntry(SystemsDisplayTests.Noon, SystemHistoryEvents.Decided, "Anna", 9, "rejected", "Wrong board.");

        Assert.Equal("Told the contributor: Wrong board.", SystemsDisplay.HistoryNote(entry));
        Assert.Null(SystemsDisplay.HistoryNote(entry with { Note = null }));
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
        SystemOverviewEntry system = SystemsDisplayTests.System() with { ViewsLast30Days = views };

        Assert.Equal("1 maintainer", SystemsDisplay.ListLine(system));
    }

    // ###########################################################################################
    // The section: the three counts with their numbers bold, the countries, the BETA-source views
    // counted apart - each line only when there is something to say.
    // ###########################################################################################
    [Fact]
    public void The_views_section_says_the_counts_the_countries_and_the_beta_views_apart()
    {
        var views = new BoardViewStatistics(12, 48, 310, 3, [new("DK", "Denmark", 120), new("DE", "Germany", 60)]);

        Assert.Equal("Views in CRT", SystemsDisplay.ViewsHeading(views));

        IReadOnlyList<ReviewNoteRun> counts = SystemsDisplay.ViewCountRuns(views);
        Assert.Equal("[12] in the last 7 days - [48] in 30 days - [310] in 12 months", string.Concat(counts.Select(run => run.Text)));
        Assert.Equal(["12", "48", "310"], counts.Where(run => run.IsCount).Select(run => run.Text));

        Assert.Equal("Most views in the last 12 months: Denmark (120), Germany (60).", SystemsDisplay.ViewCountriesLine(views));
        Assert.Equal(
            "Not counted above: [3] from CRTs downloading BETA data in the last 30 days.",
            string.Concat(SystemsDisplay.ViewBetaRuns(views)!.Select(run => run.Text)));

        Assert.Equal("A view is this board on screen in CRT for at least 10 seconds, counted every time.", SystemsDisplay.ViewExplanation);
    }

    [Fact]
    public void A_board_nobody_viewed_says_so_in_its_heading_and_nothing_else()
    {
        var none = new BoardViewStatistics(0, 0, 0, 0, []);

        Assert.True(SystemsDisplay.HasNoViews(none));
        Assert.Equal("No views of this board in CRT counted in the last 12 months.", SystemsDisplay.ViewsHeading(none));
        Assert.Null(SystemsDisplay.ViewCountriesLine(none));
        Assert.Null(SystemsDisplay.ViewBetaRuns(none));

        // Only BETA-source views is still something to show.
        Assert.False(SystemsDisplay.HasNoViews(none with { FromBetaLast30Days = 2 }));
    }
}
