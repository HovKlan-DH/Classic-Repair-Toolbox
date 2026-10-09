using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// Covers ReviewQueueDisplay - how a queue row reads: the badges, the board part by part, the
// contributor's comment and the footer (owner request, 2026-09-26).
//
// The wording is tested rather than eyeballed because the queue is the screen a maintainer decides
// from: which submission to open next, and whether anything has been waiting too long. Two of the
// rules below exist specifically to stop the list telling a comfortable lie - a missing timestamp
// rendered as 1970, and a wait rounded up into a coarser, kinder unit.
public sealed class ReviewQueueDisplayTests
{
    private static readonly DateTimeOffset Now = new(2026, 9, 21, 12, 0, 0, TimeSpan.Zero);

    private static ReviewQueueRow Row(
        long id = 42,
        string boardId = "Commodore/C64/250407",
        string summary = "Corrected R12.",
        DateTimeOffset? createdUtc = null,
        bool touchesSharedFiles = false,
        bool? isNewBoard = null,
        bool? awaitsYou = null,
        string state = "pending") =>
        new(id, boardId, state, summary, "someone@example.com", createdUtc, touchesSharedFiles, isNewBoard, awaitsYou);

    // -----------------------------------------------------------------------------------
    // Grouped by board (2026-09-26)
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** EACH BOARD ONCE, AND THE QUEUE'S ORDER KEPT. *** The queue arrives oldest first - the
    // longest wait leads, a decision (ReviewQueueTests). Grouping must not undo it: the board
    // holding the oldest submission comes first, and each board's submissions stay in queue order.
    // ###########################################################################################
    [Fact]
    public void The_queue_is_grouped_by_board_in_the_order_the_oldest_submission_waits()
    {
        IReadOnlyList<ReviewQueueGroup> groups = ReviewQueueDisplay.Group(
        [
            ReviewQueueDisplayTests.Row(id: 1, boardId: "Commodore/C64/250407"),
            ReviewQueueDisplayTests.Row(id: 2, boardId: "Commodore/C65/Prototype"),
            ReviewQueueDisplayTests.Row(id: 3, boardId: "Commodore/C64/250407"),
            ReviewQueueDisplayTests.Row(id: 4, boardId: "Amstrad/CPC464/Z70200")
        ]);

        Assert.Equal(["Commodore/C64/250407", "Commodore/C65/Prototype", "Amstrad/CPC464/Z70200"], groups.Select(group => group.BoardId));
        Assert.Equal([1L, 3L], groups[0].Rows.Select(row => row.Id));
        Assert.Equal([2L], groups[1].Rows.Select(row => row.Id));
    }

    [Fact]
    public void An_empty_queue_has_no_groups()
    {
        Assert.Empty(ReviewQueueDisplay.Group([]));
    }

    // A board is new or not as a whole - the server's answer, from whichever row carries one.
    [Fact]
    public void A_group_is_a_new_board_when_the_server_says_so()
    {
        Assert.True(ReviewQueueDisplay.Group([ReviewQueueDisplayTests.Row(isNewBoard: true)])[0].IsNewBoard);
        Assert.False(ReviewQueueDisplay.Group([ReviewQueueDisplayTests.Row(isNewBoard: false)])[0].IsNewBoard);
        Assert.Null(ReviewQueueDisplay.Group([ReviewQueueDisplayTests.Row()])[0].IsNewBoard);
    }

    [Fact]
    public void A_heading_names_the_board_part_by_part()
    {
        Assert.Equal("Commodore / C64 / 250407", ReviewQueueDisplay.BoardHeading("Commodore/C64/250407"));
    }

    // Not three parts: shown whole - and a missing one said to be missing, not left as a gap that
    // looks like a rendering fault.
    [Fact]
    public void A_malformed_or_missing_board_is_headed_as_it_is()
    {
        Assert.Equal("Commodore/C64", ReviewQueueDisplay.BoardHeading("Commodore/C64"));
        Assert.Equal("(unknown board)", ReviewQueueDisplay.BoardHeading(""));
        Assert.Equal("(unknown board)", ReviewQueueDisplay.BoardHeading(null));
    }

    // -----------------------------------------------------------------------------------
    // The parts of a row
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_board_is_named_part_by_part()
    {
        Assert.Equal(("Commodore", "C64", "250407"), ReviewQueueDisplay.BoardParts("Commodore/C64/250407"));
    }

    // Not three well-formed parts: no parts at all - the row then shows the id whole rather than
    // naming, say, "C64" as the manufacturer.
    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("Commodore/C64")]
    [InlineData("Commodore/C64/250407/extra")]
    [InlineData("Commodore//250407")]
    public void A_malformed_board_id_has_no_parts(string? boardId)
    {
        Assert.Null(ReviewQueueDisplay.BoardParts(boardId));
    }

    [Fact]
    public void The_comment_is_trimmed_and_a_missing_one_SAYS_SO()
    {
        // The summary field is optional, so blank is a real possibility - and an empty line reads
        // as a rendering fault rather than as missing information.
        Assert.Equal("Corrected R12.", ReviewQueueDisplay.Comment(ReviewQueueDisplayTests.Row(summary: "  Corrected R12.\n")));
        Assert.Equal("(no description given)", ReviewQueueDisplay.Comment(ReviewQueueDisplayTests.Row(summary: "   ")));
    }

    // "New board" is the one badge a maintainer must never miss - the highest-risk submission
    // there is. It is the server's to say; a server that did not say gets no badge, not a guess.
    // A published board - the ordinary case - is not marked at all (2026-09-26).
    [Fact]
    public void Only_a_new_board_is_badged()
    {
        Assert.Equal("New board", ReviewQueueDisplay.BoardBadge(true));
        Assert.Null(ReviewQueueDisplay.BoardBadge(false));
        Assert.Null(ReviewQueueDisplay.BoardBadge(null));
    }

    // ###########################################################################################
    // An opened submission awaits you while this account may decide it and its approval is still
    // needed - the same answer that leaves Approve on. An older server sends no approval status;
    // one approval then publishes, so it awaits whoever may publish.
    // ###########################################################################################
    [Fact]
    public void An_opened_submission_awaits_you_only_while_your_approval_is_still_needed()
    {
        static ApprovalStatus Status(bool canApprove, bool youApproved = false) =>
            new(
                Required: [ApproverRole.Maintainer, ApproverRole.Administrator],
                Given: [],
                WaitingFor: [ApproverRole.Maintainer, ApproverRole.Administrator],
                YourRole: ApproverRole.Administrator,
                CanApprove: canApprove,
                ApprovalPublishes: false,
                YouApproved: youApproved);

        Assert.True(ReviewQueueDisplay.AwaitsYou(canPublish: true, approval: null));
        Assert.True(ReviewQueueDisplay.AwaitsYou(canPublish: true, approval: Status(canApprove: true)));

        Assert.False(ReviewQueueDisplay.AwaitsYou(canPublish: true, approval: Status(canApprove: false, youApproved: true)));
        Assert.False(ReviewQueueDisplay.AwaitsYou(canPublish: false, approval: Status(canApprove: true)));
    }

    // -----------------------------------------------------------------------------------
    // The footer - the wait, and the two-approval notes
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_footer_says_how_long_it_has_waited()
    {
        Assert.Equal(
            "Waiting 3 hours",
            ReviewQueueDisplay.Footer(ReviewQueueDisplayTests.Row(createdUtc: ReviewQueueDisplayTests.Now.AddHours(-3)), ReviewQueueDisplayTests.Now));

        // "Just now" alone would not say what happened just now.
        Assert.Equal(
            "Arrived just now",
            ReviewQueueDisplay.Footer(ReviewQueueDisplayTests.Row(createdUtc: ReviewQueueDisplayTests.Now), ReviewQueueDisplayTests.Now));
    }

    [Fact]
    public void A_submission_changing_SHARED_FILES_says_so_in_its_footer()
    {
        // The one kind of change that reaches every board citing the file, and why it needs the
        // administrator's approval as well.
        Assert.Equal(
            "Waiting 2 days - replaces a shared file",
            ReviewQueueDisplay.Footer(
                ReviewQueueDisplayTests.Row(createdUtc: ReviewQueueDisplayTests.Now.AddDays(-2), touchesSharedFiles: true),
                ReviewQueueDisplayTests.Now));

        Assert.DoesNotContain("shared files", ReviewQueueDisplay.Footer(ReviewQueueDisplayTests.Row(), ReviewQueueDisplayTests.Now));
    }

    // ###########################################################################################
    // A submission NOT waiting for this account is dimmed in the list, and its footer says why:
    // this account's approval is given, the other approver's is not. It replaces "one of two
    // approvals given", which says the same from the other side.
    // ###########################################################################################
    [Fact]
    public void A_submission_not_waiting_for_you_says_it_is_with_the_other_approver()
    {
        Assert.Equal(
            "Waiting 2 days - replaces a shared file - with the other approver",
            ReviewQueueDisplay.Footer(
                ReviewQueueDisplayTests.Row(createdUtc: ReviewQueueDisplayTests.Now.AddDays(-2), touchesSharedFiles: true, awaitsYou: false, state: "approved"),
                ReviewQueueDisplayTests.Now));

        // Waiting for YOU as the second approver: the first approval is the news.
        Assert.Equal(
            "Waiting 2 days - replaces a shared file - one of two approvals given",
            ReviewQueueDisplay.Footer(
                ReviewQueueDisplayTests.Row(createdUtc: ReviewQueueDisplayTests.Now.AddDays(-2), touchesSharedFiles: true, awaitsYou: true, state: "approved"),
                ReviewQueueDisplayTests.Now));
    }

    // -----------------------------------------------------------------------------------
    // How long it has been waiting - the number that shames a backlog
    // -----------------------------------------------------------------------------------

    [Theory]
    [InlineData(0, "just now")]
    [InlineData(30, "just now")]
    [InlineData(60, "waiting 1 minute")]
    [InlineData(120, "waiting 2 minutes")]
    [InlineData(3600, "waiting 1 hour")]
    [InlineData(7200, "waiting 2 hours")]
    [InlineData(86400, "waiting 1 day")]
    [InlineData(172800, "waiting 2 days")]
    public void A_wait_is_described_in_the_unit_a_person_thinks_in(int secondsWaited, string expected)
    {
        Assert.Equal(
            expected,
            ReviewQueueDisplay.Waiting(
                ReviewQueueDisplayTests.Now.AddSeconds(-secondsWaited),
                ReviewQueueDisplayTests.Now));
    }

    [Fact]
    public void A_wait_TRUNCATES_rather_than_rounding_up_into_a_kinder_unit()
    {
        // *** THIS IS THE NUMBER THAT SHAMES A BACKLOG. *** Rounding in the comforting direction
        // is exactly the wrong behaviour: 13 days must never read as "2 weeks" and 29 days must
        // never read as "a month", because the coarser unit always makes a wait sound shorter
        // than it is, and this figure exists to be uncomfortable.
        Assert.Equal(
            "waiting 13 days",
            ReviewQueueDisplay.Waiting(ReviewQueueDisplayTests.Now.AddDays(-13.9), ReviewQueueDisplayTests.Now));

        Assert.Equal(
            "waiting 29 days",
            ReviewQueueDisplay.Waiting(ReviewQueueDisplayTests.Now.AddDays(-29.9), ReviewQueueDisplayTests.Now));

        // And the boundary itself: 59 minutes is not yet an hour.
        Assert.Equal(
            "waiting 59 minutes",
            ReviewQueueDisplay.Waiting(ReviewQueueDisplayTests.Now.AddMinutes(-59.9), ReviewQueueDisplayTests.Now));
    }

    [Fact]
    public void A_MISSING_timestamp_shows_nothing_rather_than_fifty_six_years()
    {
        // The parser deliberately keeps CreatedUtc nullable rather than defaulting it. Rendering
        // a default as "waiting 20000 days" would be a spectacular lie on a screen whose whole
        // job is telling the truth about a backlog.
        Assert.Equal(string.Empty, ReviewQueueDisplay.Waiting(null, ReviewQueueDisplayTests.Now));
        Assert.Equal(string.Empty, ReviewQueueDisplay.Footer(ReviewQueueDisplayTests.Row(), ReviewQueueDisplayTests.Now));
    }

    [Fact]
    public void A_clock_skew_reads_as_just_now_rather_than_a_negative_wait()
    {
        // The client's clock and the server's need not agree to the second. "waiting -3 minutes"
        // looks broken; "just now" is true enough and is what the maintainer would conclude anyway.
        Assert.Equal(
            "just now",
            ReviewQueueDisplay.Waiting(ReviewQueueDisplayTests.Now.AddMinutes(3), ReviewQueueDisplayTests.Now));
    }

    [Fact]
    public void A_null_row_is_refused_rather_than_drawn_blank()
    {
        Assert.Throws<ArgumentNullException>(() => ReviewQueueDisplay.Comment(null!));
        Assert.Throws<ArgumentNullException>(() => ReviewQueueDisplay.Footer(null!, ReviewQueueDisplayTests.Now));
    }
}
