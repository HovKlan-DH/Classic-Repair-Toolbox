using System;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftDiscardContract - telling the server a contributor discarded their own draft (owner
// request, 2026-09-28): which of their submissions it is told about, what its answers mean, and
// the route both applications build from one place.
// ###########################################################################################
public sealed class DraftDiscardContractTests
{
    private const string C128 = "Commodore/C128/310378";

    private static readonly DateTimeOffset DraftCreated = new(2026, 9, 20, 12, 0, 0, TimeSpan.Zero);

    private static SubmissionReceipt Receipt(long id, string state, DateTimeOffset? sent = null, string systemId = C128, DateTimeOffset? discarded = null) =>
        new()
        {
            SubmissionId = id,
            UploadToken = "tok" + id,
            SystemId = systemId,
            SentUtc = sent ?? DraftCreated.AddDays(1),
            LastKnownState = state,
            DraftDiscardedUtc = discarded
        };

    [Fact]
    public void The_route_is_under_the_submission_it_is_about()
    {
        Assert.Equal("submissions/42/draft-discarded", DraftDiscardContract.PathUnderApi(42));
    }

    // ###########################################################################################
    // *** EXACTLY THE SUBMISSIONS CRT STILL ASKS THE SERVER ABOUT. *** Waiting, approved,
    // changes requested, in BETA, taken back out of BETA, not checked yet - a maintainer may
    // still act on each. A final one (published, not accepted, replaced, expired) is past it.
    // ###########################################################################################
    [Theory]
    [InlineData("", true)]
    [InlineData("pending", true)]
    [InlineData("approved", true)]
    [InlineData("changes_requested", true)]
    [InlineData("merged", true)]
    [InlineData("returned", true)]
    [InlineData("published", false)]
    [InlineData("rejected", false)]
    [InlineData("withdrawn", false)]
    [InlineData("abandoned", false)]
    public void A_submission_is_worth_reporting_while_a_maintainer_may_still_act_on_it(string state, bool expected)
    {
        Assert.Equal(expected, DraftDiscardContract.IsWorthReporting(state));
        Assert.Equal(SubmissionReceiptPresenter.IsStillOpen(state), DraftDiscardContract.IsWorthReporting(state));
    }

    [Fact]
    public void Only_this_boards_unfinished_submissions_sent_from_this_draft_and_not_reported_yet_are_chosen()
    {
        SubmissionReceipt[] receipts =
        [
            Receipt(1, "merged"),                                         // chosen
            Receipt(2, "pending"),                                        // chosen
            Receipt(3, "published"),                                      // finished
            Receipt(4, "merged", systemId: "Commodore/C64/250407"),       // another board
            Receipt(5, "merged", sent: DraftCreated.AddDays(-1)),         // an earlier draft's
            Receipt(6, "merged", discarded: DraftCreated.AddDays(2)),     // already reported
        ];

        Assert.Equal(
            [1, 2],
            DraftDiscardContract.WhichToReport(receipts, " commodore/c128/310378 ", DraftCreated).Select(receipt => receipt.SubmissionId));
    }

    // A draft marker that does not say when it was made leaves no submission out on that ground.
    [Fact]
    public void With_no_draft_creation_time_every_unfinished_submission_of_the_board_is_chosen()
    {
        Assert.Equal(
            [5],
            DraftDiscardContract.WhichToReport([Receipt(5, "merged", sent: DraftCreated.AddDays(-9))], C128, null).Select(receipt => receipt.SubmissionId));

        Assert.Empty(DraftDiscardContract.WhichToReport([Receipt(5, "merged")], "  ", null));
    }

    // ###########################################################################################
    // *** WHICH ANSWERS FINISH A NOTICE. *** Told (2xx), or never tellable - no such submission or
    // a token that no longer matches (404), a request refused as such. Anything else - no answer,
    // a 5xx, a rate limit - is tried again at the next launch, since late beats never.
    // ###########################################################################################
    [Theory]
    [InlineData(204, DraftDiscardDelivery.Done)]
    [InlineData(200, DraftDiscardDelivery.Done)]
    [InlineData(404, DraftDiscardDelivery.Done)]
    [InlineData(400, DraftDiscardDelivery.Done)]
    [InlineData(500, DraftDiscardDelivery.TryLater)]
    [InlineData(503, DraftDiscardDelivery.TryLater)]
    [InlineData(429, DraftDiscardDelivery.TryLater)]
    [InlineData(null, DraftDiscardDelivery.TryLater)]
    public void An_answer_finishes_the_notice_or_leaves_it_for_the_next_launch(int? status, DraftDiscardDelivery expected)
    {
        Assert.Equal(expected, DraftDiscardContract.DeliveryFor(status));
    }
}
