using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// The submission "Contributor Submissions" will open on, read before the tab is shown (owner
// request, 2026-10-02: no "please wait" on the first look at the tab) - SubmissionPrefetch.
//
// What it must hold to: the answers are used only for EXACTLY the queue row they were read for, and
// only once. A held answer used for a row the queue has since changed would show the maintainer a
// submission as it was, not as it is.
// ###########################################################################################
public sealed class SubmissionPrefetchTests
{
    private static ReviewQueueRow Row(long id = 42, string state = "pending", bool? awaitsYou = true) =>
        new(id, "Commodore/C64/250407", state, "Corrected U8.", "c@example.com",
            new DateTimeOffset(2026, 10, 2, 9, 0, 0, TimeSpan.Zero), AwaitsYou: awaitsYou);

    private static PrefetchedSubmission Read(ReviewQueueRow row) =>
        new(
            row,
            new ReviewSubmissionDetail(row, true, new ReviewChangeSummaryView(false, []), [], new ReviewSubmissionAssets([]), [],
                new Dictionary<string, string>(), new Dictionary<string, string>(), [], null),
            new ReviewTableData(1, null, new SubmissionRows()));

    [Fact]
    public void With_nothing_held_the_entry_to_open_needs_reading_and_no_entry_does_not()
    {
        var prefetch = new SubmissionPrefetch();

        Assert.True(prefetch.NeedsReading(Row()));
        Assert.False(prefetch.NeedsReading(null));
        Assert.Null(prefetch.HeldRow);
    }

    // The badge's minute check asks every minute; an unchanged entry must not be read every minute.
    [Fact]
    public void The_entry_already_held_is_not_read_again()
    {
        var prefetch = new SubmissionPrefetch();
        prefetch.Keep(Read(Row()));

        Assert.False(prefetch.NeedsReading(Row()));
        Assert.Equal(Row(), prefetch.HeldRow);
    }

    // The queue row is the record of what changed: another approver acting, a state moving.
    [Fact]
    public void A_changed_queue_row_or_another_entry_is_read_again()
    {
        var prefetch = new SubmissionPrefetch();
        prefetch.Keep(Read(Row()));

        Assert.True(prefetch.NeedsReading(Row(awaitsYou: false)));
        Assert.True(prefetch.NeedsReading(Row(state: "approved")));
        Assert.True(prefetch.NeedsReading(Row(id: 43)));
    }

    [Fact]
    public void Opening_the_row_it_was_read_for_takes_it_once()
    {
        var prefetch = new SubmissionPrefetch();
        PrefetchedSubmission read = Read(Row());
        prefetch.Keep(read);

        Assert.Same(read, prefetch.Take(Row()));
        Assert.Null(prefetch.Take(Row()));
        Assert.Null(prefetch.HeldRow);
    }

    // Same submission, but the queue says something else about it now: what was read is stale.
    [Fact]
    public void A_row_changed_since_the_read_is_not_opened_from_it()
    {
        var prefetch = new SubmissionPrefetch();
        prefetch.Keep(Read(Row()));

        Assert.Null(prefetch.Take(Row(awaitsYou: false)));
    }

    // Opening another submission means the moment for the held one has passed - it is dropped.
    [Fact]
    public void Opening_another_submission_drops_what_was_held()
    {
        var prefetch = new SubmissionPrefetch();
        prefetch.Keep(Read(Row()));

        Assert.Null(prefetch.Take(Row(id: 43)));
        Assert.Null(prefetch.HeldRow);
        Assert.Null(prefetch.Take(Row()));
    }

    [Fact]
    public void Forgetting_empties_it()
    {
        var prefetch = new SubmissionPrefetch();
        prefetch.Keep(Read(Row()));

        prefetch.Forget();

        Assert.Null(prefetch.HeldRow);
        Assert.True(prefetch.NeedsReading(Row()));
    }
}
