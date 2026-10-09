using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Handlers.DataHandling;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftDiscardReporter - telling the server a contributor discarded their draft (owner request,
// 2026-09-28), for every receipt still waiting to be told. What each answer does is CRT.Data's
// DraftDiscardContract.DeliveryFor; this pins that the reporter applies it to the store: a finished
// notice is never sent again, and one with no answer is kept for the next launch.
//
// Drives the real static store through its LoadFrom seam, so it joins that store's collection.
// ###########################################################################################
[Collection("SubmissionReceipts")]
public sealed class DraftDiscardReporterTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public DraftDiscardReporterTests()
    {
        SubmissionReceiptStore.LoadFrom(Path.Combine(this.thisWorkspace.Root, "submissions.json"));
    }

    public void Dispose()
    {
        SubmissionReceiptStore.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private static void Discarded(params long[] ids)
    {
        foreach (long id in ids)
            SubmissionReceiptStore.Record(new SubmissionReceipt { SubmissionId = id, UploadToken = "tok" + id, BoardId = "Commodore/C128/310378", LastKnownState = "merged" });

        SubmissionReceiptStore.MarkDraftDiscarded(ids, new DateTimeOffset(2026, 9, 28, 9, 0, 0, TimeSpan.Zero));
    }

    [Fact]
    public async Task A_told_or_untellable_notice_is_finished_and_one_with_no_answer_waits_for_the_next_launch()
    {
        Discarded(1, 2, 3, 4);

        var answers = new Dictionary<long, int?> { [1] = 204, [2] = 404, [3] = 503, [4] = null };
        var sent = new List<(long, string)>();

        int finished = await DraftDiscardReporter.ReportPendingAsync((id, token, _) =>
        {
            sent.Add((id, token));
            return Task.FromResult(answers[id]);
        });

        Assert.Equal(2, finished);
        Assert.Equal([(1L, "tok1"), (2L, "tok2"), (3L, "tok3"), (4L, "tok4")], sent);
        Assert.Equal([3L, 4L], SubmissionReceiptStore.PendingDraftDiscardNotices.Select(receipt => receipt.SubmissionId).OrderBy(id => id));
    }

    [Fact]
    public async Task A_finished_notice_is_never_sent_again()
    {
        Discarded(1);

        await DraftDiscardReporter.ReportPendingAsync((_, _, _) => Task.FromResult<int?>(204));

        int calls = 0;
        await DraftDiscardReporter.ReportPendingAsync((_, _, _) =>
        {
            calls++;
            return Task.FromResult<int?>(204);
        });

        Assert.Equal(0, calls);
    }

    // A sender that throws is a notice not delivered - kept, and the others still go.
    [Fact]
    public async Task A_sender_that_throws_leaves_that_notice_waiting_and_still_sends_the_rest()
    {
        Discarded(1, 2);

        int finished = await DraftDiscardReporter.ReportPendingAsync((id, _, _) =>
            id == 1 ? throw new InvalidOperationException("broken") : Task.FromResult<int?>(204));

        Assert.Equal(1, finished);
        Assert.Equal([1L], SubmissionReceiptStore.PendingDraftDiscardNotices.Select(receipt => receipt.SubmissionId));
    }
}
