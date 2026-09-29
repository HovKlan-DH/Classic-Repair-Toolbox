using Handlers.DataHandling;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.Online
{
    // ###########################################################################################
    // TELLS THE SERVER A CONTRIBUTOR DISCARDED THEIR DRAFT (owner request, 2026-09-28) - see
    // CRT.Data's DraftDiscardContract for why a maintainer needs to know.
    //
    // The Drafts tab marks the receipts when the draft is discarded (SubmissionReceiptStore
    // .MarkDraftDiscarded) and runs this straight away; the launch runs it again for anything still
    // waiting. So a discard made with no network is reported the next time CRT starts with one, and
    // the maintainer is not left approving work its author threw away.
    //
    // *** RUN FROM THE UI THREAD. *** The receipts are one list that the status checks change too;
    // each await here resumes on the caller's context, so every change to it happens on that one
    // thread, as the status refresh's do.
    //
    // One run at a time: the launch and a discard can start one each, and two would send the same
    // notice twice (harmless - the server keeps the first - but pointless).
    // ###########################################################################################
    public static class DraftDiscardReporter
    {
        private static int thisRunning;

        // Sends every notice still waiting; returns how many are finished (told, or never tellable).
        // `send` is SubmissionClient.ReportDraftDiscardedAsync - a parameter so a test needs no network.
        public static async Task<int> ReportPendingAsync(
            Func<long, string, CancellationToken, Task<int?>> send,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(send);

            if (Interlocked.Exchange(ref DraftDiscardReporter.thisRunning, 1) == 1)
                return 0;

            try
            {
                int finished = 0;

                foreach (SubmissionReceipt receipt in SubmissionReceiptStore.PendingDraftDiscardNotices)
                {
                    int? status;

                    try
                    {
                        status = await send(receipt.SubmissionId, receipt.UploadToken, cancellationToken);
                    }
                    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"Could not report the discarded draft of submission [{receipt.SubmissionId}]: [{ex.Message}] - tried again at the next launch");
                        status = null;
                    }

                    if (DraftDiscardContract.DeliveryFor(status) == DraftDiscardDelivery.Done)
                    {
                        SubmissionReceiptStore.MarkDraftDiscardReported(receipt.SubmissionId);
                        finished++;

                        Logger.Info($"Reported the discarded draft of submission [{receipt.SubmissionId}] (answer [{status}])");
                    }
                }

                return finished;
            }
            finally
            {
                Interlocked.Exchange(ref DraftDiscardReporter.thisRunning, 0);
            }
        }
    }
}
