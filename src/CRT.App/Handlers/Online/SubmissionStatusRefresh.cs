using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.Online
{
    // ###########################################################################################
    // ASKS THE SERVER WHAT HAS HAPPENED TO THIS MACHINE'S SUBMISSIONS, and writes the answers
    // into the local receipts (maintainer report, 2026-09-22).
    //
    // *** WHY THIS EXISTS AS ITS OWN CLASS: THE CHECK ONLY EVER RAN FROM A BUTTON. *** Receipts
    // were loaded from disk at startup, but nothing asked the server until somebody opened "My
    // submissions" and pressed Refresh. So a reviewer could request changes and the contributor
    // would launch CRT to a stale cache, with no badge and no comment, until they happened to
    // press a button they had no reason to think was necessary. The feedback was invisible in
    // exactly the situation it was written for.
    //
    // Pulled out of MySubmissionsWindow rather than copied, so the launch check and the button
    // cannot drift into asking different questions - the same reasoning CLAUDE.md gives for every
    // other rule that lives in Handlers/ instead of in a control.
    //
    // *** ONLY SUBMISSIONS THAT CAN STILL MOVE ARE ASKED ABOUT. *** A decided one cannot change
    // again, so re-asking is a request that can only return what is already stored. That is what
    // keeps this cheap for a contributor with a long history, and it is the same
    // SubmissionReceiptPresenter.IsStillOpen rule the window has always used.
    //
    // *** EVERY FAILURE IS SILENT AND HARMLESS. *** This runs unattended at launch, so an
    // unreachable server, an expired token or a rate limit must cost nothing but a stale row -
    // never a dialog, never a delayed start, never a failed launch. The receipt simply keeps its
    // last known state, exactly as it would if the check had not run.
    // ###########################################################################################
    public static class SubmissionStatusRefresh
    {
        // ###########################################################################################
        // The lookup used to ask about one submission. Injected so tests drive this with no network
        // (CLAUDE.md test rule 6), and so the caller decides whether that is a real client.
        // ###########################################################################################
        public delegate Task<SubmissionStatus?> StatusLookup(
            long submissionId,
            string uploadToken,
            CancellationToken cancellationToken);

        // ###########################################################################################
        // Checks every still-open submission and records what the server said.
        //
        // Returns how many receipts actually CHANGED - not how many were checked. The caller uses
        // that to decide whether anything on screen needs rebuilding, and "changed" is the only
        // honest answer to that question: a check that confirmed six unchanged rows should not
        // cause a refresh of anything.
        // ###########################################################################################
        public static async Task<int> RefreshAsync(
            StatusLookup lookup,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(lookup);

            List<SubmissionReceipt> toCheck = SubmissionReceiptStore.All
                .Where(receipt => SubmissionReceiptPresenter.IsStillOpen(receipt.LastKnownState))
                .ToList();

            int changed = 0;

            foreach (SubmissionReceipt receipt in toCheck)
            {
                // Honoured between rows rather than only at the start: the app may be closing, and
                // a long queue of receipts should not hold shutdown open.
                cancellationToken.ThrowIfCancellationRequested();

                SubmissionStatus? status;

                try
                {
                    status = await lookup(receipt.SubmissionId, receipt.UploadToken, cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    // One unreachable row must not abandon the rest - the others may be on a
                    // different code path entirely (a decided one, a token the server still knows).
                    CrtLog.Warning(
                        $"Could not check submission {receipt.SubmissionId}: [{ex.Message}]");

                    continue;
                }

                if (status is null)
                {
                    continue;
                }

                // *** COMPARED BEFORE WRITING, so the return value means something. *** UpdateState
                // rewrites the row and saves the file unconditionally, so counting calls rather
                // than actual changes would report "something happened" on every launch and make
                // the caller rebuild the UI for nothing.
                bool stateMoved = !string.Equals(
                    status.State ?? string.Empty,
                    receipt.LastKnownState ?? string.Empty,
                    StringComparison.Ordinal);

                bool commentMoved = !string.Equals(
                    (status.ReviewerComment ?? string.Empty).Trim(),
                    (receipt.ReviewerComment ?? string.Empty).Trim(),
                    StringComparison.Ordinal);

                SubmissionReceiptStore.UpdateState(
                    receipt.SubmissionId,
                    status.State ?? string.Empty,
                    status.ReviewerComment ?? string.Empty,
                    now,
                    status.DecidedUtc);

                if (stateMoved || commentMoved)
                {
                    changed++;
                }
            }

            return changed;
        }

        // ###########################################################################################
        // The launch check: fire-and-forget, and it swallows everything.
        //
        // *** NOTHING ABOUT STARTUP MAY DEPEND ON THIS. *** It is started without being awaited, so
        // an exception escaping here would fault a Task nobody observes - reported arbitrarily late
        // by TaskScheduler.UnobservedTaskException, or never. The same reasoning Main.StartAsync's
        // own header gives for catching there.
        //
        // Deliberately does NOT run when there is nothing to ask about, so a user who has never
        // contributed makes no network request at all on launch.
        //
        // *** TWO CALLBACKS, BECAUSE THERE ARE TWO QUESTIONS (code review, 2026-09-25). ***
        //
        //   - onChanged: "did a receipt MOVE?" - fires only when one did. Right for redrawing a
        //     badge, which only needs work when something new arrived.
        //   - onFinished: "the check is over" - fires EVERY time, with how many moved (0 when
        //     there was nothing open to ask about, or when the server could not be reached).
        //
        // Retiring published drafts used to hang off onChanged, which is the wrong trigger for
        // work whose precondition is a STATE rather than a transition: a receipt stored as
        // "merged" by an earlier launch is never open again, so it never changes again, so
        // onChanged never fires for it - and a draft whose sync arrived one launch late was never
        // retired at all. Work like that belongs on onFinished.
        //
        // Neither fires when the check is cancelled - the app is closing.
        // ###########################################################################################
        public static async Task RefreshQuietlyAsync(
            StatusLookup lookup,
            DateTimeOffset now,
            Action? onChanged = null,
            CancellationToken cancellationToken = default,
            Action<int>? onFinished = null)
        {
            int changed = 0;

            try
            {
                if (SubmissionReceiptStore.All.Any(receipt =>
                        SubmissionReceiptPresenter.IsStillOpen(receipt.LastKnownState)))
                {
                    changed = await SubmissionStatusRefresh
                        .RefreshAsync(lookup, now, cancellationToken)
                        .ConfigureAwait(false);

                    if (changed > 0)
                    {
                        onChanged?.Invoke();
                    }
                }
            }
            catch (OperationCanceledException)
            {
                // The app is closing. Nothing to report, and nothing more to start.
                return;
            }
            catch (Exception ex)
            {
                CrtLog.Warning($"Submission status check at launch failed: [{ex.Message}]");
            }

            try
            {
                onFinished?.Invoke(changed);
            }
            catch (Exception ex)
            {
                // Same rule as the check itself: this runs unobserved, so nothing may escape.
                CrtLog.Warning($"Work after the submission status check failed: [{ex.Message}]");
            }
        }
    }
}
