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
    // into the local receipts (owner report, 2026-09-22).
    //
    // *** WHY THIS EXISTS AS ITS OWN CLASS: THE CHECK ONLY EVER RAN FROM A BUTTON. *** Receipts
    // were loaded from disk at startup, but nothing asked the server until somebody opened "My
    // submissions" and pressed Refresh. So a maintainer could request changes and the contributor
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
    // last known state, exactly as it would if the check had not run. The ONE exception is the
    // server answering "update CRT" (2026-10-04): that is handed back to be shown, since nothing
    // would otherwise ever say why the receipts stopped moving.
    // ###########################################################################################
    public static class SubmissionStatusRefresh
    {
        // ###########################################################################################
        // *** AND AGAIN WHILE CRT RUNS, FOR THE DRAFTS TAB'S BADGE (owner request, 2026-09-30: "Like
        // the "Maintainer" tab now has a badge, then I think the "Draft" also should have same
        // functionality, as the contributor likewise can receive an update"). ***
        //
        // Every PeriodicInterval, the same question as at launch - only about what is still open.
        //
        // *** ONE MINUTE, the Maintainer tab's own interval (owner decision, 2026-10-01). *** It was
        // five, on the reasoning that a review takes hours and this runs in every contributor's CRT
        // rather than in a handful of maintainers'. The project owner weighed that against the cost
        // and chose one minute anyway - "this is not a high volume thing ... not that many users
        // globally" - because the check is what retires a published draft, and a contributor
        // watching their draft disappear should not wait five minutes for it. ChecksPeriodically
        // still means nobody without an open submission asks anything at all.
        // ###########################################################################################
        public static readonly TimeSpan PeriodicInterval = TimeSpan.FromMinutes(1);

        // Whether a periodic check asks anything: not while CRT's window is minimised (no tab badge to
        // see), and not when nothing sent is still open - a contributor with nothing waiting, and
        // everybody who has never contributed, sends nothing at all.
        public static bool ChecksPeriodically(IEnumerable<SubmissionReceipt> receipts, DateTimeOffset now, bool windowMinimised)
        {
            ArgumentNullException.ThrowIfNull(receipts);

            return !windowMinimised && receipts.Any(receipt => SubmissionReceiptPresenter.IsStillOpen(receipt, now));
        }

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
        // cause a refresh of anything. A receipt the server stopped knowing, or knows again, is a
        // change too: its row now reads "No longer on the server", or no longer does.
        //
        // *** "UPDATE CRT" STOPS THE ROUND (code review, 2026-10-04). *** The server answered every
        // submission the same way, so the rest are not asked; `onOutdated` is handed the server's
        // own sentence, for the caller to show - the one failure here that is NOT silent, because
        // nothing else would ever tell the contributor why their receipts stopped moving.
        // ###########################################################################################
        public static async Task<int> RefreshAsync(
            StatusLookup lookup,
            DateTimeOffset now,
            CancellationToken cancellationToken = default,
            Action<string>? onOutdated = null)
        {
            ArgumentNullException.ThrowIfNull(lookup);

            List<SubmissionReceipt> toCheck = SubmissionReceiptStore.All
                .Where(receipt => SubmissionReceiptPresenter.IsStillOpen(receipt, now))
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
                catch (SubmissionNotFoundException)
                {
                    // The server does not know it any more - deleted with its board, for one. It
                    // was asked about every minute for ever; now once a day (code review, 2026-10-04).
                    // The first time is a change: the row now says so, and must be drawn again.
                    if (receipt.NotFoundUtc is null)
                        changed++;

                    SubmissionReceiptStore.NoteNotFound(receipt.SubmissionId, now);
                    continue;
                }
                catch (ClientOutdatedException ex)
                {
                    CrtLog.Warning($"The server asks for a newer CRT before it answers about submissions: [{ex.Message}]");
                    onOutdated?.Invoke(ex.Message);
                    break;
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
                    (status.MaintainerComment ?? string.Empty).Trim(),
                    (receipt.MaintainerComment ?? string.Empty).Trim(),
                    StringComparison.Ordinal);

                // *** AN UNCHANGED ANSWER WRITES NOTHING, OR NEARLY (code review, 2026-10-04). ***
                // The minute check saved the whole receipts file once per open receipt, only to move
                // its "last checked" date. Everything the row shows moved, or a "not found" to clear,
                // is written at once; otherwise only a stale date is (NoteChecked).
                bool knownAgain = receipt.NotFoundUtc is not null;

                bool rowMoved = stateMoved || commentMoved ||
                    (status.DecidedUtc is not null && status.DecidedUtc != receipt.DecidedUtc) ||
                    status.AmendedByMaintainer != receipt.AmendedByMaintainer ||
                    knownAgain;

                if (rowMoved)
                {
                    SubmissionReceiptStore.UpdateState(
                        receipt.SubmissionId,
                        status.State ?? string.Empty,
                        status.MaintainerComment ?? string.Empty,
                        now,
                        status.DecidedUtc,
                        status.AmendedByMaintainer);
                }
                else
                {
                    SubmissionReceiptStore.NoteChecked(receipt.SubmissionId, now);
                }

                // Known again after "not found" is a change too: the row stops saying so.
                if (stateMoved || commentMoved || knownAgain)
                {
                    changed++;
                }
            }

            return changed;
        }

        // The sentence CRT shows when the server answered "update CRT": the server's own words, said
        // to be about the submissions' status, so it is clear what could not be done.
        public static string DescribeOutdated(string serversWords) =>
            "The status of your submissions could not be checked: " + serversWords.Trim();

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
        // work whose precondition is a STATE rather than a transition: a receipt that reached its
        // last state ("published") by an earlier launch is never open again, so it never changes
        // again, so onChanged never fires for it - and a draft whose sync arrived one launch late
        // was never retired at all. Work like that belongs on onFinished.
        //
        // Neither fires when the check is cancelled - the app is closing.
        //
        // onOutdated: the server answered "update CRT" - its sentence, see RefreshAsync. Called on
        // the thread the check ran on, before onFinished.
        // ###########################################################################################
        public static async Task RefreshQuietlyAsync(
            StatusLookup lookup,
            DateTimeOffset now,
            Action? onChanged = null,
            CancellationToken cancellationToken = default,
            Action<int>? onFinished = null,
            Action<string>? onOutdated = null)
        {
            int changed = 0;

            try
            {
                if (SubmissionReceiptStore.All.Any(receipt =>
                        SubmissionReceiptPresenter.IsStillOpen(receipt, now)))
                {
                    changed = await SubmissionStatusRefresh
                        .RefreshAsync(lookup, now, cancellationToken, onOutdated)
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
                // Not "at launch": the Drafts tab's minute check runs this too.
                CrtLog.Warning($"Submission status check failed: [{ex.Message}]");
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
