using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // "THE CONTRIBUTOR DISCARDED THEIR OWN DRAFT" (owner request, 2026-09-28): "if the contributor
    // has discarded his own data, I think this needs to be validated and by sure it should be
    // clearly visible on the system and for the maintainer(s) - both in the BETA to PROD queue, but
    // also in the normal queue. In this scenario the maintainer most likely should push system back
    // to normal queue, and clarify with contributor."
    //
    // Discarding a draft in CRT is local - it deletes the contributor's own working copy and never
    // withdraws what they sent. But a maintainer about to publish that work to every CRT user should
    // know that the person who made it threw their own copy away; it may mean they changed their
    // mind. So CRT tells the server, once per live submission of that board, and the server shows it
    // beside the submission in the review queue, on "Beta > Prod" and on the Systems screen.
    //
    // THE ROUTE: POST {api}/submissions/{id}/draft-discarded with the submission's capability token
    // in X-Submission-Token, exactly as the contributor's status check proves ownership. No body -
    // the server records its own time of arrival. It answers 204, or 404 for a submission it does
    // not have or a token that does not match (indistinguishable, as on every contributor route).
    //
    // Shared by CRT (which sends it) and CRT.Server (which maps it), so the path is written once.
    // ###########################################################################################
    public static class DraftDiscardContract
    {
        // The last segment of the route, after "submissions/{id}/".
        public const string RouteSegment = "draft-discarded";

        // The path under the API base both applications use ("{base}/submissions/42/draft-discarded").
        public static string PathUnderApi(long submissionId) =>
            FormattableString.Invariant($"submissions/{submissionId}/{DraftDiscardContract.RouteSegment}");

        // ###########################################################################################
        // Which of the contributor's submissions of the discarded board are worth telling the server
        // about: every one a maintainer may still act on. That is exactly the set CRT keeps asking
        // the server about - SubmissionReceiptPresenter.IsStillOpen - so the two cannot disagree:
        // waiting, approved, changes requested, in BETA, taken back out of BETA, or not checked
        // yet. A final state (published, not accepted, replaced, expired) is past acting on.
        // ###########################################################################################
        public static bool IsWorthReporting(string? lastKnownState) =>
            SubmissionReceiptPresenter.IsStillOpen(lastKnownState);

        // ###########################################################################################
        // The receipts a discard of this system's draft is reported for: sent FROM THIS DRAFT (at or
        // after draftCreatedUtc, the marker's - an older submission belongs to an earlier draft that
        // was retired, and its board has moved on), still worth reporting, and not reported already.
        // Matched on the system id exactly as LatestForSystem matches the Drafts tab's badge, so the
        // dialog warns about precisely the submissions the row shows.
        // ###########################################################################################
        public static IReadOnlyList<SubmissionReceipt> WhichToReport(
            IEnumerable<SubmissionReceipt>? receipts,
            string? systemId,
            DateTimeOffset? draftCreatedUtc = null)
        {
            string id = systemId?.Trim() ?? string.Empty;

            if (id.Length == 0)
                return [];

            return (receipts ?? [])
                .Where(receipt => string.Equals(receipt.SystemId?.Trim(), id, StringComparison.OrdinalIgnoreCase))
                .Where(receipt => draftCreatedUtc is null || receipt.SentUtc >= draftCreatedUtc.Value)
                .Where(receipt => DraftDiscardContract.IsWorthReporting(receipt.LastKnownState))
                .Where(receipt => receipt.DraftDiscardedUtc is null)
                .OrderBy(receipt => receipt.SubmissionId)
                .ToList();
        }

        // ###########################################################################################
        // What the server's answer means for the notice. Done on 2xx, and on 404 - the submission is
        // gone or the token no longer matches, and asking again can never succeed - and on a refusal
        // of the request itself. Anything else (no answer, a 5xx, a rate limit) is tried again on the
        // next launch, since a notice is worth delivering late rather than never.
        // ###########################################################################################
        public static DraftDiscardDelivery DeliveryFor(int? status) =>
            status switch
            {
                >= 200 and < 300 => DraftDiscardDelivery.Done,
                400 or 404 or 405 or 410 or 413 or 415 or 422 => DraftDiscardDelivery.Done,
                _ => DraftDiscardDelivery.TryLater
            };
    }

    public enum DraftDiscardDelivery
    {
        Done,
        TryLater
    }
}
