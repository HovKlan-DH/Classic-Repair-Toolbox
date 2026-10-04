using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT HAS HAPPENED TO A SYSTEM, newest first, for the Maintainer tab's Systems screen
    // (owner request, 2026-09-27: "For a system in 'Systems' menu - I would like to see the date,
    // newest first, to understand what has happened to a system").
    //
    // The server sends FACTS - when, which kind of event, who did it, which submission, and the
    // event's own detail - and the Maintainer tab puts them into words (SystemsDisplay), the split every
    // other answer on the Systems screen keeps.
    //
    //   Event        one of SystemHistoryEvents - an unknown one (a newer server) is shown as itself.
    //   Who          the person who did it: the contributor who sent a submission, the maintainer who
    //                decided or changed one, the administrator who changed a pool.
    //   SubmissionId the submission it concerns, when it concerns one.
    //   Detail       the event's own words: a submission's description, the state a decision left
    //                it in (the DATABASE word - "merged" is "into BETA", whatever it became after),
    //                the address a pool change named, what a promotion or push-back did.
    //   Note         what a maintainer told the contributor with a decision.
    // ###########################################################################################
    public sealed record SystemHistoryEntry(
        DateTimeOffset AtUtc,
        string Event,
        string? Who,
        long? SubmissionId,
        string? Detail,
        string? Note = null);

    // The kinds of event, as SystemHistoryEntry.Event carries them. Most are the audit trail's own
    // action words, so an entry and the audit row it came from name the event alike.
    public static class SystemHistoryEvents
    {
        public const string Sent = "submission.sent";
        public const string Decided = "submission.decided";
        public const string Amended = "submission.amended";
        public const string PublishedToProduction = "production.published";

        // The server found production already holding the system's BETA state - copied there
        // outside CRT - and recorded it as in production (code review, 2026-09-29).
        public const string FoundInProduction = "production.found_in_place";
        public const string PushedBack = "beta.rolledback";

        // Beta > Prod's "Reject" (owner request, 2026-09-28): the same rollback, its submissions
        // rejected instead of returned to the queue - see BetaRollbackFlow.
        public const string RejectedFromBeta = "beta.rejected";
        public const string MaintainerAdded = "maintainer.granted";
        public const string MaintainerRemoved = "maintainer.revoked";
        public const string Invited = "maintainer.invited";
        public const string InvitationWithdrawn = "maintainer.invitation_withdrawn";
        public const string InvitationAccepted = "maintainer.invitation_accepted";
        public const string Placed = "system.placed";

        // The contributor discarded their own draft in CRT after sending the submission (owner
        // request, 2026-09-28) - see DraftDiscardContract. Audited under "#{id}".
        public const string DraftDiscarded = "submission.draft_discarded";

        // The administrator deleted the system from both data trees and the database (owner
        // request, 2026-10-03 - see CRT.Server's SystemDeletionFlow). Audited under the system id,
        // so a system created again under the same id shows it at the start of its history.
        public const string Deleted = "system.deleted";
    }

    // ###########################################################################################
    // *** ONE SUBMISSION IN BETA PER SYSTEM (owner decision, 2026-09-27: "it should be possible only
    // to submit ONE contributor submission to Beta per system. It can quickly become complex, if
    // this gets mixed up, having many contributions grouped and then reverting from Beta to queue
    // again"). ***
    //
    // A push-back can only ever be per SYSTEM - a publish replaces the board's rows wholesale, so one
    // contributor's work cannot be picked back out of several - so while a system has a submission
    // in BETA that has not gone to production, no other submission of it may be approved. The
    // server refuses it (ApprovePublishFlow) and the Maintainer tab turns Approve off beside the same
    // sentence, so the two say it identically.
    // ###########################################################################################
    public static class OneSubmissionInBeta
    {
        public static string BusyMessage(string? systemId) =>
            $"{(string.IsNullOrWhiteSpace(systemId) ? "This system" : systemId.Trim())} already has a submission in BETA that " +
            "has not gone to the stable source. A system takes one submission into BETA at a time: publish that one " +
            "to stable (under " + MaintainerScreenWording.BetaQueueQuoted + ") or push it back to the queue first.";

        // ###########################################################################################
        // The same rule for a change made on the Systems screen, which goes straight to BETA (owner
        // decision, 2026-10-03: "If a system is already in 'BETA > Stable' queue, then it should
        // simply disallow it, even if this is coming from a maintainer"). The server sends it as the
        // reason the table is read-only, and refuses an edit with it.
        // ###########################################################################################
        public static string NoChangeMessage(string? systemId) =>
            $"{(string.IsNullOrWhiteSpace(systemId) ? "This system" : systemId.Trim())} is waiting in BETA for the stable source, " +
            "so no change can be made to it here. Publish it to stable, push it back or reject it under " + MaintainerScreenWording.BetaQueueQuoted + " first.";
    }
}
