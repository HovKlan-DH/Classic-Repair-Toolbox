using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A CONTRIBUTOR KEEPS AFTER SENDING A SUBMISSION (NewContributeStrategy.md Phase 4,
    // task 6) - the id and the capability token, so the app can ask the server how it went.
    //
    // *** THIS EXISTS BECAUSE CONTRIBUTING HAS NO ACCOUNT. *** With no sign-in there is nothing
    // to hang a "my submissions" list on: the server cannot answer "what did I send" because it
    // has no idea who is asking. The capability token returned once at creation is the ONLY proof
    // of ownership, and until now it lived in a local variable and was gone the moment the Submit
    // dialog closed - so a contributor could never learn anything more than whatever the closing
    // screen happened to say.
    //
    // A receipt is therefore exactly what a paper receipt is: the thing you keep so you can ask
    // about the transaction later. The server stores only the token's HASH, so a lost receipt
    // cannot be recovered by anyone, including the maintainer.
    //
    // WHAT THIS IS NOT. It is not an identity, not a login, and not a list the server maintains.
    // It is per-machine by construction, which is a real limitation and is stated to the user in
    // those words rather than papered over - see the view's own empty state. A reinstall or a
    // second computer starts empty.
    //
    // *** AND THAT LIMITATION HAS NO BACKSTOP TODAY (corrected 2026-09-23). *** This comment used
    // to say the contributor "still gets told the outcome by email, which is the channel that does
    // not depend on this file existing". NO SUCH EMAIL IS SENT. EmailTemplates carries four
    // account messages (verification, already-registered, password reset, password changed) and
    // nothing for a submission outcome; no submission flow injects IEmailSender at all. The
    // contact address is stored so a REVIEWER can reply by hand, and that is its only use.
    //
    // So this file IS the channel. Losing it loses the outcome - which is why the badge rules
    // below matter more than they look, and why the notification work in Phase 6 (item 11) is
    // still outstanding. Three user-visible strings promised that email and have been corrected.
    //
    // THE TOKEN IS A SECRET and this record carries it. It authorises reading a submission, so a
    // receipts file is worth the same care as any other credential store - it lives beside the
    // user's own settings, never in the synced Data tree and never inside a draft that could be
    // submitted. See SubmissionReceiptStore for where it actually goes.
    // ###########################################################################################
    public sealed class SubmissionReceipt
    {
        // The server's own id for the submission. Small and sequential, which is exactly why the
        // token below is needed - an id alone would let anyone walk the range.
        public long SubmissionId { get; init; }

        // The capability token, returned once by the server at creation and never again.
        public string UploadToken { get; init; } = string.Empty;

        // The system this was a submission for, as its ExcelDataFile identity - what the row is
        // labelled with, and what ties a receipt back to a draft still on this machine.
        public string SystemId { get; init; } = string.Empty;

        // What the contributor typed as their summary. Kept locally so the list reads as theirs
        // ("fixed U8 pinout") rather than as a row of ids, and so it still reads that way when the
        // server cannot be reached at all.
        public string Summary { get; init; } = string.Empty;

        // When it was sent, by this machine's clock. UTC, so a receipts file that travels between
        // time zones does not reorder itself.
        public DateTimeOffset SentUtc { get; init; }

        // ###########################################################################################
        // The last state the server reported, cached so the list renders immediately and still says
        // something useful with no connection. Empty until the first successful refresh.
        //
        // DELIBERATELY A CACHE AND NOT A SOURCE OF TRUTH: the server decides the state, this only
        // remembers what it last said. Anything reading this has to be willing to be out of date,
        // which is why the view shows when it was last checked rather than presenting the value as
        // current fact.
        // ###########################################################################################
        public string LastKnownState { get; init; } = string.Empty;

        public DateTimeOffset? LastCheckedUtc { get; init; }

        // What a reviewer said, once there is a reviewer to say it. Phase 5 builds the review
        // application; until then the server has no field for this and it stays empty. Carried now
        // so that adding it server-side does not need a format change on every contributor's disk.
        public string ReviewerComment { get; init; } = string.Empty;

        // A reviewer changed some of the rows in the review application before deciding
        // (2026-09-25) - SubmissionStatus.AmendedByReviewer, cached like the state.
        public bool AmendedByReviewer { get; init; }

        // ###########################################################################################
        // The reviewer comment this contributor has SEEN, verbatim.
        //
        // *** IT STORES THE TEXT, NOT A BOOLEAN. *** A "seen" flag would have to be cleared by
        // whatever writes a new comment, and the moment one writer forgets, a second round of
        // review feedback arrives already marked as read - which is the only failure mode that
        // actually matters here, because the contributor never finds out they were asked for
        // something. Comparing the text means "is this new" is DERIVED rather than maintained, so
        // there is no flag to forget to reset.
        //
        // Empty means nothing has been acknowledged yet - correct for every receipt written before
        // this field existed, whose comment is therefore treated as unread the first time one
        // arrives.
        //
        // See SubmissionReceiptPresenter.HasUnreadComment for the rule itself.
        // ###########################################################################################
        public string AcknowledgedComment { get; init; } = string.Empty;

        // ###########################################################################################
        // The DECIDED state this contributor has seen, verbatim - the twin of AcknowledgedComment
        // and stored for exactly the same reason (maintainer report, 2026-09-23).
        //
        // *** IT EXISTS BECAUSE A DECISION WITH NO COMMENT WAS INVISIBLE. *** HasUnreadComment
        // returns false the moment the comment is empty, and a reviewer approving a submission
        // usually types nothing at all - there is nothing to say about work being accepted. So the
        // state moved to "Published", the row said so, and NOTHING told the contributor to go and
        // look: no badge, and with the draft gone no Drafts tab either. Reported by the maintainer
        // on the first real publish.
        //
        // "A comment nobody notices is a comment nobody reads" is the rule HasUnreadComment was
        // written for; this is the same rule applied to the OUTCOME, which is the part the
        // contributor was actually waiting for.
        //
        // Text, not a boolean, for the identical reason: a flag would have to be cleared by
        // whatever writes the next state, and one forgetful writer means a later decision arrives
        // pre-dismissed. Comparing the stored value makes "is this new" DERIVED.
        //
        // Empty means nothing has been acknowledged. Correct for every receipt written before this
        // field existed - their first decision then reads as unread once, which is right: nobody
        // has confirmed seeing it.
        // ###########################################################################################
        public string AcknowledgedState { get; init; } = string.Empty;

        // ###########################################################################################
        // When the reviewer actually decided - the server's own `decidedUtc`, not when this machine
        // happened to find out.
        //
        // *** DISTINCT FROM LastCheckedUtc, AND THE DIFFERENCE IS THE POINT. *** "Last checked" is
        // this computer's bookkeeping and says nothing about the submission; a contributor reading
        // a request for changes wants to know WHEN IT WAS WRITTEN, because a comment from three
        // weeks ago on a board they have since revised means something quite different from one
        // written this morning.
        //
        // Null while nothing has been decided, which is every submission still waiting for review.
        // ###########################################################################################
        public DateTimeOffset? DecidedUtc { get; init; }
    }

    // ###########################################################################################
    // The four kinds of outcome a submission state can represent - see
    // SubmissionReceiptPresenter.ClassifyState.
    //
    // Deliberately NOT one value per server state: a reader scanning a list of submissions needs
    // to know whether a row is finished and good, finished and not, waiting, or waiting on THEM.
    // The exact state is already spelled out in words beside it.
    // ###########################################################################################
    public enum SubmissionOutcomeKind
    {
        // Sent, not decided. The ordinary resting state, and the default for anything unrecognised.
        Waiting,

        // Decided, and the contribution is in.
        Good,

        // Decided, and it is not going in - or the send never completed.
        Bad,

        // *** THE CONTRIBUTOR HAS TO DO SOMETHING. *** The only bucket that is about the reader
        // rather than the submission.
        NeedsAction
    }

    // ###########################################################################################
    // Turns a receipt plus its last known state into the words the list actually shows.
    //
    // PURE, so the vocabulary is unit tested rather than trusted. That matters more here than it
    // looks: these strings are the only explanation a contributor gets inside the app, and the
    // server's own state names are database values ("pending") that would be actively misleading
    // if shown raw - "pending" reads as "not sent yet" to someone who has just sent it.
    // ###########################################################################################
    public static class SubmissionReceiptPresenter
    {
        // ###########################################################################################
        // A one-line state description in the contributor's own terms.
        //
        // Each mapping exists because the raw value would mislead:
        //   - "uploading" means the send never completed, which to the user is a failed attempt,
        //     not work in progress - nothing is uploading any more;
        //   - "pending" means QUEUED, waiting for a person - not "pending" as in unfinished;
        //   - "abandoned" is the server's word for an upload window that expired, which sounds
        //     like the contributor gave up rather than that time ran out;
        //   - an unknown value is reported as unknown rather than guessed at, because a future
        //     server state rendered as something plausible-but-wrong is worse than an honest
        //     "the server said something this version does not recognise".
        // ###########################################################################################
        public static string DescribeState(string? state)
        {
            if (string.IsNullOrWhiteSpace(state))
                return "Not checked yet";

            return state.Trim().ToLowerInvariant() switch
            {
                "uploading" => "Never finished sending",
                "pending" => "Waiting for review",
                "accepted" => "Accepted",
                // *** "BETA source" AND "source" (maintainer wording, 2026-09-25). *** The two
                // stages are named for where the data went, in the words CRT's own Configuration
                // tab uses for the two places it downloads from ("online source", "BETA source").
                "published" => "Published to source",
                "rejected" => "Not accepted",
                "abandoned" => "Expired before it was finished",

                // *** THE FOUR PHASE 5 REVIEW STATES, MISSING UNTIL 2026-09-22. *** They were added
                // to the server's own vocabulary when the review application was built and never
                // taught to this method, so the first real review round trip showed a contributor
                // "Reported as [changes_requested]" - a raw database value, complete with its
                // underscore, in the one place this class exists to prevent exactly that.
                //
                // "changes_requested" is the one that mattered: it is the state where somebody is
                // being ASKED TO DO SOMETHING, and it read as a fault in the application.
                "changes_requested" => "Changes requested",
                "approved" => "Approved, waiting to be published",

                // *** "merged" IS THE BETA SOURCE, NOT EVERYONE'S (2026-09-25). *** Since the
                // two-stage publish, a reviewer's approval writes the BETA data; the board goes out
                // to everyone when it is published to production, and the server then reports this
                // same submission as "published" (ProductionPromotionRules.ContributorFacingState)
                // - which is the row that says "Published to source".
                "merged" => "Published to BETA source",
                "withdrawn" => "Withdrawn",

                // A state this build has never heard of. Reported honestly rather than guessed at -
                // a future server value rendered as something plausible-but-wrong is worse than an
                // admission that this version does not know it.
                _ => $"Reported as [{state.Trim()}]"
            };
        }

        // ###########################################################################################
        // WHAT KIND OF OUTCOME a state represents, for anything that needs to COLOUR it
        // (maintainer request, 2026-09-22).
        //
        // *** THE CLASSIFICATION IS HERE, NOT IN THE UI, so the colour cannot disagree with the
        // words. *** DescribeState already turns a state into a sentence; this turns the same state
        // into which of four buckets it falls in, off the same switch. A tab that decided colours
        // for itself would eventually paint "Changes requested" green after somebody edited one of
        // the two lists.
        //
        // Four buckets rather than one per state, because the reader is scanning a list and only
        // needs to know: is this finished and good, finished and not, waiting, or waiting on ME.
        // ###########################################################################################
        public static SubmissionOutcomeKind ClassifyState(string? state)
        {
            if (string.IsNullOrWhiteSpace(state))
                return SubmissionOutcomeKind.Waiting;

            return state.Trim().ToLowerInvariant() switch
            {
                // Finished, and the contribution is in. "accepted" and "approved" are not yet
                // published but are past the decision, which is what the contributor cares about.
                "published" or "merged" or "accepted" or "approved" => SubmissionOutcomeKind.Good,

                // *** NEEDS THE CONTRIBUTOR TO DO SOMETHING. *** The one bucket that is about the
                // reader rather than about the submission, and the reason this classification
                // exists at all - it is what the coloured edge is for.
                "changes_requested" => SubmissionOutcomeKind.NeedsAction,

                // Finished, and it is not going in. "uploading" belongs here rather than in
                // Waiting: nothing is uploading any more, the send failed partway, and the row is
                // as dead as a rejection until the contributor sends again.
                "rejected" or "abandoned" or "withdrawn" or "uploading" => SubmissionOutcomeKind.Bad,

                // Includes "pending" and anything this build has never heard of. Neutral is the
                // safe direction: colouring an unknown state as good or bad would state something
                // this version cannot actually know.
                _ => SubmissionOutcomeKind.Waiting
            };
        }

        // ###########################################################################################
        // Whether this submission is still waiting on somebody - which is what decides if there is
        // any point asking the server about it again.
        //
        // A decided submission never changes state again, so re-checking it forever is a request
        // per row per refresh that can only ever return the same answer.
        // ###########################################################################################
        public static bool IsStillOpen(string? state)
        {
            if (string.IsNullOrWhiteSpace(state))
                return true;

            return state.Trim().ToLowerInvariant() switch
            {
                // "withdrawn" was once missing from this list, which meant every such row was
                // re-checked on every launch forever. Final states are asked about no more.
                "rejected" or "abandoned" or "accepted" or "published" or "withdrawn" => false,

                // *** "merged" IS OPEN AGAIN (2026-09-25), deliberately. *** It used to be final.
                // Since the two-stage publish it means "in the BETA data", and the server moves it
                // on to "published" once the board reaches production - so it has to be asked about
                // until then, or the row would say "in the BETA data" for ever. One request per
                // merged row per launch, for the few days between the two publishes - and no
                // longer than MergedRecheckWindow: see the receipt overload below.

                // *** "changes_requested" AND "approved" ARE DELIBERATELY STILL OPEN. *** Neither
                // is the end: a submission with changes requested can be re-reviewed after the
                // contributor acts, and an approved one is still waiting to be published. Treating
                // either as decided would freeze the row at that state and the contributor would
                // never see it move.
                _ => true
            };
        }

        // ###########################################################################################
        // How long after its decision a "merged" (in BETA) submission is still asked about.
        //
        // *** A BOUND, BECAUSE THE SECOND PUBLISH MAY NEVER COME (code review, 2026-09-25). ***
        // Publishing to production is off until the server is set up for it, and a reviewer may
        // never promote a board. Without a bound, every merged receipt was asked about on every
        // launch for ever - the very re-check "withdrawn" once caused. Thirty days is far longer
        // than BETA to production is meant to take; after it the row keeps its last answer.
        // ###########################################################################################
        public static readonly TimeSpan MergedRecheckWindow = TimeSpan.FromDays(30);

        // ###########################################################################################
        // Whether this RECEIPT is worth asking the server about AT LAUNCH: its state is still open,
        // and a "merged" one was decided less than MergedRecheckWindow ago (or sent, for a receipt
        // that never recorded its decision date). "Check for updates" in My submissions uses the
        // state alone - somebody pressed it to ask.
        // ###########################################################################################
        public static bool IsStillOpen(SubmissionReceipt receipt, DateTimeOffset nowUtc)
        {
            ArgumentNullException.ThrowIfNull(receipt);

            if (!SubmissionReceiptPresenter.IsStillOpen(receipt.LastKnownState))
                return false;

            if (!string.Equals(receipt.LastKnownState?.Trim(), "merged", StringComparison.OrdinalIgnoreCase))
                return true;

            DateTimeOffset since = receipt.DecidedUtc ?? receipt.SentUtc;

            return nowUtc - since < SubmissionReceiptPresenter.MergedRecheckWindow;
        }

        // ###########################################################################################
        // THE ONE DATE FORMAT EVERY SUBMISSION LINE USES: "2026-September-23" (maintainer request,
        // 2026-09-23).
        //
        // *** YEAR FIRST, MONTH BY NAME, DAY WITHOUT A LEADING ZERO. *** The shape was chosen to be
        // unambiguous on sight: "09-10-2026" is the tenth of September to half the world and the
        // ninth of October to the other half, and these dates are read beside one another on a list
        // where the difference decides whether a reviewer's comment is stale.
        //
        // *** INVARIANT CULTURE, DELIBERATELY. *** The month name is the one part of this that a
        // culture can change, and it did: on the maintainer's own Danish machine the previous
        // format rendered "22 september 2026" - lower case, because Danish does not capitalise
        // month names. Formatting with CultureInfo.CurrentCulture would give a different string on
        // every machine while the format string looked identical, so the format the maintainer
        // asked for would only actually appear in English locales.
        //
        // The "d" (not "dd") is what drops the leading zero, and it is the reason this cannot just
        // be a round-trip format like "yyyy-MM-dd".
        // ###########################################################################################
        public static string FormatDate(DateTimeOffset value)
        {
            return value.ToLocalTime().ToString("yyyy-MMMM-d", CultureInfo.InvariantCulture);
        }

        // ###########################################################################################
        // "Sent 2026-September-21" - the date only, never a time.
        //
        // A time would imply a precision this does not have: the clock is the contributor's own,
        // and a submission sent at 23:58 in one time zone is a different day in another. The day is
        // what someone actually recalls ("I sent that last week").
        // ###########################################################################################
        public static string DescribeSent(DateTimeOffset sentUtc)
        {
            return "Sent " + SubmissionReceiptPresenter.FormatDate(sentUtc);
        }

        // ###########################################################################################
        // "Replied 2026-September-22" - when the REVIEWER decided, not when this machine noticed.
        //
        // Empty when nothing has been decided, so the caller can leave the line out entirely rather
        // than print a label with nothing after it.
        //
        // Same date-only rule as DescribeSent, and for the same reason: the clocks belong to
        // different people in different time zones, so a time would imply a precision that is not
        // there. See that method's header.
        // ###########################################################################################
        // ###########################################################################################
        // The line "My submissions" shows when a reviewer changed the submission before deciding it
        // (2026-09-25), so a contributor comparing what was published with what they sent knows
        // where a difference came from. Empty when nobody changed it.
        // ###########################################################################################
        public static string DescribeAmended(bool amendedByReviewer) =>
            amendedByReviewer
                ? "A reviewer changed some of the details before deciding, so what is published is not exactly what you sent."
                : string.Empty;

        public static string DescribeDecided(DateTimeOffset? decidedUtc)
        {
            if (decidedUtc is null)
                return string.Empty;

            return "Replied " + SubmissionReceiptPresenter.FormatDate(decidedUtc.Value);
        }

        // ###########################################################################################
        // "Last checked 2026-September-23" - when THIS COMPUTER last asked the server, which is
        // bookkeeping and nothing more.
        //
        // *** IT SAYS NOTHING ABOUT THE SUBMISSION, and that is why it is phrased and coloured as
        // the quietest thing on the card. *** The maintainer asked what it meant, which is fair
        // warning: beside "Sent" and "Replied", both of which are facts about the contribution, a
        // third date invites the reader to look for meaning that is not there.
        //
        // It is kept because it is the only thing that distinguishes "the server says this is still
        // waiting" from "nobody has been able to ask since Tuesday" - without it, a stale row and a
        // current one look identical.
        //
        // Empty when this receipt has never been checked, so the caller drops the line entirely.
        // ###########################################################################################
        public static string DescribeLastChecked(DateTimeOffset? lastCheckedUtc)
        {
            if (lastCheckedUtc is null)
                return string.Empty;

            return "Last checked " + SubmissionReceiptPresenter.FormatDate(lastCheckedUtc.Value);
        }

        // ###########################################################################################
        // Newest first - the thing just sent is the thing being asked about.
        // ###########################################################################################
        public static IReadOnlyList<SubmissionReceipt> InDisplayOrder(IEnumerable<SubmissionReceipt>? receipts)
        {
            if (receipts is null)
                return [];

            return receipts
                .OrderByDescending(receipt => receipt.SentUtc)
                .ThenByDescending(receipt => receipt.SubmissionId)
                .ToList();
        }

        // ###########################################################################################
        // Does this receipt carry a reviewer comment the contributor has NOT yet acknowledged?
        //
        // *** THE WHOLE POINT: A COMMENT NOBODY NOTICES IS A COMMENT NOBODY READS. *** Contributing
        // needs no account, so there is no inbox and no thread - the reviewer's sentence is the
        // entire channel back to the person who did the work. Before this, it appeared only inside
        // a window reached by a button nobody has a reason to press, so being asked for changes was
        // invisible unless the contributor happened to go looking.
        //
        // *** COMPARED AS TEXT, so a SECOND round of feedback is unread again. *** This is the
        // case a simple "seen" boolean gets wrong: the contributor reads "fix U8", fixes it,
        // resubmits, and is asked for something else. With a flag, that second comment arrives
        // already marked read and is never seen. Comparing the acknowledged text against the
        // current one makes a CHANGED comment unread by construction.
        //
        // ORDINAL comparison, not culture-aware: this is an exact-match question about two copies
        // of the same stored string, and a culture-sensitive comparison could call two different
        // sentences equal.
        //
        // Trimmed on both sides so that whitespace the server may add or drop around the text
        // cannot resurface a comment the contributor has already dealt with.
        // ###########################################################################################
        public static bool HasUnreadComment(SubmissionReceipt? receipt)
        {
            if (receipt is null)
                return false;

            string comment = (receipt.ReviewerComment ?? string.Empty).Trim();

            // Nothing said. Not unread - there is nothing to read.
            if (comment.Length == 0)
                return false;

            return !string.Equals(
                comment,
                (receipt.AcknowledgedComment ?? string.Empty).Trim(),
                StringComparison.Ordinal);
        }

        // ###########################################################################################
        // Has this submission been DECIDED since the contributor last looked?
        //
        // *** THE HALF THAT WAS MISSING, AND THE COMMONEST CASE OF ALL (maintainer report,
        // 2026-09-23). *** A reviewer approving good work usually writes nothing - there is
        // nothing to say - so HasUnreadComment saw an empty comment and reported "not unread".
        // The contributor's submission went to "Published" and the application said nothing:
        // no badge, and once the draft was gone, no Drafts tab either. The one outcome everybody
        // is actually waiting for was the one outcome that arrived silently.
        //
        // *** ONLY A DECIDED STATE COUNTS. *** "pending" is not news - it is the state a
        // submission is in from the moment it is sent, and treating it as unread would badge
        // every contribution the instant it left. What is worth interrupting somebody for is the
        // transition to an outcome: published, rejected, changes requested, and the rest.
        //
        // The comparison is against the state the contributor has SEEN, so a submission that goes
        // approved -> merged raises the badge a second time. That is correct: they are different
        // facts, and the second one ("it is actually live now") is the one they were waiting for.
        // ###########################################################################################
        public static bool HasUnreadDecision(SubmissionReceipt? receipt)
        {
            if (receipt is null)
                return false;

            string state = (receipt.LastKnownState ?? string.Empty).Trim();

            // Never checked, or still queued. Nothing has happened to report.
            if (state.Length == 0 || SubmissionReceiptPresenter.IsStillOpenForBadge(state))
                return false;

            return !string.Equals(
                state,
                (receipt.AcknowledgedState ?? string.Empty).Trim(),
                StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // Whether a state is "nothing has happened yet" as far as the BADGE is concerned.
        //
        // *** DELIBERATELY NOT IsStillOpen, and the difference is the whole point. *** IsStillOpen
        // answers "is it worth asking the server again", and it counts "changes_requested" and
        // "approved" as still open because neither is the end of the road. But both are absolutely
        // news to the contributor - one is a request aimed straight at them - so reusing that rule
        // here would suppress the badge for the two states that most need it.
        //
        // Only the states where no human has yet ruled on the submission are silent.
        // ###########################################################################################
        private static bool IsStillOpenForBadge(string state)
        {
            return state.Trim().ToLowerInvariant() switch
            {
                // Queued for a reviewer, or still mid-send. Neither is a decision.
                "pending" or "uploading" => true,

                // Anything decided - and anything this build does not recognise, which is safer
                // reported than swallowed: an unknown state arriving on a receipt means the server
                // did something, and telling the contributor to look is better than hiding it.
                _ => false
            };
        }

        // ###########################################################################################
        // Is there anything on this receipt the contributor has not seen - a comment, a decision,
        // or both?
        //
        // The single question every badge and every "needs attention" surface should ask, so a new
        // kind of news cannot be added to one surface and forgotten on another.
        // ###########################################################################################
        public static bool HasUnreadNews(SubmissionReceipt? receipt)
        {
            return SubmissionReceiptPresenter.HasUnreadComment(receipt)
                || SubmissionReceiptPresenter.HasUnreadDecision(receipt);
        }

        // ###########################################################################################
        // How many receipts carry news the contributor has not seen - the number in the badge.
        //
        // Counting RECEIPTS rather than items of news, because one submission has at most one
        // current comment and one current state: the server keeps the latest decision, not a
        // thread. "(2)" therefore means two different submissions want attention, which is what
        // the contributor needs to know. A row carrying both a new decision AND a new comment is
        // still one row to go and look at.
        //
        // *** IT COUNTS DECISIONS AS WELL AS COMMENTS SINCE 2026-09-23. *** The name is kept
        // because it is what every caller already asks for and the question is unchanged - "how
        // many rows should I be told about" - but a silent approval used to count zero. See
        // HasUnreadDecision.
        // ###########################################################################################
        public static int UnreadCommentCount(IEnumerable<SubmissionReceipt>? receipts)
        {
            if (receipts is null)
                return 0;

            return receipts.Count(SubmissionReceiptPresenter.HasUnreadNews);
        }
    }
}
