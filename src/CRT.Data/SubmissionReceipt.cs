using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json.Serialization;

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
    // cannot be recovered by anyone, including the project owner.
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
    // contact address is stored so a MAINTAINER can reply by hand, and that is its only use.
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
    //
    // *** A RECORD, SO A CHANGE IS `existing with { ... }` (code review, 2026-09-29). *** It was a
    // class, and SubmissionReceiptStore rebuilt it field by field in four places - each of which had
    // to be edited whenever a field was added, and a field left out of one was silently ERASED by
    // that change (the defect class that dropped a workbook's caption three times). `with` carries
    // every field it does not name. ToString is overridden so the token never reaches a log.
    public sealed record SubmissionReceipt
    {
        // The server's own id for the submission. Small and sequential, which is exactly why the
        // token below is needed - an id alone would let anyone walk the range.
        public long SubmissionId { get; init; }

        // The capability token, returned once by the server at creation and never again.
        public string UploadToken { get; init; } = string.Empty;

        // The board this was a submission for, as its ExcelDataFile identity - what the row is
        // labelled with, and what ties a receipt back to a draft still on this machine.
        //
        // Written as "SystemId", the name it had before "system" became "board" everywhere (owner
        // decision, 2026-10-09): every receipt already on a contributor's machine carries it.
        [JsonPropertyName("SystemId")]
        public string BoardId { get; init; } = string.Empty;

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

        // When the server last answered that it does not know this submission (HTTP 404 - deleted
        // with its board, for one), or null while it does (code review, 2026-10-04). Such a receipt
        // is asked about only once per NotFoundRecheckInterval rather than every minute for ever;
        // any later answer clears it. "My submissions" keeps showing its last known state.
        public DateTimeOffset? NotFoundUtc { get; init; }

        // What a maintainer said, once there is a maintainer to say it. Phase 5 builds the review
        // application; until then the server has no field for this and it stays empty. Carried now
        // so that adding it server-side does not need a format change on every contributor's disk.
        public string MaintainerComment { get; init; } = string.Empty;

        // A maintainer changed some of the rows in the Maintainer tab before deciding
        // (2026-09-25) - SubmissionStatus.AmendedByMaintainer, cached like the state.
        public bool AmendedByMaintainer { get; init; }

        // ###########################################################################################
        // *** THE TWO NAMES A RECEIPT WAS WRITTEN WITH BEFORE THE RENAME - READ, NEVER WRITTEN (code
        // review, 2026-09-25). *** The role was "reviewer" until then, and an earlier build saved
        // ReviewerComment and AmendedByReviewer. A decided submission is never asked about again, so
        // without these its maintainer's comment would be gone from "My submissions" for good. Each
        // only fills the new field; the getters answer null, which WhenWritingNull leaves out, so the
        // next save writes the new names alone.
        // ###########################################################################################
        [JsonInclude]
        [JsonPropertyName("ReviewerComment")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        private string? LegacyReviewerComment
        {
            get => null;
            init
            {
                if (!string.IsNullOrEmpty(value) && string.IsNullOrEmpty(this.MaintainerComment))
                    this.MaintainerComment = value;
            }
        }

        [JsonInclude]
        [JsonPropertyName("AmendedByReviewer")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        private bool? LegacyAmendedByReviewer
        {
            get => null;
            init
            {
                if (value == true)
                    this.AmendedByMaintainer = true;
            }
        }

        // ###########################################################################################
        // The maintainer comment this contributor has SEEN, verbatim.
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
        // and stored for exactly the same reason (owner report, 2026-09-23).
        //
        // *** IT EXISTS BECAUSE A DECISION WITH NO COMMENT WAS INVISIBLE. *** HasUnreadComment
        // returns false the moment the comment is empty, and a maintainer approving a submission
        // usually types nothing at all - there is nothing to say about work being accepted. So the
        // state moved to "Published", the row said so, and NOTHING told the contributor to go and
        // look: no badge, and with the draft gone no Drafts tab either. Reported by the project owner
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
        // When the maintainer actually decided - the server's own `decidedUtc`, not when this machine
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

        // ###########################################################################################
        // Whether the contributor has closed the "now in the online source - switch back from
        // BETA" notice for this submission (owner request, 2026-09-27). Set only by dismissing it,
        // and only meaningful once the submission is "published" - which never changes again, so
        // unlike the two Acknowledged texts above a plain flag cannot go stale.
        // ###########################################################################################
        public bool SourceNoticeDismissed { get; init; }

        // ###########################################################################################
        // Whether the contributor has closed the "now in the BETA source - tick BETA to try it"
        // notice for this submission (owner request, 2026-10-03). Set only by dismissing it; a later
        // submission has its own.
        // ###########################################################################################
        public bool BetaNoticeDismissed { get; init; }

        // ###########################################################################################
        // THE CONTRIBUTOR DISCARDED THEIR DRAFT of this board after sending this (owner request,
        // 2026-09-28) - see DraftDiscardContract. When, by this machine's clock, and whether the
        // server has been told. Kept until it has, so a discard made offline is reported at the
        // next launch rather than lost.
        // ###########################################################################################
        public DateTimeOffset? DraftDiscardedUtc { get; init; }

        public bool DraftDiscardReported { get; init; }

        // ###########################################################################################
        // *** WHAT THE DRAFT HELD WHEN THIS WAS SENT (owner request, 2026-10-03: "the submitted is
        // identical to what is in draft now. It should not be possible to submit the same data
        // again"). *** DraftFingerprint.Compute of the draft at the moment of sending. While the
        // draft still gives the same fingerprint, its Submit button is greyed out - see
        // SubmissionReceiptPresenter.IsAlreadySent. Empty on a receipt written before this existed,
        // which therefore never greys anything out.
        // ###########################################################################################
        public string DraftFingerprint { get; init; } = string.Empty;

        // Never the generated ToString, which would print UploadToken - a secret - into any log line
        // that formats a receipt.
        public override string ToString() => $"Submission #{this.SubmissionId} ({this.BoardId})";
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
        //   - "pending" means QUEUED, waiting for a person - not "pending" as in unfinished. It
        //     says "Submitted - awaiting feedback from a maintainer" (owner wording, 2026-09-27):
        //     it is the first thing a contributor sees after sending, and it names who acts next;
        //   - "abandoned" is the server's word for an upload window that expired, which sounds
        //     like the contributor gave up rather than that time ran out;
        //   - an unknown value is reported as unknown rather than guessed at, because a future
        //     server state rendered as something plausible-but-wrong is worse than an honest
        //     "the server said something this version does not recognise".
        // ###########################################################################################
        public const string PendingWording = "Submitted - awaiting feedback from a maintainer";

        public static string DescribeState(string? state)
        {
            if (string.IsNullOrWhiteSpace(state))
                return "Not checked yet";

            return state.Trim().ToLowerInvariant() switch
            {
                "uploading" => "Never finished sending",
                "pending" => SubmissionReceiptPresenter.PendingWording,
                "accepted" => "Accepted",
                // *** "the stable source" AND "the BETA source" (owner decision, 2026-10-01). ***
                // The two stages are named for where the data went, in the words CRT's own
                // Configuration tab uses for the two places it downloads from. It was "Published to
                // source" against "Published to the BETA source", which made the finished state the
                // nameless one - a contributor could not tell which of the two it meant. See
                // AppConfig.GetOnlineSourceLabel for why "stable" rather than "production".
                "published" => "Published to the stable source",
                "rejected" => "Not accepted",
                "abandoned" => "Expired before it was finished",

                // *** THE FOUR PHASE 5 REVIEW STATES, MISSING UNTIL 2026-09-22. *** They were added
                // to the server's own vocabulary when the Maintainer tab was built and never
                // taught to this method, so the first real review round trip showed a contributor
                // "Reported as [changes_requested]" - a raw database value, complete with its
                // underscore, in the one place this class exists to prevent exactly that.
                //
                // "changes_requested" is the one that mattered: it is the state where somebody is
                // being ASKED TO DO SOMETHING, and it read as a fault in the application.
                "changes_requested" => "Changes requested",
                "approved" => "Approved, waiting to be published",

                // *** "merged" IS THE BETA SOURCE, NOT EVERYONE'S (2026-09-25). *** Since the
                // two-stage publish, a maintainer's approval writes the BETA data; the board goes out
                // to everyone when it is published to production, and the server then reports this
                // same submission as "published" (ProductionPromotionRules.ContributorFacingState)
                // - which is the row that says "Published to the stable source".
                "merged" => "Published to the BETA source",

                // *** "withdrawn" IS A SUBMISSION REPLACED BY THE CONTRIBUTOR'S NEWER ONE (owner
                // decision, 2026-09-26). *** Nothing else sets it: the server withdraws an older,
                // untouched submission of the same board when a newer one from the same
                // contributor arrives (CRT.Server's SubmissionReplacementRules), and its comment
                // says so. If a contributor could ever withdraw one by hand, that needs a state of
                // its own - these words would then be wrong.
                "withdrawn" => "Replaced by a newer submission",

                // *** "returned" IS A SUBMISSION TAKEN BACK OUT OF BETA (code review, 2026-09-27). ***
                // Not a database state: the server reports it for a `pending` submission carrying a
                // maintainer's reason, which only a BETA rollback produces
                // (ProductionPromotionRules.ContributorFacingState). The contributor had already
                // been shown "Published to the BETA source" and mailed so; a bare "Waiting for review"
                // afterwards said nothing about what had happened. The reason is the maintainer
                // comment beside it.
                "returned" => "Taken back out of BETA - waiting for review again",

                // A state this build has never heard of. Reported honestly rather than guessed at -
                // a future server value rendered as something plausible-but-wrong is worse than an
                // admission that this version does not know it.
                _ => $"Reported as [{state.Trim()}]"
            };
        }

        // ###########################################################################################
        // *** A SUBMISSION THE SERVER NO LONGER KNOWS (2026-10-04). *** The server answers "not
        // found" for a submission deleted with its board (Account > "Delete a board") or by a reset
        // of the contribution data at go-live (Account > "Reset contribution data"). The receipt keeps
        // the state it last had - and showed it, "Submitted - awaiting feedback" for ever, about a
        // submission nobody will ever look at. These two read the RECEIPT, not only its state, so a
        // receipt marked not found (SubmissionReceipt.NotFoundUtc) says so instead - in the neutral
        // colour, since nothing was decided on its merit.
        // ###########################################################################################
        public const string NotOnServerWording = "No longer on the server";

        public static string DescribeReceiptState(SubmissionReceipt receipt)
        {
            ArgumentNullException.ThrowIfNull(receipt);

            return receipt.NotFoundUtc is not null
                ? SubmissionReceiptPresenter.NotOnServerWording
                : SubmissionReceiptPresenter.DescribeState(receipt.LastKnownState);
        }

        public static SubmissionOutcomeKind ClassifyReceipt(SubmissionReceipt receipt)
        {
            ArgumentNullException.ThrowIfNull(receipt);

            return receipt.NotFoundUtc is not null
                ? SubmissionOutcomeKind.Waiting
                : SubmissionReceiptPresenter.ClassifyState(receipt.LastKnownState);
        }

        // ###########################################################################################
        // WHAT KIND OF OUTCOME a state represents, for anything that needs to COLOUR it
        // (owner request, 2026-09-22).
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

                // Taken back out of BETA because something needs attention - the owner's own words
                // for why the button exists: "so I can inform contributor if something is missing".
                // The maintainer's note says what; this is the colour that makes it read.
                "returned" => SubmissionOutcomeKind.NeedsAction,

                // Finished, and it is not going in. "uploading" belongs here rather than in
                // Waiting: nothing is uploading any more, the send failed partway, and the row is
                // as dead as a rejection until the contributor sends again.
                "rejected" or "abandoned" or "uploading" => SubmissionOutcomeKind.Bad,

                // Replaced by the contributor's own newer submission (see DescribeState): its work
                // goes on in that one, so it is NOT painted as refused - red there read as "your
                // contribution was turned down". Neutral, like anything else not decided on merit.
                "withdrawn" => SubmissionOutcomeKind.Waiting,

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

                // "returned" (taken back out of BETA) is waiting for review again, so it is open too -
                // it falls to the default arm below, deliberately.

                // *** "changes_requested" AND "approved" ARE DELIBERATELY STILL OPEN. *** Neither
                // is the end: a submission with changes requested can be re-reviewed after the
                // contributor acts, and an approved one is still waiting to be published. Treating
                // either as decided would freeze the row at that state and the contributor would
                // never see it move.
                _ => true
            };
        }

        // ###########################################################################################
        // How long after its decision a "merged" (in BETA) submission is asked about at EVERY
        // launch.
        //
        // *** A BOUND, BECAUSE THE SECOND PUBLISH MAY NEVER COME (code review, 2026-09-25). ***
        // Publishing to production is off until the server is set up for it, and a maintainer may
        // never promote a board. Without a bound, every merged receipt was asked about on every
        // launch for ever - the very re-check "withdrawn" once caused. Thirty days is far longer
        // than BETA to production is meant to take.
        //
        // *** BUT NO LONGER A POINT AFTER WHICH IT IS NEVER ASKED AGAIN (code review, 2026-09-27). ***
        // A merged submission can now CHANGE after any length of time: a BETA rollback puts it back
        // in the queue ("returned"). With the old hard stop, a rollback more than thirty days after
        // the merge never reached "My submissions", which said "Published to the BETA source" for good
        // while the server said otherwise. Past the window it is asked about once every
        // MergedLateRecheckInterval instead - still cheap, never silent.
        // ###########################################################################################
        public static readonly TimeSpan MergedRecheckWindow = TimeSpan.FromDays(30);

        public static readonly TimeSpan MergedLateRecheckInterval = TimeSpan.FromDays(7);

        // How often a receipt the server answered "not found" for is asked about again (code review,
        // 2026-10-04). Not never: a 404 from a server being redeployed must not silence a live
        // submission for good. Once a day is nothing beside the minute check it replaces.
        public static readonly TimeSpan NotFoundRecheckInterval = TimeSpan.FromDays(1);

        // ###########################################################################################
        // Whether this RECEIPT is worth asking the server about AT LAUNCH: its state is still open,
        // and a "merged" one was decided less than MergedRecheckWindow ago (or sent, for a receipt
        // that never recorded its decision date) - or, past that, was last checked at least
        // MergedLateRecheckInterval ago. "Check for updates" in My submissions uses the state alone
        // - somebody pressed it to ask.
        // ###########################################################################################
        public static bool IsStillOpen(SubmissionReceipt receipt, DateTimeOffset nowUtc)
        {
            ArgumentNullException.ThrowIfNull(receipt);

            if (!SubmissionReceiptPresenter.IsStillOpen(receipt.LastKnownState))
                return false;

            // Unknown to the server: once a day, whatever its state.
            if (receipt.NotFoundUtc is not null)
            {
                return receipt.LastCheckedUtc is not DateTimeOffset lastChecked ||
                    nowUtc - lastChecked >= SubmissionReceiptPresenter.NotFoundRecheckInterval;
            }

            if (!string.Equals(receipt.LastKnownState?.Trim(), "merged", StringComparison.OrdinalIgnoreCase))
                return true;

            DateTimeOffset since = receipt.DecidedUtc ?? receipt.SentUtc;

            if (nowUtc - since < SubmissionReceiptPresenter.MergedRecheckWindow)
                return true;

            // Past the window: now and then, so a late rollback still arrives.
            return receipt.LastCheckedUtc is not DateTimeOffset checkedUtc ||
                nowUtc - checkedUtc >= SubmissionReceiptPresenter.MergedLateRecheckInterval;
        }

        // ###########################################################################################
        // THE ONE DATE FORMAT EVERY SUBMISSION LINE USES: "2026-September-23" (owner request,
        // 2026-09-23).
        //
        // *** YEAR FIRST, MONTH BY NAME, DAY WITHOUT A LEADING ZERO. *** The shape was chosen to be
        // unambiguous on sight: "09-10-2026" is the tenth of September to half the world and the
        // ninth of October to the other half, and these dates are read beside one another on a list
        // where the difference decides whether a maintainer's comment is stale.
        //
        // *** INVARIANT CULTURE, DELIBERATELY. *** The month name is the one part of this that a
        // culture can change, and it did: on the project owner's own Danish machine the previous
        // format rendered "22 september 2026" - lower case, because Danish does not capitalise
        // month names. Formatting with CultureInfo.CurrentCulture would give a different string on
        // every machine while the format string looked identical, so the format the project owner
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
        // A board's name as a contributor reads it: "Commodore/C128/310378 Open128" ->
        // "Commodore C128 310378 Open128".
        //
        // *** A RECEIPT HOLDS THE BOARD ID ITSELF, with no file name on the end (2026-09-27). ***
        // "My submissions" used to drop the last segment as though it were the workbook's file
        // name, so a real receipt read "Commodore C128" with the board missing. A workbook path
        // (the shape older fixtures used) still loses only its ".xlsx" file.
        // ###########################################################################################
        public static string DescribeBoard(string? boardId)
        {
            if (string.IsNullOrWhiteSpace(boardId))
            {
                return "(unknown board)";
            }

            List<string> segments = boardId.Split('/', StringSplitOptions.RemoveEmptyEntries)
                .Select(segment => segment.Trim())
                .Where(segment => segment.Length > 0)
                .ToList();

            if (segments.Count > 1 && segments[^1].EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                segments.RemoveAt(segments.Count - 1);
            }

            return segments.Count == 0 ? boardId.Trim() : string.Join(' ', segments);
        }

        // ###########################################################################################
        // *** THE "SWITCH BACK FROM BETA" NOTICE (owner request, 2026-09-27). *** A contributor
        // checks their submission in the BETA data - "I do think the user must do that to confirm it
        // works" - and then needs telling when it has reached the normal online source, so they
        // stop downloading BETA. Which submissions that notice is about: published to production,
        // not yet dismissed, and only while this machine downloads from BETA at all - someone on the
        // normal source has nothing to switch.
        //
        // Derived from the receipts every time rather than raised once when a state changes, so it
        // cannot be missed: a notice raised as the app closed would otherwise be gone for good,
        // since a published submission is never asked about again.
        // ###########################################################################################
        public static IReadOnlyList<SubmissionReceipt> NeedingSourceSwitchNotice(
            IEnumerable<SubmissionReceipt>? receipts,
            bool downloadingFromBeta)
        {
            if (!downloadingFromBeta)
            {
                return [];
            }

            return (receipts ?? [])
                .Where(receipt => string.Equals(receipt.LastKnownState?.Trim(), "published", StringComparison.OrdinalIgnoreCase))
                .Where(receipt => !receipt.SourceNoticeDismissed)
                .OrderBy(receipt => receipt.SentUtc)
                .ToList();
        }

        // The notice's words, naming each board once.
        public static string DescribeSourceSwitchNotice(IReadOnlyList<SubmissionReceipt> receipts)
        {
            (string named, string verb) = SubmissionReceiptPresenter.NameBoards(receipts);

            return $"{named} {verb} now published to the stable source. You are downloading data from the BETA source - " +
                   "you can switch back to the stable source on the Configuration tab.";
        }

        // ###########################################################################################
        // *** THE "TRY IT IN BETA" NOTICE (owner request, 2026-10-03: "When a system gets published
        // to the either online source, then there should be an information to the contributor that
        // he can test this"). *** The twin of the switch-back notice above, for the step before it:
        // a submission accepted into the BETA source ("merged"), not dismissed, while this machine
        // downloads from the STABLE source - someone already on BETA gets the data with the next
        // data check, and is not told to tick a box that is ticked. Derived from the receipts each
        // time, like the other one, so it cannot be missed.
        // ###########################################################################################
        public static IReadOnlyList<SubmissionReceipt> NeedingBetaTryNotice(
            IEnumerable<SubmissionReceipt>? receipts,
            bool downloadingFromBeta)
        {
            if (downloadingFromBeta)
            {
                return [];
            }

            return (receipts ?? [])
                .Where(receipt => string.Equals(receipt.LastKnownState?.Trim(), "merged", StringComparison.OrdinalIgnoreCase))
                .Where(receipt => !receipt.BetaNoticeDismissed)
                .Where(receipt => receipt.NotFoundUtc is null)
                .OrderBy(receipt => receipt.SentUtc)
                .ToList();
        }

        // The notice's words. With "Check for new or updated data at application launch" off, the
        // BETA check box is greyed out, so the notice names that one first.
        public static string DescribeBetaTryNotice(IReadOnlyList<SubmissionReceipt> receipts, bool checkDataOnLaunch)
        {
            (string named, string verb) = SubmissionReceiptPresenter.NameBoards(receipts);

            string tick = checkDataOnLaunch
                ? $"tick \"{ConfigurationWording.BetaSourceCheckBox}\""
                : $"tick \"{ConfigurationWording.CheckDataOnLaunchCheckBox}\" and then \"{ConfigurationWording.BetaSourceCheckBox}\"";

            return $"{named} {verb} now in the BETA source. To try it before everyone else, {tick} on the Configuration tab.";
        }

        // ###########################################################################################
        // The line under a submission in "My submissions" while it is in the BETA source (2026-10-03)
        // - there after the notice is closed, gone once the state moves on. Empty otherwise.
        // ###########################################################################################
        public const string BetaTryLine = "To try it before everyone else, download data from the BETA source - see the Configuration tab.";

        public static string DescribeBetaTry(string? state) =>
            string.Equals(state?.Trim(), "merged", StringComparison.OrdinalIgnoreCase)
                ? SubmissionReceiptPresenter.BetaTryLine
                : string.Empty;

        // The same for a receipt - none for one the server no longer knows (2026-10-04): what it sent
        // is not in the BETA source any more.
        public static string DescribeReceiptBetaTry(SubmissionReceipt receipt)
        {
            ArgumentNullException.ThrowIfNull(receipt);

            return receipt.NotFoundUtc is not null
                ? string.Empty
                : SubmissionReceiptPresenter.DescribeBetaTry(receipt.LastKnownState);
        }

        // "Your submission for Commodore C64 250407" / "Your submissions for A and B", and its verb -
        // each board named once.
        private static (string Named, string Verb) NameBoards(IReadOnlyList<SubmissionReceipt> receipts)
        {
            ArgumentNullException.ThrowIfNull(receipts);

            List<string> boards = receipts
                .Select(receipt => SubmissionReceiptPresenter.DescribeBoard(receipt.BoardId))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            string named = boards.Count switch
            {
                0 => "Your submission",
                1 => $"Your submission for {boards[0]}",
                _ => $"Your submissions for {string.Join(", ", boards.Take(boards.Count - 1))} and {boards[^1]}",
            };

            return (named, boards.Count > 1 ? "are" : "is");
        }

        // ###########################################################################################
        // The LATEST submission of one board, or null when it was never sent - for the badge on
        // its Drafts tab row (owner request, 2026-09-27: "When I have submitted ... I need to see
        // that somehow").
        //
        // Matched on the board id, the one a submission is sent under
        // (BoardDescriptorRules.BoardIdFromExcelDataFile) - never on a display name. Latest by
        // the time it was sent; a newer one replaces an older one on the server too.
        //
        // *** ONLY SUBMISSIONS SENT FROM THIS DRAFT (code review, 2026-09-27). *** A contributor
        // whose first submission reached production has that draft retired, and starts a new one of
        // the same board - which then carried "Published to the stable source", in green, beside changes that
        // had not been sent at all. draftCreatedUtc (the marker's) leaves out anything sent before
        // the draft existed; null, from a marker that does not say, leaves nothing out.
        // ###########################################################################################
        public static SubmissionReceipt? LatestForBoard(
            IEnumerable<SubmissionReceipt>? receipts,
            string? boardId,
            DateTimeOffset? draftCreatedUtc = null)
        {
            string id = boardId?.Trim() ?? string.Empty;
            if (id.Length == 0)
            {
                return null;
            }

            return (receipts ?? [])
                .Where(receipt => string.Equals(receipt.BoardId?.Trim(), id, StringComparison.OrdinalIgnoreCase))
                .Where(receipt => draftCreatedUtc is null || receipt.SentUtc >= draftCreatedUtc.Value)
                .OrderByDescending(receipt => receipt.SentUtc)
                .ThenByDescending(receipt => receipt.SubmissionId)
                .FirstOrDefault();
        }

        // ###########################################################################################
        // *** HAS THE DRAFT ALREADY BEEN SENT AS IT IS NOW? (owner request, 2026-10-03: "It should
        // not be possible to submit the same data again"; cases agreed with the project owner) ***
        // True while the draft's fingerprint is the one its LATEST submission (LatestForBoard,
        // only those sent from this draft) was sent with - whatever has happened to that
        // submission since: waiting, approved, in BETA, taken back out of BETA, changes requested,
        // even not accepted. Sending the same rows again answers none of those.
        //
        // NOT when that submission never finished sending - cancelled, the connection lost, the
        // upload window expired, or never confirmed by the server (no state yet: the receipt is
        // written before the upload, and its state only once the server confirms it). Then the
        // same draft is exactly what should be sent again.
        //
        // Never blocks on an unknown: a receipt from before fingerprints existed, or a draft whose
        // fingerprint could not be worked out (DraftFingerprint gives "" then).
        //
        // NOR when the server no longer knows that submission (2026-10-04) - deleted with its board,
        // or by the reset at go-live. What was sent is gone, so sending the same draft again is
        // exactly what is wanted.
        // ###########################################################################################
        public static bool IsAlreadySent(SubmissionReceipt? latest, string? currentFingerprint)
        {
            if (latest is null ||
                latest.NotFoundUtc is not null ||
                string.IsNullOrEmpty(latest.DraftFingerprint) ||
                string.IsNullOrEmpty(currentFingerprint))
            {
                return false;
            }

            if (SubmissionReceiptPresenter.NeverFinishedSending(latest.LastKnownState))
            {
                return false;
            }

            return string.Equals(latest.DraftFingerprint, currentFingerprint, StringComparison.Ordinal);
        }

        private static bool NeverFinishedSending(string? state) =>
            (state ?? string.Empty).Trim().ToLowerInvariant() is "" or "uploading" or "abandoned";

        // The greyed-out Submit button's tooltip (case 1).
        public static string DescribeAlreadySent(SubmissionReceipt receipt)
        {
            ArgumentNullException.ThrowIfNull(receipt);

            return $"You have already sent this draft as it is now, on {SubmissionReceiptPresenter.FormatDate(receipt.SentUtc)}. " +
                   "Change something to send it again.";
        }

        // ###########################################################################################
        // The badge's tooltip: which submission it is about and where the rest is. The badge itself
        // is DescribeState's words, so it cannot say anything "My submissions" does not.
        // ###########################################################################################
        public static string DescribeLastSubmission(SubmissionReceipt receipt)
        {
            ArgumentNullException.ThrowIfNull(receipt);

            return $"Your last submission of this board, sent {SubmissionReceiptPresenter.FormatDate(receipt.SentUtc)}. " +
                   "Open \"My submissions\" above for the details.";
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
        // "Replied 2026-September-22" - when the MAINTAINER decided, not when this machine noticed.
        //
        // Empty when nothing has been decided, so the caller can leave the line out entirely rather
        // than print a label with nothing after it.
        //
        // Same date-only rule as DescribeSent, and for the same reason: the clocks belong to
        // different people in different time zones, so a time would imply a precision that is not
        // there. See that method's header.
        // ###########################################################################################
        // ###########################################################################################
        // The line "My submissions" shows when a maintainer changed the submission before deciding it
        // (2026-09-25), so a contributor comparing what was published with what they sent knows
        // where a difference came from. Empty when nobody changed it.
        // ###########################################################################################
        public static string DescribeAmended(bool amendedByMaintainer) =>
            amendedByMaintainer
                ? "A maintainer changed some of the details before deciding, so what is published is not exactly what you sent."
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
        // the quietest thing on the card. *** The project owner asked what it meant, which is fair
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
        // Does this receipt carry a maintainer comment the contributor has NOT yet acknowledged?
        //
        // *** THE WHOLE POINT: A COMMENT NOBODY NOTICES IS A COMMENT NOBODY READS. *** Contributing
        // needs no account, so there is no inbox and no thread - the maintainer's sentence is the
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

            string comment = (receipt.MaintainerComment ?? string.Empty).Trim();

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
        // *** THE HALF THAT WAS MISSING, AND THE COMMONEST CASE OF ALL (owner report,
        // 2026-09-23). *** A maintainer approving good work usually writes nothing - there is
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
                // Queued for a maintainer, or still mid-send. Neither is a decision.
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
