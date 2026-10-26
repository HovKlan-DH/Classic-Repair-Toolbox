using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THE SUBMIT DIALOG SAYS WHEN IT SENDS SEVERAL DRAFTS (CRT 4.0 step 2, slice 5 - the plan:
    // "Submit offers to send all the drafts at once, with one message"). Each draft is its own
    // submission - the server takes one board at a time - so these words say what came of EACH: sent,
    // turned down, or not sent and why. Words only, so they are tested without a display.
    // ###########################################################################################
    internal static class SubmitDraftsWording
    {
        public const string OtherDraftsHeading = "Your other drafts";

        public const string OtherDraftsIntro =
            "Tick any you want to send too. Each goes for review on its own, with the same email address and the same message.";

        // A tick box under the credit name, while another draft is ticked - so crediting the other drafts
        // is a choice of its own (code review, 2026-10-26: unticking every credit of this board still
        // credited them all).
        public const string CreditOthers =
            "Credit me on the other drafts I tick too, each for its own work";

        // ---- why a draft is not offered, or was not sent after all -----------------------------------

        public const string NothingToSend = "nothing to send yet";
        public const string AlreadySent = "already sent as it is";
        public const string DraftGone = "its draft is no longer there";
        public const string NotSaved = "its draft could not be saved - its header on the Contribute tab says why";
        public const string NotRead = "its draft could not be read";
        public const string CreditNotAdded = "its credit could not be put into its draft";
        public const string NoConnection = "it could not reach the server - check your internet connection and try again";
        public const string Unexpected = "something unexpected stopped it - the log says what";

        public static string ErrorsFirst(int errors) =>
            errors == 1 ? "1 error to fix first" : string.Create(CultureInfo.InvariantCulture, $"{errors} errors to fix first");

        // Why a draft's row says it cannot be sent, in the dialog's short form - or null when it can be.
        public static string? NotOffered(bool hasSomething, bool alreadySent, int errors) =>
            !hasSomething ? NothingToSend
            : alreadySent ? AlreadySent
            : errors > 0 ? ErrorsFirst(errors)
            : null;

        // One tick box: the board and what its draft holds, or - greyed out - why it cannot go.
        public static string Choice(string boardName, string rowSummary, string? notOffered) =>
            notOffered is null ? $"{boardName} ({rowSummary})" : $"{boardName} - {notOffered}";

        // A sentence turned into the tail of "Board: not sent - ...": no full stop, a small first
        // letter unless the first word has a capital of its own further on ("KiCad", "BETA").
        public static string AsReason(string? sentence)
        {
            string text = (sentence ?? string.Empty).Trim().TrimEnd('.').Trim();

            if (text.Length == 0)
                return Unexpected;

            string firstWord = text.Split(' ', 2)[0];
            bool keepsCapitals = firstWord.Skip(1).Any(char.IsUpper);

            return keepsCapitals ? text : char.ToLowerInvariant(text[0]) + text[1..];
        }

        // ---- while sending ----------------------------------------------------------------------------

        public static string GettingReady(string boardName) => $"Getting {boardName} ready...";

        public static string Sending(string boardName, string progress) => $"{boardName}: {progress}";

        // ---- what came of it ---------------------------------------------------------------------------

        public static string Heading(int sent, int total) =>
            sent == total ? "Contributions sent"
            : sent == 0 ? "Contributions not sent"
            : string.Create(CultureInfo.InvariantCulture, $"{sent} of {total} contributions sent");

        // `credited`: credits went into drafts before sending - so the drafts are not as they were, and
        // the credits go with the next try (code review, 2026-10-26).
        public static string Intro(int sent, bool credited = false) =>
            sent == 0
                ? credited
                    ? "Nothing was sent. Your drafts are untouched but for the credits put in, which go with the next try."
                    : "Nothing was sent. Your drafts are untouched."
                : "Each contribution sent is queued for review on its own. You get an email when a maintainer has looked at it, " +
                  "and again as it goes into the BETA source and then the stable source.\n\n" +
                  (credited ? "Your drafts are untouched but for the credits put in - keep " : "Your drafts are untouched - keep ") +
                  "using each as normal. Each tidies itself away once its work is in the stable source.";

        public static string Queued(string boardName) => $"{boardName}: queued for review";

        public static string NotAccepted(string boardName) => $"{boardName}: not accepted - the server found problems to fix first";

        public static string NotSent(string boardName, string reason) => $"{boardName}: not sent - {reason}";

        // Every sentence above, for the test that they are plain ASCII.
        internal static IEnumerable<string> All() =>
        [
            OtherDraftsHeading, OtherDraftsIntro, NothingToSend, AlreadySent, DraftGone, NotSaved,
            NotRead, CreditNotAdded, NoConnection, Unexpected, ErrorsFirst(2), Heading(1, 2), Intro(0), Intro(1),
            Intro(0, credited: true), Intro(1, credited: true), CreditOthers,
            Queued("A"), NotAccepted("A"), NotSent("A", NotRead), GettingReady("A"),
        ];
    }
}
