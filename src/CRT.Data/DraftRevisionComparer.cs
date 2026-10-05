using System;
using System.Globalization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // How a draft's base revision relates to the official data's current revision
    // (NewContributeStrategy.md Phase 2, session 2d - the drift warning).
    //
    // Three-valued rather than a boolean, and the distinction is the whole point: the app may only
    // tell the user what the data actually supports. "Changed" is an honest claim whenever the two
    // strings differ; "older" is a much stronger claim that needs both sides to parse as real
    // dates.
    // ###########################################################################################
    public enum DraftDriftState
    {
        // Nothing to say - the draft was started from the revision that is still current.
        InSync,

        // Both revisions parsed as dates, and the official one is later. The strongest claim.
        OfficialIsNewer,

        // The revisions differ, but an ORDER could not be established - one side did not parse, or
        // they parse to the same date while differing as text. Still worth warning about; just not
        // as "newer".
        Changed,

        // No comparison is possible or meaningful: a draft-only system (which has no official
        // counterpart at all), or a draft written before base revisions were recorded on every
        // path. Never warned about - an unknown base is not evidence of drift.
        Unknown,
    }

    // ###########################################################################################
    // Interprets the free-text revision strings CRT reads off a board workbook's
    // "# Revision date:" marker (see BoardDataReader), and the ONLY place in the app that does so.
    //
    // WHAT THESE STRINGS ACTUALLY LOOK LIKE, read out of the shipped boards in Assets/Data rather
    // than assumed: "2026-May-12", "2026-July-19", "2026-August-21", "2026-September-20" - and, from
    // the server's own stamp, "2026-October-4". That is yyyy-MMMM-d with the FULL ENGLISH MONTH NAME
    // and a day of one OR two digits - not ISO-8601, and not "dd" (see AcceptedFormats). The test
    // fixtures (BoardWorkbookBuilder) use "2026-01-15" instead, so both shapes are real and both are
    // parsed.
    //
    // TWO TRAPS, both of which have a test named after them:
    //
    //   1. COMPARING THESE AS STRINGS IS BACKWARDS. Lexically "2026-August-21" sorts BEFORE
    //      "2026-May-14", because "A" < "M" - while August is four months later. Any CompareTo on
    //      these values is a defect, which is why ordering here goes through DateTime and string
    //      comparison is used only to answer "are these the same text".
    //
    //   2. IT MUST FAIL SOFT. The value is hand-maintained free text in a spreadsheet cell; a
    //      contributed or hand-edited board can carry anything, including a non-English month name
    //      or nothing at all. Unparseable input degrades to "Changed", never to an exception and
    //      never to a guessed ordering.
    //
    // InvariantCulture throughout, deliberately: the month names in the data are English, so a
    // machine with a Danish or German locale must still read "August" - parsing under the current
    // culture would make the drift warning depend on the user's regional settings.
    // ###########################################################################################
    public static class DraftRevisionComparer
    {
        // Ordered most-likely-first. yyyy-MMMM-d is what every board carries - the shipped ones
        // ("2026-August-21") and the server's own stamp (BoardWorkbookStyle.FormatRevisionDate, which
        // writes no leading zero: "2026-October-4"). A single "d" reads one OR two digits; "dd" read
        // only two, so every board the server published on the 1st to the 9th failed to parse
        // (2026-10-05). yyyy-MMM-d covers an abbreviated month ("2026-Sep-20"); the two numeric forms
        // cover the test fixtures and any contributor who wrote an ISO date.
        private static readonly string[] AcceptedFormats =
        {
            "yyyy-MMMM-d",
            "yyyy-MMM-d",
            "yyyy-MM-dd",
            "yyyy-M-d",
        };

        // ###########################################################################################
        // Compares a draft's recorded base revision against the official data's current one.
        //
        // A blank base revision gives Unknown rather than Changed: it means the base was never
        // recorded (a draft written before every save path stamped it), which is not evidence that
        // anything drifted. Warning on it would be a claim the data cannot support, and silently
        // stamping today's revision would be worse - it would assert the draft started from data it
        // may well predate by several syncs.
        // ###########################################################################################
        public static DraftDriftState Compare(string? baseRevision, string? officialRevision)
        {
            string normalizedBase = baseRevision?.Trim() ?? string.Empty;
            string normalizedOfficial = officialRevision?.Trim() ?? string.Empty;

            if (normalizedBase.Length == 0)
            {
                return DraftDriftState.Unknown;
            }

            // Identical text is in sync whether or not it parses as a date - two boards stamped
            // "see changelog" are as much in sync as two stamped "2026-May-12".
            if (string.Equals(normalizedBase, normalizedOfficial, StringComparison.Ordinal))
            {
                return DraftDriftState.InSync;
            }

            // An official revision that is missing entirely cannot be ordered against anything, but
            // it does differ from the recorded base, so it is a change.
            if (normalizedOfficial.Length == 0)
            {
                return DraftDriftState.Changed;
            }

            if (TryParse(normalizedBase, out DateTime parsedBase) &&
                TryParse(normalizedOfficial, out DateTime parsedOfficial) &&
                parsedOfficial > parsedBase)
            {
                return DraftDriftState.OfficialIsNewer;
            }

            // Everything else that differs: either side unparseable, the two parsing to the same
            // date while differing as text ("2026-05-12" vs "2026-May-12"), or the official one
            // parsing EARLIER than the base. That last case means something odd happened - a
            // rollback, or a hand-edited workbook - and it is still "changed underneath you", so it
            // folds in here rather than becoming a fourth state nobody could act on differently.
            return DraftDriftState.Changed;
        }

        // ###########################################################################################
        // Whether one revision string can be read as a date at all - exposed because the UI wants to
        // show the raw strings either way, and a caller may reasonably want to know whether the
        // ordering claim behind a message is sound.
        // ###########################################################################################
        public static bool TryParse(string? revision, out DateTime parsed) =>
            DateTime.TryParseExact(
                revision?.Trim() ?? string.Empty,
                AcceptedFormats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out parsed);
    }
}
