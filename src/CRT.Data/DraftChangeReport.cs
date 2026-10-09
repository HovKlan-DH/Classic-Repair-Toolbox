using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THE "WHAT CHANGED" WINDOW SHOWS for one drafted board
    // (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
    //
    // *** THIS REPLACES DraftDriftReport, AND IT ANSWERS A NARROWER QUESTION. *** The old report
    // classified each DRAFTED row against the official data as it stands now, and could say two
    // things this cannot:
    //
    //   - GoneOfficially: "the row you edited has since disappeared from the official data", so
    //     your edit no longer applies to anything;
    //   - NowExistsOfficially: "a row you added now exists officially too", and the merge keeps
    //     the official one, so what you see is not your version.
    //
    // Both were consequences of the MERGE. BoardDraftApplier silently dropped a Modified row whose
    // key had vanished, and failed closed on an Added row whose key collided - invisible outcomes
    // that the report existed to surface. There is no merge any more, so neither outcome exists
    // and neither can be reported.
    //
    // What replaces it is plainer and, for the ordinary case, more useful: exactly what the
    // contributor has changed against the published board, derived by comparison
    // (BoardDataDiffer) rather than read from a ledger they could edit around in Excel.
    //
    // *** WHAT WAS LOST, STATED PLAINLY: *** the revision comparison still says "the published
    // board has moved on since you started", but nothing now says WHICH of your edits that
    // collides with. Recovering that needs the board as it was when the draft was seeded, and the
    // draft folder keeps no copy of it. That is a feature in its own right, not an oversight here.
    // ###########################################################################################
    public sealed class DraftChangeReport
    {
        // Whether the published board has moved since this draft was taken - the cheap revision
        // comparison, unchanged from before.
        public DraftDriftState State { get; init; }

        // The published revision this draft was seeded from, off the marker.
        public string BaseRevision { get; init; } = string.Empty;

        // The published revision as it reads NOW.
        public string OfficialRevision { get; init; } = string.Empty;

        // Every row that differs between the draft and the published board.
        public IReadOnlyList<BoardRowChange> Rows { get; init; } = [];
    }
}
