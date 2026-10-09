using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // WHERE A BOARD IS, AS THREE STAGES under its name on the Boards screen (owner request,
    // 2026-10-04: "I see this system here ... where does this sit now, as I do not think it is in
    // BETA nor stable?" - and the answer chosen: a stage line, with a BETA / Stable switch on Board
    // data and Files).
    //
    //   Submitted - its newest submission: number, state in CRT's own words, when it was sent or
    //               decided, and how many there have been;
    //   BETA      - its revision there, and whether it waits to be pushed to stable - or not there;
    //   Stable    - its revision there and when it was published - or not there, or not known when
    //               the server has no stable source to look in.
    //
    // A board "not published yet" then says why at a glance: its submission's state, and nothing
    // in either tree. Pure, so what the maintainer reads is tested.
    // ###########################################################################################
    public enum BoardStageState
    {
        // There: a submission, or the board in that tree.
        Reached,

        // Not there - said, never left blank.
        NotThere,

        // Not known yet (the submissions not read) or not knowable (no stable source).
        NotKnown
    }

    public sealed record BoardStage(string Label, string Value, string? Detail, BoardStageState State);

    // Where the board's newest work sits - the stage drawn as current (0 Submitted, 1 BETA, 2 Stable)
    // - and the sentence that says it under the stages.
    public sealed record BoardStageNow(int Stage, string Sentence);

    public static class BoardStagesDisplay
    {
        public const string SubmittedLabel = "Submitted";

        public const string BetaLabel = "BETA";

        public const string StableLabel = "Stable";

        public const string NotThere = "Not there";

        // ###########################################################################################
        // *** WHAT EACH OF THE THREE PLACES IS, said on its card (owner request, 2026-10-09: the three
        // columns were not clear enough - "please see how you can improve this so the user better
        // can understand this"). *** Numbered, since the work goes through them in that order.
        // ###########################################################################################
        public const string SubmittedNote = "Sent in by contributors, reviewed here";

        public const string BetaNote = "Published for testing - CRT's BETA source";

        public const string StableNote = "What everybody using CRT gets";

        public static string NumberedLabel(int stage, string label) =>
            $"{(stage + 1).ToString(CultureInfo.InvariantCulture)}  {label}";

        // ###########################################################################################
        // WHERE THE BOARD'S WORK IS NOW, AND WHAT COMES NEXT - the card drawn as current and the
        // sentence under the three (owner request, 2026-10-09). Null while the board's submissions are
        // not read yet, since the work may be one of them.
        //
        // EVERY SUBMISSION COUNTS, not only the newest (code review, 2026-10-09): #5 waiting for review
        // behind a newer #6 that was rejected still waits, and was said as "nothing is waiting".
        //
        // BETA ahead of the stable source comes first even when a submission waits for review too:
        // a board takes one submission into BETA at a time (OneSubmissionInBeta), so BETA has to be
        // published or pushed back before those can be approved - which the sentence then says. A
        // submission back with its contributor for changes waits on nobody here, so it is added to
        // what the trees say rather than outlining Submitted.
        // ###########################################################################################
        public static BoardStageNow? Now(BoardOverviewEntry board, IReadOnlyList<BoardSubmissionEntry>? submissions)
        {
            ArgumentNullException.ThrowIfNull(board);

            if (submissions is null)
                return null;

            List<BoardSubmissionEntry> waiting = BoardStagesDisplay.InState(submissions, "pending", "approved", "returned");
            List<BoardSubmissionEntry> withContributor = BoardStagesDisplay.InState(submissions, "changes_requested");

            if (board.IsAwaitingProduction)
            {
                string beta = $"BETA is ahead of the stable source and waits under {MaintainerScreenWording.BetaQueueQuoted} to be published to stable - or pushed back.";

                return new BoardStageNow(1, waiting.Count == 0
                    ? beta
                    : $"{beta} {BoardStagesDisplay.Numbers(waiting)} {(waiting.Count == 1 ? "waits" : "wait")} for review too, and can be approved once that is done.");
            }

            if (waiting.Count > 0)
            {
                return new BoardStageNow(0,
                    $"{BoardStagesDisplay.Numbers(waiting)} {(waiting.Count == 1 ? "waits" : "wait")} for review under {MaintainerScreenWording.ContributorQueueQuoted}.");
            }

            BoardStageNow trees = (board.InBeta, board.InProduction) switch
            {
                (true, true) => new BoardStageNow(2, "BETA and the stable source hold the same - nothing is waiting."),
                (true, false) => new BoardStageNow(1, "In BETA only - not published to the stable source yet."),
                (true, null) => new BoardStageNow(1, "In BETA. This server has no stable source to compare it with."),
                (false, true) => new BoardStageNow(2, "In the stable source, but not in BETA."),
                _ => new BoardStageNow(0, "Nothing of this board is published yet.")
            };

            return withContributor.Count == 0
                ? trees
                : trees with
                {
                    Sentence = withContributor.Count == 1
                        ? $"{trees.Sentence} {BoardStagesDisplay.Numbers(withContributor)} is back with the contributor for changes."
                        : $"{trees.Sentence} {BoardStagesDisplay.Numbers(withContributor)} are back with the contributors for changes."
                };
        }

        // The submissions in any of `states` (CRT's contributor-facing words), oldest first.
        private static List<BoardSubmissionEntry> InState(IEnumerable<BoardSubmissionEntry> submissions, params string[] states) =>
            submissions
                .Where(submission => states.Contains(submission.State?.Trim().ToLowerInvariant() ?? string.Empty, StringComparer.Ordinal))
                .OrderBy(submission => submission.CreatedUtc)
                .ThenBy(submission => submission.Id)
                .ToList();

        // "#5", "#5 and #7", "#5, #6 and #7".
        private static string Numbers(IReadOnlyList<BoardSubmissionEntry> submissions)
        {
            List<string> numbers = submissions.Select(submission => $"#{submission.Id.ToString(CultureInfo.InvariantCulture)}").ToList();

            return numbers.Count == 1
                ? numbers[0]
                : $"{string.Join(", ", numbers.Take(numbers.Count - 1))} and {numbers[^1]}";
        }

        // The three, left to right. `submissions` is the board's detail's - null until it is read.
        public static IReadOnlyList<BoardStage> For(BoardOverviewEntry board, IReadOnlyList<BoardSubmissionEntry>? submissions)
        {
            ArgumentNullException.ThrowIfNull(board);

            return [BoardStagesDisplay.Submitted(submissions), BoardStagesDisplay.Beta(board), BoardStagesDisplay.Stable(board)];
        }

        // ###########################################################################################
        // The NEWEST submission - the one that says where the latest work is. "#5 Not accepted",
        // "decided 21 Sep 2026 - the latest of 2".
        // ###########################################################################################
        public static BoardStage Submitted(IReadOnlyList<BoardSubmissionEntry>? submissions)
        {
            if (submissions is null)
                return new BoardStage(BoardStagesDisplay.SubmittedLabel, string.Empty, null, BoardStageState.NotKnown);

            BoardSubmissionEntry? latest = submissions
                .OrderByDescending(submission => submission.CreatedUtc)
                .ThenByDescending(submission => submission.Id)
                .FirstOrDefault();

            if (latest is null)
                return new BoardStage(BoardStagesDisplay.SubmittedLabel, "No submissions", null, BoardStageState.NotThere);

            string when = latest.DecidedUtc is DateTimeOffset decided
                ? $"decided {SubmissionReceiptPresenter.FormatDate(decided)}"
                : $"sent {SubmissionReceiptPresenter.FormatDate(latest.CreatedUtc)}";

            string of = submissions.Count > 1
                ? $" - the latest of {submissions.Count.ToString(CultureInfo.InvariantCulture)}"
                : string.Empty;

            return new BoardStage(
                BoardStagesDisplay.SubmittedLabel,
                $"#{latest.Id.ToString(CultureInfo.InvariantCulture)} {SubmissionReceiptPresenter.DescribeState(latest.State)}",
                when + of,
                BoardStageState.Reached);
        }

        public static BoardStage Beta(BoardOverviewEntry board)
        {
            ArgumentNullException.ThrowIfNull(board);

            if (!board.InBeta)
                return new BoardStage(BoardStagesDisplay.BetaLabel, BoardStagesDisplay.NotThere, null, BoardStageState.NotThere);

            return new BoardStage(
                BoardStagesDisplay.BetaLabel,
                BoardStagesDisplay.Revision(board.BetaRevision, "In BETA"),
                board.IsAwaitingProduction ? $"ahead of stable - waiting under {MaintainerScreenWording.BetaQueueQuoted}" : null,
                BoardStageState.Reached);
        }

        public static BoardStage Stable(BoardOverviewEntry board)
        {
            ArgumentNullException.ThrowIfNull(board);

            return board.InProduction switch
            {
                true => new BoardStage(
                    BoardStagesDisplay.StableLabel,
                    BoardStagesDisplay.Revision(board.ProductionRevision, "In the stable source"),
                    board.ProductionPublishedUtc is DateTimeOffset published
                        ? $"published {SubmissionReceiptPresenter.FormatDate(published)}"
                        : null,
                    BoardStageState.Reached),
                false => new BoardStage(BoardStagesDisplay.StableLabel, BoardStagesDisplay.NotThere, null, BoardStageState.NotThere),
                _ => new BoardStage(BoardStagesDisplay.StableLabel, "Not known", "this server has no stable source to look in", BoardStageState.NotKnown)
            };
        }

        // "Revision 2026-September-25", or the plain fact for a board with no revision on record.
        private static string Revision(string? revision, string withoutOne) =>
            string.IsNullOrWhiteSpace(revision) ? withoutOne : $"Revision {revision.Trim()}";
    }
}
