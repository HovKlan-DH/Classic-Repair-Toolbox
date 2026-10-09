using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // How a queue row reads on screen (NewContributeStrategy.md Phase 5, task 2).
    //
    // *** THE QUEUE IS GROUPED BY BOARD (2026-09-26). *** Each row was two joined lines at first -
    // "#4 Commodore/C64/250407" over the comment, the wait and the shared-file note run together -
    // then six lines with labelled Manufacturer / Hardware / Board and two badges on every row,
    // which the project owner still found "quite hard to overview". Now one heading per board, and
    // two short lines per submission under it:
    //
    //     Commodore / C64 / 250407             [New board]
    //        Corrected the pinout pictures for U8 and added U10.
    //        Waiting 10 hours - replaces a shared file
    //        Added the missing CIA pictures.
    //        Waiting 2 days - with the other approver             (dimmed)
    //
    // Submissions for one board are reviewed together, and the board is said once. Only what is
    // UNUSUAL is marked: "New board" on the heading (a published board is the ordinary case), and
    // a submission that does NOT wait for this account is dimmed and says so - where "Awaiting your
    // review" used to sit on nearly every row. Both answers are the SERVER's (ReviewQueueEntry).
    //
    // *** WHY A QUEUE ROW CARRIES NO CHANGE SUMMARY. *** Computing one needs BOTH the published board
    // and the submitted manifest, and the queue endpoint returns neither - making it do so would mean
    // loading two full BoardData per queued submission on the server, for a list the maintainer
    // scrolls past. So the queue shows what the contributor themselves said about it, which is also
    // what a maintainer scans for: "Corrected R12." says more than "3 components changed".
    //
    // Pure, so the wording is tested rather than eyeballed.
    // ###########################################################################################
    public static class ReviewQueueDisplay
    {
        public const string NewBoardBadge = "New board";

        // ###########################################################################################
        // The queue as board groups: each board once, in the order its OLDEST submission waits
        // (the queue arrives oldest first, and a group is placed by its first row), with its own
        // submissions under it in queue order. The queue's order is a decision - the longest-waiting
        // submission leads - and grouping keeps it: the board holding it comes first.
        // ###########################################################################################
        public static IReadOnlyList<ReviewQueueGroup> Group(IEnumerable<ReviewQueueRow> rows)
        {
            ArgumentNullException.ThrowIfNull(rows);

            var groups = new List<ReviewQueueGroup>();
            var byBoard = new Dictionary<string, List<ReviewQueueRow>>(StringComparer.Ordinal);

            foreach (ReviewQueueRow row in rows)
            {
                string board = row.BoardId ?? string.Empty;

                if (!byBoard.TryGetValue(board, out List<ReviewQueueRow>? members))
                {
                    members = [];
                    byBoard[board] = members;
                    groups.Add(new ReviewQueueGroup(board, members));
                }

                members.Add(row);
            }

            return groups;
        }

        // "Commodore / C64 / 250407" - or the id whole when it is not three parts, and a missing one
        // said to be missing rather than left as a gap that looks like a rendering fault.
        public static string BoardHeading(string? boardId)
        {
            if (BoardParts(boardId) is { } parts)
                return $"{parts.Manufacturer} / {parts.Hardware} / {parts.Board}";

            return string.IsNullOrWhiteSpace(boardId) ? "(unknown board)" : boardId;
        }

        // ###########################################################################################
        // The board's three parts, or null when the id is not a well-formed one - the row then
        // shows the id whole (BoardWhole) rather than inventing parts.
        // ###########################################################################################
        public static (string Manufacturer, string Hardware, string Board)? BoardParts(string? boardId)
        {
            if (!BoardDescriptorRules.IsValidBoardId(boardId))
                return null;

            string[] parts = boardId!.Split('/');

            return (parts[0], parts[1], parts[2]);
        }

        // ###########################################################################################
        // What the contributor wrote about it. *** A SUBMISSION WITH NO SUMMARY SAYS SO. *** The
        // field is optional, and an empty line reads as a rendering fault rather than as missing
        // information.
        // ###########################################################################################
        public static string Comment(ReviewQueueRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return string.IsNullOrWhiteSpace(row.Summary) ? "(no description given)" : row.Summary.Trim();
        }

        // ###########################################################################################
        // "New board" when the server says the board has no published board - the highest-risk
        // submission there is. Nothing otherwise: a published board is the ordinary case, and
        // marking it on every heading only drowned the one that matters.
        // ###########################################################################################
        public static string? BoardBadge(bool? isNewBoard) =>
            isNewBoard == true ? ReviewQueueDisplay.NewBoardBadge : null;

        // ###########################################################################################
        // Whether a submission OPENED waits for this account: it may decide it, and its approval is
        // still needed - the answer that leaves the Approve button on. The row on screen follows it,
        // since the detail judges shared files against the tree as it is now.
        // ###########################################################################################
        public static bool AwaitsYou(bool canPublish, ApprovalStatus? approval) =>
            canPublish && ApprovalWording.CanApprove(approval);

        // ###########################################################################################
        // A submission's second line: how long it has waited, and the two-approval notes.
        //
        // A shared-file change reaches every board citing the file and needs TWO approvals - the
        // board's maintainer and the administrator (2026-09-25); 'approved' in the queue is the first
        // of the two, given and waiting for the other. When it does NOT wait for this account (the
        // server's AwaitsYou), that is the reason: this account's side is done, and the other
        // approver's is not - said, since the row is dimmed for it.
        // ###########################################################################################
        public static string Footer(ReviewQueueRow row, DateTimeOffset now)
        {
            ArgumentNullException.ThrowIfNull(row);

            var parts = new List<string>();

            string waiting = ReviewQueueDisplay.Waiting(row.CreatedUtc, now);

            if (waiting.Length > 0)
                parts.Add(waiting == "just now" ? "arrived just now" : waiting);

            if (row.TouchesSharedFiles)
                parts.Add("replaces a shared file");

            if (row.AwaitsYou == false)
                parts.Add("with the other approver");
            else if (string.Equals(row.State, "approved", StringComparison.Ordinal))
                parts.Add("one of two approvals given");

            string footer = string.Join(" - ", parts);

            return footer.Length == 0 ? footer : char.ToUpperInvariant(footer[0]) + footer[1..];
        }

        // ###########################################################################################
        // How long this has been waiting, in the units a person actually thinks in.
        //
        // *** THIS IS THE NUMBER THAT SHAMES A BACKLOG, so it must not round in the comforting
        // direction. *** Truncating rather than rounding means a submission waiting 13 days reads
        // as "13 days", never "2 weeks", and 29 days never becomes "a month" - the coarser unit
        // always makes a wait sound shorter than it is.
        //
        // An absent timestamp yields an EMPTY string rather than a guess. Rendering a missing date
        // as "waiting 56 years" (the 1970 default) would be worse than saying nothing.
        // ###########################################################################################
        public static string Waiting(DateTimeOffset? createdUtc, DateTimeOffset now)
        {
            if (createdUtc is null)
                return string.Empty;

            TimeSpan waited = now - createdUtc.Value;

            // A clock difference between client and server can produce a small negative. Reading
            // "waiting -3 minutes" looks broken; "just now" is true enough and is what the
            // maintainer would conclude anyway.
            if (waited < TimeSpan.Zero)
                return "just now";

            if (waited.TotalMinutes < 1)
                return "just now";

            if (waited.TotalHours < 1)
                return ReviewQueueDisplay.Plural((int)waited.TotalMinutes, "minute");

            if (waited.TotalDays < 1)
                return ReviewQueueDisplay.Plural((int)waited.TotalHours, "hour");

            return ReviewQueueDisplay.Plural((int)waited.TotalDays, "day");
        }

        private static string Plural(int count, string unit) =>
            count == 1
                ? $"waiting 1 {unit}"
                : $"waiting {count.ToString(CultureInfo.InvariantCulture)} {unit}s";
    }

    // One board's submissions in the queue. IsNewBoard is the server's answer for the board -
    // every submission to one board gets the same one.
    public sealed record ReviewQueueGroup(string BoardId, IReadOnlyList<ReviewQueueRow> Rows)
    {
        public bool? IsNewBoard => this.Rows.Select(row => row.IsNewBoard).FirstOrDefault(isNew => isNew is not null);
    }
}
