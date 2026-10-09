using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // How the administrator's "Delete a board" reads (owner request, 2026-10-03: "as admin, I
    // should be able to have a possibility to delete a system completely, which then will remove it
    // from everywhere - including BETA and stable sources ... list all systems and have a delete
    // button on each (with a confirmation box)").
    //
    // Pure, so the words are tested - the rule BoardsDisplay and ProductionDisplay follow. The
    // counts are the SERVER's plan (CRT.Server's BoardDeletionFlow), the code that then deletes, so
    // the confirmation cannot promise less than what goes. A refusal (blockedBecause, or any error)
    // is the server's sentence, shown unchanged.
    //
    // A submission's state is said in CRT's own words (SubmissionReceiptPresenter.DescribeState),
    // as on the Boards screen - and a contributor is never "their/them/they" (owner request,
    // 2026-10-01).
    // ###########################################################################################
    public static class BoardDeletionWording
    {
        public const string ListHeading = "Delete a board";

        public const string Explanation =
            "Deleting a board removes it from everywhere: its folder in the BETA and the stable data, its row in both " +
            "drop-down lists, its record in the database with every submission, maintainer and invitation of it, and " +
            "its view statistics. Shared files it uses are kept. Pressing Delete first shows exactly what would go.";

        public const string NoBoards = "There are no boards.";

        public const string DeleteButton = "Delete";

        public const string ConfirmTitle = "Delete a board";

        public const string ConfirmButton = "Delete board";

        public const string CannotBeUndone = "This cannot be undone.";

        public const string SharedFilesKept =
            "Shared files the board uses are kept. Any that nothing else uses are then listed under Unused files.";

        public const string ReasonPrompt = "Why? Each contributor above is told this, and nothing else.";

        // "Commodore / C64 / 250407" - the Boards screen's own way of naming a board.
        public static string Name(BoardDeletePlanAnswer plan)
        {
            ArgumentNullException.ThrowIfNull(plan);

            string name = string.Join(
                " / ",
                new[] { plan.Manufacturer, plan.Hardware, plan.Board }.Where(part => !string.IsNullOrWhiteSpace(part)));

            return name.Length > 0 ? name : plan.BoardId;
        }

        public static string Headline(BoardDeletePlanAnswer plan) =>
            $"Delete {BoardDeletionWording.Name(plan)} completely?";

        // ###########################################################################################
        // What goes, one line per place - each tree, then the database, then the statistics. A place
        // holding nothing of it says so, rather than leaving the administrator to wonder whether it
        // was looked at.
        // ###########################################################################################
        public static IReadOnlyList<string> WhatGoes(BoardDeletePlanAnswer plan)
        {
            ArgumentNullException.ThrowIfNull(plan);

            return
            [
                $"The BETA data: {BoardDeletionWording.TreePart(plan.BetaFiles, plan.ListedInBeta)}",
                $"The stable data: {BoardDeletionWording.TreePart(plan.ProductionFiles, plan.ListedInProduction)}",
                $"The database: {BoardDeletionWording.DatabasePart(plan)}",
                "Its view statistics"
            ];
        }

        // ###########################################################################################
        // The open submissions' heading - deleted too, and each contributor mailed (owner decision,
        // 2026-10-03: "Delete them and mail the contributors"). Null when none is open, and the
        // reason box is then not asked for either.
        // ###########################################################################################
        public static string? OpenHeading(BoardDeletePlanAnswer plan)
        {
            ArgumentNullException.ThrowIfNull(plan);

            return plan.OpenSubmissions.Count switch
            {
                0 => null,
                1 => "1 submission to it is still open. It is deleted too, and its contributor is mailed the reason below:",
                int count => $"{count} submissions to it are still open. These are deleted too, and each contributor is mailed the reason below:"
            };
        }

        // "#14 - 2026-Oct-02 - Waiting for review - anna@example.com - Corrected U8."
        public static string OpenLine(BoardDeleteOpenSubmission submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            string contributor = string.IsNullOrWhiteSpace(submission.Contributor) ? "(no contact address)" : submission.Contributor.Trim();
            string summary = string.IsNullOrWhiteSpace(submission.Summary) ? "(no description given)" : submission.Summary.Trim();

            return $"#{submission.Id.ToString(CultureInfo.InvariantCulture)} - " +
                $"{SubmissionReceiptPresenter.FormatDate(submission.CreatedUtc)} - " +
                $"{SubmissionReceiptPresenter.DescribeState(submission.State)} - {contributor} - {summary}";
        }

        // The reason is the mail's only content, so it is asked for only when somebody is mailed.
        public static bool NeedsReason(BoardDeletePlanAnswer plan)
        {
            ArgumentNullException.ThrowIfNull(plan);

            return plan.OpenSubmissions.Count > 0;
        }

        public static string Done(BoardDeleteAnswer answer)
        {
            ArgumentNullException.ThrowIfNull(answer);

            string told = answer.ContributorsMailed switch
            {
                0 => string.Empty,
                1 => ", and 1 contributor was told",
                int count => $", and {count} contributors were told"
            };

            return $"{answer.BoardId} is deleted: {BoardDeletionWording.Files(answer.BetaFilesRemoved)} removed from BETA, " +
                $"{BoardDeletionWording.Files(answer.ProductionFilesRemoved)} from the stable data, " +
                $"{BoardDeletionWording.Count(answer.SubmissionsDeleted, "submission", "submissions")} deleted{told}.";
        }

        private static string TreePart(int files, bool listed)
        {
            if (files == 0 && !listed)
                return "nothing of it";

            if (files == 0)
                return "its row in the drop-down lists";

            return listed
                ? $"{BoardDeletionWording.Files(files)}, and its row in the drop-down lists"
                : BoardDeletionWording.Files(files);
        }

        private static string DatabasePart(BoardDeletePlanAnswer plan)
        {
            if (!plan.HasRecord)
                return "nothing - it has no record there";

            var parts = new List<string>
            {
                BoardDeletionWording.Count(plan.Submissions, "submission", "submissions"),
                BoardDeletionWording.Count(plan.Maintainers, "maintainer", "maintainers")
            };

            if (plan.Invitations > 0)
                parts.Add(BoardDeletionWording.Count(plan.Invitations, "open invitation", "open invitations"));

            return $"its record, with {string.Join(", ", parts.Take(parts.Count - 1))} and {parts[^1]}";
        }

        private static string Files(int count) => BoardDeletionWording.Count(count, "file", "files");

        private static string Count(int count, string one, string many) =>
            count == 1 ? $"1 {one}" : $"{count.ToString(CultureInfo.InvariantCulture)} {many}";
    }
}
