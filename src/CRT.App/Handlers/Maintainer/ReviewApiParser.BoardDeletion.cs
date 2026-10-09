using System.Collections.Generic;
using System.Text.Json;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // ReviewApiParser, the administrator's "Delete a board" (owner request, 2026-10-03): the plan
    // shown in the confirmation, and what the delete did. Both read into CRT.Data's own records
    // (BoardDeletePlanAnswer, BoardDeleteAnswer), which the server writes - ReviewWireContractTests
    // holds the two ends together. The rules in ReviewApiParser.cs's header hold here.
    // ###########################################################################################
    public static partial class ReviewApiParser
    {
        // ###########################################################################################
        // POST /api/admin/boards/delete/plan. Refused without a board id or a fingerprint - a plan
        // that cannot be sent back is not one the delete would accept. A missing count reads 0; a
        // missing open-submission list reads as none.
        // ###########################################################################################
        public static BoardDeletePlanAnswer? ParseBoardDeletePlan(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? boardId = ReviewApiParser.String(root, "boardId");
            string? fingerprint = ReviewApiParser.String(root, "fingerprint");

            if (string.IsNullOrWhiteSpace(boardId) || string.IsNullOrWhiteSpace(fingerprint))
                return null;

            var open = new List<BoardDeleteOpenSubmission>();

            foreach (JsonElement entry in ReviewApiParser.Objects(root, "openSubmissions"))
            {
                if (ReviewApiParser.Long(entry, "id") is not long id)
                    continue;

                open.Add(new BoardDeleteOpenSubmission(
                    id,
                    ReviewApiParser.String(entry, "state") ?? string.Empty,
                    ReviewApiParser.String(entry, "contributor") ?? string.Empty,
                    ReviewApiParser.String(entry, "summary"),
                    ReviewApiParser.Time(entry, "createdUtc") ?? default));
            }

            return new BoardDeletePlanAnswer(
                boardId,
                ReviewApiParser.String(root, "manufacturer") ?? string.Empty,
                ReviewApiParser.String(root, "hardware") ?? string.Empty,
                ReviewApiParser.String(root, "board") ?? string.Empty,
                fingerprint,
                ReviewApiParser.Count(root, "betaFiles"),
                ReviewApiParser.Count(root, "productionFiles"),
                ReviewApiParser.Bool(root, "listedInBeta") ?? false,
                ReviewApiParser.Bool(root, "listedInProduction") ?? false,
                ReviewApiParser.Bool(root, "hasRecord") ?? false,
                ReviewApiParser.Count(root, "submissions"),
                ReviewApiParser.Count(root, "maintainers"),
                ReviewApiParser.Count(root, "invitations"),
                open,
                ReviewApiParser.String(root, "blockedBecause"));
        }

        // POST /api/admin/boards/delete. Refused without the board id it is about.
        public static BoardDeleteAnswer? ParseBoardDelete(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? boardId = ReviewApiParser.String(root, "boardId");

            return string.IsNullOrWhiteSpace(boardId)
                ? null
                : new BoardDeleteAnswer(
                    boardId,
                    ReviewApiParser.Count(root, "betaFilesRemoved"),
                    ReviewApiParser.Count(root, "productionFilesRemoved"),
                    ReviewApiParser.Count(root, "submissionsDeleted"),
                    ReviewApiParser.Count(root, "contributorsMailed"));
        }

        private static int Count(JsonElement element, string name) =>
            ReviewApiParser.Long(element, name) is long value && value is >= 0 and <= int.MaxValue ? (int)value : 0;
    }
}
