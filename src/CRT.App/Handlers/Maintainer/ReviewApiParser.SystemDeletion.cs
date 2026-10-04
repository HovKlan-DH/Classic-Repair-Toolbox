using System.Collections.Generic;
using System.Text.Json;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // ReviewApiParser, the administrator's "Delete a system" (owner request, 2026-10-03): the plan
    // shown in the confirmation, and what the delete did. Both read into CRT.Data's own records
    // (SystemDeletePlanAnswer, SystemDeleteAnswer), which the server writes - ReviewWireContractTests
    // holds the two ends together. The rules in ReviewApiParser.cs's header hold here.
    // ###########################################################################################
    public static partial class ReviewApiParser
    {
        // ###########################################################################################
        // POST /api/admin/systems/delete/plan. Refused without a system id or a fingerprint - a plan
        // that cannot be sent back is not one the delete would accept. A missing count reads 0; a
        // missing open-submission list reads as none.
        // ###########################################################################################
        public static SystemDeletePlanAnswer? ParseSystemDeletePlan(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? systemId = ReviewApiParser.String(root, "systemId");
            string? fingerprint = ReviewApiParser.String(root, "fingerprint");

            if (string.IsNullOrWhiteSpace(systemId) || string.IsNullOrWhiteSpace(fingerprint))
                return null;

            var open = new List<SystemDeleteOpenSubmission>();

            foreach (JsonElement entry in ReviewApiParser.Objects(root, "openSubmissions"))
            {
                if (ReviewApiParser.Long(entry, "id") is not long id)
                    continue;

                open.Add(new SystemDeleteOpenSubmission(
                    id,
                    ReviewApiParser.String(entry, "state") ?? string.Empty,
                    ReviewApiParser.String(entry, "contributor") ?? string.Empty,
                    ReviewApiParser.String(entry, "summary"),
                    ReviewApiParser.Time(entry, "createdUtc") ?? default));
            }

            return new SystemDeletePlanAnswer(
                systemId,
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

        // POST /api/admin/systems/delete. Refused without the system id it is about.
        public static SystemDeleteAnswer? ParseSystemDelete(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? systemId = ReviewApiParser.String(root, "systemId");

            return string.IsNullOrWhiteSpace(systemId)
                ? null
                : new SystemDeleteAnswer(
                    systemId,
                    ReviewApiParser.Count(root, "betaFilesRemoved"),
                    ReviewApiParser.Count(root, "productionFilesRemoved"),
                    ReviewApiParser.Count(root, "submissionsDeleted"),
                    ReviewApiParser.Count(root, "contributorsMailed"));
        }

        private static int Count(JsonElement element, string name) =>
            ReviewApiParser.Long(element, name) is long value && value is >= 0 and <= int.MaxValue ? (int)value : 0;
    }
}
