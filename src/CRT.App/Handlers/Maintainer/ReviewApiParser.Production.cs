using System;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // ReviewApiParser, part three: "Beta > Prod" - a production plan, pushing a board back out of BETA,
    // the publish itself - and the administrator's unused files. Split out of ReviewApiParser.cs when it
    // passed the project's ~1,500 lines (code review, 2026-09-27); the rules in that file's header hold here.
    // ###########################################################################################
    public static partial class ReviewApiParser
    {
        public static ProductionPlanView? ParseProductionPlan(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? systemId = ReviewApiParser.String(root, "systemId");

            if (string.IsNullOrWhiteSpace(systemId))
                return null;

            var files = new List<PromotionFile>();

            if (root.TryGetProperty("files", out JsonElement raw) && raw.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement element in raw.EnumerateArray())
                {
                    try
                    {
                        PromotionFile? file = element.Deserialize<PromotionFile>(ReviewApiParser.FactOptions);

                        if (file is not null && !string.IsNullOrWhiteSpace(file.Path))
                            files.Add(file);
                    }
                    catch (JsonException)
                    {
                        // One unreadable entry is dropped; the count the maintainer sees is then
                        // short, which is visible, rather than the whole plan being lost.
                    }
                }
            }

            var problems = new List<ReviewFindingView>();

            if (root.TryGetProperty("problems", out JsonElement list) && list.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement item in list.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.Object)
                        continue;

                    problems.Add(new ReviewFindingView(
                        ReviewApiParser.String(item, "code") ?? string.Empty,
                        ReviewApiParser.String(item, "subject") ?? string.Empty,
                        ReviewApiParser.String(item, "message") ?? string.Empty,
                        ReviewApiParser.IsError(item)));
                }
            }

            return new ProductionPlanView(
                systemId,
                ReviewApiParser.String(root, "betaRevision"),
                ReviewApiParser.String(root, "betaContentHash") ?? string.Empty,
                ReviewApiParser.Bool(root, "touchesSharedFiles") ?? false,

                // Defaults FALSE: a button enabled on a missing answer would offer a publish the
                // server never agreed to.
                ReviewApiParser.Bool(root, "canPublish") ?? false,
                ReviewApiParser.String(root, "refusal"),
                (int)(ReviewApiParser.Long(root, "unchangedCount") ?? 0),
                files,
                problems,
                ReviewApiParser.ParseApproval(root),
                ReviewApiParser.ParseRemovals(root),
                ReviewApiParser.ParseCarrying(root),
                ReviewApiParser.ParseStrings(root, "unchangedFiles"),
                ReviewApiParser.String(root, "betaDataUrl"),
                ReviewApiParser.String(root, "productionDataUrl"),
                ReviewApiParser.ParseSizes(root, "fileSizes"));
        }

        // ###########################################################################################
        // Path -> bytes (2026-10-04, the plan's FileSizes). Null when absent - an older server - and
        // an entry that is not a whole, non-negative number is left out rather than shown wrong.
        // ###########################################################################################
        private static IReadOnlyDictionary<string, long>? ParseSizes(JsonElement root, string property)
        {
            if (!root.TryGetProperty(property, out JsonElement raw) || raw.ValueKind != JsonValueKind.Object)
                return null;

            var sizes = new Dictionary<string, long>(StringComparer.Ordinal);

            foreach (JsonProperty entry in raw.EnumerateObject())
            {
                if (entry.Value.ValueKind == JsonValueKind.Number && entry.Value.TryGetInt64(out long size) && size >= 0)
                    sizes[entry.Name] = size;
            }

            return sizes;
        }

        // ###########################################################################################
        // A submission's file tree (2026-09-28): CRT.Data's SubmissionFilesAnswer, read as that very
        // record. An entry with no path is dropped; an answer with no system is not an answer.
        // ###########################################################################################
        public static SubmissionFilesAnswer? ParseSubmissionFiles(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            try
            {
                SubmissionFilesAnswer? answer = root.Deserialize<SubmissionFilesAnswer>(ReviewApiParser.FactOptions);

                if (answer is null || string.IsNullOrWhiteSpace(answer.SystemId))
                    return null;

                return answer with
                {
                    Files = (answer.Files ?? []).Where(file => file is not null && !string.IsNullOrWhiteSpace(file.Path)).ToList()
                };
            }
            catch (JsonException)
            {
                return null;
            }
        }

        // ###########################################################################################
        // The merged submissions this promotion would carry (owner request, 2026-09-27). Absent on
        // an older server, which reads as "none named" rather than as a failure - the panel then
        // shows what it always did.
        // ###########################################################################################
        private static IReadOnlyList<CarriedSubmission> ParseCarrying(JsonElement root) =>
            ReviewApiParser.ParseCarryingFrom(root, "carrying");

        // The same shape under two names: the plan's `carrying`, and a rollback's `returning`.
        private static IReadOnlyList<CarriedSubmission> ParseCarryingFrom(JsonElement root, string property)
        {
            var carrying = new List<CarriedSubmission>();

            if (!root.TryGetProperty(property, out JsonElement list) || list.ValueKind != JsonValueKind.Array)
                return carrying;

            foreach (JsonElement item in list.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.Object)
                    continue;

                carrying.Add(new CarriedSubmission(
                    ReviewApiParser.Long(item, "id") ?? 0,
                    ReviewApiParser.String(item, "contactEmail") ?? string.Empty,
                    ReviewApiParser.String(item, "comment"),
                    ReviewApiParser.Time(item, "decidedUtc"),

                    // Its contributor discarded their own draft since (2026-09-28).
                    ReviewApiParser.Time(item, "draftDiscardedUtc")));
            }

            return carrying;
        }

        // ###########################################################################################
        // What a rollback would do, and what it did (2026-09-27). `kind` is the server's own word -
        // "restore" or "remove" - rather than a parsed enum name, so the wire does not depend on a
        // C# identifier.
        // ###########################################################################################
        public static BetaRollbackPlanView? ParseBetaRollbackPlan(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? systemId = ReviewApiParser.String(root, "systemId");

            if (string.IsNullOrWhiteSpace(systemId))
                return null;

            return new BetaRollbackPlanView(
                systemId,

                // Defaults to "restore": "remove" deletes a board from BETA, so a missing or
                // unreadable answer must never be taken as that one.
                string.Equals(ReviewApiParser.String(root, "kind"), "remove", StringComparison.OrdinalIgnoreCase)
                    ? BetaRollbackKind.RemoveFromBeta
                    : BetaRollbackKind.RestoreFromProduction,
                ReviewApiParser.Strings(root, "restored"),
                ReviewApiParser.Strings(root, "removed"),
                ReviewApiParser.ParseCarryingFrom(root, "returning"),
                ReviewApiParser.Strings(root, "sharedRestored"));
        }

        public static BetaRollbackResult? ParseBetaRollback(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? systemId = ReviewApiParser.String(root, "systemId");

            if (string.IsNullOrWhiteSpace(systemId))
                return null;

            return new BetaRollbackResult(
                systemId,
                string.Equals(ReviewApiParser.String(root, "kind"), "remove", StringComparison.OrdinalIgnoreCase)
                    ? BetaRollbackKind.RemoveFromBeta
                    : BetaRollbackKind.RestoreFromProduction,
                (int)(ReviewApiParser.Long(root, "filesRestored") ?? 0),
                (int)(ReviewApiParser.Long(root, "filesRemoved") ?? 0),
                (int)(ReviewApiParser.Long(root, "submissionsReturned") ?? 0),
                ReviewApiParser.Bool(root, "rejected") ?? false);
        }

        public static ProductionPublishResult? ParseProductionPublish(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? systemId = ReviewApiParser.String(root, "systemId");

            return string.IsNullOrWhiteSpace(systemId)
                ? null
                : new ProductionPublishResult(
                    systemId,
                    ReviewApiParser.String(root, "revision"),
                    (int)(ReviewApiParser.Long(root, "filesCopied") ?? 0),
                    ReviewApiParser.String(root, "state") ?? "published",
                    ReviewApiParser.ParseRoles(root, "waitingFor"),
                    ReviewApiParser.ParseStrings(root, "removedFiles"));
        }

        // ###########################################################################################
        // The administrator's "Unused files" list - CRT.Data's UnusedFileListing, as the server
        // wrote it. Null when the answer has no tree name, which is not a usable answer.
        // ###########################################################################################
        public static UnusedFileListing? ParseUnusedFiles(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            try
            {
                UnusedFileListing? listing = root.Deserialize<UnusedFileListing>(ReviewApiParser.FactOptions);

                return listing is null || string.IsNullOrWhiteSpace(listing.Tree)
                    ? null
                    : listing with { Problems = listing.Problems ?? [], Files = listing.Files ?? [] };
            }
            catch (JsonException)
            {
                return null;
            }
        }

        public static UnusedFileRemovalResult? ParseUnusedFileRemoval(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? tree = ReviewApiParser.String(root, "tree");

            return string.IsNullOrWhiteSpace(tree)
                ? null
                : new UnusedFileRemovalResult(
                    tree,
                    ReviewApiParser.ParseStrings(root, "removed"),
                    ReviewApiParser.ParseStrings(root, "kept"),
                    ReviewApiParser.String(root, "notDoneBecause"));
        }

        // ###########################################################################################
        // POST /api/admin/manifest/rebuild (2026-10-01). Every word shown comes from the server, so
        // a server that gains a third tree or a new reason to skip one says so with no CRT release.
        // A missing headline is the only thing refused: an answer with nothing to say is not one.
        // ###########################################################################################
        public static ManifestRebuildResult? ParseManifestRebuild(string? json)
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            string? headline = ReviewApiParser.String(root, "headline");

            if (string.IsNullOrWhiteSpace(headline))
                return null;

            var trees = new List<ManifestRebuildTreeResult>();

            if (root.TryGetProperty("trees", out JsonElement listed) && listed.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement entry in listed.EnumerateArray())
                {
                    if (entry.ValueKind != JsonValueKind.Object)
                        continue;

                    string tree = ReviewApiParser.String(entry, "tree") ?? string.Empty;
                    string message = ReviewApiParser.String(entry, "message") ?? string.Empty;

                    trees.Add(new ManifestRebuildTreeResult(
                        tree,
                        ReviewApiParser.Bool(entry, "skipped") ?? false,
                        (int)(ReviewApiParser.Long(entry, "entries") ?? -1),
                        message));
                }
            }

            return new ManifestRebuildResult(headline, trees);
        }
    }
}
