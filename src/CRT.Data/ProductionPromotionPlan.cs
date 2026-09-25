using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT PUBLISHING ONE SYSTEM FROM BETA TO PRODUCTION WILL COPY (maintainer request,
    // 2026-09-25: "it should be a two-fold process, where it is first published to BETA and then it
    // is published to the real production").
    //
    // Pure, like PublishPlan: the facts about both trees go in as delegates, and a list of copies or
    // a list of refusals comes out. The server's ProductionPromoter performs the list; the review
    // application shows the SAME records to the reviewer before they press the button, which is why
    // they live here in CRT.Data - one type on both ends of the wire, the SubmittedFileFact rule.
    //
    // *** ONLY BYTES THAT ARE ALREADY IN BETA. *** Nothing reaches Production except a file the
    // BETA tree holds right now, so the most a promotion can publish is what a reviewer could
    // already look at there. That is the whole of the argument for letting a reviewer do it.
    //
    // *** PER SYSTEM, NOT PER SUBMISSION. *** BETA is one tree: two submissions merged into one
    // board cannot be promoted separately, because the board's workbook already holds both.
    //
    // WHAT IS COPIED:
    //   - every file under the system's own BETA folder ("<Manufacturer>/<Hardware>/<Board>/...")
    //     that Production lacks or holds different bytes for - the workbook and sidecar of every
    //     generation, images, attachments, KiCad files. Never a system.json: that file is retired
    //     (2026-09-25) and a copy an earlier build left in BETA is not carried;
    //   - every SHARED file the BETA board cites ("<Manufacturer>/Shared files/...", "Generic
    //     shared files/...") that Production lacks or holds differently. Those reach every board
    //     citing them, so a promotion that changes one is the ADMINISTRATOR's (TouchesSharedFiles).
    //
    // WHAT IS REFUSED:
    //   - a cited file of ANOTHER board that is not already identical in Production: the board
    //     would go out pointing at a file Production does not have. That board goes first.
    //   - a cited file of this board or a shared folder that BETA itself does not have: the BETA
    //     board is broken, and a broken board is not promoted.
    //   - a path that differs from one already in Production only by capitalisation, for the
    //     reason PublishedTreeView gives: every Windows and macOS client folds the two together.
    //
    // THIS PLAN ONLY COPIES. What the promotion REMOVES from production - files the board stops
    // citing that nothing else there uses - is ProductionPromotionFlow.PreviewRemovals, shown beside
    // this list and sent back with the approval (2026-09-25).
    //
    // THE ORDER is the publish's, and for the same reason - an interrupted promotion must leave a
    // board that loads: content files first, then the board's workbooks and sidecars, which make
    // those files reachable. Re-running the same promotion repairs an interrupted one, because
    // every copy overwrites and is verified.
    // ###########################################################################################
    public static class ProductionPromotionPlan
    {
        // ###########################################################################################
        // Builds the list.
        //
        //   ownFiles     - every file under the system's folder in BETA, data-root-relative with
        //                  forward slashes, as the server's walk found them.
        //   citedFiles   - every file the BETA board's rows cite (SubmissionManifestBuilder.
        //                  CollectReferencedFiles over the BETA board).
        //   beta         - hashes in BETA; production - hashes and case variants in Production.
        // ###########################################################################################
        public static ProductionPromotionResult Build(
            string manufacturer,
            string hardware,
            string board,
            IEnumerable<string> ownFiles,
            IEnumerable<string> citedFiles,
            PublishedTreeView beta,
            PublishedTreeView production)
        {
            ArgumentNullException.ThrowIfNull(ownFiles);
            ArgumentNullException.ThrowIfNull(citedFiles);
            ArgumentNullException.ThrowIfNull(beta);
            ArgumentNullException.ThrowIfNull(production);

            string systemFolder = $"{manufacturer}/{hardware}/{board}";

            var problems = new List<ValidationFinding>();
            var copies = new Dictionary<string, PromotionFile>(StringComparer.Ordinal);
            int unchanged = 0;

            string[] own = ownFiles
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => path.Replace('\\', '/'))
                .Where(path => !ProductionPromotionPlan.IsHidden(path))
                .Where(path => !ProductionPromotionPlan.IsRetired(systemFolder, path))
                .Distinct(StringComparer.Ordinal)
                .ToArray();

            // For "is this cited file one the walk found?" below - once per cited file, so a set
            // rather than a scan of `own` (a large board cites ~2,000 of its ~2,200 files).
            var ownSet = new HashSet<string>(own, StringComparer.Ordinal);

            if (own.Length == 0)
            {
                problems.Add(ProductionPromotionPlan.Error(
                    "promote.not_in_beta",
                    systemFolder,
                    $"[{systemFolder}] has nothing in the BETA data, so there is nothing to publish to production."));

                return new ProductionPromotionResult([], 0, problems, TouchesSharedFiles: false);
            }

            // ---- The system's own folder -------------------------------------------------------
            foreach (string path in own)
            {
                if (!path.StartsWith(systemFolder + "/", StringComparison.Ordinal))
                {
                    // The walk handed over something outside the folder it was asked for. Not
                    // this system's to promote, whatever it is.
                    problems.Add(ProductionPromotionPlan.Error(
                        "promote.outside_system", path, $"[{path}] is not inside [{systemFolder}]."));
                    continue;
                }

                if (ProductionPromotionPlan.TryPlan(path, beta, production, problems, out PromotionFile? file))
                {
                    if (file is null)
                        unchanged++;
                    else
                        copies[path] = file;
                }
            }

            // ---- What the board cites outside its own folder -----------------------------------
            foreach (string cited in citedFiles
                         .Where(path => !string.IsNullOrWhiteSpace(path))
                         .Select(path => path.Replace('\\', '/'))
                         .Distinct(StringComparer.Ordinal))
            {
                SubmissionFileScope scope = SubmissionFileScopes.Classify(manufacturer, hardware, board, cited);

                if (scope == SubmissionFileScope.Own)
                {
                    // Already walked. A cited own file the walk did not find is a BETA board
                    // pointing at nothing.
                    if (!ownSet.Contains(cited))
                    {
                        problems.Add(ProductionPromotionPlan.Error(
                            "promote.cited_missing_in_beta",
                            cited,
                            $"The BETA board cites [{cited}], which is not in the BETA data. Fix the board in BETA first."));
                    }

                    continue;
                }

                if (ProductionPromotionPlan.IsHidden(cited))
                {
                    problems.Add(ProductionPromotionPlan.Error(
                        "promote.hidden_path", cited, $"[{cited}] is a hidden path and is never published."));
                    continue;
                }

                if (scope == SubmissionFileScope.Foreign)
                {
                    string? inBeta = beta.HashOf(cited);
                    string? inProduction = production.HashOf(cited);

                    if (inBeta is null || !string.Equals(inBeta, inProduction, StringComparison.Ordinal))
                    {
                        problems.Add(ProductionPromotionPlan.Error(
                            "promote.foreign_not_in_production",
                            cited,
                            $"The board uses [{cited}], which belongs to another board and is not yet in production " +
                            "in the same form. Publish that board to production first."));
                    }

                    continue;
                }

                // Shared: copied like an own file, and noted, because it reaches every board.
                if (beta.HashOf(cited) is null)
                {
                    problems.Add(ProductionPromotionPlan.Error(
                        "promote.cited_missing_in_beta",
                        cited,
                        $"The BETA board cites [{cited}], which is not in the BETA data. Fix the board in BETA first."));
                    continue;
                }

                if (ProductionPromotionPlan.TryPlan(cited, beta, production, problems, out PromotionFile? shared))
                {
                    if (shared is null)
                        unchanged++;
                    else
                        copies[cited] = shared with { IsShared = true };
                }
            }

            IReadOnlyList<PromotionFile> ordered = copies.Values
                .Select(file => file with { Stage = ProductionPromotionPlan.StageOf(systemFolder, file.Path) })
                .OrderBy(file => file.Stage)
                .ThenBy(file => file.Path, StringComparer.Ordinal)
                .ToList();

            return new ProductionPromotionResult(
                problems.Count == 0 ? ordered : [],
                unchanged,
                problems,
                TouchesSharedFiles: ordered.Any(file => file.IsShared));
        }

        // ###########################################################################################
        // One file: null out-param for "already identical in Production", a PromotionFile for a
        // copy, false (with a problem recorded) for a refusal.
        // ###########################################################################################
        private static bool TryPlan(
            string path,
            PublishedTreeView beta,
            PublishedTreeView production,
            List<ValidationFinding> problems,
            out PromotionFile? file)
        {
            file = null;

            string? betaHash = beta.HashOf(path);

            if (betaHash is null)
            {
                problems.Add(ProductionPromotionPlan.Error(
                    "promote.unreadable_in_beta", path, $"[{path}] could not be read in the BETA data."));
                return false;
            }

            string? productionHash = production.HashOf(path);

            if (string.Equals(betaHash, productionHash, StringComparison.Ordinal))
                return true;

            if (productionHash is null)
            {
                string? variant = production.FindCaseVariant(path);

                if (variant is not null && !string.Equals(variant, path, StringComparison.Ordinal))
                {
                    problems.Add(ProductionPromotionPlan.Error(
                        "promote.case_collision",
                        path,
                        $"[{path}] differs only by capitalisation from [{variant}], which production already has. " +
                        "On Windows and macOS the two would be the same file."));
                    return false;
                }
            }

            file = new PromotionFile(
                path,
                betaHash,
                productionHash is null ? PromotionChange.Added : PromotionChange.Replaced,
                PromotionStage.Content,
                IsShared: false);

            return true;
        }

        // ###########################################################################################
        // Which of the two write stages a file belongs to. The board's workbooks and sidecars are
        // the .xlsx and .json files at the folder's TOP level, and go after everything they cite.
        // ###########################################################################################
        internal static PromotionStage StageOf(string systemFolder, string path)
        {
            if (!path.StartsWith(systemFolder + "/", StringComparison.Ordinal))
                return PromotionStage.Content;

            string rest = path[(systemFolder.Length + 1)..];

            if (rest.Contains('/'))
                return PromotionStage.Content;

            return rest.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                   rest.EndsWith(".json", StringComparison.OrdinalIgnoreCase)
                ? PromotionStage.Board
                : PromotionStage.Content;
        }

        // A dot-segment anywhere: temporary files a publish writes beside its target, and anything
        // else hidden. Never published, so never promoted.
        private static bool IsHidden(string path) =>
            path.Split('/').Any(segment => segment.StartsWith('.'));

        // The retired system.json at the board folder's top level - see SystemDescriptorStore. The
        // promotion removes one from production instead of carrying one there.
        private static bool IsRetired(string systemFolder, string path) =>
            string.Equals(path, $"{systemFolder}/{SystemDescriptorStore.FileName}", StringComparison.Ordinal);

        private static ValidationFinding Error(string code, string subject, string message) =>
            new()
            {
                Severity = ValidationSeverity.Error,
                Code = code,
                Subject = subject,
                Message = message
            };
    }

    // ###########################################################################################
    // One file a promotion will copy. On the wire to the review application as this very record,
    // so the two cannot disagree about its fields. The enums travel as NAMES for the reason
    // SubmissionFileScope gives.
    // ###########################################################################################
    public sealed record PromotionFile(
        string Path,
        string Sha256,
        PromotionChange Change,
        PromotionStage Stage,
        bool IsShared);

    [JsonConverter(typeof(JsonStringEnumConverter<PromotionChange>))]
    public enum PromotionChange
    {
        Added,
        Replaced
    }

    // In write order.
    [JsonConverter(typeof(JsonStringEnumConverter<PromotionStage>))]
    public enum PromotionStage
    {
        Content,
        Board
    }

    public sealed record ProductionPromotionResult(
        IReadOnlyList<PromotionFile> Files,
        int UnchangedCount,
        IReadOnlyList<ValidationFinding> Problems,
        bool TouchesSharedFiles)
    {
        public bool CanPromote => this.Problems.Count == 0;
    }
}
