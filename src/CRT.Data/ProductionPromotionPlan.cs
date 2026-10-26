using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT PUBLISHING ONE BOARD FROM BETA TO PRODUCTION WILL COPY (owner request,
    // 2026-09-25: "it should be a two-fold process, where it is first published to BETA and then it
    // is published to the real production").
    //
    // Pure, like PublishPlan: the facts about both trees go in, a list of copies or a list of
    // refusals comes out. The server's ProductionPromoter performs the list; the Maintainer tab shows
    // the SAME records to the maintainer before they press the button, which is why they live here
    // in CRT.Data - one type on both ends of the wire, the SubmittedFileFact rule.
    //
    // *** ONLY BYTES THAT ARE ALREADY IN BETA. *** Nothing reaches Production except a file the
    // BETA tree holds right now, so the most a promotion can publish is what a maintainer could
    // already look at there. That is the whole of the argument for letting a maintainer do it.
    //
    // *** PER BOARD, NOT PER SUBMISSION. *** BETA is one tree: two submissions merged into one
    // board cannot be promoted separately, because the board's file already holds both.
    //
    // *** WHAT IS COPIED IN THE CRT 4.0 TREE (2026-10-25). *** A board is its board file
    // ("<board id>/board.json") and the stored files it lists ("Files/<ab>/<sha256>.<ext>",
    // DataTreeLayout). So:
    //   - every stored file the BETA board lists that Production does not hold yet. A stored file
    //     never changes - its name is its checksum - so one Production holds is already identical,
    //     and a file the board REPLACED is a new stored file beside the old;
    //   - the board file, when Production's differs - last, because it is what makes the copied
    //     files reachable (the publish's order: an interrupted promotion leaves files nothing points
    //     at, never a board pointing at files that are not there). Re-running repairs it.
    // No shared files and no other board's files any more (every board owns its files), so a
    // promotion never touches another board and needs no second approval for one.
    //
    // WHAT IS REFUSED:
    //   - a BETA board without a readable board file: there is nothing to publish;
    //   - a stored file the BETA board lists that BETA does not hold: the BETA board is broken, and
    //     a broken board is not promoted;
    //   - a board folder that differs from one already in Production only by capitalisation, for
    //     the reason PublishedTreeView gives: every Windows and macOS client folds the two together.
    //
    // THIS PLAN ONLY COPIES. What the promotion REMOVES from production - stored files no board
    // there lists once this board's file is copied - is ProductionPromotionFlow.PreviewRemovals,
    // shown beside this list and sent back with the approval (2026-09-25).
    // ###########################################################################################
    public static class ProductionPromotionPlan
    {
        // ###########################################################################################
        // Builds the list.
        //
        //   betaFiles            - the BETA board's files by name (its board file's list); null when
        //                          BETA has no readable board file.
        //   betaBoardFile        - the SHA-256 of BETA's board file; productionBoardFile the same
        //                          in Production, null when it has none.
        //   inBeta, inProduction - whether a tree holds a stored file (data-root-relative).
        //   productionVariant    - the folder Production holds that differs from a path only by
        //                          capitalisation, or null (PublishedTreeView.FindCaseVariant).
        // ###########################################################################################
        public static ProductionPromotionResult Build(
            string manufacturer,
            string hardware,
            string board,
            IReadOnlyList<BoardDocument.StoredFileRow>? betaFiles,
            string? betaBoardFile,
            string? productionBoardFile,
            Func<string, bool> inBeta,
            Func<string, bool> inProduction,
            Func<string, string?>? productionVariant = null)
        {
            ArgumentNullException.ThrowIfNull(inBeta);
            ArgumentNullException.ThrowIfNull(inProduction);

            string boardFolder = $"{manufacturer}/{hardware}/{board}";
            string boardFile = DataTreeLayout.BoardFilePath(boardFolder);

            var problems = new List<ValidationFinding>();

            if (betaFiles is null || string.IsNullOrWhiteSpace(betaBoardFile))
            {
                problems.Add(ProductionPromotionPlan.Error(
                    "promote.not_in_beta",
                    boardFolder,
                    $"[{boardFolder}] has nothing in the BETA data, so there is nothing to publish to the stable source."));

                return new ProductionPromotionResult([], 0, problems, TouchesSharedFiles: false);
            }

            var copies = new Dictionary<string, PromotionFile>(StringComparer.Ordinal);

            // Already identical in production - by path, for the file tree (owner request,
            // 2026-09-28), and counted for the summary.
            var unchanged = new List<string>();

            foreach (BoardDocument.StoredFileRow file in betaFiles)
            {
                if (!DataTreeLayout.IsSha256(file.Sha256))
                {
                    problems.Add(ProductionPromotionPlan.Error(
                        "promote.unreadable_in_beta", file.Name, $"The BETA board lists [{file.Name}] without a usable checksum. Fix the board in BETA first."));
                    continue;
                }

                string stored = DataTreeLayout.StoredFilePath(file.Sha256, file.Name);

                if (copies.ContainsKey(stored) || unchanged.Contains(stored, StringComparer.Ordinal))
                    continue;

                if (!inBeta(stored))
                {
                    problems.Add(ProductionPromotionPlan.Error(
                        "promote.cited_missing_in_beta",
                        file.Name,
                        $"The BETA board lists [{file.Name}], which is not in the BETA data. Fix the board in BETA first."));
                    continue;
                }

                if (inProduction(stored))
                {
                    unchanged.Add(stored);
                    continue;
                }

                copies[stored] = new PromotionFile(stored, file.Sha256, PromotionChange.Added, PromotionStage.Content, IsShared: false, Name: file.Name);
            }

            // ---- The board file, last -------------------------------------------------------------
            if (string.Equals(betaBoardFile, productionBoardFile, StringComparison.OrdinalIgnoreCase))
            {
                unchanged.Add(boardFile);
            }
            else
            {
                string? variant = productionBoardFile is null ? productionVariant?.Invoke(boardFolder) : null;

                if (variant is not null && !string.Equals(variant, boardFolder, StringComparison.Ordinal))
                {
                    problems.Add(ProductionPromotionPlan.Error(
                        "promote.case_collision",
                        boardFolder,
                        $"[{boardFolder}] differs only by capitalisation from [{variant}], which the stable source already has. " +
                        "On Windows and macOS the two would be the same folder."));
                }
                else
                {
                    copies[boardFile] = new PromotionFile(
                        boardFile,
                        betaBoardFile!.ToLowerInvariant(),
                        productionBoardFile is null ? PromotionChange.Added : PromotionChange.Replaced,
                        PromotionStage.Board,
                        IsShared: false,
                        Name: DataTreeLayout.BoardFileName);
                }
            }

            IReadOnlyList<PromotionFile> ordered = copies.Values
                .OrderBy(file => file.Stage)
                .ThenBy(file => file.Path, StringComparer.Ordinal)
                .ToList();

            return new ProductionPromotionResult(
                problems.Count == 0 ? ordered : [],
                unchanged.Count,
                problems,
                TouchesSharedFiles: false,
                Unchanged: unchanged.Order(StringComparer.Ordinal).ToList(),
                Names: ProductionPromotionPlan.NamesOf(boardFile, betaFiles));
        }

        // The board's name for each tree path of its files - what the file tree shows.
        internal static IReadOnlyDictionary<string, string> NamesOf(string boardFile, IReadOnlyList<BoardDocument.StoredFileRow> files)
        {
            var names = new Dictionary<string, string>(StringComparer.Ordinal) { [boardFile] = DataTreeLayout.BoardFileName };

            foreach (BoardDocument.StoredFileRow file in files.Where(file => DataTreeLayout.IsSha256(file.Sha256)))
                names.TryAdd(DataTreeLayout.StoredFilePath(file.Sha256, file.Name), file.Name);

            return names;
        }

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
    // One file a promotion will copy. On the wire to the Maintainer tab as this very record,
    // so the two cannot disagree about its fields. The enums travel as NAMES for the reason
    // SubmissionFileScope gives.
    //
    // `Name` (2026-10-25): the board's name for the file - Path is where it lies in the tree, which
    // for a 4.0 stored file is its checksum. Optional: an older server sends none.
    // ###########################################################################################
    public sealed record PromotionFile(
        string Path,
        string Sha256,
        PromotionChange Change,
        PromotionStage Stage,
        bool IsShared,
        string? Name = null);

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

    // `Unchanged`: the paths counted by UnchangedCount - the file tree's "already the same" (owner
    // request, 2026-09-28). Empty when the plan was refused before it looked at any file.
    // `Names`: the board's name for each tree path of its files (2026-10-25) - see PromotionFile.Name.
    public sealed record ProductionPromotionResult(
        IReadOnlyList<PromotionFile> Files,
        int UnchangedCount,
        IReadOnlyList<ValidationFinding> Problems,
        bool TouchesSharedFiles,
        IReadOnlyList<string>? Unchanged = null,
        IReadOnlyDictionary<string, string>? Names = null)
    {
        public bool CanPromote => this.Problems.Count == 0;
    }
}
