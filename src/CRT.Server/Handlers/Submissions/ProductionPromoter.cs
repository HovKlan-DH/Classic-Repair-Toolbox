using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // PERFORMS a production promotion: copies the files ProductionPromotionPlan listed from BETA
    // to Production (2026-09-25). The I/O half; the plan decided WHAT and in which order.
    //
    // *** EVERYTHING IS CHECKED BEFORE THE FIRST BYTE LANDS, as PublishExecutor does. *** Every
    // source and destination resolves inside its own tree through SubmissionPathRules, and neither
    // path passes through a symbolic link (PublishPathSafety) - a link in either tree would make a
    // perfectly contained path read from, or write to, somewhere else. A refusal found at file 600
    // would otherwise leave Production half-promoted.
    //
    // *** EACH COPY IS VERIFIED AGAINST THE HASH THE PLAN SAW. *** The plan read BETA; a file that
    // changed in BETA since (a hand edit, a publish racing without the lock) must not reach every
    // user under the name of the file the maintainer checked. VerifiedFileCopy hashes as it writes
    // and replaces the real file only on a match.
    //
    // NOT ATOMIC, and it cannot be - the same position PublishExecutor is in. The plan's ORDER is
    // what makes an interruption recoverable (content, then workbooks, then system.json), and
    // running the same promotion again repairs it.
    // ###########################################################################################
    public static class ProductionPromoter
    {
        // `canWriteFolder` is for tests - TreeWriteAccess's real probe otherwise.
        public static async Task<PromotionCopyOutcome> CopyAsync(
            IReadOnlyList<PromotionFile> files,
            string betaRoot,
            string productionRoot,
            CancellationToken cancellationToken = default,
            Func<string, bool>? canWriteFolder = null)
        {
            ArgumentNullException.ThrowIfNull(files);
            ArgumentException.ThrowIfNullOrWhiteSpace(betaRoot);
            ArgumentException.ThrowIfNullOrWhiteSpace(productionRoot);

            // ---- 0. Every path, before anything is written -----------------------------------
            var resolved = new List<(PromotionFile File, string Source, string Destination)>();

            foreach (PromotionFile file in files)
            {
                if (!SubmissionPathRules.TryResolve(betaRoot, file.Path, out string source, out string sourceReason))
                    return PromotionCopyOutcome.Failed($"[{file.Path}] cannot be read from BETA: {sourceReason} Nothing was copied.");

                if (!SubmissionPathRules.TryResolve(productionRoot, file.Path, out string destination, out string destinationReason))
                    return PromotionCopyOutcome.Failed($"[{file.Path}] cannot be written to the stable source: {destinationReason} Nothing was copied.");

                string? link =
                    PublishPathSafety.FindLinkOnPath(betaRoot, source, PublishPathSafety.IsLink) ??
                    PublishPathSafety.FindLinkOnPath(productionRoot, destination, PublishPathSafety.IsLink);

                if (link is not null)
                {
                    return PromotionCopyOutcome.Failed(
                        $"There is a symbolic link at [{link}], and copying through it could read or write outside " +
                        "the data trees. Nothing was copied. Remove the link on the server.");
                }

                if (!File.Exists(source))
                    return PromotionCopyOutcome.Failed($"[{file.Path}] is no longer in BETA. Nothing was copied.");

                resolved.Add((file, source, destination));
            }

            // Every folder it copies into may be written - a folder copied into production by hand
            // as another user refuses them all, and finding that out at file 600 leaves production
            // half-promoted (owner report, 2026-09-28). After the link check; see TreeWriteAccess.
            IReadOnlyList<string> refusing = TreeWriteAccess.FoldersRefusing(
                productionRoot, resolved.Select(entry => entry.Destination), canWriteFolder);

            if (refusing.Count > 0)
                return PromotionCopyOutcome.Refused(TreeWriteAccess.RefusalMessage("production", productionRoot, refusing), refusing);

            // ---- 1. The copies, in the plan's order ------------------------------------------
            int copied = 0;

            foreach ((PromotionFile file, string source, string destination) in resolved)
            {
                cancellationToken.ThrowIfCancellationRequested();

                VerifiedCopyResult result;

                try
                {
                    result = await VerifiedFileCopy
                        .CopyAsync(source, file.Sha256, destination, cancellationToken: cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    return PromotionCopyOutcome.Failed(
                        $"[{file.Path}] could not be written to the stable source ({ex.Message}). " +
                        ProductionPromoter.PartialSentence(copied),
                        copied);
                }

                if (result != VerifiedCopyResult.Copied)
                {
                    return PromotionCopyOutcome.Failed(
                        $"[{file.Path}] changed in BETA after it was checked, so it was not copied. " +
                        ProductionPromoter.PartialSentence(copied),
                        copied);
                }

                copied++;
            }

            return PromotionCopyOutcome.Done(copied);
        }

        private static string PartialSentence(int copied) =>
            copied == 0
                ? "Nothing was copied."
                : $"{copied} file(s) before it were already copied; publishing to the stable source again completes it.";
    }

    public sealed record PromotionCopyOutcome(bool IsDone, int FilesCopied, string? Error)
    {
        public static PromotionCopyOutcome Done(int copied) => new(true, copied, null);

        public static PromotionCopyOutcome Failed(string error, int copied = 0) => new(false, copied, error);

        // Refused before anything was copied because these folders may not be written - kept so
        // the flow can log the command that fixes them.
        public static PromotionCopyOutcome Refused(string error, IReadOnlyList<string> foldersRefusing) =>
            new(false, 0, error) { FoldersRefusing = foldersRefusing };

        public IReadOnlyList<string> FoldersRefusing { get; init; } = [];
    }
}
