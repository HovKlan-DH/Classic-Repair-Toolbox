using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // PERFORMS a BETA rollback (owner decision, 2026-09-27): copies production's files over BETA's
    // and deletes the ones only BETA had. The I/O half; BetaRollbackPlan decided what.
    //
    // *** EVERY PATH IS RESOLVED BEFORE THE FIRST BYTE MOVES, as ProductionPromoter does. *** Both
    // sides resolve inside their own tree through SubmissionPathRules and neither passes through a
    // symbolic link, so a link in either tree cannot make a contained path read from or write to
    // somewhere else. A refusal found at file 600 would otherwise leave BETA half-rolled-back.
    //
    // *** EVERY RESTORE IS A VERIFIED COPY (code review, 2026-09-27). *** The first version used a
    // bare File.Copy onto the served path: a CRT syncing mid-copy downloaded a truncated scan, a
    // crash left that truncated file in BETA, and a production file changed by hand between the
    // plan and the copy went out unseen. VerifiedFileCopy - the promotion's and the publish's own -
    // writes a hidden temporary file, hashes it as it writes, and replaces the real file in one
    // rename only when the bytes are the ones the plan compared. Anything else refuses.
    //
    // *** RESTORE FIRST, THEN REMOVE. *** Stopping midway through the restores leaves files that are
    // each one of two known-good versions; stopping midway through the removals leaves files nothing
    // cites - untidy, but every board still loads. Re-running repairs either.
    //
    // *** NO SHARED FILE IS EVER REMOVED HERE (owner decision, 2026-09-27). *** A shared file is
    // only ever RESTORED to production's bytes; one the rolled-back submissions added stays, unused,
    // for Account > Unused files (AutomaticRemovalScope). The system's own BETA-only files are removed
    // directly: production's workbook, now in place, does not cite them.
    // ###########################################################################################
    public static class BetaRollbackWriter
    {
        public const string TemporaryTag = "crt-rollback";

        public static async Task<BetaRollbackWriteOutcome> ApplyAsync(
            BetaRollbackFilePlan filePlan,
            string betaRoot,
            string productionRoot,
            SystemRecord system,
            ILogger logger,
            CancellationToken cancellationToken = default,
            Func<string, bool>? canWriteFolder = null)
        {
            ArgumentNullException.ThrowIfNull(filePlan);
            ArgumentException.ThrowIfNullOrWhiteSpace(betaRoot);
            ArgumentException.ThrowIfNullOrWhiteSpace(productionRoot);
            ArgumentNullException.ThrowIfNull(system);
            ArgumentNullException.ThrowIfNull(logger);

            BetaRollbackPlanResult plan = filePlan.Plan;

            // The flow refuses this before calling; never act on it here either.
            if (plan.Kind == BetaRollbackKind.ProductionUnreadable)
                return BetaRollbackWriteOutcome.Failed(BetaRollbackFlow.ProductionUnreadableMessage);

            // ---- 0. Every path, before anything is written ------------------------------------
            var restores = new List<(string Path, string Source, string Destination, string Hash)>();

            foreach (string path in plan.AllRestored)
            {
                if (!SubmissionPathRules.TryResolve(productionRoot, path, out string source, out string sourceWhy))
                    return BetaRollbackWriteOutcome.Failed($"[{path}] cannot be read from the stable source: {sourceWhy} Nothing was changed.");

                if (!SubmissionPathRules.TryResolve(betaRoot, path, out string destination, out string destinationWhy))
                    return BetaRollbackWriteOutcome.Failed($"[{path}] cannot be written to BETA: {destinationWhy} Nothing was changed.");

                string? link =
                    PublishPathSafety.FindLinkOnPath(productionRoot, source, PublishPathSafety.IsLink) ??
                    PublishPathSafety.FindLinkOnPath(betaRoot, destination, PublishPathSafety.IsLink);

                if (link is not null)
                {
                    return BetaRollbackWriteOutcome.Failed(
                        $"There is a symbolic link at [{link}], and copying through it could read or write outside " +
                        "the data trees. Nothing was changed. Remove the link on the server.");
                }

                if (!filePlan.ProductionHashes.TryGetValue(path, out string? hash) || !File.Exists(source))
                    return BetaRollbackWriteOutcome.Failed($"[{path}] is no longer in the stable source, or cannot be read there. Nothing was changed.");

                restores.Add((path, source, destination, hash));
            }

            var deletions = new List<string>();

            foreach (string path in plan.Removed)
            {
                if (!SubmissionPathRules.TryResolve(betaRoot, path, out string target, out string why))
                    return BetaRollbackWriteOutcome.Failed($"[{path}] cannot be removed from BETA: {why} Nothing was changed.");

                if (PublishPathSafety.FindLinkOnPath(betaRoot, target, PublishPathSafety.IsLink) is string link)
                {
                    return BetaRollbackWriteOutcome.Failed(
                        $"There is a symbolic link at [{link}], and deleting through it could remove a file outside " +
                        "the BETA tree. Nothing was changed. Remove the link on the server.");
                }

                deletions.Add(target);
            }

            // ###########################################################################################
            // Every folder it restores into or removes from may be written (owner report,
            // 2026-09-28) - a push-back is the one writer whose SOURCE is the production tree, so it
            // is the first to meet a folder copied from there into BETA by hand. After the link
            // check; see TreeWriteAccess. `canWriteFolder` is for tests.
            // ###########################################################################################
            IReadOnlyList<string> refusing = TreeWriteAccess.FoldersRefusing(
                betaRoot, restores.Select(restore => restore.Destination).Concat(deletions), canWriteFolder);

            if (refusing.Count > 0)
            {
                logger.LogError(
                    "Push-back of {SystemId} refused before changing anything: the service may not write into {Folders}. "
                    + "Files copied in by hand as another user do this. Give the service its access back with: {Command}",
                    system.SystemId,
                    string.Join(", ", refusing),
                    TreeWriteAccess.FixCommand(refusing));

                return BetaRollbackWriteOutcome.Failed(TreeWriteAccess.RefusalMessage("BETA", betaRoot, refusing));
            }

            // ---- 1. The restores, verified -----------------------------------------------------
            int restored = 0;

            foreach ((string path, string source, string destination, string hash) in restores)
            {
                VerifiedCopyResult result;

                try
                {
                    result = await VerifiedFileCopy.CopyAsync(
                        source, hash, destination, temporaryTag: BetaRollbackWriter.TemporaryTag, cancellationToken: cancellationToken);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    return BetaRollbackWriteOutcome.Failed(
                        $"[{path}] could not be restored into BETA ({ex.Message}). " + BetaRollbackWriter.PartialSentence(restored),
                        restored);
                }

                if (result != VerifiedCopyResult.Copied)
                {
                    return BetaRollbackWriteOutcome.Failed(
                        result == VerifiedCopyResult.HashMismatch
                            ? $"[{path}] changed in the stable source after it was checked, so it was not restored. " + BetaRollbackWriter.PartialSentence(restored)
                            : $"[{path}] is no longer in the stable source. " + BetaRollbackWriter.PartialSentence(restored),
                        restored);
                }

                // Written inside the same file-time tick as the hash the plan read - see Forget.
                BetaRollbackFiles.Forget(destination);
                restored++;
            }

            // ---- 2. The system's own BETA-only files -------------------------------------------
            int removed = 0;

            foreach (string target in deletions)
            {
                try
                {
                    if (File.Exists(target))
                    {
                        File.Delete(target);
                        BetaRollbackFiles.Forget(target);
                        removed++;
                    }
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    // A file that will not delete leaves the board correct but untidy, so this is
                    // logged rather than failed - unlike a restore, which leaves data wrong.
                    logger.LogWarning(ex, "A file could not be removed from BETA during a rollback: {Path}", target);
                }
            }

            // ###########################################################################################
            // A system that was never promoted leaves BETA entirely, so its now-empty folders go too.
            // ###########################################################################################
            if (plan.Kind == BetaRollbackKind.RemoveFromBeta &&
                SubmissionPathRules.TryResolve(betaRoot, BetaRollbackFiles.SystemFolder(system), out string folderPath, out _))
            {
                try
                {
                    if (Directory.Exists(folderPath) && !Directory.EnumerateFiles(folderPath, "*", SearchOption.AllDirectories).Any())
                        Directory.Delete(folderPath, recursive: true);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    logger.LogWarning(ex, "The emptied system folder could not be removed from BETA: {Path}", folderPath);
                }
            }

            return BetaRollbackWriteOutcome.Done(restored, removed);
        }

        private static string PartialSentence(int restored) =>
            restored == 0
                ? "Nothing was changed."
                : $"{restored} file(s) before it were already restored; pushing back again completes it.";
    }

    public sealed record BetaRollbackWriteOutcome(
        bool IsDone,
        int Restored,
        int Removed,
        string? Error)
    {
        public static BetaRollbackWriteOutcome Done(int restored, int removed) =>
            new(true, restored, removed, null);

        public static BetaRollbackWriteOutcome Failed(string error, int restored = 0) => new(false, restored, 0, error);
    }
}
