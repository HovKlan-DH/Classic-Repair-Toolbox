using System.Security.Cryptography;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Performs a publish: copies the submitted blobs into the data tree, generates the board
    // workbook, writes system.json, and records the new revision (NewContributeStrategy.md
    // Phase 5, task 6).
    //
    // *** THIS WRITES THE BETA TREE AND NEVER PRODUCTION. *** Open question 6 settled it: BETA to
    // Production is a manual file copy the maintainer performs, there is no automation to fit
    // into, and Phase 3 step 0 denies this service write permission on the Production tree at the
    // FILESYSTEM level. The kernel is the real guard; this comment is so nobody spends an
    // afternoon wondering why a path is refused. Do not add a "publish to production" option.
    //
    // *** WHAT IT DOES NOT DECIDE. *** PublishPlan already decided every path, refused every
    // unsafe one, and resolved which workbook GENERATION is being written. This class performs
    // that plan and nothing more - it does not re-derive a path, and it must not, because
    // re-joining the original strings is exactly the traversal hole SubmissionPathRules closes.
    //
    // *** PUBLISHING IS NOT ATOMIC AND CANNOT BE MADE SO HERE. *** A system is hundreds of files
    // across a tree that a sync reads concurrently; there is no rename that swaps them all at
    // once. What this does instead is order the writes so a partial publish is RECOVERABLE:
    //
    //   1. files first, which are inert until something references them;
    //   2. the workbook next, which is what makes those files reachable;
    //   3. system.json last, which is what advertises the system as being at this revision;
    //   4. the database row last of all.
    //
    // An interrupted publish therefore leaves a tree carrying files nothing points at - wasted
    // space, not corruption - rather than a workbook referencing files that never arrived, which
    // is a board that fails to load. RE-RUNNING THE SAME PUBLISH IS SAFE and is the recovery: each
    // step overwrites, and the blobs are content-addressed so a re-copy is byte-identical.
    //
    // The maintainer decided against retained revisions (open question 5), so there is no previous
    // version to roll back to. That is what makes the ordering above matter rather than being
    // merely tidy.
    // ###########################################################################################
    public sealed class PublishExecutor
    {
        private readonly BlobStore thisBlobs;
        private readonly ISubmissionStore thisStore;
        private readonly ILogger<PublishExecutor> thisLogger;

        public PublishExecutor(BlobStore blobs, ISubmissionStore store, ILogger<PublishExecutor> logger)
        {
            this.thisBlobs = blobs;
            this.thisStore = store;
            this.thisLogger = logger;
        }

        // ###########################################################################################
        // Performs the plan. Returns what was written, or the reason it stopped.
        //
        // boardData is the merged board the submission describes - official rows with the
        // contributed ones applied - which the caller builds, because assembling it needs the
        // submission's rows and this class's job is the writing.
        // ###########################################################################################
        public async Task<PublishOutcome> ExecuteAsync(
            PublishPlanDetail plan,
            BoardData boardData,
            IReadOnlyList<KiCadCalibrationEntry> calibrations,
            long? submissionId,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(plan);
            ArgumentNullException.ThrowIfNull(boardData);

            // *** REFUSED RATHER THAN DEFAULTED TO EMPTY. *** A caller passing null has almost
            // certainly failed to read the submission's calibrations rather than genuinely meaning
            // "this board has none", and treating the two alike publishes a board with a
            // contributor's calibration work silently removed.
            ArgumentNullException.ThrowIfNull(calibrations);

            // ---- 1. The files ------------------------------------------------------------------
            //
            // Inert until the workbook references them, so a failure here leaves nothing broken.
            foreach (PlannedFile file in plan.Files)
            {
                cancellationToken.ThrowIfCancellationRequested();

                bool copied;

                try
                {
                    copied = await this.thisBlobs
                        .TryCopyToAsync(file.Sha256, file.AbsolutePath, cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    // ###########################################################################################
                    // *** A FILESYSTEM REFUSAL IS A STOPPED PUBLISH, NOT A CRASH (maintainer report,
                    // 2026-09-23). ***
                    //
                    // TryCopyToAsync returns false for a MISSING blob, which the branch below
                    // handles - but it THROWS when the destination cannot be written. Nothing
                    // caught that, so it escaped through ApprovePublishFlow and the endpoint, and
                    // "Approve and publish" answered a bare 500 with nothing to act on.
                    //
                    // The real case: the service account could not overwrite an existing board
                    // image in the published tree ("Permission denied" on
                    // Board Layout 250407 NTSC.png). That is a deployment fault with a precise
                    // remedy, and a 500 is the one response that hides which file and why.
                    //
                    // Stopping here is safe for the same reason the missing-blob branch is: the
                    // workbook has not been touched, so nothing downstream reads the partially
                    // copied files. The publish is genuinely incomplete either way - what changes
                    // is that the reviewer is told WHICH file and WHY.
                    //
                    // Narrow on purpose: an IOException or a permission refusal is an environment
                    // problem to report, whereas anything else here is a defect that should not be
                    // quietly turned into a friendly message.
                    // ###########################################################################################
                    this.thisLogger.LogError(
                        ex,
                        "Publish of {SystemId} stopped: [{Path}] could not be written.",
                        plan.SystemId,
                        file.RelativePath);

                    return PublishOutcome.Failed(
                        $"[{file.RelativePath}] could not be written to the published tree " +
                        $"({ex.Message}). Nothing was changed. This is a server configuration " +
                        "problem rather than a fault in the submission.");
                }

                if (!copied)
                {
                    // The blob is missing from the store although the submission was finalised,
                    // which means the store lost it or the plan named a hash nothing uploaded.
                    // Stopping here leaves the tree as it was for everything downstream, because
                    // the workbook has not been touched.
                    this.thisLogger.LogError(
                        "Publish of {SystemId} stopped: blob {Hash} for {Path} is not in the store.",
                        plan.SystemId,
                        file.Sha256,
                        file.RelativePath);

                    return PublishOutcome.Failed(
                        $"The content for [{file.RelativePath}] is no longer in the blob store, so the publish was stopped before any board file was changed.");
                }
            }

            // ---- 2. The workbook ---------------------------------------------------------------
            //
            // Generated from the submitted rows, never uploaded - PublishPlan refuses a submitted
            // workbook outright, so that a contributor cannot overwrite a frozen generation.
            cancellationToken.ThrowIfCancellationRequested();

            string workbookHash;

            try
            {
                // ###########################################################################################
                // *** THE REVISION DATE IS STAMPED WITH THE PUBLISH DATE (maintainer request,
                // 2026-09-23). ***
                //
                // It is the "# Revision date:" line at the top of the Board schematics sheet, and it
                // means "when this board's data was last published" - so the publish is the only
                // event that can set it truthfully. Carrying the SUBMITTED value through would date
                // a freshly published board to whenever the contributor happened to start their
                // draft, which on a submission that sat in the queue for a week is simply wrong.
                //
                // Deliberately NOT done when CRT writes a draft: a draft is not published data, and
                // stamping it locally would move the value DraftRevisionComparer uses to detect that
                // the published board has drifted - every draft would immediately look as though it
                // had a newer revision than the board it came from.
                //
                // Formatted as "2026-May-12" (BoardWorkbookStyle.FormatRevisionDate), matching what
                // the shipped boards already carry.
                // ###########################################################################################
                BoardData publishedBoard = boardData.WithRevisionDate(
                    BoardWorkbookStyle.FormatRevisionDate(nowUtc));

                BoardWorkbookWriter.Write(plan.WorkbookPath, publishedBoard);

                workbookHash = await PublishExecutor
                    .ComputeFileHashAsync(plan.WorkbookPath, cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // ###########################################################################################
                // *** THE SAME FILESYSTEM REFUSAL, AT THE STAGE WHERE IT ACTUALLY HURTS. ***
                //
                // Step 1 stopping is clean - the workbook is untouched, so nothing reads the files
                // it managed to copy. Stopping HERE is not clean: the images have been replaced and
                // the workbook has not, so the board on disk is a mixture until the publish is
                // re-run. Reported as such rather than thrown, because a reviewer needs to know the
                // tree is mid-publish, which a 500 does not tell them.
                //
                // Re-running the publish repairs it: every step overwrites unconditionally, so the
                // second attempt writes the same files again from the same blobs.
                // ###########################################################################################
                this.thisLogger.LogError(
                    ex,
                    "Publish of {SystemId} stopped while writing the workbook [{Path}]. The board's "
                    + "FILES were already replaced, so the tree is part-published until this is re-run.",
                    plan.SystemId,
                    plan.WorkbookPath);

                return PublishOutcome.Failed(
                    $"The board workbook could not be written ({ex.Message}). The board's image and "
                    + "attachment files were already replaced, so this system is part-published - "
                    + "re-run the publish once the server configuration is fixed.");
            }

            // ---- 2b. The SIDECAR ---------------------------------------------------------------
            //
            // *** A BOARD IS TWO FILES, AND OMITTING THIS ONE IS SILENT DATA LOSS. *** The `.json`
            // beside the workbook holds every COMPONENT HIGHLIGHT - the rectangles the whole app is
            // built around - and every KiCad calibration. BoardWorkbookWriter does not write them,
            // so a publish that stopped at the workbook would produce a board whose highlights had
            // simply vanished, with no retained revision to restore from.
            //
            // Written immediately after the workbook and before anything advertises the revision,
            // so the two halves of a board become reachable together. An interruption between them
            // leaves a workbook with a stale sidecar, which re-running the publish repairs.
            //
            // Calibrations are NOT part of BoardData - they live only in the sidecar - so they are
            // passed separately rather than read off the merged board.
            cancellationToken.ThrowIfCancellationRequested();

            // Takes the WORKBOOK path and derives the sidecar itself, through the same helper the
            // reader uses - so this call site cannot derive it differently from where it is read.
            BoardSidecarWriter.Write(plan.WorkbookPath, boardData.ComponentHighlights, calibrations);

            string sidecarHash = await PublishExecutor
                .ComputeFileHashAsync(plan.SidecarPath, cancellationToken)
                .ConfigureAwait(false);

            // ---- 3. system.json ----------------------------------------------------------------
            //
            // DescriptorWithWorkbook, not the planned descriptor: the planned one covers the
            // uploaded files only, and a rows-only change uploads nothing - so without the
            // workbook and sidecar folded in, the content hash would not move and no client would
            // re-download the board that just changed. A submission that only MOVES A HIGHLIGHT
            // touches neither a row nor an uploaded file, so the sidecar's hash is the only thing
            // that can move it at all.
            SystemDescriptor descriptor = plan.DescriptorWithWorkbook(workbookHash, sidecarHash);

            SystemDescriptorStore.Write(plan.SystemFolder, descriptor);

            // ---- 4. The database ---------------------------------------------------------------
            //
            // Last, so a row claiming a revision can never exist before the tree actually holds
            // it. The reverse order would leave a contributor's next submission diffing against a
            // revision that was never written.
            await this.thisStore
                .SetSystemPublishedAsync(
                    plan.SystemId,
                    descriptor.Revision,
                    descriptor.ContentHash,
                    nowUtc,
                    cancellationToken)
                .ConfigureAwait(false);

            if (submissionId is long id)
            {
                await this.thisStore
                    .SetStateAsync(id, SubmissionState.Merged, nowUtc, cancellationToken)
                    .ConfigureAwait(false);
            }

            this.thisLogger.LogInformation(
                "Published {SystemId} revision {Revision} ({FileCount} files, workbook {Workbook}).",
                plan.SystemId,
                descriptor.Revision,
                plan.Files.Count,
                plan.WorkbookFileName);

            return PublishOutcome.Published(descriptor, workbookHash, plan.Files.Count);
        }

        // ###########################################################################################
        // The SHA-256 of a file on disk, streamed.
        //
        // Streamed rather than read into memory because a board workbook for a large system is not
        // small, and the strategy document's own performance note is explicit: never load a whole
        // system into memory to compute a hash.
        //
        // Lower-case hex, matching SystemDescriptorRules.ComputeContentHash and
        // dataChecksums.json - a hash that differs only in case compares unequal everywhere it is
        // used, which reads as every client being permanently out of date.
        // ###########################################################################################
        private static async Task<string> ComputeFileHashAsync(string path, CancellationToken cancellationToken)
        {
            await using FileStream stream = new(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                81920,
                useAsync: true);

            byte[] hash = await SHA256.HashDataAsync(stream, cancellationToken).ConfigureAwait(false);

            return Convert.ToHexStringLower(hash);
        }
    }

    // ###########################################################################################
    // What a publish did, or why it stopped.
    //
    // The descriptor is returned rather than merely written, because the caller records the
    // revision in its own audit trail and would otherwise have to read the file back to learn
    // what it just published.
    // ###########################################################################################
    public sealed class PublishOutcome
    {
        private PublishOutcome(bool isPublished, SystemDescriptor? descriptor, string workbookSha256, int fileCount, string? failure)
        {
            this.IsPublished = isPublished;
            this.Descriptor = descriptor;
            this.WorkbookSha256 = workbookSha256;
            this.FileCount = fileCount;
            this.Failure = failure;
        }

        public bool IsPublished { get; }

        public SystemDescriptor? Descriptor { get; }

        public string WorkbookSha256 { get; }

        public int FileCount { get; }

        // Written for a human reading a log or a review screen, not for a machine to branch on.
        public string? Failure { get; }

        public static PublishOutcome Published(SystemDescriptor descriptor, string workbookSha256, int fileCount) =>
            new(true, descriptor, workbookSha256, fileCount, null);

        public static PublishOutcome Failed(string failure) =>
            new(false, null, string.Empty, 0, failure);
    }
}
