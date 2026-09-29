using System.Security.Cryptography;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Performs a publish: copies the submitted blobs into the data tree, generates the board
    // workbook and its sidecar, and records the new revision (NewContributeStrategy.md Phase 5,
    // task 6). It no longer writes system.json - retired 2026-09-25, see RetiredSystemDescriptor -
    // and removes one an earlier build left in the board's folder.
    //
    // *** THIS WRITES THE BETA TREE AND NEVER PRODUCTION. *** Since 2026-09-25 Production IS
    // written - by ProductionPromoter, and only by copying files that are already in BETA, once a
    // maintainer has checked them there (the project owner's two-stage publish). A submission never
    // goes straight to Production, and nothing in this class may learn to: the BETA stage is the
    // one a person looks at before everybody gets it.
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
    //   0. EVERYTHING CHECKED FIRST (security review, 2026-09-25) - every blob present, intact and
    //      of the type its name claims, and no symbolic link on the way to any destination. A
    //      replaced image is NOT inert (the old workbook already cites it), so a refusal found
    //      half way through used to leave half a board replaced;
    //   1. files next, each hashed as it is copied and renamed into place only on a match;
    //   2. the workbook next, which is what makes those files reachable;
    //   3. the database row last of all, which is what records the system as being at this
    //      revision.
    //
    // An interrupted publish therefore leaves a tree carrying files nothing points at - wasted
    // space, not corruption - rather than a workbook referencing files that never arrived, which
    // is a board that fails to load. RE-RUNNING THE SAME PUBLISH IS SAFE and is the recovery: each
    // step overwrites, and the blobs are content-addressed so a re-copy is byte-identical.
    //
    // The project owner decided against retained revisions (open question 5), so there is no previous
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

        // Whether a folder may be written - TreeWriteAccess's real probe unless a test says
        // otherwise, since a folder the service may not write cannot be made on every OS a test
        // runs on.
        internal Func<string, bool>? CanWriteFolderForTests { get; init; }

        // ###########################################################################################
        // Performs the plan. Returns what was written, or the reason it stopped.
        //
        // boardData is the merged board the submission describes - official rows with the
        // contributed ones applied - which the caller builds, because assembling it needs the
        // submission's rows and this class's job is the writing.
        // ###########################################################################################
        public Task<PublishOutcome> ExecuteAsync(
            PublishPlanDetail plan,
            BoardData boardData,
            IReadOnlyList<KiCadCalibrationEntry> calibrations,
            long? submissionId,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default) =>
            this.ExecuteAsync(plan, boardData, calibrations, submissionId, nowUtc, listing: null, cancellationToken);

        // `listing`: a NEW system's row for the main Excel data file (2026-09-27), written after the
        // board and before the database - see step 2c. Null for everything already listed.
        public async Task<PublishOutcome> ExecuteAsync(
            PublishPlanDetail plan,
            BoardData boardData,
            IReadOnlyList<KiCadCalibrationEntry> calibrations,
            long? submissionId,
            DateTimeOffset nowUtc,
            MasterRowInsert? listing,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(plan);
            ArgumentNullException.ThrowIfNull(boardData);

            // *** REFUSED RATHER THAN DEFAULTED TO EMPTY. *** A caller passing null has almost
            // certainly failed to read the submission's calibrations rather than genuinely meaning
            // "this board has none", and treating the two alike publishes a board with a
            // contributor's calibration work silently removed.
            ArgumentNullException.ThrowIfNull(calibrations);

            // ---- 0. Everything checked BEFORE anything is written (security review, 2026-09-25) --
            //
            // Replacing a board image is NOT inert - the old workbook already points at it - so a
            // refusal discovered at file 600 of 1,200 would leave half a board replaced. Every blob
            // is therefore proved present, intact and of the type its name claims, and every
            // destination free of symbolic links, before the first byte lands.
            string? refusal = await this.CheckBeforeWritingAsync(plan, listing, cancellationToken).ConfigureAwait(false);

            if (refusal is not null)
                return PublishOutcome.Failed(refusal);

            // ---- 1. The files ------------------------------------------------------------------
            //
            // Each copy is hashed as it is written and moved into place only on a match - see
            // BlobStore.TryCopyToAsync - so a blob that changed since the check above still cannot
            // reach the tree.
            foreach (PlannedFile file in plan.Files)
            {
                cancellationToken.ThrowIfCancellationRequested();

                BlobCopyResult copied;

                try
                {
                    copied = await this.thisBlobs
                        .TryCopyToAsync(file.Sha256, file.AbsolutePath, cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    // ###########################################################################################
                    // *** A FILESYSTEM REFUSAL IS A STOPPED PUBLISH, NOT A CRASH (owner report,
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
                    // is that the maintainer is told WHICH file and WHY.
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

                if (copied == BlobCopyResult.HashMismatch)
                {
                    return PublishOutcome.Failed(
                        $"The stored content for [{file.RelativePath}] no longer matches what was submitted, so " +
                        "it was not published. Files before it in this publish were already replaced; " +
                        "re-running the publish is safe once the submission is sent again.");
                }

                if (copied == BlobCopyResult.Missing)
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
                // *** THE REVISION DATE IS STAMPED WITH THE PUBLISH DATE (owner request,
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
                //
                // *** TAKEN FROM THE PLAN, NOT COMPUTED AGAIN (2026-09-26). *** It used to stamp
                // FormatRevisionDate(nowUtc) here while the plan's descriptor - which is what
                // reaches `systems.current_revision` at step 4 below - was built from the SUBMITTED
                // date. So the workbook and the database row disagreed, and the row is what the
                // next draft re-bases against. ApprovePublishFlow.BuildPlan now stamps it once and
                // both read that one value.
                // ###########################################################################################
                BoardData publishedBoard = boardData.WithRevisionDate(plan.Descriptor.Revision);

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
                // re-run. Reported as such rather than thrown, because a maintainer needs to know the
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

            string sidecarHash;

            try
            {
                // Takes the WORKBOOK path and derives the sidecar itself, through the same helper
                // the reader uses - so this call site cannot derive it differently from where it is
                // read.
                //
                // Only when its content changes (owner request, 2026-09-28): a sidecar already
                // holding these highlights keeps its bytes, so "Publish to production" does not list
                // it as a file to copy when nothing in it changed - see WriteIfChanged.
                if (!BoardSidecarWriter.WriteIfChanged(plan.WorkbookPath, boardData.ComponentHighlights, calibrations))
                    this.thisLogger.LogInformation("The highlight file of {SystemId} is unchanged and was kept as it is.", plan.SystemId);

                sidecarHash = await PublishExecutor
                    .ComputeFileHashAsync(plan.SidecarPath, cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // ###########################################################################################
                // *** REPORTED, NOT A 500 (owner report, 2026-09-28). *** This escaped once: a
                // sidecar copied into BETA by hand could not be opened for writing, and "Approve and
                // publish" answered a bare 500 with the new workbook already written. The sidecar
                // is now REPLACED rather than opened (FileReplacer) and the folder is checked before
                // anything is written, so this should be unreachable - but if it is reached, the
                // maintainer is told the board is part-published and that approving again repairs
                // it, exactly as for the workbook above.
                // ###########################################################################################
                this.thisLogger.LogError(
                    ex,
                    "Publish of {SystemId} stopped while writing the highlight file [{Path}]. The board's "
                    + "FILES and WORKBOOK were already replaced, so the tree is part-published until this is re-run.",
                    plan.SystemId,
                    plan.SidecarPath);

                return PublishOutcome.Failed(
                    $"The board's highlight file could not be written ({ex.Message}). The board's files and "
                    + "workbook were already replaced, so this system is part-published - approve again once the "
                    + "server configuration is fixed.");
            }

            // ---- 2c. A NEW system's row in the main Excel data file (owner request, 2026-09-27) --
            //
            // What makes the board appear in CRT's drop-downs at all - CRT finds boards only through
            // that file. After the board's two files, so the row never names a board that is not
            // there; before the database, so the publish is not recorded until it is listed. A
            // failure here leaves the board written but unlisted - invisible, not broken - and
            // re-running the publish repairs it: the insert updates a row already there.
            if (listing is not null)
            {
                cancellationToken.ThrowIfCancellationRequested();

                MasterListingEdit edit = MasterListing.Insert(listing.MasterPath, listing.Row, listing.AfterExcelDataFile);

                if (!edit.IsDone)
                {
                    this.thisLogger.LogError(
                        "Publish of {SystemId} stopped: the board is written but could not be added to [{Master}]: {Failure}",
                        plan.SystemId, listing.MasterPath, edit.Failure);

                    return PublishOutcome.Failed(
                        $"The board was written, but it could not be added to the drop-down lists: {edit.Failure} " +
                        "Nobody sees it until it is - re-run the publish once that is fixed.");
                }
            }

            // ---- 3. What this publish IS -----------------------------------------------------
            //
            // Held in memory and recorded in the database; never written to the tree since
            // system.json was retired (2026-09-25).
            //
            // DescriptorWithWorkbook, not the planned descriptor: the planned one covers the
            // uploaded files only, and a rows-only change uploads nothing - so without the
            // workbook and sidecar folded in, the content hash would not move, and the production
            // promotion - which compares exactly this hash to decide what is waiting - would never
            // offer the board. A submission that only MOVES A HIGHLIGHT touches neither a row nor
            // an uploaded file, so the sidecar's hash is the only thing that can move it at all.
            SystemDescriptor descriptor = plan.DescriptorWithWorkbook(workbookHash, sidecarHash);

            // A system.json an earlier build wrote into this folder: not relevant to users, and
            // no longer kept. See RetiredSystemDescriptor.
            RetiredSystemDescriptor.TryRemove(plan.DataRoot, plan.SystemFolder, this.thisLogger);

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
        // The pre-write check. Null when everything may be written, otherwise the reason it may not.
        //
        // Also covers the two GENERATED files - workbook and sidecar - for links:
        // they are written through the same folders, and a link there redirects them just as well.
        // ###########################################################################################
        private async Task<string?> CheckBeforeWritingAsync(PublishPlanDetail plan, MasterRowInsert? listing, CancellationToken cancellationToken)
        {
            IEnumerable<string> destinations = plan.Files
                .Select(file => file.AbsolutePath)
                .Append(plan.WorkbookPath)
                .Append(plan.SidecarPath);

            // The main Excel data file too, for a new system: written through the same root.
            if (listing is not null)
            {
                destinations = destinations.Append(listing.MasterPath);

                // Checked before the first byte, like everything else here: a row placed after one
                // that has since left the list must not be found out after the board is written.
                if (!MasterListing.TryRead(listing.MasterPath, out IReadOnlyList<MasterListingRow> rows, out string why))
                    return $"The main Excel data file could not be read ({why}), so the new system could not be listed. Nothing was changed.";

                if (MasterListing.IndexOfSystem(rows, listing.Row.SystemId) < 0 &&
                    !SystemListingRules.AfterIsListed(rows, listing.AfterExcelDataFile))
                {
                    return $"The new system was placed after [{listing.AfterExcelDataFile}], which is no longer in the " +
                        "drop-down lists. Place it again in the Systems screen. Nothing was changed.";
                }

                if (MasterListing.NamesTakenBy(rows, listing.Row.SystemId, listing.Row.HardwareName, listing.Row.BoardName) is MasterListingRow taken)
                    return $"{MasterListing.NamesTakenMessage(taken)} Nothing was changed.";
            }

            foreach (string destination in destinations)
            {
                string? link = PublishPathSafety.FindLinkOnPath(plan.DataRoot, destination, PublishPathSafety.IsLink);

                if (link is not null)
                {
                    this.thisLogger.LogError(
                        "Publish of {SystemId} refused: [{Link}] is a symbolic link on the way to [{Destination}].",
                        plan.SystemId, link, destination);

                    return $"The published tree contains a symbolic link at [{link}], and publishing through it " +
                        "could write outside the data tree. Nothing was changed. Remove the link on the server.";
                }
            }

            // ###########################################################################################
            // *** EVERY FOLDER THIS WRITES INTO MAY BE WRITTEN (owner report, 2026-09-28). *** A folder
            // copied into BETA by hand as another user refuses every write, and finding that out at
            // the third file leaves the board half-replaced. After the link check, which a probe -
            // a real file - must never get ahead of. See TreeWriteAccess.
            // ###########################################################################################
            IReadOnlyList<string> refusing = TreeWriteAccess.FoldersRefusing(plan.DataRoot, destinations, this.CanWriteFolderForTests);

            if (refusing.Count > 0)
            {
                this.thisLogger.LogError(
                    "Publish of {SystemId} refused before writing anything: the service may not write into {Folders}. "
                    + "Files copied in by hand as another user do this. Give the service its access back with: {Command}",
                    plan.SystemId,
                    string.Join(", ", refusing),
                    TreeWriteAccess.FixCommand(refusing));

                return TreeWriteAccess.RefusalMessage("BETA", plan.DataRoot, refusing);
            }

            foreach (PlannedFile file in plan.Files)
            {
                cancellationToken.ThrowIfCancellationRequested();

                BlobVerification blob = await this.thisBlobs
                    .VerifyAsync(file.Sha256, SubmissionContentRules.SampleBytes, cancellationToken)
                    .ConfigureAwait(false);

                if (!blob.Exists)
                {
                    this.thisLogger.LogError(
                        "Publish of {SystemId} stopped: blob {Hash} for {Path} is not in the store.",
                        plan.SystemId, file.Sha256, file.RelativePath);

                    return $"The content for [{file.RelativePath}] is no longer in the blob store, so the publish " +
                        "was stopped before any board file was changed.";
                }

                if (!blob.HashMatches)
                {
                    this.thisLogger.LogError(
                        "Publish of {SystemId} stopped: blob {Hash} for {Path} no longer hashes to its name.",
                        plan.SystemId, file.Sha256, file.RelativePath);

                    return $"The stored content for [{file.RelativePath}] no longer matches what was submitted, so " +
                        "the publish was stopped before any board file was changed.";
                }

                if (!SubmissionContentRules.Matches(file.RelativePath, blob.Head, out string reason))
                {
                    return $"[{file.RelativePath}] cannot be published: {reason} Nothing was changed.";
                }
            }

            return null;
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
    // A new system's row for the main Excel data file: which file, the row, and the row it goes
    // after (null: first). See MasterListing.
    // ###########################################################################################
    public sealed record MasterRowInsert(string MasterPath, MasterListingRow Row, string? AfterExcelDataFile);

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
