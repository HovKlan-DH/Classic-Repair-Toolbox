using System.Security.Cryptography;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Performs a publish: stores the submitted blobs in the data tree, writes the board file, and
    // records the new revision (NewContributeStrategy.md Phase 5, task 6). It no longer writes
    // system.json - retired 2026-09-25, see RetiredBoardDescriptor - and removes one an earlier
    // build left in the board's folder.
    //
    // *** THE CRT 4.0 TREE (2026-10-25). *** A board is its board file, "<board id>/board.json",
    // holding everything the Excel generations kept in a workbook and its sidecar - the rows, the
    // highlights, the KiCad calibrations, the schematic renames - and the board's list of files by
    // name. Each file's bytes are stored once under their checksum ("Files/<ab>/<sha256>.<ext>"). So
    // a publish stores the bytes the tree does not hold yet, then writes the one board file.
    //
    // *** THIS WRITES THE BETA TREE AND NEVER PRODUCTION. *** Since 2026-09-25 Production IS
    // written - by ProductionPromoter, and only by copying files that are already in BETA, once a
    // maintainer has checked them there (the project owner's two-stage publish). A submission never
    // goes straight to Production, and nothing in this class may learn to: the BETA stage is the
    // one a person looks at before everybody gets it.
    //
    // *** WHAT IT DOES NOT DECIDE. *** PublishPlan already decided every place and refused every
    // unsafe one. This class performs that plan and nothing more - it does not re-derive a path,
    // and it must not, because re-joining the original strings is exactly the traversal hole
    // SubmissionPathRules closes.
    //
    // *** PUBLISHING IS NOT ATOMIC AND CANNOT BE MADE SO HERE. *** What this does instead is order
    // the writes so a partial publish is RECOVERABLE:
    //
    //   0. EVERYTHING CHECKED FIRST (security review, 2026-09-25) - every blob present, intact and
    //      of the type its name claims, and no symbolic link on the way to any destination;
    //   1. the stored files next, each hashed as it is copied and renamed into place only on a
    //      match. A stored file is NEW bytes beside the old - nothing the published board points at
    //      changes - so stopping here leaves the board exactly as it was;
    //   2. the board file next, replaced whole, which is what makes those files part of the board;
    //   3. a new board's row in the main file, so the drop-down lists name it;
    //   4. the database row last of all, which is what records the board as being at this revision.
    //
    // An interrupted publish therefore leaves a tree carrying stored files nothing points at -
    // wasted space, not corruption - never a board file naming bytes that never arrived. RE-RUNNING
    // THE SAME PUBLISH IS SAFE and is the recovery: the stored files are content-addressed, so a
    // re-copy is byte-identical, and the board file is replaced again.
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

        // `listing`: a NEW board's row for the main file (2026-09-27), written after the board and
        // before the database - see step 3. Null for everything already listed.
        public Task<PublishOutcome> ExecuteAsync(
            PublishPlanDetail plan,
            BoardData boardData,
            IReadOnlyList<KiCadCalibrationEntry> calibrations,
            long? submissionId,
            DateTimeOffset nowUtc,
            MasterRowInsert? listing,
            CancellationToken cancellationToken = default) =>
            this.ExecuteAsync(plan, boardData, calibrations, renames: null, submissionId, nowUtc, listing, cancellationToken: cancellationToken);

        // ###########################################################################################
        // `renames`: the board's whole "Schematic renames" list as its board file is to hold it - the
        // approval's (ApprovePublishFlow.RenameLogFor: what it had, plus this publish's renames;
        // owner decision 2026-10-10). Null keeps the list the board file has now, unchanged.
        //
        // `publishedBoardFile`: the board file as the approval read it once (PublishedBoardRead - code
        // review, 2026-10-26), whose unknown fields are kept rather than the file being parsed again.
        // Null, or one that could not be read, or of another path: read here, in full, as before.
        // ###########################################################################################
        public async Task<PublishOutcome> ExecuteAsync(
            PublishPlanDetail plan,
            BoardData boardData,
            IReadOnlyList<KiCadCalibrationEntry> calibrations,
            IReadOnlyList<SchematicRenameRecord>? renames,
            long? submissionId,
            DateTimeOffset nowUtc,
            MasterRowInsert? listing,
            PublishedBoardRead? publishedBoardFile = null,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(plan);
            ArgumentNullException.ThrowIfNull(boardData);

            // *** REFUSED RATHER THAN DEFAULTED TO EMPTY. *** A caller passing null has almost
            // certainly failed to read the submission's calibrations rather than genuinely meaning
            // "this board has none", and treating the two alike publishes a board with a
            // contributor's calibration work silently removed.
            ArgumentNullException.ThrowIfNull(calibrations);

            // Read before anything is written - the board file is replaced in step 2.
            IReadOnlyList<SchematicRenameRecord> renameLog =
                renames ?? BoardComponentHighlightStorage.LoadSchematicRenames(plan.BoardFilePath);

            // ---- 0. Everything checked BEFORE anything is written (security review, 2026-09-25) --
            string? refusal = await this.CheckBeforeWritingAsync(plan, listing, cancellationToken).ConfigureAwait(false);

            if (refusal is not null)
                return PublishOutcome.Failed(refusal);

            // ###########################################################################################
            // *** CANCELLABLE ONLY UNTIL HERE (code review, 2026-10-13). *** From the first write on,
            // it runs to the end whatever the client does - a maintainer's client giving up (a closed
            // laptop, a proxy timeout) must not stop it half way. The rule BoardOrderFlow and
            // BoardDeletionFlow already follow.
            // ###########################################################################################
            cancellationToken = CancellationToken.None;

            IReadOnlyList<PlannedFile> stored = plan.StoredFiles;

            // ---- 1. The stored files -------------------------------------------------------------
            //
            // Each copy is hashed as it is written and moved into place only on a match - see
            // BlobStore.TryCopyToAsync - so a blob that changed since the check above still cannot
            // reach the tree.
            foreach (PlannedFile file in stored)
            {
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
                    // 2026-09-23). *** TryCopyToAsync THROWS when the destination cannot be written.
                    // A deployment fault with a precise remedy, and a 500 is the one response that
                    // hides which file and why. Stopping here is safe: only new stored files were
                    // written, and the board file - what points at them - is untouched.
                    // ###########################################################################################
                    this.thisLogger.LogError(
                        ex,
                        "Publish of {BoardId} stopped: [{Name}] could not be stored at [{Path}].",
                        plan.BoardId,
                        file.Name,
                        file.StoredPath);

                    // The exception's own text names the server's paths - it is in the log line
                    // above, never in what the maintainer is shown (code review, 2026-10-13).
                    return PublishOutcome.Failed(
                        $"[{file.Name}] could not be written to the BETA data. Nothing of the board was changed. " +
                        "This is a server configuration problem rather than a fault in the submission - the server's log says more.");
                }

                if (copied == BlobCopyResult.HashMismatch)
                {
                    return PublishOutcome.Failed(
                        $"The stored content for [{file.Name}] no longer matches what was submitted, so it was not " +
                        "published. Nothing of the board was changed; publishing again is safe once the submission is sent again.");
                }

                if (copied == BlobCopyResult.Missing)
                {
                    this.thisLogger.LogError(
                        "Publish of {BoardId} stopped: blob {Hash} for {Name} is not in the store.",
                        plan.BoardId,
                        file.Sha256,
                        file.Name);

                    return PublishOutcome.Failed(
                        $"The content for [{file.Name}] is no longer in the blob store, so the publish was stopped before any board file was changed.");
                }
            }

            // ---- 2. The board file ---------------------------------------------------------------
            //
            // Generated from the submitted rows, never uploaded - SubmissionFileRules refuses a .json
            // outright, so a contributor cannot write a board file of their own.
            string boardFileHash;

            try
            {
                // ###########################################################################################
                // *** THE REVISION DATE IS STAMPED WITH THE PUBLISH DATE (owner request,
                // 2026-09-23). *** It means "when this board's data was last published" - so the
                // publish is the only event that can set it truthfully. TAKEN FROM THE PLAN, NOT
                // COMPUTED AGAIN (2026-09-26): the plan's descriptor is what reaches
                // `boards.current_revision` at step 4, and the next draft re-bases against that row,
                // so the board file and the row carry ONE value.
                //
                // Fields this server does not know - at the top of the board file and on every row -
                // are kept: a later CRT's additions survive an edit made through this one
                // (BoardDocumentUnknowns). So the board file there is read IN FULL or not at all
                // (code review, 2026-10-25): one that cannot be read - a newer format, a hand edit
                // gone wrong - stops the publish rather than being replaced by what this server knows.
                // ###########################################################################################
                BoardData publishedBoard = boardData.WithRevisionDate(plan.Descriptor.Revision);

                BoardDocument? existing = publishedBoardFile is { Document: { } readOnce } && publishedBoardFile.IsAt(plan.BoardFilePath)
                    ? readOnce
                    : BoardFiles.ReadInFull(plan.BoardFilePath);

                BoardDocument document = BoardDocumentMapping.ToDocument(
                    new BoardContent(publishedBoard, calibrations, renameLog),
                    name => name,
                    plan.FileList);

                BoardDocumentUnknowns.CarryOver(existing, document);

                BoardFiles.Write(plan.BoardFilePath, document);

                boardFileHash = await PublishExecutor
                    .ComputeFileHashAsync(plan.BoardFilePath, cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // ###########################################################################################
                // *** REPORTED, NOT A 500. *** The board file is replaced whole (FileReplacer), so a
                // failure leaves the OLD board file as it was, and the stored files copied in step 1
                // are new bytes nothing points at yet. Approving again repairs it.
                // ###########################################################################################
                this.thisLogger.LogError(
                    ex,
                    "Publish of {BoardId} stopped while writing the board file [{Path}].",
                    plan.BoardId,
                    plan.BoardFilePath);

                return PublishOutcome.Failed(
                    "The board file could not be written - the server's log says why. The board is as it was; " +
                    "approve again once the server configuration is fixed.");
            }
            catch (InvalidOperationException ex)
            {
                this.thisLogger.LogError(
                    ex,
                    "Publish of {BoardId} stopped: the board file [{Path}] there could not be read in full.",
                    plan.BoardId,
                    plan.BoardFilePath);

                return PublishOutcome.Failed(
                    "The board file in the BETA data could not be read in full - it may have been written by a newer " +
                    "CRT, or edited by hand - so it was not written over. The board is as it was; the server's log says more.");
            }

            // ---- 3. A NEW board's row in the main file (owner request, 2026-09-27) ----------------
            //
            // What makes the board appear in CRT's drop-downs at all - CRT finds boards only through
            // that file. After the board file, so the row never names a board that is not there;
            // before the database, so the publish is not recorded until it is listed. A failure here
            // leaves the board written but unlisted - invisible, not broken - and re-running the
            // publish repairs it: the insert updates a row already there.
            if (listing is not null)
            {
                MasterListingEdit edit = MasterListing.Insert(listing.MasterPath, listing.Row, listing.AfterExcelDataFile, nowUtc);

                if (!edit.IsDone)
                {
                    this.thisLogger.LogError(
                        "Publish of {BoardId} stopped: the board is written but could not be added to [{Master}]: {Failure}",
                        plan.BoardId, listing.MasterPath, edit.Failure);

                    return PublishOutcome.Failed(
                        $"The board was written, but it could not be added to the drop-down lists: {edit.Failure} " +
                        "Nobody sees it until it is - re-run the publish once that is fixed.");
                }
            }

            // ---- What this publish IS ------------------------------------------------------------
            //
            // DescriptorWithBoardFile, not the planned descriptor: the planned one covers the files
            // only, and a rows-only change stores nothing - so without the board file folded in the
            // content hash would not move, and the production promotion, which compares exactly this
            // hash to decide what is waiting, would never offer the board.
            BoardDescriptor descriptor = plan.DescriptorWithBoardFile(boardFileHash);

            // A system.json an earlier build wrote into this folder: not relevant to users, and
            // no longer kept. See RetiredBoardDescriptor.
            RetiredBoardDescriptor.TryRemove(plan.DataRoot, plan.BoardFolder, this.thisLogger);

            // ---- 4. The database -----------------------------------------------------------------
            //
            // Last, so a row claiming a revision can never exist before the tree actually holds
            // it. The reverse order would leave a contributor's next submission diffing against a
            // revision that was never written.
            await this.thisStore
                .SetBoardPublishedAsync(
                    plan.BoardId,
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
                "Published {BoardId} revision {Revision} ({FileCount} files stored, board file {BoardFile}).",
                plan.BoardId,
                descriptor.Revision,
                stored.Count,
                plan.BoardFileRelative);

            return PublishOutcome.Published(descriptor, boardFileHash, stored.Count);
        }

        // ###########################################################################################
        // The pre-write check. Null when everything may be written, otherwise the reason it may not.
        //
        // Also covers the GENERATED board file for links: it is written through the same folders,
        // and a link there redirects it just as well.
        // ###########################################################################################
        private async Task<string?> CheckBeforeWritingAsync(PublishPlanDetail plan, MasterRowInsert? listing, CancellationToken cancellationToken)
        {
            IReadOnlyList<PlannedFile> stored = plan.StoredFiles;

            IEnumerable<string> destinations = stored
                .Select(file => file.AbsolutePath)
                .Append(plan.BoardFilePath);

            // The main file too, for a new board: written through the same root.
            if (listing is not null)
            {
                destinations = destinations.Append(listing.MasterPath);

                // Checked before the first byte, like everything else here: a row placed after one
                // that has since left the list must not be found out after the board is written.
                if (!MasterListing.TryRead(listing.MasterPath, out IReadOnlyList<MasterListingRow> rows, out string why))
                    return $"The main data file could not be read ({why}), so the new board could not be listed. Nothing was changed.";

                if (MasterListing.IndexOfBoard(rows, listing.Row.BoardId) < 0 &&
                    !BoardListingRules.AfterIsListed(rows, listing.AfterExcelDataFile))
                {
                    return $"The new board was placed after [{BoardDescriptorRules.BoardIdFromExcelDataFile(listing.AfterExcelDataFile)}], which is no longer in the " +
                        "drop-down lists. Place it again in the Boards screen. Nothing was changed.";
                }

                if (MasterListing.NamesTakenBy(rows, listing.Row.BoardId, listing.Row.HardwareName, listing.Row.BoardName) is MasterListingRow taken)
                    return $"{MasterListing.NamesTakenMessage(taken)} Nothing was changed.";
            }

            List<string> destinationList = destinations.ToList();

            foreach (string destination in destinationList)
            {
                string? link = PublishPathSafety.FindLinkOnPath(plan.DataRoot, destination, PublishPathSafety.IsLink);

                if (link is not null)
                {
                    this.thisLogger.LogError(
                        "Publish of {BoardId} refused: [{Link}] is a symbolic link on the way to [{Destination}].",
                        plan.BoardId, link, destination);

                    // Named from the data root, never as the server's absolute path (code review,
                    // 2026-10-13) - the log line above has the whole path for the administrator.
                    return $"The BETA data contains a symbolic link at [{TreeWriteAccess.DisplayName(plan.DataRoot, link)}], and publishing through it " +
                        "could write outside the data. Nothing was changed. Remove the link on the server.";
                }
            }

            // ###########################################################################################
            // *** EVERY FOLDER THIS WRITES INTO MAY BE WRITTEN (owner report, 2026-09-28). *** A folder
            // copied into BETA by hand as another user refuses every write, and finding that out half
            // way through is the half-published board this ordering exists to avoid. After the link
            // check, which a probe - a real file - must never get ahead of. See TreeWriteAccess.
            // ###########################################################################################
            IReadOnlyList<string> refusing = TreeWriteAccess.FoldersRefusing(plan.DataRoot, destinationList, this.CanWriteFolderForTests);

            if (refusing.Count > 0)
            {
                this.thisLogger.LogError(
                    "Publish of {BoardId} refused before writing anything: the service may not write into {Folders}. "
                    + "Files copied in by hand as another user do this. Give the service its access back with: {Command}",
                    plan.BoardId,
                    string.Join(", ", refusing),
                    TreeWriteAccess.FixCommand(refusing));

                return TreeWriteAccess.RefusalMessage("BETA", plan.DataRoot, refusing);
            }

            foreach (PlannedFile file in stored)
            {
                cancellationToken.ThrowIfCancellationRequested();

                BlobVerification blob = await this.thisBlobs
                    .VerifyAsync(file.Sha256, SubmissionContentRules.SampleBytes, cancellationToken)
                    .ConfigureAwait(false);

                if (!blob.Exists)
                {
                    this.thisLogger.LogError(
                        "Publish of {BoardId} stopped: blob {Hash} for {Name} is not in the store.",
                        plan.BoardId, file.Sha256, file.Name);

                    return $"The content for [{file.Name}] is no longer in the blob store, so the publish " +
                        "was stopped before any board file was changed.";
                }

                if (!blob.HashMatches)
                {
                    this.thisLogger.LogError(
                        "Publish of {BoardId} stopped: blob {Hash} for {Name} no longer hashes to its name.",
                        plan.BoardId, file.Sha256, file.Name);

                    return $"The stored content for [{file.Name}] no longer matches what was submitted, so " +
                        "the publish was stopped before any board file was changed.";
                }

                // The finalise-time rules again - a KiCad schematic's sheets read from the whole
                // file (KiCadSheetReferences) - so a submission queued before a rule existed cannot
                // carry past it into BETA. Every name of the same bytes is judged by its own type.
                foreach (PlannedFile named in plan.Files.Where(other => string.Equals(other.StoredPath, file.StoredPath, StringComparison.Ordinal)))
                {
                    if (!SubmissionContentRules.MatchesWholeFile(
                            named.Name, blob.Head, () => this.thisBlobs.OpenText(named.Sha256), out string reason))
                    {
                        return $"[{named.Name}] cannot be published: {reason} Nothing was changed.";
                    }
                }
            }

            return null;
        }

        // ###########################################################################################
        // The SHA-256 of a file on disk, streamed. Lower-case hex, matching
        // BoardDescriptorRules.ComputeContentHash and dataChecksums.json - a hash that differs only
        // in case compares unequal everywhere it is used, which reads as every client being
        // permanently out of date.
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
    // A new board's row for the main file: which file, the row, and the row it goes after (null:
    // first). See MasterListing.
    // ###########################################################################################
    public sealed record MasterRowInsert(string MasterPath, MasterListingRow Row, string? AfterExcelDataFile);

    // ###########################################################################################
    // What a publish did, or why it stopped.
    //
    // The descriptor is returned rather than merely written, because the caller records the
    // revision in its own audit trail and would otherwise have to read the file back to learn
    // what it just published. WorkbookSha256 is the board file's.
    // ###########################################################################################
    public sealed class PublishOutcome
    {
        private PublishOutcome(bool isPublished, BoardDescriptor? descriptor, string workbookSha256, int fileCount, string? failure)
        {
            this.IsPublished = isPublished;
            this.Descriptor = descriptor;
            this.WorkbookSha256 = workbookSha256;
            this.FileCount = fileCount;
            this.Failure = failure;
        }

        public bool IsPublished { get; }

        public BoardDescriptor? Descriptor { get; }

        public string WorkbookSha256 { get; }

        public int FileCount { get; }

        // Written for a human reading a log or a review screen, not for a machine to branch on.
        public string? Failure { get; }

        public static PublishOutcome Published(BoardDescriptor descriptor, string workbookSha256, int fileCount) =>
            new(true, descriptor, workbookSha256, fileCount, null);

        public static PublishOutcome Failed(string failure) =>
            new(false, null, string.Empty, 0, failure);
    }
}
