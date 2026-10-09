using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // APPROVING a submission: the whole chain, from a maintainer's decision to a published board
    // (NewContributeStrategy.md Phase 5, tasks 5 and 6).
    //
    // *** THIS IS THE ONLY IRREVERSIBLE OPERATION IN THE BOARD. *** Task 7 was struck by the
    // project owner, so no publish history is retained: this overwrites the published board in place
    // and the only way back is to publish a correction. Every decision below is made on that
    // basis.
    //
    // *** THE ORDER OF THE CHECKS IS THE DESIGN, not incidental. *** Cheap refusals first, so the
    // dangerous work only ever runs after everything that could refuse has:
    //
    //   1. AUTHORITY   - may this account publish anything at all? (an administrator, or a
    //                    maintainer of at least one board)
    //   2. EXISTENCE   - is there such a submission?
    //   3. AUTHORITY   - over THIS submission's board (ReviewDecisionRules, through
    //      AND STATE     ReviewAuthority), and is it still undecided? (the double-publish
    //                    interlock)
    //   4. PAYLOAD     - can its rows actually be read?
    //   5. PLAN        - are all the paths safe, and which generation is being written?
    //   6. APPROVALS   - is this the approval that publishes, or the first of two? A first one is
    //                    recorded here and nothing is written - only once everything above has
    //                    passed, so no approval is ever recorded for what cannot be published.
    //   7. WRITE       - only now does anything touch the tree.
    //
    // Every refusal before step 6 leaves the tree and the submission exactly as they were, so a
    // refused approval can simply be retried once whatever was wrong is fixed.
    //
    // *** A MISSING PAYLOAD IS REFUSED, and that check is load-bearing rather than defensive. ***
    // Without the payload there are no rows, and a board built from no rows is EMPTY. Publishing
    // it would delete the board's entire contents - every component, every highlight - over a
    // board that was working, irreversibly. It is the single most damaging thing this class could
    // be made to do, and it would look like an ordinary successful publish.
    //
    // A COORDINATOR, not a worker: PublishMerge decides what the board becomes, PublishPlan
    // decides where it may be written, PublishExecutor writes it, ReviewDecisionRules decides who
    // may ask. This class only sequences them, which is why it has no logic worth unit testing in
    // isolation - the tests drive the real chain against a real temp tree.
    // ###########################################################################################
    public sealed class ApprovePublishFlow
    {
        private readonly PublishExecutor thisExecutor;
        private readonly PublishedBoardReader thisPublishedBoards;
        private readonly ISubmissionStore thisStore;
        private readonly IAccountStore thisAccounts;
        private readonly ILogger<ApprovePublishFlow> thisLogger;
        private readonly PublishLock thisLock;

        // `publishLock` is shared with ProductionPromotionFlow (a DI singleton): a promotion copies
        // OUT of BETA, and must never see half of a publish INTO it. Optional only so a test that
        // exercises the publish alone need not build one.
        private readonly bool thisOneSubmissionInBeta;

        // Asked when step 3b finds the board "waiting in BETA": production may already hold it,
        // copied there by hand - see ProductionPromotionFlow.RecordIfProductionAlreadyHoldsAsync.
        // Optional so a test of the publish alone need not build one.
        private readonly ProductionPromotionFlow? thisProduction;
        private readonly ServerOptions? thisOptions;

        public ApprovePublishFlow(
            PublishExecutor executor,
            PublishedBoardReader publishedBoards,
            ISubmissionStore store,
            IAccountStore accounts,
            ILogger<ApprovePublishFlow> logger,
            PublishLock? publishLock = null,
            ServerOptions? options = null,
            ProductionPromotionFlow? production = null)
        {
            this.thisProduction = production;
            this.thisOptions = options;
            // One submission in BETA per board (step 3b) - only where there IS a production to
            // publish to. Without one nothing ever leaves BETA, and the rule would close every
            // board to further approvals after its first.
            this.thisOneSubmissionInBeta = options?.IsProductionPublishingConfigured == true;
            this.thisExecutor = executor;
            this.thisPublishedBoards = publishedBoards;
            this.thisStore = store;
            this.thisAccounts = accounts;
            this.thisLogger = logger;
            this.thisLock = publishLock ?? new PublishLock();
        }

        // ###########################################################################################
        // Approves a submission and publishes it.
        //
        // Returns an outcome rather than throwing, because every caller is an HTTP endpoint that
        // has to turn a failure into a status code and a sentence a maintainer can act on.
        // ###########################################################################################
        public Task<ApproveOutcome> ApproveAsync(
            long submissionId,
            ReviewAccess? access,
            string? dataTreeRoot,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default) =>
            this.ApproveAsync(submissionId, access, dataTreeRoot, nowUtc, shownRemovals: null, cancellationToken);

        // `shownRemovals` is the list of files the maintainer was shown this publish would remove
        // (FileRemovalPreview). The publish is refused when the list it would now remove differs -
        // see step 5b. Null means none were shown, which matches only "nothing to remove".
        public async Task<ApproveOutcome> ApproveAsync(
            long submissionId,
            ReviewAccess? access,
            string? dataTreeRoot,
            DateTimeOffset nowUtc,
            IReadOnlyCollection<string>? shownRemovals,
            CancellationToken cancellationToken = default)
        {
            // ---- 1. Authority, the cheap half ---------------------------------------------------
            //
            // An account that may publish NOTHING is refused before anything is loaded. Which
            // board it may publish is asked at step 3, once the submission says which it is.
            if (!ReviewAuthority.CanReviewAnything(access))
            {
                return ApproveOutcome.Refused(
                    ReviewAuthority.DescribeRefusal(access, null),
                    isForbidden: true);
            }

            if (string.IsNullOrWhiteSpace(dataTreeRoot))
            {
                // A missing setting must never make the service's own working directory the data
                // tree - that is where its configuration and binaries live.
                this.thisLogger.LogError(
                    "Approval of submission {SubmissionId} refused: no data tree root is configured.",
                    submissionId);

                return ApproveOutcome.Refused("The server has no data tree configured to publish into.");
            }

            // Everything from here to the database row happens under the one publish lock - see
            // PublishLock. A production promotion reading this board's BETA folder waits.
            using IDisposable gate = await this.thisLock.EnterAsync(cancellationToken).ConfigureAwait(false);

            // ---- 2. Existence ------------------------------------------------------------------
            SubmissionRecord? record = await this.thisStore
                .FindAsync(submissionId, cancellationToken)
                .ConfigureAwait(false);

            if (record is null)
                return ApproveOutcome.NotFound();

            // ---- 3. Authority over THIS board, then state -------------------------------------
            //
            // The board half first, and as a FORBIDDEN: a maintainer of another board is refused
            // whatever the state.
            if (!ReviewAuthority.CanPublish(access, record))
                return ApproveOutcome.Refused(ReviewAuthority.DescribeRefusal(access, record), isForbidden: true);

            // Then the double-publish interlock. Re-read from the store rather than trusted from
            // the caller: the app's belief that a submission is still pending can be seconds out
            // of date, and with no revision history a second publish is not a harmless no-op.
            if (!ReviewDecisionRules.CanApprove(access, record, out string why))
                return ApproveOutcome.Refused(why, isConflict: true);

            // ---- 3b. One submission in BETA per board (owner decision, 2026-09-27) --------------
            //
            // "It should be possible only to submit ONE contributor submission to Beta per board":
            // a push-back returns everything merged since the last promotion, and a publish replaces
            // the board's rows wholesale, so work from two contributors in BETA at once could only be
            // pushed back together. So while the board waits in BETA for production, no other
            // submission of it is approved - not even the first of two approvals, which would only
            // promise a publish this rule will refuse. Reviewing, changing, requesting changes and
            // rejecting all go on as normal. See OneSubmissionInBeta.
            if (this.thisOneSubmissionInBeta)
            {
                BoardRecord? boardRecord = await this.thisStore.FindBoardAsync(record.BoardId, cancellationToken).ConfigureAwait(false);

                // Unless production already holds that BETA state - a board copied there by hand
                // was never promoted, and would otherwise block the board for ever (code review,
                // 2026-09-29). Then it is recorded as in production and the approval goes on.
                if (ProductionPromotionRules.IsAwaitingProduction(boardRecord) &&
                    !(this.thisProduction is not null && this.thisOptions is not null &&
                      await this.thisProduction.RecordIfProductionAlreadyHoldsAsync(
                          access, boardRecord!, this.thisOptions, nowUtc, useCache: false, cancellationToken).ConfigureAwait(false)))
                {
                    return ApproveOutcome.Refused(OneSubmissionInBeta.BusyMessage(record.BoardId), isConflict: true);
                }
            }

            // ---- 4. The payload ----------------------------------------------------------------
            //
            // See the class header: an unreadable payload means no rows, and publishing no rows
            // deletes the board.
            SubmissionManifest? manifest = await this.thisStore
                .LoadPayloadAsync(submissionId, cancellationToken)
                .ConfigureAwait(false);

            if (manifest is null)
            {
                this.thisLogger.LogError(
                    "Approval of submission {SubmissionId} refused: its payload could not be loaded, " +
                    "so publishing would write an empty board over a working one.",
                    submissionId);

                return ApproveOutcome.Refused(
                    "This submission's contents could not be loaded, so it cannot be published.");
            }

            // ---- 5. The plan -------------------------------------------------------------------
            //
            // Decides every path and refuses every unsafe one BEFORE a byte is written, which is
            // what makes an irreversible operation safe to start.
            BoardData? published = await this.thisPublishedBoards
                .TryReadAsync(dataTreeRoot, manifest, cancellationToken)
                .ConfigureAwait(false);

            BoardData board = PublishMerge.Build(manifest, published);

            PublishPlanResult plan = ApprovePublishFlow.BuildPlan(
                dataTreeRoot, manifest, published, nowUtc);

            if (!plan.IsPlanned)
            {
                string reasons = string.Join(" ", plan.Problems.Select(problem => problem.Message));

                this.thisLogger.LogWarning(
                    "Approval of submission {SubmissionId} refused by the publish plan: {Reasons}",
                    submissionId, reasons);

                return ApproveOutcome.Refused($"This submission cannot be published: {reasons}");
            }

            // ---- 5b. A NEW board must have its place in the drop-down lists (owner request,
            //          2026-09-27): "This must be done before it can be pushed to BETA." -----------
            //
            // Before any approval is recorded, so nobody approves what cannot be listed. See
            // ListingForPublishAsync.
            ListingForPublish listing = await ApprovePublishFlow
                .ListingForPublishAsync(dataTreeRoot, manifest, plan.Plan!, this.thisStore, cancellationToken)
                .ConfigureAwait(false);

            if (!listing.IsReady)
                return ApproveOutcome.Refused(listing.Problem, isConflict: true);

            // A board with no "# Hardware:" / "# Board:" caption of its own - a new board, or one
            // published before the caption was kept - takes the names it is listed under.
            board = PublishMerge.CaptionedAs(board, listing.ListedAs?.HardwareName, listing.ListedAs?.BoardName);

            // ---- 6. Is this the approval that publishes? ---------------------------------------
            //
            // A submission changing a shared file needs a maintainer of the board AND an
            // administrator (owner decision, 2026-09-25); the first of them is recorded here
            // and the submission waits, in the queue, as 'approved'. Nothing is published until
            // the second arrives. See ApprovalRules.
            //
            // AFTER the payload and the plan (code review, 2026-09-25): both are read-only, and an
            // approval recorded for something the plan then refuses told the other approver their
            // approval was needed for a submission nobody could publish.
            //
            // Whether a shared file changes is decided against the tree AS IT IS NOW, not only as
            // it was when the submission arrived - see TouchesSharedFilesNowAsync.
            bool touchesSharedFiles = await ApprovePublishFlow
                .TouchesSharedFilesNowAsync(record, manifest, dataTreeRoot, this.thisStore, cancellationToken)
                .ConfigureAwait(false);

            ApprovalStatus approval = await ApprovePublishFlow.ApprovalStatusAsync(
                access, record, touchesSharedFiles, this.thisStore, this.thisAccounts, cancellationToken).ConfigureAwait(false);

            if (!approval.CanApprove)
                return ApproveOutcome.Refused(ApprovePublishFlow.WhyNot(approval, "this"), isConflict: true);

            if (!approval.ApprovalPublishes)
            {
                await this.thisStore
                    .AddApprovalAsync(submissionId, approval.YourRole!.Value, access!.Account.Id, ApprovePublishFlow.Label(access), nowUtc, cancellationToken)
                    .ConfigureAwait(false);

                await this.thisStore
                    .SetStateAsync(submissionId, SubmissionState.Approved, nowUtc, cancellationToken)
                    .ConfigureAwait(false);

                IReadOnlyList<ApproverRole> stillWaiting = approval.WaitingFor
                    .Where(role => role != approval.YourRole)
                    .ToList();

                this.thisLogger.LogInformation(
                    "Submission {SubmissionId} approved by {Account} as {Role}; waiting for {Waiting}.",
                    submissionId, access.Account.Email, approval.YourRole, string.Join(", ", stillWaiting));

                return ApproveOutcome.AwaitingApproval(stillWaiting);
            }

            // ---- 6b. What it removes, and is that what the maintainer was shown? ----------------
            //
            // Files the board stops citing that nothing else in the tree uses are removed after
            // the write (owner decision, 2026-09-25: no orphan files). The maintainer saw that
            // list before approving; if it has changed since - another publish started or stopped
            // citing a shared file - they look again rather than having files removed they were
            // never shown. No list at all is an older maintainer application - see
            // RemovalsNotSentMessage.
            FileRemovalPreview removals = ApprovePublishFlow.PreviewRemovals(dataTreeRoot, published, board, plan);

            if (!removals.Matches(shownRemovals))
            {
                return shownRemovals is null
                    ? ApproveOutcome.Refused(ApprovePublishFlow.RemovalsNotSentMessage)
                    : ApproveOutcome.Refused(ApprovePublishFlow.RemovalsChangedMessage, isConflict: true);
            }

            // ---- 6c. What it changes, while the board before still exists (owner request,
            //          2026-10-04) ---------------------------------------------------------------
            //
            // The History view's summary of this submission. Taken HERE because nothing else can:
            // the write below replaces BETA's board, and no copy of it is kept. Never stops the
            // publish - see SummariseChanges.
            SubmissionChanges? changes = ApprovePublishFlow.SummariseChanges(
                dataTreeRoot, manifest, published, board, plan.Plan!, this.thisLogger, submissionId);

            // ---- 7. Write --------------------------------------------------------------------
            //
            // The first step that touches the tree. Everything above can refuse for free.
            PublishOutcome outcome = await this.thisExecutor
                .ExecuteAsync(
                    plan.Plan!,
                    board,
                    PublishMerge.CalibrationsOf(manifest),
                    submissionId,
                    nowUtc,
                    listing.Insert,
                    cancellationToken)
                .ConfigureAwait(false);

            if (!outcome.IsPublished)
            {
                // The submission stays PENDING deliberately, so it can be retried once whatever
                // stopped it is fixed - a missing blob being the likely cause. Marking it failed
                // would need somebody to resurrect it by hand.
                return ApproveOutcome.Refused(outcome.Failure ?? "The publish did not complete.");
            }

            // ---- 8. Remove what the board no longer uses ---------------------------------------
            //
            // Only the files the maintainer was shown, and only those the tree - read again, now
            // that the new workbook is written - still does not use. Never fails the publish: it
            // has already happened.
            UnusedFileRemoval removal = UnusedFileRemover.Remove(dataTreeRoot, removals.Files, this.thisLogger);

            // The approval that published it, beside any earlier one.
            await this.thisStore
                .AddApprovalAsync(submissionId, approval.YourRole!.Value, access!.Account.Id, ApprovePublishFlow.Label(access), nowUtc, cancellationToken)
                .ConfigureAwait(false);

            // The decision row, recording WHO approved it. PublishExecutor already moved the state
            // to merged; this adds the maintainer and the audit trail, which SetStateAsync cannot.
            await this.thisStore
                .SetDecisionAsync(
                    submissionId,
                    SubmissionState.Merged,
                    access!.Account.Id,
                    comment: null,
                    nowUtc,
                    cancellationToken)
                .ConfigureAwait(false);

            // What it changed, for the board's History view - with the files step 8 removed. A
            // record that cannot be written costs the history one summary, never the publish.
            if (changes is not null)
            {
                try
                {
                    await this.thisStore
                        .SetChangesAsync(submissionId, SubmissionChangeFacts.WithRemovedFiles(changes, removal.Removed), nowUtc, cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    this.thisLogger.LogWarning(ex, "What submission {SubmissionId} changed could not be recorded for its history.", submissionId);
                }
            }

            this.thisLogger.LogInformation(
                "Submission {SubmissionId} was published to {BoardId} at revision {Revision} by {Account}.",
                submissionId, manifest.BoardId, outcome.Descriptor!.Revision, access.Account.Email);

            return ApproveOutcome.Published(outcome.Descriptor) with { RemovedFiles = removal.Removed };
        }

        // ###########################################################################################
        // WHAT A PUBLISH CHANGES (owner request, 2026-10-04: "summarize each submission change in a
        // textual form ... to get an idea, besides the sometimes vague description from the
        // contributor"), worked out just before the write: the review's own comparison of BETA's
        // board with the one about to replace it (ReviewSummary - rows, highlights, calibrations),
        // and which of the plan's files are new at their path and which replace other bytes. The
        // files it removes are added after step 8 (SubmissionChangeFacts.WithRemovedFiles).
        //
        // *** NEVER STOPS THE PUBLISH. *** The summary is history, not a check: a comparison that
        // throws is logged and the submission simply has no summary - the same as one published
        // before summaries existed.
        // ###########################################################################################
        internal static SubmissionChanges? SummariseChanges(
            string dataTreeRoot,
            SubmissionManifest manifest,
            BoardData? published,
            BoardData board,
            PublishPlanDetail plan,
            ILogger logger,
            long submissionId)
        {
            try
            {
                ReviewChangeSummary rows = ReviewSummary.Compare(
                    published,
                    board,
                    manifest.Renames,
                    publishedCalibrations: ReviewEndpoints.PublishedCalibrations(dataTreeRoot, manifest, published),
                    submittedCalibrations: PublishMerge.CalibrationsOf(manifest));

                (IReadOnlyList<string> added, IReadOnlyList<string> replaced) = SubmissionChangeFacts.SplitWrites(
                    plan.Files.Select(file => (file.RelativePath, file.AbsolutePath)));

                return SubmissionChangeFacts.Build(rows, added, replaced);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                logger.LogWarning(ex, "What submission {SubmissionId} changes could not be worked out; it is published without a summary.", submissionId);
                return null;
            }
        }

        public const string RemovalsChangedMessage =
            "The files this publish would remove have changed since you opened it - another publish has " +
            "changed what uses them. Open the submission again and check the list, then approve.";

        // ###########################################################################################
        // WHETHER THIS PUBLISH MAY GO AHEAD AS FAR AS THE DROP-DOWN LISTS ARE CONCERNED, and the row
        // it must add (owner request, 2026-09-27). CRT finds boards ONLY through the newest main Excel
        // data file, so a board it does not list is published to nobody.
        //
        //   - listed already: nothing to add - every existing board, the ordinary case;
        //   - not listed: a maintainer must have PLACED it (the Boards screen) - refused until then,
        //     and the publish adds its row where they placed it;
        //   - no readable file: a board that is NEW to the tree is refused, since it could never be
        //     listed; a board already in the tree is let through exactly as before this existed -
        //     the file was never part of publishing one, and must not start blocking it.
        // ###########################################################################################
        internal static async Task<ListingForPublish> ListingForPublishAsync(
            string dataTreeRoot,
            SubmissionManifest manifest,
            PublishPlanDetail plan,
            ISubmissionStore store,
            CancellationToken cancellationToken)
        {
            bool isNewToTree = !PublishedBoardLocator.Locate(dataTreeRoot, manifest).Exists;

            string? master = MasterListing.NewestMasterPath(dataTreeRoot);

            if (master is null || !MasterListing.TryRead(master, out IReadOnlyList<MasterListingRow> rows, out _))
            {
                return isNewToTree
                    ? ListingForPublish.Refused(master is null
                        ? BoardListingRules.NoMasterMessage
                        : "BETA's main Excel data file could not be read, so the new board could not be added to the drop-down lists.")
                    : ListingForPublish.NothingToAdd;
            }

            int listed = MasterListing.IndexOfBoard(rows, manifest.BoardId);

            if (listed >= 0)
                return ListingForPublish.Listed(rows[listed]);

            BoardPlacement? placement = await store.GetPlacementAsync(manifest.BoardId, cancellationToken).ConfigureAwait(false);

            if (placement is null)
                return ListingForPublish.Refused(BoardListingRules.NotPlacedMessage);

            // Another board listed under the same names since this one was placed.
            if (MasterListing.NamesTakenBy(rows, manifest.BoardId, placement.HardwareName, placement.BoardName) is MasterListingRow taken)
                return ListingForPublish.Refused(MasterListing.NamesTakenMessage(taken));

            if (!BoardListingRules.AfterIsListed(rows, placement.AfterExcelDataFile))
            {
                return ListingForPublish.Refused(
                    $"It was placed after [{placement.AfterExcelDataFile}], which is no longer in the drop-down lists. " +
                    "Place it again in the Boards screen.");
            }

            return ListingForPublish.Add(new MasterRowInsert(
                master,
                BoardListingRules.RowFor(placement, dataTreeRoot, plan.WorkbookPath),
                placement.AfterExcelDataFile));
        }

        // ###########################################################################################
        // The refusal for a client (an older CRT) that sent NO list at all (code review, 2026-09-25) -
        // shared with ProductionPromotionFlow. The current application always sends its list, empty
        // or not, so none at all means one built before the list existed. It was told the list had
        // "changed since you opened it", which was false and which reopening could never fix.
        // ###########################################################################################
        public const string RemovalsNotSentMessage =
            "This would remove files that nothing uses any more, and your copy of CRT did not send " +
            "the list of them it showed you - it is older than the server. Update CRT, " +
            "then open this again and check the list before approving.";

        // ###########################################################################################
        // THE FILES THIS PUBLISH WOULD REMOVE, shown to the maintainer before they approve and
        // checked again when they do - the same computation both times.
        //
        // The candidates are what the board stops citing; the tree is then read with the workbook
        // the plan WRITES standing in for its new citations, so a file another board - or an older
        // generation of this one - still cites is kept. Nothing to preview when the board is new
        // or the plan is refused.
        // ###########################################################################################
        internal static FileRemovalPreview PreviewRemovals(
            string dataTreeRoot,
            BoardData? published,
            BoardData board,
            PublishPlanResult plan)
        {
            if (published is null || !plan.IsPlanned)
                return FileRemovalPreview.Nothing;

            string workbook = Path
                .GetRelativePath(Path.GetFullPath(dataTreeRoot), plan.Plan!.WorkbookPath)
                .Replace(Path.DirectorySeparatorChar, '/');

            // Only inside the board's own folder (owner decision, 2026-09-27) - a shared file or
            // another board's the board stops citing stays, for Account > Unused files.
            IReadOnlyList<string> after = SubmissionManifestBuilder.CollectReferencedFiles(board);
            IReadOnlyList<string> candidates = AutomaticRemovalScope.Within(
                BoardDescriptorRules.BoardIdFromExcelDataFile(workbook),
                DataTreeUsage.NoLongerCited(SubmissionManifestBuilder.CollectReferencedFiles(published), after));

            if (candidates.Count == 0)
                return FileRemovalPreview.Nothing;

            return FileRemovalPreview.Compute(
                dataTreeRoot,
                candidates,
                new Dictionary<string, IReadOnlyCollection<string>> { [workbook] = after.ToList() },
                ApprovePublishFlow.PreviewReads);
        }

        // ###########################################################################################
        // The preview for the review screen: the same plan the approval would build - but only when
        // the board stops citing something (code review, 2026-09-25). The submission detail asks
        // for this on every click on a queue row; building the plan hashes files, and reading the
        // tree parses every workbook in it, so both are skipped when there is nothing to remove,
        // which is nearly every submission. What IS read goes through PreviewReads.
        // ###########################################################################################
        internal static FileRemovalPreview PreviewRemovals(
            string dataTreeRoot,
            SubmissionManifest manifest,
            BoardData? published,
            DateTimeOffset nowUtc)
        {
            if (published is null)
                return FileRemovalPreview.Nothing;

            BoardData board = PublishMerge.Build(manifest, published);

            IReadOnlyList<string> candidates = DataTreeUsage.NoLongerCited(
                SubmissionManifestBuilder.CollectReferencedFiles(published),
                SubmissionManifestBuilder.CollectReferencedFiles(board));

            if (candidates.Count == 0)
                return FileRemovalPreview.Nothing;

            return ApprovePublishFlow.PreviewRemovals(
                dataTreeRoot,
                published,
                board,
                ApprovePublishFlow.BuildPlan(dataTreeRoot, manifest, published, nowUtc));
        }

        // ###########################################################################################
        // The workbook reads every removal PREVIEW shares - the submission detail, the approval's
        // check against what was shown, the production plan and the administrator's list of unused
        // files - so a workbook is parsed once per version rather than once per click. Never used
        // to decide a deletion: UnusedFileRemover reads the tree afresh. See WorkbookReadCache.
        // ###########################################################################################
        internal static WorkbookReadCache PreviewReads { get; } = new();

        // ###########################################################################################
        // Where a submission's approval stands for this account - the rule is ApprovalRules; this
        // gathers what it needs: whether the board has maintainers, what has been given, and the
        // account's role. `touchesSharedFiles` is TouchesSharedFilesNow's answer, which the caller
        // already holds. Shared with the detail endpoint, so the Maintainer tab shows exactly
        // what an approval here would do.
        // ###########################################################################################
        internal static async Task<ApprovalStatus> ApprovalStatusAsync(
            ReviewAccess? access,
            SubmissionRecord record,
            bool touchesSharedFiles,
            ISubmissionStore store,
            IAccountStore accounts,
            CancellationToken cancellationToken)
        {
            // A maintainer who can actually give the maintainer half - see
            // ReviewAuthority.CanGiveMaintainerApproval for the rows that cannot.
            bool hasMaintainers = touchesSharedFiles &&
                (await accounts.GetMaintainersOfBoardAsync(record.BoardId, cancellationToken).ConfigureAwait(false))
                    .Any(ReviewAuthority.CanGiveMaintainerApproval);

            IReadOnlyList<GivenApproval> given = await store
                .GetApprovalsAsync(record.Id, cancellationToken)
                .ConfigureAwait(false);

            return ApprovalRules.Status(
                ApprovalRules.Required(touchesSharedFiles, hasMaintainers),
                given,
                ReviewAuthority.RoleIn(access, record.BoardId),
                access?.Account.Id);
        }

        // ###########################################################################################
        // *** DOES THIS SUBMISSION CHANGE A SHARED FILE - JUDGED AGAINST THE TREE AS IT IS NOW
        // (code review, 2026-09-25). ***
        //
        // The flag stored at create compared the submission with the tree of THAT moment. A
        // submission citing a shared file unchanged was an ordinary one-approval item for ever -
        // and once another publish changed that file, its now-old bytes would revert the file for
        // every board that cites it, on a single maintainer's approval, with the administrator never
        // asked. So the stored flag is only a floor: true stays true, and the submission is checked
        // afresh against the published tree (SubmissionSharedFiles, the create-time rule).
        //
        // No manifest to check (an unreadable payload) leaves the stored flag - nothing else can be
        // known, and such a submission is refused before it publishes anyway.
        // ###########################################################################################
        // ###########################################################################################
        // Does approving this REPLACE a shared file as the tree stands NOW (SubmissionSharedFiles)?
        //
        // *** DECIDED AFRESH, NOT FROM THE FLAG STORED AT CREATE (2026-09-27). *** Until then the
        // stored flag was a floor, only ever raised - and it was set for merely ADDING a shared
        // file, which no longer needs the administrator (owner decision). Trusting it would keep
        // every submission queued before that change waiting for a second approval nobody needs.
        // What the approval would do to the tree now is the question, so that is what is asked;
        // the flag is only the answer when the payload cannot be read.
        // ###########################################################################################
        internal static bool TouchesSharedFilesNow(SubmissionRecord record, SubmissionManifest? manifest, string? dataTreeRoot) =>
            manifest is null
                ? record.TouchesSharedFiles
                : SubmissionSharedFiles.TouchesSharedFiles(manifest, PublishedTreeProbe.For(dataTreeRoot));

        // The same, and the answer is STORED when it differs from the flag, so the queue marks the
        // submission as needing the administrator - or no longer needing them - in every window,
        // not only in this one.
        private static async Task<bool> TouchesSharedFilesNowAsync(
            SubmissionRecord record,
            SubmissionManifest manifest,
            string dataTreeRoot,
            ISubmissionStore store,
            CancellationToken cancellationToken)
        {
            bool touches = ApprovePublishFlow.TouchesSharedFilesNow(record, manifest, dataTreeRoot);

            if (touches != record.TouchesSharedFiles)
                await store.SetTouchesSharedFilesAsync(record.Id, touches, cancellationToken).ConfigureAwait(false);

            return touches;
        }

        // ###########################################################################################
        // Why this account's approval cannot be added: this account gave it already; somebody else
        // gave the approval in this account's role (a board's second maintainer - told who, rather
        // than "you have already approved" for something they never did); or not a role this item
        // needs (a maintainer on a board whose shared-file change is the administrator's alone -
        // which only a pool changed mid-request can produce).
        // ###########################################################################################
        internal static string WhyNot(ApprovalStatus approval, string what)
        {
            string waiting = ApprovePublishFlow.Describe(approval.WaitingFor);

            if (approval.YouApproved)
                return $"You have already approved {what}. It is waiting for {waiting}.";

            GivenApproval? sameRole = approval.YourRole is null
                ? null
                : approval.Given.FirstOrDefault(item => item.Role == approval.YourRole);

            if (sameRole is not null)
            {
                string role = sameRole.Role == ApproverRole.Administrator ? "the administrator" : "a maintainer of the board";

                return $"{sameRole.By} has already approved {what} as {role}. It is waiting for {waiting}.";
            }

            return $"This needs the approval of {waiting}.";
        }

        // "the administrator", "a maintainer of the board", or both - for a sentence.
        internal static string Describe(IReadOnlyList<ApproverRole> roles) =>
            string.Join(" and ", roles.Select(role => role == ApproverRole.Administrator ? "the administrator" : "a maintainer of the board"));

        internal static string Label(ReviewAccess access) =>
            $"{access.Account.DisplayName} ({access.Account.Email})";

        // ###########################################################################################
        // The publish plan for this submission.
        //
        // *** ORIGIN IS SET ONCE AND NEVER RECOMPUTED. *** A board that arrived through this
        // pipeline stays "contributed" however many times it is later revised, including by the
        // project owner - BoardDescriptorRules says so outright. So an EXISTING descriptor's origin
        // wins, and only a board with none at all is classified here.
        //
        // *** THE REVISION IS THE PUBLISH DATE, STAMPED HERE (owner confirmation, 2026-09-26: "when
        // the maintainer publish it to BETA, the revision date gets updated from server. Same
        // happens when it gets published to real production, so server always wins, and what is
        // typed by user is not important"). ***
        //
        // PublishExecutor has stamped the WORKBOOK with this since 2026-09-23, but the plan was
        // still built from the SUBMITTED date, so the two disagreed: the board read
        // "2026-September-21" while `boards.current_revision` kept whatever the contributor's
        // draft happened to carry. That row is what the next draft re-bases against
        // (DraftBaseRevision), so the drift check was comparing against a revision no board ever
        // held. One value, computed once, used by both - which is why it is passed IN to the
        // executor rather than each working it out.
        //
        // It also unblocks a NEW BOARD. Its seeded workbook carries no revision date at all
        // (DraftSeeder.CreateNewBoard - there is nothing to inherit one from, and CRT never asks),
        // PublishMerge's fallback to the published board finds none either, and PublishPlan then
        // refused `publish.no-revision` at the one irreversible step - reported by the project
        // owner, 2026-09-26. With the server stamping it, a new board has a revision by
        // construction and that refusal becomes unreachable through this path. It stays in
        // PublishPlan as a guard for any other caller.
        // ###########################################################################################
        internal static PublishPlanResult BuildPlan(
            string dataTreeRoot,
            SubmissionManifest manifest,
            BoardData? published,
            DateTimeOffset nowUtc)
        {
            PublishedBoardLocation location = PublishedBoardLocator.Locate(dataTreeRoot, manifest);

            string boardFolder = location.Exists
                ? location.BoardFolder
                : Path.Combine(
                    dataTreeRoot,
                    manifest.Manufacturer.Trim(),
                    manifest.Hardware.Trim(),
                    manifest.Board.Trim());

            // What is already on disk, so the plan can spot a case-only collision against a file
            // it is not itself writing.
            IEnumerable<string> existing = Directory.Exists(boardFolder)
                ? Directory.EnumerateFiles(boardFolder)
                    .Select(Path.GetFileName)
                    .Where(name => !string.IsNullOrEmpty(name))
                    .Select(name => name!)
                : [];

            return PublishPlan.Build(
                dataTreeRoot,
                boardFolder,
                existing,
                ApprovePublishFlow.BoardStem(manifest, location),
                manifest,
                BoardWorkbookStyle.FormatRevisionDate(nowUtc),
                nowUtc,

                // No maintainers and a fixed origin: the descriptor these would fill is no longer
                // written anywhere (system.json was retired, 2026-09-25). The maintainers are the
                // `maintainers` table and the origin is `boards.origin`, both in the database.
                maintainers: null,
                BoardDescriptorRules.BoardOrigin.Contributed,

                // What is published NOW, read at the moment of publishing rather than trusted from
                // when the submission arrived: another board's file it cites must still be
                // byte-identical, and no case-variant may have appeared since. (2026-09-25)
                PublishedTreeProbe.For(dataTreeRoot));
        }

        // ###########################################################################################
        // The board file's stem - "Data C64 250407", the part before the generation suffix.
        //
        // *** READ OFF THE EXISTING FILE WHERE THERE IS ONE, because it does NOT follow the folder
        // names mechanically. *** "Data C128DCR 250477" lives under C128/250477, so rebuilding it
        // from the identity would write a second, differently-named board beside the real one and
        // the board would then carry two. Only a genuinely new board falls back to the
        // convention.
        // ###########################################################################################
        private static string BoardStem(SubmissionManifest manifest, PublishedBoardLocation location)
        {
            if (location.Exists)
            {
                string name = Path.GetFileNameWithoutExtension(location.WorkbookPath);

                // Strip the generation suffix, leaving the stem the plan will re-add one to.
                Version? generation = DataGenerationRules.TryReadGeneration(
                    Path.GetFileName(location.WorkbookPath));

                if (generation is not null)
                {
                    int suffix = name.LastIndexOf(" v", StringComparison.Ordinal);

                    if (suffix > 0)
                        name = name[..suffix];
                }

                if (!string.IsNullOrWhiteSpace(name))
                    return name;
            }

            return $"Data {manifest.Hardware.Trim()} {manifest.Board.Trim()}";
        }
    }

    // What ListingForPublishAsync decided: go ahead (adding Insert when it is not null), or not.
    // ListedAs is the row the board is, or will be, listed under - the caption for a board that
    // has none (PublishMerge.CaptionedAs); null when the file could not be read.
    public sealed record ListingForPublish(bool IsReady, string Problem, MasterRowInsert? Insert, MasterListingRow? ListedAs = null)
    {
        public static readonly ListingForPublish NothingToAdd = new(true, string.Empty, null);

        public static ListingForPublish Listed(MasterListingRow row) => new(true, string.Empty, null, row);

        public static ListingForPublish Add(MasterRowInsert insert) => new(true, string.Empty, insert, insert.Row);

        public static ListingForPublish Refused(string problem) => new(false, problem, null);
    }

    // ###########################################################################################
    // What approving produced, or why it did not.
    //
    // IsForbidden and IsConflict are distinguished because the endpoint turns them into different
    // status codes, and the difference matters to the reader: "your account may not do this" and
    // "somebody else already decided this" send a maintainer to completely different places.
    // ###########################################################################################
    public sealed record ApproveOutcome(
        bool IsPublished,
        string Error,
        BoardDescriptor? Descriptor,
        bool IsNotFound = false,
        bool IsForbidden = false,
        bool IsConflict = false)
    {
        public static ApproveOutcome Published(BoardDescriptor descriptor) =>
            new(true, string.Empty, descriptor);

        public static ApproveOutcome Refused(
            string error, bool isForbidden = false, bool isConflict = false) =>
            new(false, error, null, IsForbidden: isForbidden, IsConflict: isConflict);

        public static ApproveOutcome NotFound() =>
            new(false, "No such submission.", null, IsNotFound: true);

        // The first of two approvals: recorded, nothing published, waiting for these roles.
        public static ApproveOutcome AwaitingApproval(IReadOnlyList<ApproverRole> waitingFor) =>
            new(false, string.Empty, null) { WaitingFor = waitingFor };

        public IReadOnlyList<ApproverRole> WaitingFor { get; init; } = [];

        // The files a publish removed because nothing in the tree used them any more.
        public IReadOnlyList<string> RemovedFiles { get; init; } = [];

        public bool IsAwaitingApproval => this.WaitingFor.Count > 0;
    }
}
