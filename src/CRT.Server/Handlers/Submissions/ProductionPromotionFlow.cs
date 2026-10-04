using System.Collections.Concurrent;
using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // PUBLISHING A SYSTEM FROM BETA TO PRODUCTION (owner request, 2026-09-25): "first
    // published to BETA and then it is published to the real production. The maintainer is still
    // allowed to do this, but only after he has checked that the data looks correct in BETA."
    //
    // A COORDINATOR, like ApprovePublishFlow: ProductionPromotionPlan decides what is copied,
    // ProductionPromoter copies it, ReviewAuthority decides who may. The order of the checks is the
    // design, cheapest refusal first:
    //
    //   1. CONFIGURED    - is publishing to production switched on at all?
    //   2. AUTHORITY     - may this account publish anything?
    //   3. EXISTENCE     - is there such a system, and is its BETA state ahead of production?
    //   4. WHAT WAS SEEN - is BETA still exactly what the maintainer checked? (the content hash)
    //   5. PLAN          - what would be copied, and is anything refused?
    //   6. APPROVAL      - does this account's approval publish it? A plan that changes a shared
    //                      file needs a maintainer of the board AND the administrator (owner
    //                      decision, 2026-09-25): the first is recorded against this BETA state
    //                      and nothing is copied until the second arrives. See ApprovalRules.
    //   7. COPY, then record it.
    //
    // *** "CHECKED IN BETA" IS MADE CONCRETE BY STEP 4. *** A maintainer cannot be proved to have
    // looked, but they can be held to promoting exactly what they could have looked at: the
    // request names the BETA content hash the Maintainer tab showed them, and a publish that
    // has landed in BETA since changes that hash and refuses the promotion. The Maintainer tab
    // adds the human half - a box the maintainer ticks to say they checked it in CRT.
    //
    // Steps 3 to 7 run under PublishLock, so no publish can land in BETA between the check in
    // step 4 and the copy in step 7.
    // ###########################################################################################
    public sealed class ProductionPromotionFlow
    {
        private readonly ISubmissionStore thisStore;
        private readonly IAccountStore thisAccounts;
        private readonly PublishedBoardReader thisBoards;
        private readonly PublishLock thisLock;
        private readonly ILogger<ProductionPromotionFlow> thisLogger;

        // When a system was last found NOT to be in production already, per BETA and production
        // state (see RecordIfProductionAlreadyHoldsAsync) - so the list, read every minute by
        // every open Maintainer tab, does not rebuild the same plan each time.
        private readonly ConcurrentDictionary<string, DateTimeOffset> thisFoundDifferentUtc = new(StringComparer.Ordinal);

        // How long a "production differs" answer is trusted by the list before the trees are
        // compared again - a hand copy made meanwhile is noticed within this.
        internal static readonly TimeSpan DifferentRecheck = TimeSpan.FromMinutes(10);

        public ProductionPromotionFlow(
            ISubmissionStore store,
            IAccountStore accounts,
            PublishedBoardReader boards,
            PublishLock publishLock,
            ILogger<ProductionPromotionFlow> logger)
        {
            this.thisStore = store;
            this.thisAccounts = accounts;
            this.thisBoards = boards;
            this.thisLock = publishLock;
            this.thisLogger = logger;
        }

        public const string NotConfiguredMessage =
            "Publishing to the stable source is not switched on for this server. Until it is, BETA is copied to the stable source by hand.";

        // ###########################################################################################
        // The systems whose BETA state is ahead of production, that this account may promote -
        // leaving out only what the plan alone can reveal (a shared-file change), which the plan
        // step reports when the system is opened.
        // ###########################################################################################
        //
        // With `options` (the endpoint always passes them), a system whose record says it waits but
        // whose production copy already matches BETA - copied there by hand - is recorded as in
        // production and left out; see RecordIfProductionAlreadyHoldsAsync.
        // ###########################################################################################
        public async Task<IReadOnlyList<SystemRecord>> ListAwaitingAsync(
            ReviewAccess? access,
            ServerOptions? options = null,
            DateTimeOffset? nowUtc = null,
            CancellationToken cancellationToken = default)
        {
            if (!ReviewAuthority.CanReviewAnything(access))
                return [];

            IReadOnlyList<SystemRecord> systems = await this.thisStore.ListSystemsAsync(cancellationToken);

            var awaiting = new List<SystemRecord>();

            foreach (SystemRecord system in systems
                         .Where(ProductionPromotionRules.IsAwaitingProduction)
                         .Where(system => ReviewAuthority.CanPublish(access, system.SystemId)))
            {
                if (options is not null &&
                    await this.RecordIfProductionAlreadyHoldsAsync(
                        access, system, options, nowUtc ?? DateTimeOffset.UtcNow, useCache: true, cancellationToken))
                {
                    continue;
                }

                awaiting.Add(system);
            }

            return awaiting;
        }

        // ###########################################################################################
        // THE "BETA > PROD" LIST AS THE MAINTAINER TAB READS IT: each waiting system with whether it
        // waits for THIS account and whether it carries a draft its contributor discarded.
        //
        // *** A FIXED NUMBER OF QUERIES, HOWEVER MANY SYSTEMS WAIT (code review, 2026-09-29). *** Every
        // open Maintainer tab reads this every minute, and it used to ask three queries per waiting
        // system - approvals, merged submissions and their discards - only to set two booleans. The
        // two facts are now asked for the whole list at once (GetProductionApprovalsForAsync,
        // GetSystemsCarryingDiscardedDraftsAsync), with the same bounds the panel and the mails use.
        //
        // The discard mark NEVER fails the list, as it never failed the plan: it is context.
        // ###########################################################################################
        public async Task<IReadOnlyList<ProductionListEntry>> ListEntriesAsync(
            ReviewAccess? access,
            ServerOptions options,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default)
        {
            IReadOnlyList<SystemRecord> systems = await this.ListAwaitingAsync(access, options, nowUtc, cancellationToken);

            if (systems.Count == 0 || access is null)
                return [];

            IReadOnlyDictionary<string, IReadOnlyList<GivenApproval>> approvals =
                await this.thisStore.GetProductionApprovalsForAsync(
                    systems
                        .Where(system => !string.IsNullOrWhiteSpace(system.ContentHash))
                        .Select(system => (system.SystemId, system.ContentHash!))
                        .ToList(),
                    cancellationToken);

            IReadOnlySet<string> discarded;

            try
            {
                discarded = await this.thisStore.GetSystemsCarryingDiscardedDraftsAsync(
                    systems.Select(system => (system.SystemId, system.ProductionPublishedUtc)).ToList(),
                    nowUtc,
                    cancellationToken);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                this.thisLogger.LogWarning(ex, "The Beta > Prod list could not read which systems carry a discarded draft.");
                discarded = new HashSet<string>(StringComparer.Ordinal);
            }

            return systems
                .Select(system => new ProductionListEntry(
                    system.SystemId,
                    system.Manufacturer,
                    system.Hardware,
                    system.Board,
                    BetaRevision: system.CurrentRevision,
                    BetaContentHash: system.ContentHash,
                    system.ProductionRevision,
                    system.ProductionPublishedUtc,

                    // Whether it waits for THIS account or for the other approver - the BETA badge
                    // (2026-09-27).
                    AwaitsYou: ProductionPromotionRules.AwaitsAccount(
                        approvals.TryGetValue(system.SystemId, out IReadOnlyList<GivenApproval>? given) ? given : [],
                        access.Account.Id),

                    // A contributor whose work this carries discarded their own draft (2026-09-28) -
                    // marked on the list itself, so it is seen before the system is even opened.
                    CarriesDiscardedDraft: discarded.Contains(system.SystemId)))
                .ToList();
        }

        // ###########################################################################################
        // *** A BOARD COPIED TO PRODUCTION BY HAND IS NOT WAITING FOR PRODUCTION (code review,
        // 2026-09-29). *** The record says a system waits while its BETA content hash differs from
        // the one its last PROMOTION recorded - and a board the project owner copied into production
        // by hand as root (DEPLOYMENT.md's own practice) was never promoted, so it stayed "waiting"
        // for ever: listed in Beta > Prod, and - through the one-in-BETA rule - blocking every new
        // approval of that system, with nothing telling the maintainer that a no-op "Publish to
        // production" would clear it.
        //
        // So the TREES are asked when the record says "waiting": production already holds this
        // BETA state when publishing it would copy nothing (every file the plan looks at is
        // unchanged, and there is at least one), remove nothing, and add no row to production's
        // drop-down lists. Then it is recorded as in production - the same record a promotion
        // writes, at `nowUtc` - and audited, so the system's history says how it got there.
        //
        // Never a false "yes" from a half-written BETA: a publish lands new files in BETA before its
        // record moves, and a file production lacks is a copy. A publish landing after the check
        // moves BETA's content hash past the one recorded here, so the system waits again.
        //
        // `useCache`: the list's minute check trusts a "differs" answer for DifferentRecheck, keyed
        // by both hashes, so it does not rebuild plans for systems genuinely waiting.
        // ###########################################################################################
        internal async Task<bool> RecordIfProductionAlreadyHoldsAsync(
            ReviewAccess? access,
            SystemRecord system,
            ServerOptions options,
            DateTimeOffset nowUtc,
            bool useCache,
            CancellationToken cancellationToken = default)
        {
            if (!options.IsProductionPublishingConfigured || !ProductionPromotionRules.IsAwaitingProduction(system))
                return false;

            string key = $"{system.SystemId}|{system.ContentHash}|{system.ProductionContentHash}";

            if (useCache &&
                this.thisFoundDifferentUtc.TryGetValue(key, out DateTimeOffset checkedUtc) &&
                nowUtc - checkedUtc < ProductionPromotionFlow.DifferentRecheck)
            {
                return false;
            }

            ProductionPromotionResult plan = await this.BuildPlanAsync(system, options, cancellationToken);

            bool holds = plan.CanPromote &&
                plan.Files.Count == 0 &&
                plan.UnchangedCount > 0 &&
                ProductionPromotionFlow.PreviewRemovals(options.DataTreeRoot!, options.ProductionDataTreeRoot!, system).Files.Count == 0 &&
                ProductionPromotionFlow.ListingForProduction(options.DataTreeRoot!, options.ProductionDataTreeRoot!, system.SystemId) is { IsReady: true, Insert: null };

            if (!holds)
            {
                this.thisFoundDifferentUtc[key] = nowUtc;
                return false;
            }

            this.thisFoundDifferentUtc.TryRemove(key, out _);

            await this.thisStore.SetSystemInProductionAsync(
                system.SystemId, system.CurrentRevision, system.ContentHash, nowUtc, cancellationToken);

            await this.thisAccounts.WriteAuditAsync(
                new AuditEntry(
                    access?.Account.Id,
                    access?.Account.Email ?? string.Empty,
                    SystemHistoryEvents.FoundInProduction,
                    system.SystemId,
                    $"revision {system.CurrentRevision}; production already held every file ({plan.UnchangedCount})",
                    nowUtc),
                cancellationToken);

            this.thisLogger.LogInformation(
                "{SystemId} revision {Revision} was already in production (copied there outside CRT) - recorded as published to production.",
                system.SystemId, system.CurrentRevision);

            return true;
        }

        // ###########################################################################################
        // What promoting this system would copy - shown to the maintainer before they press the
        // button, from the same code that performs it.
        // ###########################################################################################
        public async Task<PromotionPlanOutcome> PlanAsync(
            ReviewAccess? access,
            string? systemId,
            ServerOptions options,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(options);

            if (!options.IsProductionPublishingConfigured)
                return PromotionPlanOutcome.NotConfigured();

            if (!ReviewAuthority.CanReviewAnything(access))
                return PromotionPlanOutcome.Forbidden(ReviewAuthority.DescribeRefusal(access, null));

            SystemRecord? system = string.IsNullOrWhiteSpace(systemId)
                ? null
                : await this.thisStore.FindSystemAsync(systemId, cancellationToken);

            if (system is null)
                return PromotionPlanOutcome.NotFound("No such system.");

            if (!ReviewAuthority.CanPublish(access, system.SystemId))
                return PromotionPlanOutcome.Forbidden($"This account is not a maintainer of {system.SystemId}.");

            ProductionPromotionResult plan = await this.BuildPlanAsync(system, options, cancellationToken);
            ApprovalStatus approval = await this.ApprovalStatusAsync(access, system, plan, cancellationToken);

            // What the promotion would REMOVE from production, shown before anyone approves it.
            FileRemovalPreview removals = plan.CanPromote
                ? ProductionPromotionFlow.PreviewRemovals(options.DataTreeRoot!, options.ProductionDataTreeRoot!, system)
                : FileRemovalPreview.Nothing;

            // A system production does not list, and BETA does not list either, is refused here -
            // before anyone presses the button - rather than published to nobody.
            ListingForPublish listing = ProductionPromotionFlow.ListingForProduction(
                options.DataTreeRoot!, options.ProductionDataTreeRoot!, system.SystemId);

            string? refusal = ProductionPromotionFlow.RefusalFor(plan) ?? (listing.IsReady ? null : listing.Problem);

            return PromotionPlanOutcome.Planned(system, plan, refusal, approval) with { Removals = removals };
        }

        // ###########################################################################################
        // Promotes. See the class header for the order of the checks.
        // ###########################################################################################
        public Task<PromotionOutcome> PromoteAsync(
            ReviewAccess? access,
            string? systemId,
            string? expectedBetaContentHash,
            ServerOptions options,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default) =>
            this.PromoteAsync(access, systemId, expectedBetaContentHash, options, nowUtc, shownRemovals: null, cancellationToken);

        // `shownRemovals`: the files the maintainer was shown this would remove from production. The
        // publishing approval is refused when the list differs now - the same rule as a BETA
        // publish (ApprovePublishFlow step 5b).
        public async Task<PromotionOutcome> PromoteAsync(
            ReviewAccess? access,
            string? systemId,
            string? expectedBetaContentHash,
            ServerOptions options,
            DateTimeOffset nowUtc,
            IReadOnlyCollection<string>? shownRemovals,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(options);

            // ---- 1 and 2 ----------------------------------------------------------------------
            if (!options.IsProductionPublishingConfigured)
                return PromotionOutcome.NotConfigured();

            if (!ReviewAuthority.CanReviewAnything(access))
                return PromotionOutcome.Forbidden(ReviewAuthority.DescribeRefusal(access, null));

            if (string.IsNullOrWhiteSpace(systemId))
                return PromotionOutcome.NotFound();

            using IDisposable gate = await this.thisLock.EnterAsync(cancellationToken);

            // ---- 3. The system, re-read under the lock ---------------------------------------
            SystemRecord? system = await this.thisStore.FindSystemAsync(systemId, cancellationToken);

            if (system is null)
                return PromotionOutcome.NotFound();

            if (!ReviewAuthority.CanPublish(access, system.SystemId))
                return PromotionOutcome.Forbidden($"This account is not a maintainer of {system.SystemId}.");

            if (!ProductionPromotionRules.IsAwaitingProduction(system))
                return PromotionOutcome.Conflict("The stable source already has this system as it is in BETA. There is nothing to publish.");

            // ---- 4. What the maintainer checked ------------------------------------------------
            if (!string.Equals(system.ContentHash, expectedBetaContentHash?.Trim(), StringComparison.Ordinal))
            {
                return PromotionOutcome.Conflict(
                    "This system has changed in BETA since you opened it - another contribution has been published " +
                    "there. Check it again in CRT with the BETA data, then publish it to the stable source.");
            }

            // ---- 5. The plan -----------------------------------------------------------------
            ProductionPromotionResult plan = await this.BuildPlanAsync(system, options, cancellationToken);

            string? refusal = ProductionPromotionFlow.RefusalFor(plan);

            if (refusal is not null)
                return PromotionOutcome.Refused(refusal);

            // ---- 5b. Its row in production's drop-down lists (owner request, 2026-09-27) --------
            //
            // Decided before a file is copied, like every other refusal here.
            ListingForPublish listing = ProductionPromotionFlow.ListingForProduction(
                options.DataTreeRoot!, options.ProductionDataTreeRoot!, system.SystemId);

            if (!listing.IsReady)
                return PromotionOutcome.Refused(listing.Problem);

            // ---- 6. Is this the approval that publishes? --------------------------------------
            ApprovalStatus approval = await this.ApprovalStatusAsync(access, system, plan, cancellationToken);

            if (!approval.CanApprove)
                return PromotionOutcome.Conflict(ApprovePublishFlow.WhyNot(approval, "publishing this to the stable source"));

            if (!approval.ApprovalPublishes)
            {
                await this.thisStore.AddProductionApprovalAsync(
                    system.SystemId, system.ContentHash!, approval.YourRole!.Value,
                    access!.Account.Id, ApprovePublishFlow.Label(access), nowUtc, cancellationToken);

                IReadOnlyList<ApproverRole> stillWaiting = approval.WaitingFor
                    .Where(role => role != approval.YourRole)
                    .ToList();

                this.thisLogger.LogInformation(
                    "{Account} approved publishing {SystemId} to production as {Role}; waiting for {Waiting}.",
                    access.Account.Email, system.SystemId, approval.YourRole, string.Join(", ", stillWaiting));

                return PromotionOutcome.AwaitingApproval(system, stillWaiting);
            }

            // ---- 7. Copy, then record ---------------------------------------------------------
            // What this removes from production, and is it what the maintainer was shown?
            FileRemovalPreview removals = ProductionPromotionFlow.PreviewRemovals(
                options.DataTreeRoot!, options.ProductionDataTreeRoot!, system);

            // No list at all is an older maintainer application, not a changed list - see
            // ApprovePublishFlow.RemovalsNotSentMessage.
            if (!removals.Matches(shownRemovals))
            {
                return shownRemovals is null
                    ? PromotionOutcome.Refused(ApprovePublishFlow.RemovalsNotSentMessage)
                    : PromotionOutcome.Conflict(ProductionPromotionFlow.RemovalsChangedMessage);
            }

            PromotionCopyOutcome copy = await ProductionPromoter.CopyAsync(
                plan.Files,
                options.DataTreeRoot!,
                options.ProductionDataTreeRoot!,
                cancellationToken);

            if (!copy.IsDone)
            {
                this.thisLogger.LogError(
                    "Publishing {SystemId} to production stopped after {Copied} file(s): {Error}",
                    system.SystemId, copy.FilesCopied, copy.Error);

                // Folders copied into production by hand: the administrator needs the command.
                if (copy.FoldersRefusing.Count > 0)
                {
                    this.thisLogger.LogError(
                        "Give the service its access to production's folders back with: {Command}",
                        TreeWriteAccess.FixCommand(copy.FoldersRefusing));
                }

                return PromotionOutcome.Refused(copy.Error ?? "The copy did not complete.");
            }

            // ---- 7b. The row in production's main Excel data file --------------------------------
            //
            // After the files, so it never names a board production does not have; before the
            // promotion is recorded, so a failure is retried by publishing again - the copy is
            // verified and the insert updates a row already there.
            if (listing.Insert is MasterRowInsert insert)
            {
                MasterListingEdit edit = MasterListing.Insert(insert.MasterPath, insert.Row, insert.AfterExcelDataFile, nowUtc);

                if (!edit.IsDone)
                {
                    this.thisLogger.LogError(
                        "Publishing {SystemId} to production: the files were copied but the row could not be added to [{Master}]: {Failure}",
                        system.SystemId, insert.MasterPath, edit.Failure);

                    return PromotionOutcome.Refused(
                        $"The files were copied, but the system could not be added to the stable source's drop-down lists: {edit.Failure} " +
                        "Publish it again once that is fixed.");
                }
            }

            // A system.json in production's copy of this board - retired, and possibly carried
            // across by hand from BETA before promotions existed. See RetiredSystemDescriptor.
            if (SubmissionPathRules.TryResolve(
                    options.ProductionDataTreeRoot!,
                    $"{system.Manufacturer}/{system.Hardware}/{system.Board}",
                    out string productionSystemFolder,
                    out _))
            {
                RetiredSystemDescriptor.TryRemove(options.ProductionDataTreeRoot!, productionSystemFolder, this.thisLogger);
            }

            // What the board no longer uses in production - only the files shown, and only those
            // still unused now the new workbooks are in place. Never fails the promotion.
            UnusedFileRemoval removal = UnusedFileRemover.Remove(options.ProductionDataTreeRoot!, removals.Files, this.thisLogger);

            await this.thisStore.AddProductionApprovalAsync(
                system.SystemId, system.ContentHash!, approval.YourRole!.Value,
                access!.Account.Id, ApprovePublishFlow.Label(access), nowUtc, cancellationToken);

            await this.thisStore.SetSystemInProductionAsync(
                system.SystemId, system.CurrentRevision, system.ContentHash, nowUtc, cancellationToken);

            await this.thisAccounts.WriteAuditAsync(
                new AuditEntry(
                    access!.Account.Id,
                    access.Account.Email,
                    ProductionPromotionFlow.PublishedAction,
                    system.SystemId,
                    $"revision {system.CurrentRevision}; {copy.FilesCopied} file(s) copied; {plan.UnchangedCount} unchanged; " +
                    $"{removal.Removed.Count} unused file(s) removed",
                    nowUtc),
                cancellationToken);

            this.thisLogger.LogInformation(
                "{Account} published {SystemId} revision {Revision} to production ({Copied} file(s) copied).",
                access.Account.Email, system.SystemId, system.CurrentRevision, copy.FilesCopied);

            return PromotionOutcome.Published(system, copy.FilesCopied, system.ProductionPublishedUtc) with { RemovedFiles = removal.Removed };
        }

        public const string RemovalsChangedMessage =
            "The files this would remove from the stable source have changed since you opened it - another publish has " +
            "changed what uses them. Refresh, check the list, then publish.";

        // ###########################################################################################
        // THE FILES A PROMOTION WOULD REMOVE FROM PRODUCTION.
        //
        // After the promotion the system's workbooks in production are BETA's, so the candidates
        // are what production's workbooks cite that BETA's do not, and production is then read
        // with BETA's citations standing in for its own. A file another production board still
        // cites is kept; so is anything an older generation cites. A workbook that cannot be read
        // in either tree blocks the removal rather than guessing.
        // ###########################################################################################
        internal static FileRemovalPreview PreviewRemovals(string betaRoot, string productionRoot, SystemRecord system)
        {
            string folder = $"{system.Manufacturer}/{system.Hardware}/{system.Board}";

            if (!ProductionPromotionFlow.TryReadCitations(productionRoot, folder, out Dictionary<string, IReadOnlyCollection<string>> before, out string? why) ||
                !ProductionPromotionFlow.TryReadCitations(betaRoot, folder, out Dictionary<string, IReadOnlyCollection<string>> after, out why))
            {
                return new FileRemovalPreview([], "Nothing is removed, because " + why);
            }

            // Only inside the system's own folder (owner decision, 2026-09-27) - see
            // AutomaticRemovalScope.
            IReadOnlyList<string> candidates = AutomaticRemovalScope.Within(
                folder,
                DataTreeUsage.NoLongerCited(
                    before.Values.SelectMany(files => files),
                    after.Values.SelectMany(files => files)));

            return FileRemovalPreview.Compute(productionRoot, candidates, after, ApprovePublishFlow.PreviewReads);
        }

        // Each board workbook at the top of the system's folder -> what it cites.
        private static bool TryReadCitations(
            string root,
            string folder,
            out Dictionary<string, IReadOnlyCollection<string>> citations,
            out string? why)
        {
            citations = new Dictionary<string, IReadOnlyCollection<string>>(StringComparer.OrdinalIgnoreCase);
            why = null;

            if (!SubmissionPathRules.TryResolve(root, folder, out string full, out _) || !Directory.Exists(full))
                return true;

            try
            {
                foreach (string file in Directory.EnumerateFiles(full, "*" + DataGenerationRules.WorkbookExtension))
                {
                    string relative = $"{folder}/{Path.GetFileName(file)}";

                    if (!DataTreeUsage.IsBoardFolderWorkbook(relative))
                        continue;

                    // Through the preview cache: these citations only choose what to SHOW; the
                    // removal itself re-reads the tree (UnusedFileRemover).
                    if (!ApprovePublishFlow.PreviewReads.TryGetCitations(file, out IReadOnlyCollection<string> cited))
                    {
                        why = $"the board workbook [{relative}] could not be read.";
                        return false;
                    }

                    citations[relative] = cited.ToList();
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                why = $"the folder [{folder}] could not be read.";
                return false;
            }

            return true;
        }

        public const string PublishedAction = "production.published";

        // ###########################################################################################
        // WHAT PRODUCTION'S DROP-DOWN LISTS NEED FROM THIS PROMOTION (owner decision, 2026-09-27:
        // "insert at the same place"): nothing when production lists the system already; its BETA
        // row, after the same neighbours, when only BETA does (MasterListing.TryResolvePlacement);
        // and a refusal when neither does - it has not been placed yet.
        //
        // Either file missing or unreadable changes nothing and refuses nothing: promotions worked
        // without these files before, and must not start failing on them.
        // ###########################################################################################
        internal static ListingForPublish ListingForProduction(string betaRoot, string productionRoot, string systemId)
        {
            string? betaMaster = MasterListing.NewestMasterPath(betaRoot);
            string? productionMaster = MasterListing.NewestMasterPath(productionRoot);

            if (betaMaster is null || productionMaster is null ||
                !MasterListing.TryRead(betaMaster, out IReadOnlyList<MasterListingRow> betaRows, out _) ||
                !MasterListing.TryRead(productionMaster, out IReadOnlyList<MasterListingRow> productionRows, out _))
            {
                return ListingForPublish.NothingToAdd;
            }

            if (MasterListing.IndexOfSystem(productionRows, systemId) >= 0)
                return ListingForPublish.NothingToAdd;

            int inBeta = MasterListing.IndexOfSystem(betaRows, systemId);

            if (inBeta < 0)
            {
                return ListingForPublish.Refused(
                    "This system is not in the drop-down lists in BETA or in the stable source, so nobody would see it. " +
                    "Place it in the Systems screen first.");
            }

            // Production can list another system under these names that BETA does not.
            if (MasterListing.NamesTakenBy(productionRows, systemId, betaRows[inBeta].HardwareName, betaRows[inBeta].BoardName) is MasterListingRow taken)
                return ListingForPublish.Refused(MasterListing.NamesTakenMessage(taken));

            MasterListing.TryResolvePlacement(betaRows, systemId, productionRows, out string? after);

            return ListingForPublish.Add(new MasterRowInsert(productionMaster, betaRows[inBeta], after));
        }

        // ###########################################################################################
        // Why this plan cannot be carried out at all, or null when it can.
        // ###########################################################################################
        private static string? RefusalFor(ProductionPromotionResult plan) =>
            plan.CanPromote ? null : string.Join(" ", plan.Problems.Select(problem => problem.Message));

        // ###########################################################################################
        // Where publishing this system to production stands for this account: the same rule as a
        // submission (ApprovalRules), with "changes a shared file" read off the PLAN and the
        // approvals given against the BETA content hash the plan was made from.
        // ###########################################################################################
        private async Task<ApprovalStatus> ApprovalStatusAsync(
            ReviewAccess? access,
            SystemRecord system,
            ProductionPromotionResult plan,
            CancellationToken cancellationToken)
        {
            bool hasMaintainers = plan.TouchesSharedFiles &&
                (await this.thisAccounts.GetMaintainersOfSystemAsync(system.SystemId, cancellationToken))
                    .Any(ReviewAuthority.CanGiveMaintainerApproval);

            IReadOnlyList<GivenApproval> given = string.IsNullOrWhiteSpace(system.ContentHash)
                ? []
                : await this.thisStore.GetProductionApprovalsAsync(system.SystemId, system.ContentHash, cancellationToken);

            return ApprovalRules.Status(
                ApprovalRules.Required(plan.TouchesSharedFiles, hasMaintainers),
                given,
                ReviewAuthority.RoleIn(access, system.SystemId),
                access?.Account.Id);
        }

        // ###########################################################################################
        // The plan, from the two trees as they are now. The BETA board is read for what it CITES,
        // which is how its shared files are found; the walk of its folder is what it OWNS.
        // ###########################################################################################
        private async Task<ProductionPromotionResult> BuildPlanAsync(
            SystemRecord system,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            string betaRoot = options.DataTreeRoot ?? string.Empty;
            string productionRoot = options.ProductionDataTreeRoot ?? string.Empty;

            PublishedTreeView? beta = PublishedTreeProbe.For(betaRoot);
            PublishedTreeView? production = PublishedTreeProbe.For(productionRoot);

            if (beta is null || production is null)
            {
                return new ProductionPromotionResult(
                    [],
                    0,
                    [new ValidationFinding
                    {
                        Severity = ValidationSeverity.Error,
                        Code = "promote.tree_unavailable",
                        Subject = system.SystemId,
                        Message = "The BETA or stable data tree could not be read on the server."
                    }],
                    TouchesSharedFiles: false);
            }

            var identity = new SubmissionManifest
            {
                SystemId = system.SystemId,
                Manufacturer = system.Manufacturer,
                Hardware = system.Hardware,
                Board = system.Board
            };

            BoardData? board = await this.thisBoards.TryReadAsync(betaRoot, identity, cancellationToken);

            IReadOnlyList<string> cited = board is null
                ? []
                : SubmissionManifestBuilder.CollectReferencedFiles(board);

            return ProductionPromotionPlan.Build(
                system.Manufacturer,
                system.Hardware,
                system.Board,
                ProductionPromotionFlow.WalkSystemFolder(betaRoot, system),
                cited,
                beta,
                production);
        }

        // Every file under the system's BETA folder, data-root-relative with forward slashes.
        // Empty when the folder does not resolve or is not there - which the plan refuses.
        internal static IReadOnlyList<string> WalkSystemFolder(string betaRoot, SystemRecord system)
        {
            string relative = $"{system.Manufacturer}/{system.Hardware}/{system.Board}";

            if (!SubmissionPathRules.TryResolve(betaRoot, relative, out string folder, out _) || !Directory.Exists(folder))
                return [];

            string fullRoot = Path.GetFullPath(betaRoot);

            try
            {
                return Directory
                    .EnumerateFiles(folder, "*", SearchOption.AllDirectories)
                    .Select(path => Path.GetRelativePath(fullRoot, path).Replace(Path.DirectorySeparatorChar, '/'))
                    .ToList();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return [];
            }
        }
    }

    public sealed record PromotionPlanOutcome(
        SystemRecord? System,
        ProductionPromotionResult? Plan,
        string? Refusal,
        bool IsNotConfigured = false,
        bool IsForbidden = false,
        bool IsNotFound = false,
        ApprovalStatus? Approval = null)
    {
        public static PromotionPlanOutcome Planned(SystemRecord system, ProductionPromotionResult plan, string? refusal, ApprovalStatus approval) =>
            new(system, plan, refusal, Approval: approval);

        public static PromotionPlanOutcome NotConfigured() =>
            new(null, null, ProductionPromotionFlow.NotConfiguredMessage, IsNotConfigured: true);

        public static PromotionPlanOutcome Forbidden(string why) => new(null, null, why, IsForbidden: true);

        public static PromotionPlanOutcome NotFound(string why) => new(null, null, why, IsNotFound: true);

        // What promoting would remove from production - see ProductionPromotionFlow.PreviewRemovals.
        public FileRemovalPreview Removals { get; init; } = FileRemovalPreview.Nothing;
    }

    public sealed record PromotionOutcome(
        bool IsPublished,
        string Error,
        SystemRecord? System = null,
        int FilesCopied = 0,
        DateTimeOffset? PreviousProductionPublishedUtc = null,
        bool IsNotConfigured = false,
        bool IsForbidden = false,
        bool IsNotFound = false,
        bool IsConflict = false)
    {
        public static PromotionOutcome Published(SystemRecord system, int filesCopied, DateTimeOffset? previous) =>
            new(true, string.Empty, system, filesCopied, previous);

        public static PromotionOutcome NotConfigured() =>
            new(false, ProductionPromotionFlow.NotConfiguredMessage, IsNotConfigured: true);

        public static PromotionOutcome Forbidden(string why) => new(false, why, IsForbidden: true);

        public static PromotionOutcome NotFound() => new(false, "No such system.", IsNotFound: true);

        public static PromotionOutcome Conflict(string why) => new(false, why, IsConflict: true);

        public static PromotionOutcome Refused(string why) => new(false, why);

        // The first of two approvals: recorded against this BETA state, nothing copied.
        public static PromotionOutcome AwaitingApproval(SystemRecord system, IReadOnlyList<ApproverRole> waitingFor) =>
            new(false, string.Empty, system) { WaitingFor = waitingFor };

        public IReadOnlyList<ApproverRole> WaitingFor { get; init; } = [];

        public bool IsAwaitingApproval => this.WaitingFor.Count > 0;

        // The files the promotion removed from production because nothing there used them any more.
        public IReadOnlyList<string> RemovedFiles { get; init; } = [];
    }
}
