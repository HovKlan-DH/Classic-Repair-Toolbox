using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // ROLLING A BETA BOARD BACK, AND RETURNING ITS SUBMISSIONS TO THE QUEUE (owner decision,
    // 2026-09-27): "I have no possibility to 'push back to queue', which I think I should be able
    // to, so I can inform contributor if something is missing."
    //
    // The mirror of ProductionPromotionFlow, and deliberately shaped like it - the same lock, the
    // same authority, refusals cheapest first:
    //
    //   1. CONFIGURED - production publishing switched on? (a restore reads production's tree)
    //   2. AUTHORITY  - may this account decide this system?
    //   3. EXISTENCE  - is there such a system, and is its BETA state actually ahead?
    //   4. PLAN       - what would be written, and what returns to the queue?
    //   5. WRITE the tree, then the BETA checksum manifest, then the bookkeeping.
    //
    // *** WHY THIS IS POSSIBLE - it was first judged impossible and that was wrong. *** The
    // reasoning was that a publish is irreversible and the old bytes are collected. The second half
    // is false, and the project owner said so: SubmissionCollectionStates.Live contains Merged, so
    // DeleteRetiredPayloadsAsync never touches a merged submission - "everything stays shadowed ...
    // data is still there". Production also holds a COMPLETE board, so overwriting BETA from it is
    // a defined restore rather than a guess. BetaRollbackPlan's header records both.
    //
    // *** PER SYSTEM, NEVER PER SUBMISSION - and that shapes the whole feature. ***
    // ProductionPromotionPlan says it in the other direction: "two submissions merged into one
    // board cannot be promoted separately, because the board's workbook already holds both". So a
    // rollback reverts EVERY submission merged since the last promotion, and all of them go back to
    // `pending` with the maintainer's comment (owner decision). Silently discarding the other
    // contributors' accepted work would be the worst version of this.
    //
    // *** NOT AN APPROVAL-GATED OPERATION, deliberately. *** Publishing needs two approvals for a
    // shared-file change because it pushes data OUT to everyone. This only ever puts back bytes
    // production already serves - reviewed and public - so the dangerous direction is the one it
    // undoes. That holds for the shared files it touches too: they go back to production's bytes,
    // and only when BETA still holds exactly what a returning submission wrote (BetaRollbackPlan).
    // It is still audited, and still refused to anyone without authority over the board.
    //
    // *** "REJECT" IS THE SAME ROLLBACK (owner request, 2026-09-28: "a direct 'Reject' button also
    // - just like the normal queue. Then there is no need to push it back and then reject it"). ***
    // The tree moves exactly as for a push-back; only what the submissions become differs -
    // `rejected` with the comment, rather than `pending` - and the contributor gets the ordinary
    // rejection mail. Being the same code is the point: a second way of taking a board out of BETA
    // would be a second set of shared-file and manifest rules to keep right.
    // ###########################################################################################
    public sealed class BetaRollbackFlow
    {
        public const string RolledBackAction = SystemHistoryEvents.PushedBack;

        public const string RejectedAction = SystemHistoryEvents.RejectedFromBeta;

        public const string NothingToRollBackMessage =
            "This system's BETA data is the same as the stable source's, so there is nothing to roll back.";

        // A promoted system whose production folder lists nothing (BetaRollbackPlan's header).
        public const string ProductionUnreadableMessage =
            "This system was published to the stable source, but its folder there cannot be read - it is missing, renamed, " +
            "or the stable data folder is not reachable. Nothing was changed: pushing back now would take the " +
            "board out of BETA as if it had never been in the stable source. Check the stable data folder, then try again.";

        public const string NotRecordedMessage =
            "BETA WAS rolled back, but recording it failed, so its submissions are not back in the queue yet and " +
            "nobody has been told. Push back again to finish - the files are already in place, so nothing moves twice.";

        private readonly ISubmissionStore thisStore;
        private readonly IAccountStore thisAccounts;
        private readonly PublishLock thisLock;
        private readonly ILogger<BetaRollbackFlow> thisLogger;

        public BetaRollbackFlow(
            ISubmissionStore store,
            IAccountStore accounts,
            PublishLock publishLock,
            ILogger<BetaRollbackFlow> logger)
        {
            this.thisStore = store;
            this.thisAccounts = accounts;
            this.thisLock = publishLock;
            this.thisLogger = logger;
        }

        // ###########################################################################################
        // What a rollback WOULD do - shown to the maintainer before they press anything, from the
        // same code that performs it. Its hashes are cached per file version (BetaRollbackFiles),
        // so the pass under the lock that follows reads again only what changed in between.
        // ###########################################################################################
        public async Task<BetaRollbackOutcome> PlanAsync(
            ReviewAccess? access,
            string? systemId,
            ServerOptions options,
            CancellationToken cancellationToken = default)
        {
            (BetaRollbackOutcome? refusal, SystemRecord? system) =
                await this.CheckAsync(access, systemId, options, cancellationToken);

            if (refusal is not null)
                return refusal;

            BetaRollbackFilePlan plan = await this.BuildPlanAsync(system!, options, cancellationToken);

            if (plan.Plan.Kind == BetaRollbackKind.ProductionUnreadable)
                return BetaRollbackOutcome.Conflict(BetaRollbackFlow.ProductionUnreadableMessage);

            return BetaRollbackOutcome.Planned(system!, plan.Plan);
        }

        // ###########################################################################################
        // Performs it, under the publish lock - so no approval can land in BETA between the plan
        // being built and the tree being written. `reject` rejects the submissions it takes out
        // instead of returning them to the queue - see the class header.
        // ###########################################################################################
        public async Task<BetaRollbackOutcome> RollBackAsync(
            ReviewAccess? access,
            string? systemId,
            string? comment,
            ServerOptions options,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default,
            bool reject = false)
        {
            (BetaRollbackOutcome? refusal, SystemRecord? system) =
                await this.CheckAsync(access, systemId, options, cancellationToken);

            if (refusal is not null)
                return refusal;

            // ###########################################################################################
            // *** THE COMMENT IS REQUIRED. *** It is the contributor's ONLY feedback - contributing
            // needs no account, so there is no inbox and no thread (SubmissionRecord.DecisionComment
            // says so). A rollback with no reason tells someone their accepted work was taken out of
            // BETA and says nothing about why, which is worse than not offering the button.
            // ###########################################################################################
            string reason = (comment ?? string.Empty).Trim();

            if (reason.Length == 0)
            {
                return BetaRollbackOutcome.Refused(reject
                    ? "Say why this is being rejected - it is the only message the contributor receives."
                    : "Say why this is being rolled back - it is the only message the contributor receives.");
            }

            using IDisposable held = await this.thisLock.EnterAsync(cancellationToken);

            // Re-read inside the lock: the system may have been promoted or published meanwhile.
            SystemRecord? current = await this.thisStore.FindSystemAsync(system!.SystemId, cancellationToken);

            if (current is null)
                return BetaRollbackOutcome.NotFound($"There is no system [{system.SystemId}].");

            if (!ProductionPromotionRules.IsAwaitingProduction(current))
                return BetaRollbackOutcome.Conflict(BetaRollbackFlow.NothingToRollBackMessage);

            BetaRollbackFilePlan filePlan = await this.BuildPlanAsync(current, options, cancellationToken);
            BetaRollbackPlanResult plan = filePlan.Plan;

            // Checked again under the lock: the folder may have gone since the plan was shown.
            if (plan.Kind == BetaRollbackKind.ProductionUnreadable)
                return BetaRollbackOutcome.Conflict(BetaRollbackFlow.ProductionUnreadableMessage);

            BetaRollbackWriteOutcome written = await BetaRollbackWriter.ApplyAsync(
                filePlan,
                options.DataTreeRoot!,
                options.ProductionDataTreeRoot!,
                current,
                this.thisLogger,
                cancellationToken);

            // ###########################################################################################
            // *** THE BETA MANIFEST IS REBUILT THE MOMENT THE TREE HAS MOVED (code review,
            // 2026-09-27). *** Every other writer of the BETA tree does this - a publish, the unused
            // files screen - and the first version of the rollback did not. A CRT already holding the
            // rolled-back bytes then saw its hash match the stale manifest and never downloaded
            // production's again; a fresh one downloaded the restored file, failed the checksum and
            // refused it; and a removed file stayed advertised and 404'd on every sync.
            //
            // HERE rather than in the endpoint, and BEFORE the bookkeeping - also after a refusal
            // part-way, since some files may already have moved: the manifest has to describe the
            // tree whatever happens next.
            // ###########################################################################################
            // A system that was never in production leaves BETA entirely - and with it, its row in
            // BETA's drop-down lists (2026-09-27), or every BETA user would be offered a board that
            // is not there. Only when the files went; its stored placement is kept, so the next
            // publish puts it back in the same place. Never fails the rollback - the tree has moved.
            if (written.IsDone && plan.Kind == BetaRollbackKind.RemoveFromBeta)
                this.RemoveFromBetaList(options, current.SystemId);

            this.RegenerateBetaManifest(options, current.SystemId);

            if (!written.IsDone)
                return BetaRollbackOutcome.Refused(written.Error ?? "The rollback did not complete.");

            // ###########################################################################################
            // *** ONE TRANSACTION, CAUGHT (code review, 2026-09-27). *** The submissions back to
            // pending, their approvals cleared, the system's BETA state following the tree - all or
            // nothing. A failure is answered as what it is rather than escaping as a 500 that reads
            // "it did not happen": the tree HAS been rolled back, and pushing back again finds it
            // already level with production, changes no file, and records it.
            //
            // AFTER the tree, deliberately: a submission moved back to `pending` while its data is
            // still in BETA would put a board in the queue that is already live.
            //
            // A restored system is level with production; a removed one has no BETA state at all.
            // ###########################################################################################
            try
            {
                await this.thisStore.RecordRollbackAsync(
                    current.SystemId,
                    plan.Returning.Select(submission => submission.Id).ToList(),
                    access!.Account.Id,
                    reason,
                    plan.Kind == BetaRollbackKind.RestoreFromProduction ? current.ProductionRevision : null,
                    plan.Kind == BetaRollbackKind.RestoreFromProduction ? current.ProductionContentHash : null,
                    nowUtc,
                    reject,
                    cancellationToken);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                this.thisLogger.LogError(
                    ex,
                    "{SystemId} WAS rolled back in BETA, but recording it failed - its submissions are still merged.",
                    current.SystemId);

                return BetaRollbackOutcome.Refused(BetaRollbackFlow.NotRecordedMessage);
            }

            await this.thisAccounts.WriteAuditAsync(
                new AuditEntry(
                    access!.Account.Id,
                    access.Account.Email,
                    reject ? BetaRollbackFlow.RejectedAction : BetaRollbackFlow.RolledBackAction,
                    current.SystemId,
                    $"{plan.Kind}; {written.Restored} file(s) restored, {written.Removed} removed; " +
                    $"{plan.Returning.Count} submission(s) {(reject ? "rejected" : "returned to the queue")}; {reason}",
                    nowUtc),
                cancellationToken);

            this.thisLogger.LogInformation(
                "{Account} rolled {SystemId} back from BETA ({Kind}): {Restored} restored, {Removed} removed, " +
                "{Returned} submission(s) {What}.",
                access.Account.Email, current.SystemId, plan.Kind, written.Restored, written.Removed, plan.Returning.Count,
                reject ? "rejected" : "returned to the queue");

            return BetaRollbackOutcome.RolledBack(current, plan, written, reject);
        }

        // The checks both entry points share, cheapest first.
        private async Task<(BetaRollbackOutcome? Refusal, SystemRecord? System)> CheckAsync(
            ReviewAccess? access,
            string? systemId,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(options);

            if (!options.IsProductionPublishingConfigured)
                return (BetaRollbackOutcome.NotConfigured(ProductionPromotionFlow.NotConfiguredMessage), null);

            if (string.IsNullOrWhiteSpace(systemId))
                return (BetaRollbackOutcome.NotFound("No system was named."), null);

            SystemRecord? system = await this.thisStore.FindSystemAsync(systemId, cancellationToken);

            if (system is null)
                return (BetaRollbackOutcome.NotFound($"There is no system [{systemId}]."), null);

            // The same authority a promotion needs: an administrator, or a maintainer of this board.
            if (!ReviewAuthority.CanPublish(access, system.SystemId))
                return (BetaRollbackOutcome.Forbidden("You do not maintain this system."), null);

            if (!ProductionPromotionRules.IsAwaitingProduction(system))
                return (BetaRollbackOutcome.Conflict(BetaRollbackFlow.NothingToRollBackMessage), null);

            return (null, system);
        }

        private async Task<BetaRollbackFilePlan> BuildPlanAsync(
            SystemRecord system,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            IReadOnlyList<SubmissionRecord> merged = await this.thisStore.GetMergedSubmissionsAsync(
                system.SystemId,
                system.ProductionPublishedUtc,
                DateTimeOffset.UtcNow,
                cancellationToken);

            // Every file each returning submission carried - where the shared files it changed, and
            // the exact bytes it put there, come from.
            var files = new List<SubmissionFileRecord>();

            foreach (SubmissionRecord submission in merged)
                files.AddRange(await this.thisStore.GetFilesAsync(submission.Id, cancellationToken));

            // Each contributor's address - the account's when they were signed in (2026-09-29): the
            // confirmation names them, and the push-back mail is sent to them.
            IReadOnlyDictionary<long, string> addresses =
                await ContributorAddresses.ResolveAsync(this.thisAccounts, merged, cancellationToken);

            return await BetaRollbackFiles.PlanAsync(
                options.DataTreeRoot!,
                options.ProductionDataTreeRoot!,
                system,
                ProductionPromotionRules.Carrying(merged, addresses: addresses),
                files,
                cancellationToken);
        }

        private void RemoveFromBetaList(ServerOptions options, string systemId)
        {
            string? master = MasterListing.NewestMasterPath(options.DataTreeRoot);

            if (master is null)
                return;

            MasterListingEdit edit = MasterListing.Remove(master, systemId);

            if (!edit.IsDone)
            {
                this.thisLogger.LogError(
                    "{SystemId} was removed from BETA, but its row could not be taken out of [{Master}]: {Failure} - " +
                    "BETA users are offered a board that is not there until it is.",
                    systemId, master, edit.Failure);
            }
            else if (edit.Changed)
            {
                this.thisLogger.LogInformation("{SystemId} was taken out of BETA's drop-down lists.", systemId);
            }
        }

        // Never fails the rollback: a failed rebuild leaves the previous manifest, and is logged.
        private void RegenerateBetaManifest(ServerOptions options, string systemId)
        {
            int written;

            try
            {
                written = DataChecksumManifest.Write(
                    options.DataTreeRoot ?? string.Empty,
                    options.PublicDataBaseUrl ?? string.Empty,
                    options.ManifestPath ?? string.Empty);
            }
            catch (Exception ex)
            {
                this.thisLogger.LogWarning(ex, "The BETA checksum manifest could not be rebuilt after rolling back {SystemId}.", systemId);
                return;
            }

            if (written < 0)
            {
                this.thisLogger.LogWarning(
                    "{SystemId} was rolled back in BETA but the checksum manifest at [{ManifestPath}] could not be " +
                    "rebuilt - clients will not see the rollback until it is.",
                    systemId, options.ManifestPath);
            }
        }
    }

    // ###########################################################################################
    // What a rollback would do, or did. Shaped like PromotionOutcome so ProductionEndpoints maps
    // it onto status codes the same way.
    // ###########################################################################################
    public sealed record BetaRollbackOutcome(
        SystemRecord? System,
        BetaRollbackPlanResult? Plan,
        string? Error = null,
        bool IsNotConfigured = false,
        bool IsForbidden = false,
        bool IsNotFound = false,
        bool IsConflict = false,
        bool IsDone = false,
        int FilesRestored = 0,
        int FilesRemoved = 0,
        bool Rejected = false)
    {
        public bool IsPlanned => this.Plan is not null && this.Error is null;

        public static BetaRollbackOutcome Planned(SystemRecord system, BetaRollbackPlanResult plan) =>
            new(system, plan);

        public static BetaRollbackOutcome RolledBack(
            SystemRecord system,
            BetaRollbackPlanResult plan,
            BetaRollbackWriteOutcome written,
            bool rejected = false) =>
            new(system, plan, IsDone: true, FilesRestored: written.Restored, FilesRemoved: written.Removed, Rejected: rejected);

        public static BetaRollbackOutcome NotConfigured(string error) => new(null, null, error, IsNotConfigured: true);

        public static BetaRollbackOutcome Forbidden(string error) => new(null, null, error, IsForbidden: true);

        public static BetaRollbackOutcome NotFound(string error) => new(null, null, error, IsNotFound: true);

        public static BetaRollbackOutcome Conflict(string error) => new(null, null, error, IsConflict: true);

        public static BetaRollbackOutcome Refused(string error) => new(null, null, error);
    }
}
