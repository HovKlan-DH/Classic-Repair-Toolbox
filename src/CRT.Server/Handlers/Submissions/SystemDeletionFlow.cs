using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // DELETING A SYSTEM COMPLETELY (owner request, 2026-10-03): "as admin, I should be able to have
    // a possibility to delete a system completely, which then will remove it from everywhere -
    // including BETA and stable sources ... part of the 'Admin' menu ... with a confirmation box".
    // Prompted by test systems left in the Systems list as "Not published yet" with nothing behind
    // them, but it deletes any system the rules below let go.
    //
    // ADMINISTRATOR ONLY, and like BetaRollbackFlow a plan first and the act second, refusals
    // cheapest first:
    //
    //   1. CONFIGURED - production publishing switched on? "Everywhere" includes the stable data,
    //                   which this service may only write once it is (ServerOptions).
    //   2. AUTHORITY  - an administrator.
    //   3. EXISTENCE  - a database record, or anything of it in either tree.
    //   4. PLAN       - SystemDeletionFiles.Survey of both trees, and the submissions; refused when
    //                   it is blocked (an older main Excel data file lists it, another board uses a
    //                   file in its folder, a link, an unreadable tree).
    //   5. UNDER THE PUBLISH LOCK: the plan again, held to the FINGERPRINT the administrator was
    //      shown; every folder it writes probed; then the drop-down rows (stable first), the files
    //      (stable, then BETA), both checksum manifests, and - only once both trees are clean - the
    //      database record, the board views and the audit row. The contributors of the OPEN
    //      submissions are mailed last, outside the lock.
    //
    // *** THE ORDER IS THE CRASH STORY. *** Rows before files: a board whose files are going is never
    // offered without them - a row gone while its files stay is merely unlisted. The database record
    // LAST: while a file is left, the record, the submissions and the mails stay, so pressing Delete
    // again (which plans again) finishes the job and mails once. Every step is safe to repeat.
    //
    // *** WHAT IS DELETED WITH THE RECORD *** is the schema's ON DELETE CASCADE - ISubmissionStore.
    // DeleteSystemAsync: every submission whatever its state, the pool, invitations, approvals. The
    // AUDIT rows stay (nothing deletes audit), and this delete writes one more. The board views go
    // too (owner decision, 2026-10-03: "Delete them"), so the Fun facts page stops counting it.
    //
    // *** OPEN SUBMISSIONS ARE DELETED AND THEIR CONTRIBUTORS MAILED (owner decision, 2026-10-03,
    // asked: "Delete them and mail the contributors"). *** SystemDeletionRules.IsOpen says which; the
    // reason is then required, as it is the mail's only content (the BETA rollback's rule). CRT's
    // "My submissions" shows such a submission as "No longer on the server" - the server answers
    // "not found" - which says it is gone but not why, so the mail is how the contributor learns.
    //
    // *** NOT DELETED: SHARED FILES. *** Anything outside the system's own folder stays, and is an
    // unused file for Account > Unused files when nothing else cites it (AutomaticRemovalScope's owner
    // decision). Nor older main Excel data files, which are frozen - a system one lists is refused.
    // ###########################################################################################
    public sealed class SystemDeletionFlow
    {
        public const string DeletedAction = SystemHistoryEvents.Deleted;

        private readonly ISubmissionStore thisStore;
        private readonly IAccountStore thisAccounts;
        private readonly IBoardViewStore thisBoardViews;
        private readonly SubmissionNotifier thisNotifier;
        private readonly PublishLock thisLock;
        private readonly BlobStore thisBlobs;
        private readonly ILogger<SystemDeletionFlow> thisLogger;

        public SystemDeletionFlow(
            ISubmissionStore store,
            IAccountStore accounts,
            IBoardViewStore boardViews,
            SubmissionNotifier notifier,
            PublishLock publishLock,
            BlobStore blobs,
            ILogger<SystemDeletionFlow> logger)
        {
            this.thisStore = store;
            this.thisAccounts = accounts;
            this.thisBoardViews = boardViews;
            this.thisNotifier = notifier;
            this.thisLock = publishLock;
            this.thisBlobs = blobs;
            this.thisLogger = logger;
        }

        // ###########################################################################################
        // What a delete WOULD remove - shown in the confirmation, from the code that then removes it.
        // A blocked plan is still a plan: it says why, and the Account screen shows that instead of a
        // confirmation.
        // ###########################################################################################
        public async Task<SystemDeletionOutcome> PlanAsync(
            ReviewAccess? access,
            string? systemId,
            ServerOptions options,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default,
            Func<string, bool>? isLink = null)
        {
            if (SystemDeletionFlow.Check(access, systemId, options) is SystemDeletionOutcome refusal)
                return refusal;

            string id = systemId!.Trim();
            SystemDeletionPlan? plan = await this.BuildPlanAsync(id, options, nowUtc, isLink, cancellationToken);

            return plan is null
                ? SystemDeletionOutcome.NotFound(SystemDeletionRules.NoSuchSystemMessage(id))
                : SystemDeletionOutcome.Planned(plan);
        }

        // ###########################################################################################
        // Performs it. `fingerprint` is the plan the administrator confirmed; `reason` is quoted to
        // the contributors of the open submissions. `canWriteFolder`, `isLink` and `removeListing`
        // are for tests.
        // ###########################################################################################
        public async Task<SystemDeletionOutcome> DeleteAsync(
            ReviewAccess? access,
            string? systemId,
            string? fingerprint,
            string? reason,
            ServerOptions options,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default,
            Func<string, bool>? canWriteFolder = null,
            Func<string, bool>? isLink = null,
            Func<SystemTreeSurvey, MasterListingEdit>? removeListing = null)
        {
            if (SystemDeletionFlow.Check(access, systemId, options) is SystemDeletionOutcome refusal)
                return refusal;

            string id = systemId!.Trim();
            string because = (reason ?? string.Empty).Trim();

            SystemDeletionPlan plan;
            SystemDeletionOutcome done;

            using (await this.thisLock.EnterAsync(cancellationToken))
            {
                // Planned again INSIDE the lock: nothing may publish, promote or push back between
                // the plan being checked and the trees being written.
                SystemDeletionPlan? current = await this.BuildPlanAsync(id, options, nowUtc, isLink, cancellationToken);

                if (current is null)
                    return SystemDeletionOutcome.NotFound(SystemDeletionRules.NoSuchSystemMessage(id));

                if (current.BlockedBecause is string blocked)
                    return SystemDeletionOutcome.Conflict(blocked);

                if (!string.Equals(current.Fingerprint, fingerprint?.Trim(), StringComparison.Ordinal))
                    return SystemDeletionOutcome.Conflict(SystemDeletionRules.ChangedMessage);

                if (current.Open.Count > 0 && because.Length == 0)
                    return SystemDeletionOutcome.Refused(SystemDeletionRules.ReasonRequiredMessage);

                plan = current;
                done = await this.ApplyAsync(access!, plan, because, options, nowUtc, canWriteFolder, isLink, removeListing, cancellationToken);
            }

            if (!done.IsDone || plan.Open.Count == 0)
                return done;

            // ###########################################################################################
            // *** AFTER THE FACT, AND IT MAY NOT FAIL THE DELETE. *** Everything is gone; a mail that
            // cannot be sent is logged by the notifier. Outside the lock - a slow mailer must not
            // hold up every publish. Not cancellable: the submissions are deleted, so a request given
            // up now would leave their contributors told nothing, with no way to tell them later.
            // ###########################################################################################
            int mailed = 0;

            try
            {
                mailed = await this.thisNotifier.NotifySystemDeletedAsync(
                    plan.Open.Select(open => open.Recipient), plan.SystemId, because, CancellationToken.None);
            }
            catch (Exception ex)
            {
                this.thisLogger.LogWarning(ex, "{SystemId} was deleted, but its contributors could not be told.", plan.SystemId);
            }

            return done with { ContributorsMailed = mailed };
        }

        // Under the lock, with a plan that is not blocked and matches what was confirmed.
        private async Task<SystemDeletionOutcome> ApplyAsync(
            ReviewAccess access,
            SystemDeletionPlan plan,
            string reason,
            ServerOptions options,
            DateTimeOffset nowUtc,
            Func<string, bool>? canWriteFolder,
            Func<string, bool>? isLink,
            Func<SystemTreeSurvey, MasterListingEdit>? removeListing,
            CancellationToken cancellationToken)
        {
            // ---- 1. Every folder it writes, in both trees, before anything moves ---------------
            foreach (SystemTreeSurvey survey in new[] { plan.Production, plan.Beta })
            {
                IReadOnlyList<string> refusing = SystemDeletionFiles.FoldersRefusing(survey, canWriteFolder);

                if (refusing.Count == 0)
                    continue;

                this.thisLogger.LogError(
                    "Deleting {SystemId} refused before changing anything: the service may not write into {Folders}. "
                    + "Files copied in by hand as another user do this. Give the service its access back with: {Command}",
                    plan.SystemId,
                    string.Join(", ", refusing),
                    TreeWriteAccess.FixCommand(refusing));

                return SystemDeletionOutcome.Refused(TreeWriteAccess.RefusalMessage(survey.TreeName, survey.Root!, refusing));
            }

            // ---- 2. The drop-down rows: stable first, then BETA --------------------------------
            bool rowRemoved = false;

            foreach (SystemTreeSurvey survey in new[] { plan.Production, plan.Beta })
            {
                MasterListingEdit edit = removeListing?.Invoke(survey) ?? SystemDeletionFiles.RemoveListing(survey, plan.SystemId);

                if (!edit.IsDone)
                {
                    this.thisLogger.LogError(
                        "Deleting {SystemId}: its row could not be taken out of the {Tree} data's main Excel data file: {Failure}",
                        plan.SystemId, survey.TreeName, edit.Failure);

                    // A main Excel data file already rewritten has a new checksum, and a manifest
                    // still naming the old one makes every CRT that downloads it throw it away
                    // (code review, 2026-10-04) - so the manifests follow even on this way out.
                    if (rowRemoved)
                        await this.RegenerateManifestsAsync(options, plan.SystemId);

                    return SystemDeletionOutcome.Refused(
                        SystemDeletionRules.ListingNotRemovedMessage(survey.TreeName, edit.Failure ?? string.Empty, rowRemoved));
                }

                rowRemoved |= edit.Changed;
            }

            // ---- 3. The files: stable, then BETA -----------------------------------------------
            SystemTreeRemoval production = SystemDeletionFiles.RemoveFiles(plan.Production, this.thisLogger, isLink);
            SystemTreeRemoval beta = SystemDeletionFiles.RemoveFiles(plan.Beta, this.thisLogger, isLink);

            // ---- 4. Both manifests describe the trees as they now are - whatever happened in 3 --
            await this.RegenerateManifestsAsync(options, plan.SystemId);

            int left = production.Left + beta.Left;

            if (left > 0)
                return SystemDeletionOutcome.Refused(SystemDeletionRules.FilesLeftMessage(left));

            // ---- 5. The record, and everything the schema hangs off it --------------------------
            IReadOnlyList<long> deletedSubmissions = [];

            if (plan.HasRecord)
            {
                try
                {
                    deletedSubmissions = await this.thisStore.DeleteSystemAsync(plan.SystemId, cancellationToken);
                }
                catch (Exception ex) when (ex is not OperationCanceledException)
                {
                    this.thisLogger.LogError(ex, "{SystemId}'s files were deleted, but its database record could not be.", plan.SystemId);
                    return SystemDeletionOutcome.Refused(SystemDeletionRules.NotRecordedMessage);
                }
            }

            // ###########################################################################################
            // *** FROM HERE ON THE DELETE IS DONE, AND NOTHING BELOW MAY UNDO THAT ANSWER (code review,
            // 2026-10-04). *** The record is gone, so a retry answers "There is no system": a failure
            // or a cancelled request past this point would leave the deletion unaudited and the open
            // submissions' contributors never mailed. So no token, and every step only logs.
            // ###########################################################################################

            // Its submissions' partial uploads - every submission the cascade took, uploading ones
            // included, which the plan does not list. The abandoned-upload sweeper finds partials
            // through the database rows just deleted, so nothing else would ever clear them.
            foreach (long submissionId in deletedSubmissions)
                this.thisBlobs.ClearPartials(submissionId);

            // The views are statistics: a failure is logged, never the delete's.
            try
            {
                await this.thisBoardViews.DeleteForSystemAsync(plan.SystemId, CancellationToken.None);
            }
            catch (Exception ex)
            {
                this.thisLogger.LogWarning(ex, "{SystemId} was deleted, but its board views could not be.", plan.SystemId);
            }

            try
            {
                await this.thisAccounts.WriteAuditAsync(
                    new AuditEntry(
                        access.Account.Id,
                        access.Account.Email,
                        SystemDeletionFlow.DeletedAction,
                        plan.SystemId,
                        $"{beta.Removed} file(s) removed from BETA, {production.Removed} from the stable data; " +
                        $"{plan.Submissions.Count} submission(s) deleted, {plan.Open.Count} of them open" +
                        (reason.Length > 0 ? $"; {reason}" : string.Empty),
                        nowUtc),
                    CancellationToken.None);
            }
            catch (Exception ex)
            {
                this.thisLogger.LogError(ex, "{SystemId} was deleted, but the deletion could not be audited.", plan.SystemId);
            }

            this.thisLogger.LogInformation(
                "{Account} deleted {SystemId}: {Beta} file(s) from BETA, {Production} from the stable data, {Submissions} submission(s).",
                access.Account.Email, plan.SystemId, beta.Removed, production.Removed, plan.Submissions.Count);

            return SystemDeletionOutcome.Deleted(plan, beta.Removed, production.Removed);
        }

        private static SystemDeletionOutcome? Check(ReviewAccess? access, string? systemId, ServerOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);

            if (!options.IsProductionPublishingConfigured)
                return SystemDeletionOutcome.NotConfigured(ProductionPromotionFlow.NotConfiguredMessage);

            if (!ReviewAuthority.CanAdminister(access))
                return SystemDeletionOutcome.Forbidden("Only an administrator can delete a system.");

            if (string.IsNullOrWhiteSpace(systemId))
                return SystemDeletionOutcome.NotFound(SystemDeletionRules.NoSystemNamedMessage);

            if (!SystemDescriptorRules.IsValidSystemId(systemId.Trim()))
                return SystemDeletionOutcome.NotFound(SystemDeletionRules.NoSuchSystemMessage(systemId.Trim()));

            return null;
        }

        // ###########################################################################################
        // Both trees, the record and its submissions. Null when there is nothing of it anywhere - a
        // tree that could not be surveyed is NOT "nothing", it is a blocked plan that says why.
        // ###########################################################################################
        private async Task<SystemDeletionPlan?> BuildPlanAsync(
            string systemId,
            ServerOptions options,
            DateTimeOffset nowUtc,
            Func<string, bool>? isLink,
            CancellationToken cancellationToken)
        {
            SystemRecord? record = await this.thisStore.FindSystemAsync(systemId, cancellationToken);

            // Workbooks are read for the "another board" question - seconds on a real tree.
            (SystemTreeSurvey beta, SystemTreeSurvey production) = await Task.Run(
                () => (
                    SystemDeletionFiles.Survey(SystemDeletionRules.BetaTreeName, options.DataTreeRoot, systemId, isLink),
                    SystemDeletionFiles.Survey(SystemDeletionRules.ProductionTreeName, options.ProductionDataTreeRoot, systemId, isLink)),
                cancellationToken);

            bool anything = record is not null ||
                beta.HoldsAnything || production.HoldsAnything ||
                beta.BlockedBecause is not null || production.BlockedBecause is not null;

            if (!anything)
                return null;

            IReadOnlyList<SystemSubmissionRecord> submissions = record is null
                ? []
                : await this.thisStore.GetSubmissionsForSystemAsync(systemId, SystemOverviewFlow.StoreLimit, cancellationToken);

            IReadOnlyDictionary<long, DateTimeOffset> returns = submissions.Count == 0
                ? new Dictionary<long, DateTimeOffset>()
                : await this.thisStore.GetBetaReturnsAsync(submissions.Select(item => item.Submission.Id).ToList(), cancellationToken);

            List<(SubmissionRecord Submission, string State)> open = submissions
                .Select(item => (
                    item.Submission,
                    State: ProductionPromotionRules.ContributorFacingState(
                        item.Submission.State,
                        item.Submission.DecidedUtc,
                        record?.ProductionPublishedUtc,
                        returns.TryGetValue(item.Submission.Id, out DateTimeOffset returned) ? returned : null)))
                .Where(item => SystemDeletionRules.IsOpen(item.State))
                .OrderBy(item => item.Submission.Id)
                .ToList();

            IReadOnlyDictionary<long, MailRecipient> recipients = await ContributorAddresses.ResolveRecipientsAsync(
                this.thisAccounts, open.Select(item => item.Submission), cancellationToken);

            int maintainers = record is null ? 0 : (await this.thisAccounts.GetMaintainersOfSystemAsync(systemId, cancellationToken)).Count;

            int invitations = record is null
                ? 0
                : (await this.thisAccounts.ListOpenInvitationsAsync(nowUtc, cancellationToken))
                    .Count(invitation => string.Equals(invitation.SystemId, systemId, StringComparison.Ordinal));

            string[] parts = systemId.Split('/');

            return new SystemDeletionPlan(
                systemId,
                record?.Manufacturer ?? parts[0],
                record?.Hardware ?? parts[1],
                record?.Board ?? parts[2],
                beta,
                production,
                record is not null,
                submissions,
                open.Select(item => new SystemDeletionOpenSubmission(
                    item.Submission,
                    item.State,
                    recipients.TryGetValue(item.Submission.Id, out MailRecipient? recipient) ? recipient : new MailRecipient(string.Empty))).ToList(),
                maintainers,
                invitations,
                SystemDeletionRules.Fingerprint(
                    beta.Files,
                    production.Files,
                    beta.IsListed,
                    production.IsListed,
                    record is not null,
                    submissions.Select(item => (item.Submission.Id, item.Submission.State))));
        }

        // ###########################################################################################
        // Both manifests, off the request thread - each hashes every file of its tree (~11,000), and
        // this runs under the publish lock (code review, 2026-10-04; SystemOrderFlow and
        // ManifestRebuildFlow do the same). Not cancellable: once a tree has changed, a request
        // given up must still leave its manifest describing it.
        // ###########################################################################################
        private Task RegenerateManifestsAsync(ServerOptions options, string systemId) =>
            Task.Run(() =>
            {
                this.RegenerateManifest(options.DataTreeRoot, options.PublicDataBaseUrl, options.ManifestPath, SystemDeletionRules.BetaTreeName, systemId);
                this.RegenerateManifest(options.ProductionDataTreeRoot, options.ProductionPublicDataBaseUrl, options.ProductionManifestPath, SystemDeletionRules.ProductionTreeName, systemId);
            }, CancellationToken.None);

        // Never fails the delete: a failed rebuild leaves the previous manifest, and is logged.
        private void RegenerateManifest(string? root, string? publicUrl, string? manifestPath, string treeName, string systemId)
        {
            int written;

            try
            {
                written = DataChecksumManifest.Write(root ?? string.Empty, publicUrl ?? string.Empty, manifestPath ?? string.Empty);
            }
            catch (Exception ex)
            {
                this.thisLogger.LogWarning(ex, "The {Tree} checksum manifest could not be rebuilt after deleting {SystemId}.", treeName, systemId);
                return;
            }

            if (written < 0)
            {
                this.thisLogger.LogWarning(
                    "{SystemId} was deleted from the {Tree} data but its checksum manifest at [{ManifestPath}] could not be " +
                    "rebuilt - CRT keeps offering its files until it is (" + MaintainerScreenWording.Account + " > Rebuild checksum manifests).",
                    systemId, treeName, manifestPath);
            }
        }
    }

    // ###########################################################################################
    // What a delete would remove. Open: the submissions still in play, with the contributor-facing
    // state and who is mailed. BlockedBecause: the first tree's refusal, stable or BETA.
    // ###########################################################################################
    public sealed record SystemDeletionPlan(
        string SystemId,
        string Manufacturer,
        string Hardware,
        string Board,
        SystemTreeSurvey Beta,
        SystemTreeSurvey Production,
        bool HasRecord,
        IReadOnlyList<SystemSubmissionRecord> Submissions,
        IReadOnlyList<SystemDeletionOpenSubmission> Open,
        int Maintainers,
        int Invitations,
        string Fingerprint)
    {
        public string? BlockedBecause => this.Beta.BlockedBecause ?? this.Production.BlockedBecause;

        public SystemDeletePlanAnswer ToAnswer() =>
            new(
                this.SystemId,
                this.Manufacturer,
                this.Hardware,
                this.Board,
                this.Fingerprint,
                this.Beta.Files.Count,
                this.Production.Files.Count,
                this.Beta.IsListed,
                this.Production.IsListed,
                this.HasRecord,
                this.Submissions.Count,
                this.Maintainers,
                this.Invitations,
                [.. this.Open.Select(open => new SystemDeleteOpenSubmission(
                    open.Submission.Id,
                    open.State,
                    open.Recipient.Email,
                    open.Submission.Summary,
                    open.Submission.CreatedUtc))],
                this.BlockedBecause);
    }

    public sealed record SystemDeletionOpenSubmission(SubmissionRecord Submission, string State, MailRecipient Recipient);

    // ###########################################################################################
    // What a plan or a delete came to. Shaped like BetaRollbackOutcome, so the endpoint maps it onto
    // status codes the same way.
    // ###########################################################################################
    public sealed record SystemDeletionOutcome(
        SystemDeletionPlan? Plan,
        string? Error = null,
        bool IsNotConfigured = false,
        bool IsForbidden = false,
        bool IsNotFound = false,
        bool IsConflict = false,
        bool IsDone = false,
        int BetaFilesRemoved = 0,
        int ProductionFilesRemoved = 0,
        int ContributorsMailed = 0)
    {
        public bool IsPlanned => this.Plan is not null && this.Error is null;

        public static SystemDeletionOutcome Planned(SystemDeletionPlan plan) => new(plan);

        public static SystemDeletionOutcome Deleted(SystemDeletionPlan plan, int betaFilesRemoved, int productionFilesRemoved) =>
            new(plan, IsDone: true, BetaFilesRemoved: betaFilesRemoved, ProductionFilesRemoved: productionFilesRemoved);

        public static SystemDeletionOutcome NotConfigured(string error) => new(null, error, IsNotConfigured: true);

        public static SystemDeletionOutcome Forbidden(string error) => new(null, error, IsForbidden: true);

        public static SystemDeletionOutcome NotFound(string error) => new(null, error, IsNotFound: true);

        public static SystemDeletionOutcome Conflict(string error) => new(null, error, IsConflict: true);

        public static SystemDeletionOutcome Refused(string error) => new(null, error);
    }
}
