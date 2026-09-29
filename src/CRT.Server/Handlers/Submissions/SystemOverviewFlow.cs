using System.Data.Common;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // THE "SYSTEMS" SCREEN (owner request, 2026-09-27): every system, and one system's facts - who
    // maintains it, who has contributed to it and how that went, and its recent submissions.
    //
    // *** ANY MAINTAINER SEES EVERY SYSTEM, CONTRIBUTORS INCLUDED (owner decision, 2026-09-27:
    // "Everything for everyone"). *** The rule is CanReviewAnything - the same "may open the
    // Maintainer tab at all" the queue asks - and NOT CanReview(system): a maintainer of the
    // C64 sees who contributes to the Amiga. That widens who sees a contributor's contact address,
    // which is why it was asked rather than assumed; it is recorded in NewContributeStrategy.md.
    //
    // WHICH SYSTEMS: the `systems` rows unioned with the boards in the BETA tree - the list the
    // administrator's Maintainers screen uses (MaintainerAssignmentFlows.ListSystemsAsync), for the
    // same reason: a shipped board nobody has submitted to has no row, and is still a system.
    //
    // The shape SubmissionFlows and ContributorHistory keep: the async methods only fetch, and every
    // rule - what a system's state is, who counts as one contributor, which state word a submission
    // shows - is a pure static below, tested without a database.
    // ###########################################################################################
    public static class SystemOverviewFlow
    {
        // The most submissions one system's screen lists - the newest. Contributors are counted over
        // up to StoreLimit, which is every submission any board here has had.
        public const int ListedSubmissions = 50;
        public const int StoreLimit = 1000;

        public static async Task<SystemOverviewOutcome> ListAsync(
            ReviewAccess? access,
            IReadOnlyList<PublishedSystemLister.KnownSystem> inBeta,
            IReadOnlyList<PublishedSystemLister.KnownSystem>? inProduction,
            ISubmissionStore submissions,
            IAccountStore accounts,
            CancellationToken cancellationToken = default,
            IBoardViewStore? boardViews = null,
            DateTimeOffset? now = null)
        {
            ArgumentNullException.ThrowIfNull(inBeta);
            ArgumentNullException.ThrowIfNull(submissions);
            ArgumentNullException.ThrowIfNull(accounts);

            if (!ReviewAuthority.CanReviewAnything(access))
                return SystemOverviewOutcome.Forbidden();

            IReadOnlyList<SystemRecord> rows = await submissions.ListSystemsAsync(cancellationToken);
            IReadOnlyList<MaintainerRecord> maintainers = await accounts.ListMaintainersAsync(cancellationToken);

            // How often each board was looked at in CRT (2026-09-27) - "views in 30 days" on its line.
            IReadOnlyDictionary<string, int>? views = await SystemOverviewFlow.ReadViewsAsync(
                () => boardViews!.CountSinceBySystemAsync(
                    BoardViewStatisticsRules.WindowStart(now ?? DateTimeOffset.UtcNow, 30), cancellationToken),
                boardViews is not null);

            return SystemOverviewOutcome.Listed(
                SystemOverviewFlow.WithViews(SystemOverviewFlow.Entries(rows, inBeta, inProduction, maintainers), views));
        }

        // ###########################################################################################
        // Each entry with its view count, or as it was when there are no counts (no store, or it
        // could not be read). A system nobody looked at reads 0 - it IS counted, and is not.
        // ###########################################################################################
        public static IReadOnlyList<SystemOverviewEntry> WithViews(
            IReadOnlyList<SystemOverviewEntry> entries,
            IReadOnlyDictionary<string, int>? viewsLast30Days)
        {
            ArgumentNullException.ThrowIfNull(entries);

            return viewsLast30Days is null
                ? entries
                : entries.Select(entry => entry with { ViewsLast30Days = viewsLast30Days.GetValueOrDefault(entry.SystemId) }).ToList();
        }

        // ###########################################################################################
        // *** THE VIEW COUNTS NEVER FAIL THE SCREEN. *** They are statistics beside the facts a
        // maintainer opens this screen for; a database error reading them leaves them out (null,
        // which the Maintainer tab shows as nothing at all) rather than answering a 500 for
        // the whole list.
        // ###########################################################################################
        private static async Task<T?> ReadViewsAsync<T>(Func<Task<T>> read, bool canRead)
            where T : class
        {
            if (!canRead)
                return null;

            try
            {
                return await read();
            }
            catch (DbException)
            {
                return null;
            }
        }

        public static async Task<SystemOverviewOutcome> DetailAsync(
            ReviewAccess? access,
            string? systemId,
            IReadOnlyList<PublishedSystemLister.KnownSystem> inBeta,
            IReadOnlyList<PublishedSystemLister.KnownSystem>? inProduction,
            ISubmissionStore submissions,
            IAccountStore accounts,
            CancellationToken cancellationToken = default,
            DateTimeOffset? now = null,
            IBoardViewStore? boardViews = null)
        {
            ArgumentNullException.ThrowIfNull(inBeta);
            ArgumentNullException.ThrowIfNull(submissions);
            ArgumentNullException.ThrowIfNull(accounts);

            if (!ReviewAuthority.CanReviewAnything(access))
                return SystemOverviewOutcome.Forbidden();

            if (string.IsNullOrWhiteSpace(systemId))
                return SystemOverviewOutcome.NotFound();

            SystemRecord? row = await submissions.FindSystemAsync(systemId, cancellationToken);
            PublishedSystemLister.KnownSystem? known = inBeta.FirstOrDefault(system => string.Equals(system.SystemId, systemId, StringComparison.Ordinal));

            // A system is a row or a board in the tree - neither is no system at all.
            if (row is null && known is null)
                return SystemOverviewOutcome.NotFound();

            List<MaintainerRecord> pool = (await accounts.ListMaintainersAsync(cancellationToken))
                .Where(maintainer => string.Equals(maintainer.SystemId, systemId, StringComparison.Ordinal))
                .ToList();

            IReadOnlyList<SystemSubmissionRecord> sent =
                await submissions.GetSubmissionsForSystemAsync(systemId, SystemOverviewFlow.StoreLimit, cancellationToken);

            // The display names of the contributors who were signed in - a handful of lookups.
            var named = new Dictionary<long, AccountRecord>();

            // ...and of the maintainers who decided them, whom the history names.
            foreach (long accountId in sent.Select(record => record.Submission.AccountId)
                         .Concat(sent.Select(record => record.DecidedByAccountId))
                         .OfType<long>()
                         .Distinct())
            {
                if (await accounts.FindByIdAsync(accountId, cancellationToken) is AccountRecord account)
                    named[accountId] = account;
            }

            // How often CRT users look at it (2026-09-27), from the last year's views.
            DateTimeOffset at = now ?? DateTimeOffset.UtcNow;

            IReadOnlyList<BoardViewFact>? facts = await SystemOverviewFlow.ReadViewsAsync(
                () => boardViews!.FactsForSystemAsync(systemId, BoardViewStatisticsRules.WindowStart(at, 365), cancellationToken),
                boardViews is not null);

            BoardViewStatistics? views = facts is null ? null : BoardViewStatisticsRules.Build(facts, at);

            SystemOverviewEntry entry = SystemOverviewFlow.Entry(
                systemId, row, known, SystemOverviewFlow.Holds(inProduction, systemId), pool.Count) with
            {
                ViewsLast30Days = views?.Last30Days
            };

            // The invitations nobody has accepted yet (2026-09-27) - for the ADMINISTRATOR, the one
            // person who can invite or withdraw one. Null for everybody else, who has no use for them.
            List<MaintainerInvitationEntry>? invitations = null;

            if (ReviewAuthority.CanAdminister(access))
            {
                invitations = (await accounts.ListOpenInvitationsAsync(at, cancellationToken))
                    .Where(invitation => string.Equals(invitation.SystemId, systemId, StringComparison.Ordinal))
                    .Select(invitation => new MaintainerInvitationEntry(invitation.Id, invitation.Email, invitation.CreatedUtc, invitation.ExpiresUtc))
                    .ToList();
            }

            // What has happened to it (2026-09-27): its submissions' own dates, and the audit rows
            // naming the system or one of those submissions.
            IReadOnlyList<AuditEntry> audit = await accounts.GetAuditForSubjectsAsync(
                [systemId, .. sent.Select(record => SystemHistoryRules.SubmissionSubject(record.Submission.Id))],
                SystemHistoryRules.Limit,
                cancellationToken);

            // Whose contributor discarded their own draft since sending it (2026-09-28).
            IReadOnlyDictionary<long, DateTimeOffset> discarded = await submissions.GetDraftDiscardsAsync(
                sent.Select(record => record.Submission.Id).ToList(), cancellationToken);

            // Which a BETA rollback returned to the queue (migration 0015) - "returned", not "pending".
            IReadOnlyDictionary<long, DateTimeOffset> returns = await submissions.GetBetaReturnsAsync(
                sent.Select(record => record.Submission.Id).ToList(), cancellationToken);

            return SystemOverviewOutcome.Described(new SystemDetailAnswer(
                entry,
                pool.Select(maintainer => new PoolMaintainerEntry(maintainer.AccountId, maintainer.DisplayName, maintainer.Email)).ToList(),
                SystemOverviewFlow.Contributors(sent, named),
                SystemOverviewFlow.Submissions(sent, row?.ProductionPublishedUtc, named, discarded, returns),
                invitations,
                SystemHistoryRules.Build(sent, named, audit),
                views));
        }

        // ###########################################################################################
        // The list: every row and every board in the BETA tree, once each, ordered by id - the
        // order the administrator's Maintainers screen uses, so the two lists read alike.
        // ###########################################################################################
        public static IReadOnlyList<SystemOverviewEntry> Entries(
            IReadOnlyList<SystemRecord> rows,
            IReadOnlyList<PublishedSystemLister.KnownSystem> inBeta,
            IReadOnlyList<PublishedSystemLister.KnownSystem>? inProduction,
            IReadOnlyList<MaintainerRecord> maintainers)
        {
            ArgumentNullException.ThrowIfNull(rows);
            ArgumentNullException.ThrowIfNull(inBeta);
            ArgumentNullException.ThrowIfNull(maintainers);

            Dictionary<string, SystemRecord> byRow = rows
                .GroupBy(row => row.SystemId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

            Dictionary<string, PublishedSystemLister.KnownSystem> byTree = inBeta
                .GroupBy(system => system.SystemId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

            Dictionary<string, int> poolSizes = maintainers
                .GroupBy(maintainer => maintainer.SystemId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

            return byRow.Keys
                .Union(byTree.Keys, StringComparer.Ordinal)
                .OrderBy(id => id, StringComparer.Ordinal)
                .Select(id => SystemOverviewFlow.Entry(
                    id,
                    byRow.GetValueOrDefault(id),
                    byTree.GetValueOrDefault(id),
                    SystemOverviewFlow.Holds(inProduction, id),
                    poolSizes.GetValueOrDefault(id)))
                .ToList();
        }

        // ###########################################################################################
        // One system's facts. The row names it where there is one (it is the truth once anything has
        // been submitted); the tree names a shipped board nobody has touched.
        // ###########################################################################################
        public static SystemOverviewEntry Entry(
            string systemId,
            SystemRecord? row,
            PublishedSystemLister.KnownSystem? inBeta,
            bool? inProduction,
            int maintainerCount)
        {
            return new SystemOverviewEntry(
                systemId,
                row?.Manufacturer ?? inBeta?.Manufacturer ?? string.Empty,
                row?.Hardware ?? inBeta?.Hardware ?? string.Empty,
                row?.Board ?? inBeta?.Board ?? string.Empty,
                InBeta: inBeta is not null,
                InProduction: inProduction,
                IsAwaitingProduction: ProductionPromotionRules.IsAwaitingProduction(row),

                // No row yet means nothing has closed it - the default the column has.
                IsAccepting: row?.IsAccepting ?? true,
                BetaRevision: row?.CurrentRevision,
                ProductionRevision: row?.ProductionRevision,
                ProductionPublishedUtc: row?.ProductionPublishedUtc,
                MaintainerCount: maintainerCount);
        }

        // ###########################################################################################
        // Who has contributed to this system, and how it went - newest contributor first.
        //
        // One contributor is SubmissionReplacementRules' person: the same account, or the same
        // contact email (trimmed, any case) among submissions sent without one. So one address used
        // both signed in and anonymously is TWO contributors - deliberately, as in ContributorHistory:
        // an anonymous sender's email is unverified, and merging would let anybody add submissions
        // to an account holder's record (code review, 2026-09-27). The counts follow
        // ContributorHistory: Accepted is merged (in BETA or production), a rejection counts only
        // when a maintainer made it, and a submission the contributor's own newer one replaced is not
        // counted - it is an earlier copy of the same work. Someone whose only submissions were
        // replaced ones is still LISTED: they did contribute.
        // ###########################################################################################
        public static IReadOnlyList<SystemContributorEntry> Contributors(
            IEnumerable<SystemSubmissionRecord> sent,
            IReadOnlyDictionary<long, AccountRecord>? accounts = null)
        {
            ArgumentNullException.ThrowIfNull(sent);

            return sent
                .GroupBy(record => SystemOverviewFlow.ContributorKey(record.Submission), StringComparer.Ordinal)
                .Where(group => group.Key.Length > 0)
                .Select(group =>
                {
                    SubmissionRecord first = group.First().Submission;

                    AccountRecord? account = first.AccountId is long id && accounts is not null
                        ? accounts.GetValueOrDefault(id)
                        : null;

                    string? email = account?.Email ?? group
                        .Select(record => record.Submission.ContactEmail?.Trim())
                        .FirstOrDefault(address => !string.IsNullOrWhiteSpace(address));

                    string? name = account is not null && !string.IsNullOrWhiteSpace(account.DisplayName)
                        ? account.DisplayName.Trim()
                        : null;

                    int Count(Func<SystemSubmissionRecord, bool> which) => group.Count(which);

                    return new SystemContributorEntry(
                        Email: string.IsNullOrWhiteSpace(email) ? null : email,
                        Name: name,
                        Accepted: Count(record => record.Submission.State == SubmissionState.Merged),
                        Waiting: Count(record => record.Submission.State is SubmissionState.Pending or SubmissionState.Approved),
                        ChangesRequested: Count(record => record.Submission.State == SubmissionState.ChangesRequested),
                        Rejected: Count(record => record.Submission.State == SubmissionState.Rejected && record.DecidedByMaintainer),
                        LastSubmittedUtc: group.Max(record => record.Submission.CreatedUtc));
                })
                .OrderByDescending(contributor => contributor.LastSubmittedUtc)
                .ThenBy(contributor => contributor.Email, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ###########################################################################################
        // The newest submissions, each with the word its CONTRIBUTOR is told
        // (ProductionPromotionRules.ContributorFacingState) - so the maintainer and the contributor
        // describe one submission the same way: "merged" is in BETA, "published" has reached
        // production, "returned" was pushed back out of BETA.
        //
        // A signed-in contributor's submission stores no contact email; their account's address is
        // shown instead, as the contributor list does.
        // ###########################################################################################
        //
        // `discarded` (2026-09-28): when each contributor discarded their own draft since sending.
        public static IReadOnlyList<SystemSubmissionEntry> Submissions(
            IEnumerable<SystemSubmissionRecord> sent,
            DateTimeOffset? systemProductionPublishedUtc,
            IReadOnlyDictionary<long, AccountRecord>? accounts = null,
            IReadOnlyDictionary<long, DateTimeOffset>? discarded = null,
            IReadOnlyDictionary<long, DateTimeOffset>? returns = null)
        {
            ArgumentNullException.ThrowIfNull(sent);

            return sent
                .Select(record => record.Submission)
                .OrderByDescending(submission => submission.Id)
                .Take(SystemOverviewFlow.ListedSubmissions)
                .Select(submission => new SystemSubmissionEntry(
                    submission.Id,
                    submission.AccountId is long id && accounts?.GetValueOrDefault(id) is AccountRecord account
                        ? account.Email
                        : submission.ContactEmail?.Trim(),
                    submission.Summary,
                    ProductionPromotionRules.ContributorFacingState(
                        submission.State, submission.DecidedUtc, systemProductionPublishedUtc,
                        returns is not null && returns.TryGetValue(submission.Id, out DateTimeOffset returned) ? returned : null),
                    submission.CreatedUtc,
                    submission.DecidedUtc,
                    submission.DecisionComment,
                    discarded is not null && discarded.TryGetValue(submission.Id, out DateTimeOffset at) ? at : null))
                .ToList();
        }

        // "a:12" for a signed-in contributor, "e:dennis@example.com" otherwise; empty for a submission
        // naming nobody, which cannot be grouped with anyone and is left out of the contributor list.
        private static string ContributorKey(SubmissionRecord submission)
        {
            if (submission.AccountId is long accountId)
                return $"a:{accountId}";

            string email = submission.ContactEmail?.Trim() ?? string.Empty;

            return email.Length == 0 ? string.Empty : $"e:{email.ToLowerInvariant()}";
        }

        // Whether the production tree holds the board - null when there is no production tree to ask.
        private static bool? Holds(IReadOnlyList<PublishedSystemLister.KnownSystem>? tree, string systemId) =>
            tree is null ? null : tree.Any(system => string.Equals(system.SystemId, systemId, StringComparison.Ordinal));
    }

    public sealed record SystemOverviewOutcome(
        IReadOnlyList<SystemOverviewEntry>? Systems,
        SystemDetailAnswer? Detail,
        bool IsForbidden = false,
        bool IsNotFound = false)
    {
        public static SystemOverviewOutcome Listed(IReadOnlyList<SystemOverviewEntry> systems) => new(systems, null);

        public static SystemOverviewOutcome Described(SystemDetailAnswer detail) => new(null, detail);

        public static SystemOverviewOutcome Forbidden() => new(null, null, IsForbidden: true);

        public static SystemOverviewOutcome NotFound() => new(null, null, IsNotFound: true);
    }
}
