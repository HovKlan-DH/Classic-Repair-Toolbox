using System.Data.Common;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // THE "BOARDS" SCREEN (owner request, 2026-09-27): every board, and one board's facts - who
    // maintains it, who has contributed to it and how that went, and its recent submissions.
    //
    // *** ANY MAINTAINER SEES EVERY BOARD, CONTRIBUTORS INCLUDED (owner decision, 2026-09-27:
    // "Everything for everyone"). *** The rule is CanReviewAnything - the same "may open the
    // Maintainer tab at all" the queue asks - and NOT CanReview(board): a maintainer of the
    // C64 sees who contributes to the Amiga.
    //
    // *** BUT ONLY ITS OWN MAINTAINERS SEE ITS EMAIL ADDRESSES (owner request, 2026-10-05: "I do not
    // think that normal maintainer should be able to see other email addresses if they are not set
    // as maintainer for that system. They should be able to see all mail addresses for their own
    // system(s)"). *** ReviewAuthority.CanSeeAddressesOf - the administrator and the board's pool.
    // Anybody else gets the same detail with names and no address anywhere in it: its maintainers'
    // (an empty Email), its contributors' and submissions' (null), and its history's (the name on the
    // account where there is one) - and AddressesHidden, so the tab can say why. Both decisions are
    // in CLAUDE.md ("Maintainer" and the server's rules). The first is also in the last version of
    // NewContributeStrategy.md in git history (deleted 2026-10-05), which still says "contributors'
    // addresses included" - written before the second narrowed it.
    //
    // WHICH BOARDS: the `boards` rows unioned with the boards in the BETA tree - the list the
    // administrator's Maintainers screen uses (MaintainerAssignmentFlows.ListBoardsAsync), for the
    // same reason: a shipped board nobody has submitted to has no row, and is still a board - and,
    // since 2026-10-04, anything the STABLE source holds or lists that BETA does not (owner request:
    // "I do not expect there should be cases where something can only be listed in stable? If so, it
    // must be flagged in the left-sided menu 'Boards' list"). Such a board was missing from the
    // screen altogether; it is now on it, and the Maintainer tab marks it as off.
    //
    // THE DROP-DOWN LISTS (2026-10-04): each entry says whether BETA's and the stable source's newest
    // main Excel data file list it (ListedInBeta / ListedInStable, BoardListings) - read by the
    // endpoint, null where a list could not be read.
    //
    // The shape SubmissionFlows and ContributorHistory keep: the async methods only fetch, and every
    // rule - what a board's state is, who counts as one contributor, which state word a submission
    // shows - is a pure static below, tested without a database.
    // ###########################################################################################
    public static class BoardOverviewFlow
    {
        // The most submissions one board's screen lists - the newest. Its History view shows the
        // board's WHOLE history (owner request, 2026-10-04), one card per submission, so this is a
        // bound no board comes near rather than a page size. Contributors are counted over up to
        // StoreLimit, which is every submission any board here has had.
        public const int ListedSubmissions = 500;
        public const int StoreLimit = 1000;

        public static async Task<BoardOverviewOutcome> ListAsync(
            ReviewAccess? access,
            IReadOnlyList<PublishedBoardLister.KnownBoard> inBeta,
            IReadOnlyList<PublishedBoardLister.KnownBoard>? inProduction,
            ISubmissionStore submissions,
            IAccountStore accounts,
            CancellationToken cancellationToken = default,
            IBoardViewStore? boardViews = null,
            DateTimeOffset? now = null,
            BoardListings? listings = null)
        {
            ArgumentNullException.ThrowIfNull(inBeta);
            ArgumentNullException.ThrowIfNull(submissions);
            ArgumentNullException.ThrowIfNull(accounts);

            if (!ReviewAuthority.CanReviewAnything(access))
                return BoardOverviewOutcome.Forbidden();

            IReadOnlyList<BoardRecord> rows = await submissions.ListBoardsAsync(cancellationToken);
            IReadOnlyList<MaintainerRecord> maintainers = await accounts.ListMaintainersAsync(cancellationToken);

            // How often each board was looked at in CRT (2026-09-27) - "views in 30 days" on its line.
            IReadOnlyDictionary<string, int>? views = await BoardOverviewFlow.ReadViewsAsync(
                () => boardViews!.CountSinceByBoardAsync(
                    BoardViewStatisticsRules.WindowStart(now ?? DateTimeOffset.UtcNow, 30), cancellationToken),
                boardViews is not null);

            // How each board's submissions went (2026-10-09) - "19 submissions in total; ...".
            IReadOnlyList<SubmissionStateCount> states = await submissions.GetSubmissionStateCountsAsync(cancellationToken);

            return BoardOverviewOutcome.Listed(
                BoardOverviewFlow.WithSubmissionCounts(
                    BoardOverviewFlow.WithViews(BoardOverviewFlow.Entries(rows, inBeta, inProduction, maintainers, listings), views),
                    states));
        }

        // ###########################################################################################
        // Each entry with how its submissions went - every board counted, one with none at 0.
        // ###########################################################################################
        public static IReadOnlyList<BoardOverviewEntry> WithSubmissionCounts(
            IReadOnlyList<BoardOverviewEntry> entries,
            IEnumerable<SubmissionStateCount> states)
        {
            ArgumentNullException.ThrowIfNull(entries);
            ArgumentNullException.ThrowIfNull(states);

            Dictionary<string, List<SubmissionStateCount>> byBoard = states
                .GroupBy(state => state.BoardId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);

            return entries
                .Select(entry => entry with
                {
                    SubmissionCounts = BoardOverviewFlow.SubmissionCounts(byBoard.GetValueOrDefault(entry.BoardId) ?? [])
                })
                .ToList();
        }

        // ###########################################################################################
        // How one board's submissions went (BoardSubmissionCounts): each in the word its contributor
        // is told - merged is in BETA, or in stable once a publish to stable came at or after its
        // decision - waiting for a decision or a change, or rejected by a maintainer. A rejection the
        // server made on its own is no submission anybody reviewed, and is not counted, as on the
        // Contributor view.
        // ###########################################################################################
        public static BoardSubmissionCounts SubmissionCounts(IEnumerable<SubmissionStateCount> states)
        {
            ArgumentNullException.ThrowIfNull(states);

            int waiting = 0, rejected = 0, inBeta = 0, inStable = 0;

            foreach (SubmissionStateCount state in states)
            {
                int count = Math.Max(0, state.Count);

                switch (state.State)
                {
                    case SubmissionState.Merged when state.InStable:
                        inStable += count;
                        break;

                    case SubmissionState.Merged:
                        inBeta += count;
                        break;

                    case SubmissionState.Rejected when state.DecidedByMaintainer:
                        rejected += count;
                        break;

                    case SubmissionState.Pending or SubmissionState.Approved or SubmissionState.ChangesRequested:
                        waiting += count;
                        break;
                }
            }

            return new BoardSubmissionCounts(waiting + rejected + inBeta + inStable, waiting, rejected, inBeta, inStable);
        }

        // ###########################################################################################
        // Each entry with its view count, or as it was when there are no counts (no store, or it
        // could not be read). A board nobody looked at reads 0 - it IS counted, and is not.
        // ###########################################################################################
        public static IReadOnlyList<BoardOverviewEntry> WithViews(
            IReadOnlyList<BoardOverviewEntry> entries,
            IReadOnlyDictionary<string, int>? viewsLast30Days)
        {
            ArgumentNullException.ThrowIfNull(entries);

            return viewsLast30Days is null
                ? entries
                : entries.Select(entry => entry with { ViewsLast30Days = viewsLast30Days.GetValueOrDefault(entry.BoardId) }).ToList();
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

        public static async Task<BoardOverviewOutcome> DetailAsync(
            ReviewAccess? access,
            string? boardId,
            IReadOnlyList<PublishedBoardLister.KnownBoard> inBeta,
            IReadOnlyList<PublishedBoardLister.KnownBoard>? inProduction,
            ISubmissionStore submissions,
            IAccountStore accounts,
            CancellationToken cancellationToken = default,
            DateTimeOffset? now = null,
            IBoardViewStore? boardViews = null,
            BoardListings? listings = null,
            bool oneSubmissionInBeta = false)
        {
            ArgumentNullException.ThrowIfNull(inBeta);
            ArgumentNullException.ThrowIfNull(submissions);
            ArgumentNullException.ThrowIfNull(accounts);

            if (!ReviewAuthority.CanReviewAnything(access))
                return BoardOverviewOutcome.Forbidden();

            if (string.IsNullOrWhiteSpace(boardId))
                return BoardOverviewOutcome.NotFound();

            BoardRecord? row = await submissions.FindBoardAsync(boardId, cancellationToken);
            PublishedBoardLister.KnownBoard? known = inBeta.FirstOrDefault(board => string.Equals(board.BoardId, boardId, StringComparison.Ordinal));
            PublishedBoardLister.KnownBoard? inStable = inProduction?.FirstOrDefault(board => string.Equals(board.BoardId, boardId, StringComparison.Ordinal));

            // A board is a row, a board in either tree, or a row in the stable source's drop-down
            // list (2026-10-04) - none of them is no board at all.
            if (row is null && known is null && inStable is null && listings?.Stable?.Contains(boardId) != true)
                return BoardOverviewOutcome.NotFound();

            // Its people's addresses only for its own maintainers and the administrator (2026-10-05).
            bool showAddresses = ReviewAuthority.CanSeeAddressesOf(access, boardId);

            List<MaintainerRecord> pool = (await accounts.ListMaintainersAsync(cancellationToken))
                .Where(maintainer => string.Equals(maintainer.BoardId, boardId, StringComparison.Ordinal))
                .ToList();

            IReadOnlyList<BoardSubmissionRecord> sent =
                await submissions.GetSubmissionsForBoardAsync(boardId, BoardOverviewFlow.StoreLimit, cancellationToken);

            // How often CRT users look at it (2026-09-27), from the last year's views.
            DateTimeOffset at = now ?? DateTimeOffset.UtcNow;

            IReadOnlyList<BoardViewFact>? facts = await BoardOverviewFlow.ReadViewsAsync(
                () => boardViews!.FactsForBoardAsync(boardId, BoardViewStatisticsRules.WindowStart(at, 365), cancellationToken),
                boardViews is not null);

            BoardViewStatistics? views = facts is null ? null : BoardViewStatisticsRules.Build(facts, at);

            BoardOverviewEntry entry = BoardOverviewFlow.Entry(
                boardId, row, known, BoardOverviewFlow.Holds(inProduction, boardId), pool.Count, listings, inStable) with
            {
                ViewsLast30Days = views?.Last30Days
            };

            // The invitations nobody has accepted yet (2026-09-27) - for the ADMINISTRATOR, the one
            // person who can invite or withdraw one. Null for everybody else, who has no use for them.
            List<MaintainerInvitationEntry>? invitations = null;

            if (ReviewAuthority.CanAdminister(access))
            {
                invitations = (await accounts.ListOpenInvitationsAsync(at, cancellationToken))
                    .Where(invitation => string.Equals(invitation.BoardId, boardId, StringComparison.Ordinal))
                    .Select(invitation => new MaintainerInvitationEntry(invitation.Id, invitation.Email, invitation.CreatedUtc, invitation.ExpiresUtc))
                    .ToList();
            }

            // What has happened to it (2026-09-27): its submissions' own dates, and the audit rows
            // naming the board or one of those submissions.
            IReadOnlyList<AuditEntry> audit = await accounts.GetAuditForSubjectsAsync(
                [boardId, .. sent.Select(record => BoardHistoryRules.SubmissionSubject(record.Submission.Id))],
                BoardHistoryRules.Limit,
                cancellationToken);

            // The accounts the screen names, in ONE query (code review, 2026-10-09): the contributors
            // who were signed in and the maintainers who decided them, and - for the history, which
            // names people by their accounts (2026-10-09; only with no address shown before that,
            // 2026-10-05) - whoever did each audited thing and whoever a pool change named.
            IReadOnlyDictionary<long, AccountRecord> named = await accounts.FindByIdsAsync(
                sent.Select(record => record.Submission.AccountId)
                    .Concat(sent.Select(record => record.DecidedByAccountId))
                    .Concat(audit.Select(entry => entry.ActorAccountId))
                    .Concat(audit.Select(BoardHistoryRules.AccountNamedBy))
                    .OfType<long>()
                    .ToHashSet(),
                cancellationToken);

            // Whose contributor discarded their own draft since sending it (2026-09-28).
            IReadOnlyDictionary<long, DateTimeOffset> discarded = await submissions.GetDraftDiscardsAsync(
                sent.Select(record => record.Submission.Id).ToList(), cancellationToken);

            // Which a BETA rollback returned to the queue (migration 0015) - "returned", not "pending".
            IReadOnlyDictionary<long, DateTimeOffset> returns = await submissions.GetBetaReturnsAsync(
                sent.Select(record => record.Submission.Id).ToList(), cancellationToken);

            // What each changed as it went into BETA (2026-10-04, migration 0017) - the History view's
            // summaries. One that cannot be read leaves the summaries out, never the screen.
            IReadOnlyDictionary<long, SubmissionChanges>? changes = await BoardOverviewFlow.ReadViewsAsync(
                () => submissions.GetChangesAsync(sent.Select(record => record.Submission.Id).ToList(), cancellationToken),
                canRead: true);

            // Whether this account may change its BETA data now, and why not (code review,
            // 2026-10-09) - the table's own rule (BoardEditFlow), so the Boards screen says it above
            // every view. Nothing to say for a board BETA does not hold: its table cannot be read.
            string? mayNotEditReason = known is null
                ? null
                : await BoardEditFlow.WhyNotEditableAsync(access, boardId, submissions, oneSubmissionInBeta, cancellationToken);

            return BoardOverviewOutcome.Described(new BoardDetailAnswer(
                entry,
                BoardOverviewFlow.Maintainers(pool, showAddresses),
                BoardOverviewFlow.Contributors(sent, named, showAddresses),
                BoardOverviewFlow.Submissions(sent, row?.ProductionPublishedUtc, named, discarded, returns, changes, showAddresses),
                invitations,
                BoardHistoryRules.Build(sent, named, audit, showAddresses),
                views,
                AddressesHidden: !showAddresses,
                MayEdit: known is null ? null : mayNotEditReason is null,
                MayNotEditReason: mayNotEditReason));
        }

        // ###########################################################################################
        // Who maintains it - with their addresses, or (2026-10-05) with an EMPTY address for an
        // account that does not maintain it: PoolMaintainerEntry.Email is not nullable on the wire,
        // and an empty one is what CRT already reads as "no address to show".
        // ###########################################################################################
        public static IReadOnlyList<PoolMaintainerEntry> Maintainers(IEnumerable<MaintainerRecord> pool, bool showAddresses = true)
        {
            ArgumentNullException.ThrowIfNull(pool);

            return pool
                .Select(maintainer => new PoolMaintainerEntry(
                    maintainer.AccountId,
                    maintainer.DisplayName,
                    showAddresses ? maintainer.Email : string.Empty))
                .ToList();
        }

        // ###########################################################################################
        // The list: every row, every board in the BETA tree - and every board the stable source holds
        // or lists that BETA does not (2026-10-04), so what is off there shows - once each, ordered
        // by id: the order the administrator's Maintainers screen uses, so the two lists read alike.
        // ###########################################################################################
        public static IReadOnlyList<BoardOverviewEntry> Entries(
            IReadOnlyList<BoardRecord> rows,
            IReadOnlyList<PublishedBoardLister.KnownBoard> inBeta,
            IReadOnlyList<PublishedBoardLister.KnownBoard>? inProduction,
            IReadOnlyList<MaintainerRecord> maintainers,
            BoardListings? listings = null)
        {
            ArgumentNullException.ThrowIfNull(rows);
            ArgumentNullException.ThrowIfNull(inBeta);
            ArgumentNullException.ThrowIfNull(maintainers);

            Dictionary<string, BoardRecord> byRow = rows
                .GroupBy(row => row.BoardId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

            Dictionary<string, PublishedBoardLister.KnownBoard> byTree = inBeta
                .GroupBy(board => board.BoardId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

            Dictionary<string, int> poolSizes = maintainers
                .GroupBy(maintainer => maintainer.BoardId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

            Dictionary<string, PublishedBoardLister.KnownBoard> byStable = (inProduction ?? [])
                .GroupBy(board => board.BoardId, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

            // A row of the stable list naming a board nothing else knows - by its id, which is how
            // the file names it (MasterListingRow.BoardId). The lists match ids in any case, so one
            // differing only in case from a known board is that board, not another.
            var known = new HashSet<string>(byRow.Keys.Concat(byTree.Keys).Concat(byStable.Keys), StringComparer.OrdinalIgnoreCase);
            IEnumerable<string> listedOnly = ((IEnumerable<string>?)listings?.Stable ?? []).Where(id => !known.Contains(id));

            return byRow.Keys
                .Union(byTree.Keys, StringComparer.Ordinal)
                .Union(byStable.Keys, StringComparer.Ordinal)
                .Union(listedOnly, StringComparer.Ordinal)
                .OrderBy(id => id, StringComparer.Ordinal)
                .Select(id => BoardOverviewFlow.Entry(
                    id,
                    byRow.GetValueOrDefault(id),
                    byTree.GetValueOrDefault(id),
                    BoardOverviewFlow.Holds(inProduction, id),
                    poolSizes.GetValueOrDefault(id),
                    listings,
                    byStable.GetValueOrDefault(id)))
                .ToList();
        }

        // ###########################################################################################
        // One board's facts. The row names it where there is one (it is the truth once anything has
        // been submitted); the tree names a shipped board nobody has touched.
        // ###########################################################################################
        //
        // `inStable` names a board the stable source holds and BETA does not; a board nothing but the
        // stable source's drop-down list knows is named from its id's three folders.
        public static BoardOverviewEntry Entry(
            string boardId,
            BoardRecord? row,
            PublishedBoardLister.KnownBoard? inBeta,
            bool? inProduction,
            int maintainerCount,
            BoardListings? listings = null,
            PublishedBoardLister.KnownBoard? inStable = null)
        {
            string[] parts = boardId.Split('/');
            string Part(int index) => parts.Length == 3 ? parts[index] : string.Empty;

            return new BoardOverviewEntry(
                boardId,
                row?.Manufacturer ?? inBeta?.Manufacturer ?? inStable?.Manufacturer ?? Part(0),
                row?.Hardware ?? inBeta?.Hardware ?? inStable?.Hardware ?? Part(1),
                row?.Board ?? inBeta?.Board ?? inStable?.Board ?? Part(2),
                InBeta: inBeta is not null,
                InProduction: inProduction,
                IsAwaitingProduction: ProductionPromotionRules.IsAwaitingProduction(row),

                // No row yet means nothing has closed it - the default the column has.
                IsAccepting: row?.IsAccepting ?? true,
                BetaRevision: row?.CurrentRevision,
                ProductionRevision: row?.ProductionRevision,
                ProductionPublishedUtc: row?.ProductionPublishedUtc,
                MaintainerCount: maintainerCount,

                // What BETA holds, as last published or pushed back - the Boards screen reads its
                // table again when this moves (2026-10-04).
                BetaContentHash: row?.ContentHash,

                // Whether each source's drop-down list names it (2026-10-04) - null when unknown.
                ListedInBeta: listings?.Beta?.Contains(boardId),
                ListedInStable: listings?.Stable?.Contains(boardId));
        }

        // ###########################################################################################
        // WHICH BOARDS EACH SOURCE'S DROP-DOWN LISTS NAME (2026-10-04) - BETA's and the stable
        // source's newest main Excel data file, by board id, any case. A list that cannot be read is
        // null, never empty: "not listed" is only said after looking. The boards overview is read
        // every minute while its screen is shown, so each file is read once per version
        // (MasterListingIds) rather than on every request.
        // ###########################################################################################
        public static BoardListings ReadListings(string? betaRoot, string? stableRoot) =>
            new(BoardOverviewFlow.ListedIds(betaRoot), BoardOverviewFlow.ListedIds(stableRoot));

        private static IReadOnlySet<string>? ListedIds(string? root)
        {
            if (string.IsNullOrWhiteSpace(root) || MasterListing.NewestMasterPath(root) is not string master)
                return null;

            return BoardOverviewFlow.ListingReads.Read(master);
        }

        private static readonly MasterListingIds ListingReads = new();

        // ###########################################################################################
        // Who has contributed to this board, and how it went - newest contributor first.
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
        //
        // `showAddresses` false (2026-10-05): every Email is null - grouped exactly as with them, so a
        // contributor without an account is still one entry, just not named by an address.
        // ###########################################################################################
        public static IReadOnlyList<BoardContributorEntry> Contributors(
            IEnumerable<BoardSubmissionRecord> sent,
            IReadOnlyDictionary<long, AccountRecord>? accounts = null,
            bool showAddresses = true)
        {
            ArgumentNullException.ThrowIfNull(sent);

            return sent
                .GroupBy(record => BoardOverviewFlow.ContributorKey(record.Submission), StringComparer.Ordinal)
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

                    int Count(Func<BoardSubmissionRecord, bool> which) => group.Count(which);

                    return new BoardContributorEntry(
                        Email: !showAddresses || string.IsNullOrWhiteSpace(email) ? null : email,
                        Name: name,
                        Accepted: Count(record => record.Submission.State == SubmissionState.Merged),
                        Waiting: Count(record => record.Submission.State is SubmissionState.Pending or SubmissionState.Approved),
                        ChangesRequested: Count(record => record.Submission.State == SubmissionState.ChangesRequested),
                        Rejected: Count(record => record.Submission.State == SubmissionState.Rejected && record.DecidedByMaintainer),
                        LastSubmittedUtc: group.Max(record => record.Submission.CreatedUtc));
                })
                .OrderByDescending(contributor => contributor.LastSubmittedUtc)
                .ThenBy(contributor => contributor.Email, StringComparer.OrdinalIgnoreCase)
                .ThenBy(contributor => contributor.Name, StringComparer.OrdinalIgnoreCase)
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
        // `changes` (2026-10-04): what each changed as it went into BETA, where that was recorded.
        // `showAddresses` false (2026-10-05): no ContactEmail on any of them.
        public static IReadOnlyList<BoardSubmissionEntry> Submissions(
            IEnumerable<BoardSubmissionRecord> sent,
            DateTimeOffset? boardProductionPublishedUtc,
            IReadOnlyDictionary<long, AccountRecord>? accounts = null,
            IReadOnlyDictionary<long, DateTimeOffset>? discarded = null,
            IReadOnlyDictionary<long, DateTimeOffset>? returns = null,
            IReadOnlyDictionary<long, SubmissionChanges>? changes = null,
            bool showAddresses = true)
        {
            ArgumentNullException.ThrowIfNull(sent);

            return sent
                .Select(record => record.Submission)
                .OrderByDescending(submission => submission.Id)
                .Take(BoardOverviewFlow.ListedSubmissions)
                .Select(submission => new BoardSubmissionEntry(
                    submission.Id,
                    !showAddresses
                        ? null
                        : submission.AccountId is long id && accounts?.GetValueOrDefault(id) is AccountRecord account
                            ? account.Email
                            : submission.ContactEmail?.Trim(),
                    submission.Summary,
                    ProductionPromotionRules.ContributorFacingState(
                        submission.State, submission.DecidedUtc, boardProductionPublishedUtc,
                        returns is not null && returns.TryGetValue(submission.Id, out DateTimeOffset returned) ? returned : null),
                    submission.CreatedUtc,
                    submission.DecidedUtc,
                    submission.DecisionComment,
                    discarded is not null && discarded.TryGetValue(submission.Id, out DateTimeOffset at) ? at : null,
                    changes?.GetValueOrDefault(submission.Id)))
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
        private static bool? Holds(IReadOnlyList<PublishedBoardLister.KnownBoard>? tree, string boardId) =>
            tree is null ? null : tree.Any(board => string.Equals(board.BoardId, boardId, StringComparison.Ordinal));
    }

    // Which boards each source's drop-down lists name - null for a list that could not be read, or a
    // source the server does not have (BoardOverviewFlow.ReadListings).
    public sealed record BoardListings(IReadOnlySet<string>? Beta, IReadOnlySet<string>? Stable);

    // ###########################################################################################
    // A main Excel data file's board ids, read once per version of the file (its path, size and
    // last write) - the overview asks every minute, and the file changes a few times a month.
    // ###########################################################################################
    public sealed class MasterListingIds
    {
        private readonly System.Collections.Concurrent.ConcurrentDictionary<string, (DateTime Written, long Size, IReadOnlySet<string> Ids)> thisRead =
            new(StringComparer.Ordinal);

        public IReadOnlySet<string>? Read(string masterPath)
        {
            try
            {
                var file = new FileInfo(masterPath);

                if (!file.Exists)
                    return null;

                if (this.thisRead.TryGetValue(masterPath, out var cached) &&
                    cached.Written == file.LastWriteTimeUtc &&
                    cached.Size == file.Length)
                {
                    return cached.Ids;
                }

                if (!MasterListing.TryRead(masterPath, out IReadOnlyList<MasterListingRow> rows, out _))
                    return null;

                IReadOnlySet<string> ids = rows
                    .Select(row => row.BoardId)
                    .Where(id => id.Length > 0)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                this.thisRead[masterPath] = (file.LastWriteTimeUtc, file.Length, ids);
                return ids;
            }
            catch (IOException)
            {
                return null;
            }
            catch (UnauthorizedAccessException)
            {
                return null;
            }
        }
    }

    public sealed record BoardOverviewOutcome(
        IReadOnlyList<BoardOverviewEntry>? Boards,
        BoardDetailAnswer? Detail,
        bool IsForbidden = false,
        bool IsNotFound = false)
    {
        public static BoardOverviewOutcome Listed(IReadOnlyList<BoardOverviewEntry> boards) => new(boards, null);

        public static BoardOverviewOutcome Described(BoardDetailAnswer detail) => new(null, detail);

        public static BoardOverviewOutcome Forbidden() => new(null, null, IsForbidden: true);

        public static BoardOverviewOutcome NotFound() => new(null, null, IsNotFound: true);
    }
}
