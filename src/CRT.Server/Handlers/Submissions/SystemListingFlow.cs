using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // PLACING A NEW SYSTEM IN THE DROP-DOWN LISTS (owner request, 2026-09-27): "When a system is added
    // to BETA, and it is a NEW system, can you then make sure it gets added also to the main Excel
    // data file in the Data root? The maintainer should order the new system, so it becomes visible
    // in the right location for the drop-down lists. This must be done before it can be pushed to
    // BETA." Done in the Maintainer tab's Systems screen, by dragging it into the full list.
    //
    // WHAT A PLACEMENT IS: one row of BETA's newest main Excel data file - the two drop-down names,
    // the Overview notes, and the row it goes after. It is STORED (systems.listing_*, migration 0011)
    // and WRITTEN:
    //   - at once, when the system's board is already in BETA (a new system published before this
    //     existed - Open128 was one), so it appears for BETA users straight away;
    //   - otherwise by the publish to BETA (ApprovePublishFlow), which refuses a new system that
    //     has not been placed;
    //   - into production's file by the promotion (ProductionPromotionFlow), next to the same
    //     neighbours.
    // A board of the tree that is not new never involves the file at all - it is listed already.
    //
    // WHO: seeing the list is anybody who may open the Maintainer tab (the Systems screen's
    // "everything for everyone"); PLACING a system is who may publish it - a maintainer of that
    // system, or the administrator - since placing one already in BETA publishes its row.
    //
    // The shape ApprovePublishFlow keeps: this class fetches, checks authority and writes under the
    // one PublishLock; every rule is SystemListingRules, pure and tested without a tree.
    // ###########################################################################################
    public sealed class SystemListingFlow
    {
        private readonly ISubmissionStore thisStore;
        private readonly PublishLock thisLock;
        private readonly ILogger<SystemListingFlow> thisLogger;

        // Where a saved placement is recorded (2026-09-27), for the system's history. Optional so a
        // test that does not look at the audit trail need not build one; the service always has it.
        private readonly IAccountStore? thisAccounts;

        public const string PlacedAction = SystemHistoryEvents.Placed;

        public SystemListingFlow(ISubmissionStore store, PublishLock publishLock, ILogger<SystemListingFlow> logger, IAccountStore? accounts = null)
        {
            this.thisStore = store;
            this.thisLock = publishLock;
            this.thisLogger = logger;
            this.thisAccounts = accounts;
        }

        // ###########################################################################################
        // The list as CRT shows it from BETA, and every system not in it that needs a place: its
        // board is in BETA already, or a submission for it is waiting. A new system whose every
        // submission was turned down is not offered - there is nothing to place yet.
        // ###########################################################################################
        public async Task<SystemListingOutcome> ListAsync(
            ReviewAccess? access,
            string? betaRoot,
            IReadOnlyList<PublishedSystemLister.KnownSystem> inBeta,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(inBeta);

            if (!ReviewAuthority.CanReviewAnything(access))
                return SystemListingOutcome.Forbidden();

            IReadOnlyList<MasterListingRow> rows = [];
            bool hasList = false;

            string? master = MasterListing.NewestMasterPath(betaRoot);

            if (master is not null)
            {
                if (MasterListing.TryRead(master, out IReadOnlyList<MasterListingRow> read, out string why))
                {
                    rows = read;
                    hasList = true;
                }
                else
                {
                    this.thisLogger.LogWarning("The main Excel data file [{Master}] could not be read: {Why}", master, why);
                }
            }

            IReadOnlyList<SystemRecord> systems = await this.thisStore.ListSystemsAsync(cancellationToken);

            var candidates = new Dictionary<string, (string Manufacturer, string Hardware, string Board, bool InBeta)>(StringComparer.Ordinal);

            foreach (SystemRecord system in systems)
                candidates[system.SystemId] = (system.Manufacturer, system.Hardware, system.Board, false);

            foreach (PublishedSystemLister.KnownSystem board in inBeta)
                candidates[board.SystemId] = (board.Manufacturer, board.Hardware, board.Board, true);

            var unlisted = new List<UnlistedSystemEntry>();

            foreach ((string systemId, (string manufacturer, string hardware, string board, bool isInBeta)) in candidates.OrderBy(pair => pair.Key, StringComparer.Ordinal))
            {
                // Only with a list to be missing from: without one, nothing can be said to be missing.
                if (!hasList || MasterListing.IndexOfSystem(rows, systemId) >= 0)
                    continue;

                if (!isInBeta)
                {
                    IReadOnlyList<SystemSubmissionRecord> sent = await this.thisStore
                        .GetSubmissionsForSystemAsync(systemId, SystemOverviewFlow.StoreLimit, cancellationToken);

                    if (!sent.Any(record => SystemListingRules.IsWaiting(record.Submission.State)))
                        continue;
                }

                unlisted.Add(new UnlistedSystemEntry(
                    systemId,
                    manufacturer,
                    hardware,
                    board,
                    isInBeta,
                    ReviewAuthority.CanPublish(access, systemId),
                    await this.thisStore.GetPlacementAsync(systemId, cancellationToken),
                    SystemListingRules.Suggest(rows, manufacturer, hardware, board)));
            }

            return SystemListingOutcome.Listed(new SystemListingAnswer(
                hasList,
                rows.Select(row => new SystemListingRow(row.SystemId, row.HardwareName, row.BoardName, row.ExcelDataFile)).ToList(),
                unlisted));
        }

        // ###########################################################################################
        // Saves where a system goes - and, when its board is already in BETA, writes the row into
        // BETA's main Excel data file at once (the caller then rebuilds the sync manifest).
        // ###########################################################################################
        public async Task<SetPlacementOutcome> SetAsync(
            ReviewAccess? access,
            SetPlacementRequest? request,
            string? betaRoot,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default)
        {
            if (!ReviewAuthority.CanReviewAnything(access))
                return SetPlacementOutcome.Forbidden(ReviewAuthority.DescribeRefusal(access, null));

            if (!SystemListingRules.TryNormalise(request, out string systemId, out SystemPlacement placement, out string problem))
                return SetPlacementOutcome.BadRequest(problem);

            if (!ReviewAuthority.CanPublish(access, systemId))
                return SetPlacementOutcome.Forbidden($"This account is not a maintainer of {systemId}.");

            if (string.IsNullOrWhiteSpace(betaRoot))
                return SetPlacementOutcome.Conflict("The server has no data tree configured.");

            // Under the publish lock: placing a system that is already in BETA writes BETA's main
            // Excel data file, which a publish or a promotion must not be reading half-way.
            using IDisposable gate = await this.thisLock.EnterAsync(cancellationToken).ConfigureAwait(false);

            string? master = MasterListing.NewestMasterPath(betaRoot);

            if (master is null)
                return SetPlacementOutcome.Conflict(SystemListingRules.NoMasterMessage);

            if (!MasterListing.TryRead(master, out IReadOnlyList<MasterListingRow> rows, out string why))
                return SetPlacementOutcome.Conflict($"The main Excel data file could not be read: {why}");

            if (MasterListing.IndexOfSystem(rows, systemId) >= 0)
                return SetPlacementOutcome.Conflict(SystemListingRules.AlreadyListedMessage);

            if (MasterListing.NamesTakenBy(rows, systemId, placement.HardwareName, placement.BoardName) is MasterListingRow taken)
                return SetPlacementOutcome.Conflict(MasterListing.NamesTakenMessage(taken));

            if (!SystemListingRules.AfterIsListed(rows, placement.AfterExcelDataFile))
                return SetPlacementOutcome.BadRequest($"[{placement.AfterExcelDataFile}] is not in the list. Reload the Systems screen and place it again.");

            SystemRecord? system = await this.thisStore.FindSystemAsync(systemId, cancellationToken);
            PublishedBoardLocation board = PublishedBoardLocator.LocateSystem(betaRoot, systemId);

            if (system is null && !board.Exists)
                return SetPlacementOutcome.NotFound();

            // A board in BETA with no row yet - copied there by hand - gets one, so the placement has
            // somewhere to be kept. It came from nowhere this pipeline knows, so it is "shipped".
            if (system is null)
            {
                string[] parts = systemId.Split('/');

                await this.thisStore.EnsureSystemAsync(
                    systemId, parts[0], parts[1], parts[2],
                    SystemDescriptorRules.SystemOrigin.Shipped, nowUtc, cancellationToken);
            }

            if (!await this.thisStore.SetPlacementAsync(systemId, placement, access!.Account.Id, nowUtc, cancellationToken))
                return SetPlacementOutcome.NotFound();

            if (this.thisAccounts is not null)
            {
                await this.thisAccounts.WriteAuditAsync(
                    new AuditEntry(access.Account.Id, access.Account.Email, SystemListingFlow.PlacedAction, systemId,
                        $"{placement.HardwareName} / {placement.BoardName}", nowUtc),
                    cancellationToken);
            }

            if (!board.Exists)
            {
                return SetPlacementOutcome.Saved(new SetPlacementAnswer(
                    placement,
                    ListedInBeta: false,
                    "Saved. It is added to the drop-down lists when it is published to BETA."));
            }

            MasterListingEdit edit = MasterListing.Insert(
                master,
                SystemListingRules.RowFor(placement, betaRoot, board.WorkbookPath),
                placement.AfterExcelDataFile,
                nowUtc);

            if (!edit.IsDone)
            {
                this.thisLogger.LogWarning(
                    "Placement of {SystemId} was saved but could not be written to [{Master}]: {Failure}",
                    systemId, master, edit.Failure);

                return SetPlacementOutcome.Conflict($"The placement was saved, but it could not be added to BETA's list yet: {edit.Failure}");
            }

            this.thisLogger.LogInformation(
                "{Account} added {SystemId} to BETA's drop-down lists as [{Hardware}] / [{Board}].",
                access.Account.Email, systemId, placement.HardwareName, placement.BoardName);

            return SetPlacementOutcome.Saved(new SetPlacementAnswer(
                placement,
                ListedInBeta: true,
                "Saved and added to the drop-down lists in BETA."));
        }
    }

    // ###########################################################################################
    // The rules, pure.
    // ###########################################################################################
    public static class SystemListingRules
    {
        public const string NoMasterMessage =
            "BETA has no versioned main Excel data file to add the system to. Create it on the server first.";

        public const string AlreadyListedMessage =
            "This system is already in the drop-down lists. Only a system not listed yet can be placed.";

        public const string NotPlacedMessage =
            "This is a new system, and it has not been placed in the drop-down lists yet. Place it in the " +
            "Systems screen first - it is added to the lists when it is published to BETA.";

        // The states a submission waits in - a new system with one of these needs a place.
        public static bool IsWaiting(string? state) =>
            state is SubmissionState.Pending or SubmissionState.Approved or SubmissionState.ChangesRequested;

        // ###########################################################################################
        // Where a system goes untouched: after the last board of the same hardware FOLDER, under that
        // hardware's own name - so Commodore/C128/310378 Open128 joins "Commodore 128" rather than
        // appearing as a hardware of its own called "C128". A hardware nobody has listed yet goes at
        // the end, under its folder name.
        // ###########################################################################################
        public static SystemPlacement Suggest(IReadOnlyList<MasterListingRow> rows, string manufacturer, string hardware, string board)
        {
            ArgumentNullException.ThrowIfNull(rows);

            string prefix = $"{manufacturer}/{hardware}/";

            MasterListingRow? sameHardware = rows.LastOrDefault(row =>
                (row.SystemId + "/").StartsWith(prefix, StringComparison.OrdinalIgnoreCase));

            if (sameHardware is not null)
                return new SystemPlacement(sameHardware.HardwareName, board, string.Empty, sameHardware.ExcelDataFile);

            return new SystemPlacement(hardware, board, string.Empty, rows.Count == 0 ? null : rows[^1].ExcelDataFile);
        }

        // ###########################################################################################
        // A request made into a placement fit to store and write: a well-formed system id, names that
        // can be a row (MasterListing.IsWritableRow - the writer's own rule, so what is saved can
        // always be written), trimmed, and no "after" at all rather than a blank one.
        // ###########################################################################################
        public static bool TryNormalise(SetPlacementRequest? request, out string systemId, out SystemPlacement placement, out string problem)
        {
            systemId = request?.SystemId?.Trim() ?? string.Empty;
            placement = new SystemPlacement(string.Empty, string.Empty, string.Empty, null);

            if (!SystemDescriptorRules.IsValidSystemId(systemId))
            {
                problem = "That is not a system.";
                return false;
            }

            string after = request!.AfterExcelDataFile?.Trim() ?? string.Empty;

            placement = new SystemPlacement(
                request.HardwareName?.Trim() ?? string.Empty,
                request.BoardName?.Trim() ?? string.Empty,
                request.Notes?.Trim() ?? string.Empty,
                after.Length == 0 ? null : after);

            return MasterListing.IsWritableRow(
                new MasterListingRow(placement.HardwareName, placement.BoardName, systemId + "/x.xlsx", placement.Notes),
                out problem);
        }

        public static bool AfterIsListed(IReadOnlyList<MasterListingRow> rows, string? afterExcelDataFile) =>
            string.IsNullOrWhiteSpace(afterExcelDataFile) ||
            rows.Any(row => string.Equals(row.ExcelDataFile, afterExcelDataFile.Trim(), StringComparison.OrdinalIgnoreCase));

        // The row a placement becomes, naming the board workbook by its path from the data root.
        public static MasterListingRow RowFor(SystemPlacement placement, string dataRoot, string workbookPath)
        {
            ArgumentNullException.ThrowIfNull(placement);

            string relative = Path.GetRelativePath(Path.GetFullPath(dataRoot), Path.GetFullPath(workbookPath))
                .Replace(Path.DirectorySeparatorChar, '/');

            return new MasterListingRow(placement.HardwareName, placement.BoardName, relative, placement.Notes);
        }
    }

    public sealed record SystemListingOutcome(bool IsForbidden, SystemListingAnswer? Answer)
    {
        public static SystemListingOutcome Forbidden() => new(true, null);

        public static SystemListingOutcome Listed(SystemListingAnswer answer) => new(false, answer);
    }

    public sealed record SetPlacementOutcome(
        SetPlacementAnswer? Answer,
        string Error,
        bool IsForbidden = false,
        bool IsNotFound = false,
        bool IsConflict = false)
    {
        public static SetPlacementOutcome Saved(SetPlacementAnswer answer) => new(answer, string.Empty);

        public static SetPlacementOutcome Forbidden(string error) => new(null, error, IsForbidden: true);

        public static SetPlacementOutcome NotFound() => new(null, "No such system.", IsNotFound: true);

        public static SetPlacementOutcome Conflict(string error) => new(null, error, IsConflict: true);

        public static SetPlacementOutcome BadRequest(string error) => new(null, error);
    }
}
