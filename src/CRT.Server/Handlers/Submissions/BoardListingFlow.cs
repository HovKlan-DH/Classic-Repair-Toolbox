using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // PLACING A NEW BOARD IN THE DROP-DOWN LISTS (owner request, 2026-09-27): "When a system is added
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
    // THE NOTES (owner request, 2026-10-05): a placement nobody has saved starts with the notes the
    // contributor wrote in CRT's "Create board" (BoardListingRules.SuggestedNotes). They reach the
    // main Excel data file's "Hardware notes in "Overview" tab" column only as part of the placement
    // the maintainer saves - who may correct them first.
    //
    // WHO: seeing the list is anybody who may open the Maintainer tab (the Boards screen's
    // "everything for everyone"); PLACING a board is who may publish it - a maintainer of that
    // board, or the administrator - since placing one already in BETA publishes its row.
    //
    // The shape ApprovePublishFlow keeps: this class fetches, checks authority and writes under the
    // one PublishLock; every rule is BoardListingRules, pure and tested without a tree.
    // ###########################################################################################
    public sealed class BoardListingFlow
    {
        private readonly ISubmissionStore thisStore;
        private readonly PublishLock thisLock;
        private readonly ILogger<BoardListingFlow> thisLogger;

        // Where a saved placement is recorded (2026-09-27), for the board's history. Optional so a
        // test that does not look at the audit trail need not build one; the service always has it.
        private readonly IAccountStore? thisAccounts;

        public const string PlacedAction = BoardHistoryEvents.Placed;

        public BoardListingFlow(ISubmissionStore store, PublishLock publishLock, ILogger<BoardListingFlow> logger, IAccountStore? accounts = null)
        {
            this.thisStore = store;
            this.thisLock = publishLock;
            this.thisLogger = logger;
            this.thisAccounts = accounts;
        }

        // ###########################################################################################
        // The list as CRT shows it from BETA, and every board not in it that needs a place: its
        // board is in BETA already, or a submission for it is waiting. A new board whose every
        // submission was turned down is not offered - there is nothing to place yet.
        // ###########################################################################################
        public async Task<BoardListingOutcome> ListAsync(
            ReviewAccess? access,
            string? betaRoot,
            IReadOnlyList<PublishedBoardLister.KnownBoard> inBeta,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(inBeta);

            if (!ReviewAuthority.CanReviewAnything(access))
                return BoardListingOutcome.Forbidden();

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

            IReadOnlyList<BoardRecord> boards = await this.thisStore.ListBoardsAsync(cancellationToken);

            var candidates = new Dictionary<string, (string Manufacturer, string Hardware, string Board, bool InBeta)>(StringComparer.Ordinal);

            foreach (BoardRecord board in boards)
                candidates[board.BoardId] = (board.Manufacturer, board.Hardware, board.Board, false);

            foreach (PublishedBoardLister.KnownBoard board in inBeta)
                candidates[board.BoardId] = (board.Manufacturer, board.Hardware, board.Board, true);

            var unlisted = new List<UnlistedBoardEntry>();

            foreach ((string boardId, (string manufacturer, string hardware, string board, bool isInBeta)) in candidates.OrderBy(pair => pair.Key, StringComparer.Ordinal))
            {
                // Only with a list to be missing from: without one, nothing can be said to be missing.
                if (!hasList || MasterListing.IndexOfBoard(rows, boardId) >= 0)
                    continue;

                if (!isInBeta)
                {
                    IReadOnlyList<BoardSubmissionRecord> sent = await this.thisStore
                        .GetSubmissionsForBoardAsync(boardId, BoardOverviewFlow.StoreLimit, cancellationToken);

                    if (!sent.Any(record => BoardListingRules.IsWaiting(record.Submission.State)))
                        continue;
                }

                // The contributor's own notes from "Create board" (owner request, 2026-10-05) - the
                // placement starts with them, and saving it puts them in the main Excel data file.
                string notes = BoardListingRules.SuggestedNotes(
                    await this.thisStore.GetHardwareNotesAsync(boardId, cancellationToken));

                unlisted.Add(new UnlistedBoardEntry(
                    boardId,
                    manufacturer,
                    hardware,
                    board,
                    isInBeta,
                    ReviewAuthority.CanPublish(access, boardId),
                    await this.thisStore.GetPlacementAsync(boardId, cancellationToken),
                    BoardListingRules.Suggest(rows, manufacturer, hardware, board, notes)));
            }

            return BoardListingOutcome.Listed(new BoardListingAnswer(
                hasList,
                rows.Select(row => new BoardListingRow(row.BoardId, row.HardwareName, row.BoardName, row.ExcelDataFile)).ToList(),
                unlisted));
        }

        // ###########################################################################################
        // Saves where a board goes - and, when its board is already in BETA, writes the row into
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

            if (!BoardListingRules.TryNormalise(request, out string boardId, out BoardPlacement placement, out string problem))
                return SetPlacementOutcome.BadRequest(problem);

            if (!ReviewAuthority.CanPublish(access, boardId))
                return SetPlacementOutcome.Forbidden($"This account is not a maintainer of {boardId}.");

            if (string.IsNullOrWhiteSpace(betaRoot))
                return SetPlacementOutcome.Conflict("The server has no data tree configured.");

            // Under the publish lock: placing a board that is already in BETA writes BETA's main
            // Excel data file, which a publish or a promotion must not be reading half-way.
            using IDisposable gate = await this.thisLock.EnterAsync(cancellationToken).ConfigureAwait(false);

            string? master = MasterListing.NewestMasterPath(betaRoot);

            if (master is null)
                return SetPlacementOutcome.Conflict(BoardListingRules.NoMasterMessage);

            if (!MasterListing.TryRead(master, out IReadOnlyList<MasterListingRow> rows, out string why))
                return SetPlacementOutcome.Conflict($"The main Excel data file could not be read: {why}");

            if (MasterListing.IndexOfBoard(rows, boardId) >= 0)
                return SetPlacementOutcome.Conflict(BoardListingRules.AlreadyListedMessage);

            if (MasterListing.NamesTakenBy(rows, boardId, placement.HardwareName, placement.BoardName) is MasterListingRow taken)
                return SetPlacementOutcome.Conflict(MasterListing.NamesTakenMessage(taken));

            if (!BoardListingRules.AfterIsListed(rows, placement.AfterExcelDataFile))
                return SetPlacementOutcome.BadRequest($"[{placement.AfterExcelDataFile}] is not in the list. Reload the Boards screen and place it again.");

            BoardRecord? boardRecord = await this.thisStore.FindBoardAsync(boardId, cancellationToken);
            PublishedBoardLocation board = PublishedBoardLocator.LocateBoard(betaRoot, boardId);

            if (boardRecord is null && !board.Exists)
                return SetPlacementOutcome.NotFound();

            // A board in BETA with no row yet - copied there by hand - gets one, so the placement has
            // somewhere to be kept. It came from nowhere this pipeline knows, so it is "shipped".
            if (boardRecord is null)
            {
                string[] parts = boardId.Split('/');

                await this.thisStore.EnsureBoardAsync(
                    boardId, parts[0], parts[1], parts[2],
                    BoardDescriptorRules.BoardOrigin.Shipped, nowUtc, cancellationToken);
            }

            if (!await this.thisStore.SetPlacementAsync(boardId, placement, access!.Account.Id, nowUtc, cancellationToken))
                return SetPlacementOutcome.NotFound();

            if (this.thisAccounts is not null)
            {
                await this.thisAccounts.WriteAuditAsync(
                    new AuditEntry(access.Account.Id, access.Account.Email, BoardListingFlow.PlacedAction, boardId,
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
                BoardListingRules.RowFor(placement, betaRoot, board.WorkbookPath),
                placement.AfterExcelDataFile,
                nowUtc);

            if (!edit.IsDone)
            {
                this.thisLogger.LogWarning(
                    "Placement of {BoardId} was saved but could not be written to [{Master}]: {Failure}",
                    boardId, master, edit.Failure);

                return SetPlacementOutcome.Conflict($"The placement was saved, but it could not be added to BETA's list yet: {edit.Failure}");
            }

            this.thisLogger.LogInformation(
                "{Account} added {BoardId} to BETA's drop-down lists as [{Hardware}] / [{Board}].",
                access.Account.Email, boardId, placement.HardwareName, placement.BoardName);

            return SetPlacementOutcome.Saved(new SetPlacementAnswer(
                placement,
                ListedInBeta: true,
                "Saved and added to the drop-down lists in BETA."));
        }
    }

    // ###########################################################################################
    // The rules, pure.
    // ###########################################################################################
    public static class BoardListingRules
    {
        public const string NoMasterMessage =
            "BETA has no versioned main Excel data file to add the board to. Create it on the server first.";

        public const string AlreadyListedMessage =
            "This board is already in the drop-down lists. Only a board not listed yet can be placed.";

        public const string NotPlacedMessage =
            "This is a new board, and it has not been placed in the drop-down lists yet. Place it in the " +
            "Boards screen first - it is added to the lists when it is published to BETA.";

        // The states a submission waits in - a new board with one of these needs a place.
        public static bool IsWaiting(string? state) =>
            state is SubmissionState.Pending or SubmissionState.Approved or SubmissionState.ChangesRequested;

        // ###########################################################################################
        // Where a board goes untouched: after the last board of the same hardware FOLDER, under that
        // hardware's own name - so Commodore/C128/310378 Open128 joins "Commodore 128" rather than
        // appearing as a hardware of its own called "C128". A hardware nobody has listed yet goes at
        // the end, under its folder name.
        //
        // `notes` are the contributor's (SuggestedNotes, 2026-10-05) - the Notes box starts with them.
        // ###########################################################################################
        public static BoardPlacement Suggest(IReadOnlyList<MasterListingRow> rows, string manufacturer, string hardware, string board, string? notes = null)
        {
            ArgumentNullException.ThrowIfNull(rows);

            string prefix = $"{manufacturer}/{hardware}/";
            string suggestedNotes = notes?.Trim() ?? string.Empty;

            MasterListingRow? sameHardware = rows.LastOrDefault(row =>
                (row.BoardId + "/").StartsWith(prefix, StringComparison.OrdinalIgnoreCase));

            if (sameHardware is not null)
                return new BoardPlacement(sameHardware.HardwareName, board, suggestedNotes, sameHardware.ExcelDataFile);

            return new BoardPlacement(hardware, board, suggestedNotes, rows.Count == 0 ? null : rows[^1].ExcelDataFile);
        }

        // ###########################################################################################
        // *** WHICH NOTES THE PLACEMENT STARTS WITH (owner request, 2026-10-05). *** The newest notes
        // of a submission that is still waiting or was published - never of one rejected, withdrawn
        // or abandoned, whose contributor's words were turned down with it, nor of one still
        // uploading, which nobody has seen. A newer submission's notes replace an older one's, as its
        // data does. Empty when none count.
        // ###########################################################################################
        public static string SuggestedNotes(IReadOnlyList<SubmissionNotes>? notes) =>
            notes?
                .Where(entry => (BoardListingRules.IsWaiting(entry.State) || entry.State == SubmissionState.Merged) &&
                    !string.IsNullOrWhiteSpace(entry.HardwareNotes))
                .OrderByDescending(entry => entry.CreatedUtc)
                .ThenByDescending(entry => entry.SubmissionId)
                .Select(entry => entry.HardwareNotes.Trim())
                .FirstOrDefault()
            ?? string.Empty;

        // ###########################################################################################
        // A request made into a placement fit to store and write: a well-formed board id, names that
        // can be a row (MasterListing.IsWritableRow - the writer's own rule, so what is saved can
        // always be written), trimmed, and no "after" at all rather than a blank one.
        // ###########################################################################################
        public static bool TryNormalise(SetPlacementRequest? request, out string boardId, out BoardPlacement placement, out string problem)
        {
            boardId = request?.BoardId?.Trim() ?? string.Empty;
            placement = new BoardPlacement(string.Empty, string.Empty, string.Empty, null);

            if (!BoardDescriptorRules.IsValidBoardId(boardId))
            {
                problem = "That is not a board.";
                return false;
            }

            string after = request!.AfterExcelDataFile?.Trim() ?? string.Empty;

            placement = new BoardPlacement(
                request.HardwareName?.Trim() ?? string.Empty,
                request.BoardName?.Trim() ?? string.Empty,
                request.Notes?.Trim() ?? string.Empty,
                after.Length == 0 ? null : after);

            return MasterListing.IsWritableRow(
                new MasterListingRow(placement.HardwareName, placement.BoardName, boardId + "/x.xlsx", placement.Notes),
                out problem);
        }

        public static bool AfterIsListed(IReadOnlyList<MasterListingRow> rows, string? afterExcelDataFile) =>
            string.IsNullOrWhiteSpace(afterExcelDataFile) ||
            rows.Any(row => string.Equals(row.ExcelDataFile, afterExcelDataFile.Trim(), StringComparison.OrdinalIgnoreCase));

        // The row a placement becomes, naming the board workbook by its path from the data root.
        public static MasterListingRow RowFor(BoardPlacement placement, string dataRoot, string workbookPath)
        {
            ArgumentNullException.ThrowIfNull(placement);

            string relative = Path.GetRelativePath(Path.GetFullPath(dataRoot), Path.GetFullPath(workbookPath))
                .Replace(Path.DirectorySeparatorChar, '/');

            return new MasterListingRow(placement.HardwareName, placement.BoardName, relative, placement.Notes);
        }
    }

    public sealed record BoardListingOutcome(bool IsForbidden, BoardListingAnswer? Answer)
    {
        public static BoardListingOutcome Forbidden() => new(true, null);

        public static BoardListingOutcome Listed(BoardListingAnswer answer) => new(false, answer);
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

        public static SetPlacementOutcome NotFound() => new(null, "No such board.", IsNotFound: true);

        public static SetPlacementOutcome Conflict(string error) => new(null, error, IsConflict: true);

        public static SetPlacementOutcome BadRequest(string error) => new(null, error);
    }
}
