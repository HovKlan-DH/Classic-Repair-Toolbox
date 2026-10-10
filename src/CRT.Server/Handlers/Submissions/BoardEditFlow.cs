using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A BOARD'S BOARD DATA ON THE "BOARDS" SCREEN, AND A MAINTAINER'S EDIT OF IT (owner request,
    // 2026-10-03: "all the same functionalities, as the 'Contributor Submissions' has ... Board data
    // (and I should be able to do the same edits)").
    //
    // *** THE EDIT GOES STRAIGHT TO BETA (owner decision, 2026-10-03: "I do not think it should be
    // necessary for that change to first go to 'Contributor Submissions' queue - instead it should go
    // directly to the next queue, 'BETA > Stable', so it can directly be tested in BETA"). *** It
    // replaced the same day's "Becomes a submission". The edit is still made into a submission from
    // the maintainer's account - built HERE, validated, stored and queued by SubmissionFlows' own
    // create and finalise, with the reason the maintainer gave as its description - and then
    // APPROVED AT ONCE, as that maintainer, by ApprovePublishFlow itself. So nothing reaches BETA by
    // a path of its own: every rule an approval keeps (the files it removes, shared files, one
    // submission in BETA per board, the publish lock) holds for this too, and BETA > Stable, the
    // board's history and a push-back all see an ordinary merged submission.
    //
    // *** NOT WHILE THE BOARD WAITS IN BETA > STABLE (owner decision, same day: "If a system is
    // already in 'BETA > Stable' queue, then it should simply disallow it, even if this is coming
    // from a maintainer"). *** One submission in BETA per board, exactly as for an approval - and so
    // only where publishing to stable is configured, the rule's own condition. The table opens
    // read-only with the reason (WhyNotEditableAsync).
    //
    // THE ORDER OF THE CHECKS, cheap refusals first and nothing made until all have passed:
    //
    //   1. AUTHORITY - may this account review anything, and THIS board? Anybody may READ every
    //                  board (owner decision, "Everything for everyone"); only the board's
    //                  maintainers and the administrator may change it - the same question a
    //                  submission's own table asks (AmendSubmissionFlow) - and not a board closed
    //                  to contributions, nor one waiting in BETA, nor while this account's last
    //                  change of it still waits under Contributor Submissions.
    //   2. BETA      - the board's data must be in BETA: an edit is of the board as published
    //                  there. A new board's first submission is edited in its own table.
    //   3. VERSION   - BETA's board must still be the one the table was opened on (Fingerprint).
    //                  Publishing a table read before another publish would put that publish's
    //                  changes back as they were - silently, since a board's rows are replaced
    //                  whole (PublishMerge).
    //   4. ROWS      - only the table's nine sheets come from the request; highlights, calibrations
    //                  and the revision date stay as BETA has them (SubmissionRowsBoard.
    //                  WithTableSections, the amendment's own rule). An edit that changes nothing
    //                  is refused.
    //   5. FILES     - every file the rows cite must be in BETA already, and is taken from there
    //                  (imported by the create, as any submission's unchanged files are). A
    //                  maintainer edits rows here and cannot bring in a file nobody has sent.
    //   6. REMOVALS  - what the publish would remove, which CheckAsync answers BEFORE the reason is
    //                  asked, and which the edit must send back unchanged - the approval's own rule
    //                  (FileRemovalPreview), checked here too so nothing is made for a publish the
    //                  approval would refuse.
    //   7. QUEUE     - SubmissionFlows.CreateAsync and FinaliseAsync: the same path, file, row and
    //                  content checks as any submission.
    //   8. APPROVE   - ApprovePublishFlow.ApproveAsync, as this maintainer, with the removals shown.
    //                  Should it refuse after all - something changed between the checks above and
    //                  its lock - the submission simply stays in the queue, where it can be approved
    //                  like any other: nothing of the maintainer's work is lost.
    //
    // Pure over its collaborators (the store, the blob store, the board reader, the approval), like
    // every flow - the endpoint (BoardEndpoints) is a rim.
    // ###########################################################################################
    public static class BoardEditFlow
    {
        // ###########################################################################################
        // BETA's board as the table opens on it: its rows, the KiCad calibrations from its highlight
        // file (which a board's rows do not carry - see PublishMerge.CalibrationsOf), and a
        // fingerprint of both for the edit to be sent back with.
        //
        // `oneSubmissionInBeta` is whether a board waiting in BETA takes no change - the server's
        // IsProductionPublishingConfigured, as for ApprovePublishFlow's step 3b.
        // ###########################################################################################
        public static async Task<BoardTableOutcome> ReadTableAsync(
            ReviewAccess? access,
            string? boardId,
            string? dataTreeRoot,
            ISubmissionStore store,
            PublishedBoardReader boards,
            string? betaDataUrl,
            bool oneSubmissionInBeta,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(boards);

            if (!ReviewAuthority.CanReviewAnything(access))
                return BoardTableOutcome.Forbidden(ReviewAuthority.DescribeRefusal(access, null));

            if (!BoardDescriptorRules.IsValidBoardId(boardId) || string.IsNullOrWhiteSpace(dataTreeRoot))
                return BoardTableOutcome.NotFound(BoardEditFlow.NotInBetaMessage);

            string root = Path.GetFullPath(dataTreeRoot);

            (BoardData Board, SubmissionRows Rows)? beta = await BoardEditFlow.ReadBetaAsync(root, boardId!, boards, cancellationToken)
                .ConfigureAwait(false);

            if (beta is null)
                return BoardTableOutcome.NotFound(BoardEditFlow.NotInBetaMessage);

            string? refusal = await BoardEditFlow.WhyNotEditableAsync(access, boardId!, store, oneSubmissionInBeta, cancellationToken)
                .ConfigureAwait(false);

            return BoardTableOutcome.Read(new BoardTableAnswer(
                boardId!,
                BoardEditFlow.Fingerprint(beta.Value.Rows),
                beta.Value.Rows,
                MayEdit: refusal is null,
                MayNotEditReason: refusal,
                BetaDataUrl: betaDataUrl));
        }

        // ###########################################################################################
        // THE STABLE SOURCE'S BOARD (owner request, 2026-10-04: a BETA / Stable switch on Board data -
        // "Shouldn't there be somewhere a possibility to see what we actually do have in BETA or
        // stable"). For any maintainer, like BETA's, and NEVER editable: the stable source changes
        // only when BETA is published to it. Not found when the server has no stable source, or it
        // holds nothing of the board.
        // ###########################################################################################
        public static async Task<BoardTableOutcome> ReadStableTableAsync(
            ReviewAccess? access,
            string? boardId,
            string? productionTreeRoot,
            PublishedBoardReader boards,
            string? productionDataUrl,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(boards);

            if (!ReviewAuthority.CanReviewAnything(access))
                return BoardTableOutcome.Forbidden(ReviewAuthority.DescribeRefusal(access, null));

            if (string.IsNullOrWhiteSpace(productionTreeRoot))
                return BoardTableOutcome.NotFound(BoardEditFlow.NoStableSourceMessage);

            if (!BoardDescriptorRules.IsValidBoardId(boardId))
                return BoardTableOutcome.NotFound(BoardEditFlow.NotInStableMessage);

            (BoardData Board, SubmissionRows Rows)? stable = await BoardEditFlow
                .ReadBetaAsync(Path.GetFullPath(productionTreeRoot), boardId!, boards, cancellationToken)
                .ConfigureAwait(false);

            if (stable is null)
                return BoardTableOutcome.NotFound(BoardEditFlow.NotInStableMessage);

            return BoardTableOutcome.Read(new BoardTableAnswer(
                boardId!,
                BoardEditFlow.Fingerprint(stable.Value.Rows),
                stable.Value.Rows,
                MayEdit: false,
                MayNotEditReason: BoardEditFlow.StableReadOnlyMessage,
                BetaDataUrl: null,
                ProductionDataUrl: productionDataUrl));
        }

        // ###########################################################################################
        // Steps 1 to 6 with nothing made: what publishing the edit would remove from BETA, for the
        // maintainer to see beside the reason they are asked for - or why it cannot be sent at all,
        // so they are not asked for a reason only to be refused.
        // ###########################################################################################
        public static async Task<BoardEditOutcome> CheckAsync(
            ReviewAccess? access,
            BoardEditRequest? request,
            string? dataTreeRoot,
            ISubmissionStore store,
            PublishedBoardReader boards,
            bool oneSubmissionInBeta,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(boards);

            (Prepared? prepared, BoardEditOutcome? refusal) = await BoardEditFlow
                .PrepareAsync(access, request, dataTreeRoot, store, boards, oneSubmissionInBeta, nowUtc, cancellationToken)
                .ConfigureAwait(false);

            return refusal ?? BoardEditOutcome.Checked(prepared!.Removals.Files);
        }

        // ###########################################################################################
        // The edit, published to BETA as a submission from this account that it approves at once -
        // see the header for the order.
        // ###########################################################################################
        public static async Task<BoardEditOutcome> PublishAsync(
            ReviewAccess? access,
            BoardEditRequest? request,
            string? dataTreeRoot,
            ISubmissionStore store,
            BlobStore blobs,
            PublishedBoardReader boards,
            ApprovePublishFlow approvals,
            bool oneSubmissionInBeta,
            string? ipAddress,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(blobs);
            ArgumentNullException.ThrowIfNull(boards);
            ArgumentNullException.ThrowIfNull(approvals);

            // ---- 1 to 6 -------------------------------------------------------------------------
            (Prepared? prepared, BoardEditOutcome? refusal) = await BoardEditFlow
                .PrepareAsync(access, request, dataTreeRoot, store, boards, oneSubmissionInBeta, nowUtc, cancellationToken)
                .ConfigureAwait(false);

            if (refusal is not null)
                return refusal;

            // The reason goes with the change into BETA, BETA > Stable and the board's history, as a
            // contribution's description does - so a change is not published without one.
            if (string.IsNullOrWhiteSpace(request!.Summary))
                return BoardEditOutcome.Refused(BoardEditFlow.NoReasonMessage, []);

            // The files the maintainer was shown, or a list that has moved since the check.
            if (!prepared!.Removals.Matches(request.ExpectedRemovals))
                return BoardEditOutcome.Conflict(BoardEditFlow.RemovalsChangedMessage);

            // ---- 7. Queue -----------------------------------------------------------------------
            //
            // From the account, and trusted: a maintainer is exempt from the per-address limit by the
            // database rows (SubmissionRateLimitPolicy), exactly as when sending from CRT signed in.
            Submitter submitter = Submitter.SignedIn(access!.Account.Id, ipAddress, isTrusted: true);

            SubmissionCreationOutcome created = await SubmissionFlows
                .CreateAsync(prepared.Manifest, submitter, prepared.Root, store, blobs, nowUtc, cancellationToken, prepared.Tree)
                .ConfigureAwait(false);

            if (!created.IsAccepted)
            {
                return created.IsNoRoom
                    ? BoardEditOutcome.Refused("The server is short of storage space, so the change was not sent.", created.Findings)
                    : BoardEditOutcome.Refused("The change cannot be sent:", created.Findings);
            }

            HashNegotiationResponse negotiation = created.Negotiation!;

            // Every file came from BETA, so nothing should be missing - unless one changed on the
            // server between being hashed and being imported. The row then waits in "uploading",
            // and the abandoned-upload sweep collects it.
            if (negotiation.MissingHashes.Count > 0)
            {
                return BoardEditOutcome.Refused(
                    "A file in BETA changed while the change was being sent. Open the board again and send it again.", []);
            }

            SubmissionResult queued = await SubmissionFlows
                .FinaliseAsync(negotiation.SubmissionId, negotiation.UploadToken, store, blobs, nowUtc, cancellationToken)
                .ConfigureAwait(false);

            if (!queued.IsAccepted)
                return BoardEditOutcome.Refused("The change cannot be sent:", queued.Findings);

            IReadOnlyList<ValidationFinding> warnings = queued.Findings
                .Where(finding => finding.Severity == ValidationSeverity.Warning)
                .ToList();

            // ---- 8. Approve ---------------------------------------------------------------------
            ApproveOutcome approved = await approvals
                .ApproveAsync(negotiation.SubmissionId, access, prepared.Root, nowUtc, prepared.Removals.Files, cancellationToken)
                .ConfigureAwait(false);

            if (approved.IsPublished)
            {
                return BoardEditOutcome.Published(
                    negotiation.SubmissionId, approved.Descriptor!.Revision, approved.RemovedFiles, warnings);
            }

            string why = approved.IsAwaitingApproval
                ? $"It needs the approval of {ApprovePublishFlow.Describe(approved.WaitingFor)} too."
                : approved.Error;

            return BoardEditOutcome.SavedNotPublished(negotiation.SubmissionId, why, warnings);
        }

        // What steps 1 to 6 found: where the tree is, the submission the edit becomes, and what its
        // publish would remove.
        private sealed record Prepared(string Root, PublishedTreeView? Tree, SubmissionManifest Manifest, FileRemovalPreview Removals);

        private static async Task<(Prepared? Prepared, BoardEditOutcome? Refusal)> PrepareAsync(
            ReviewAccess? access,
            BoardEditRequest? request,
            string? dataTreeRoot,
            ISubmissionStore store,
            PublishedBoardReader boards,
            bool oneSubmissionInBeta,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken)
        {
            // ---- 1. Authority ------------------------------------------------------------------
            if (!ReviewAuthority.CanReviewAnything(access))
                return (null, BoardEditOutcome.Forbidden(ReviewAuthority.DescribeRefusal(access, null)));

            string? boardId = request?.BoardId?.Trim();

            if (!BoardDescriptorRules.IsValidBoardId(boardId) || string.IsNullOrWhiteSpace(dataTreeRoot))
                return (null, BoardEditOutcome.NotFound(BoardEditFlow.NotInBetaMessage));

            string? refusal = await BoardEditFlow.WhyNotEditableAsync(access, boardId!, store, oneSubmissionInBeta, cancellationToken)
                .ConfigureAwait(false);

            if (refusal is not null)
                return (null, BoardEditOutcome.Forbidden(refusal));

            if (request!.Rows is null)
                return (null, BoardEditOutcome.Refused("The request carried no rows.", []));

            if ((request.Summary?.Trim().Length ?? 0) > SubmissionFormat.MaximumSummaryLength)
            {
                return (null, BoardEditOutcome.Refused(
                    $"The reason is longer than {SubmissionFormat.MaximumSummaryLength} characters. Shorten it.", []));
            }

            string root = Path.GetFullPath(dataTreeRoot!);

            // ---- 2. BETA -----------------------------------------------------------------------
            (BoardData Board, SubmissionRows Rows)? beta = await BoardEditFlow.ReadBetaAsync(root, boardId!, boards, cancellationToken)
                .ConfigureAwait(false);

            if (beta is null)
                return (null, BoardEditOutcome.NotFound(BoardEditFlow.NotInBetaMessage));

            // ---- 3. Version --------------------------------------------------------------------
            string current = BoardEditFlow.Fingerprint(beta.Value.Rows);

            if (!string.Equals(current, request.Fingerprint?.Trim(), StringComparison.Ordinal))
                return (null, BoardEditOutcome.Conflict(BoardEditFlow.ChangedSinceMessage));

            // ---- 4. Rows -----------------------------------------------------------------------
            SubmissionRows edited = SubmissionRowsBoard.WithTableSections(beta.Value.Rows, request.Rows);

            if (string.Equals(BoardEditFlow.Fingerprint(edited), current, StringComparison.Ordinal))
                return (null, BoardEditOutcome.Refused(BoardEditFlow.NothingChangedMessage, []));

            // ---- 5. Files ----------------------------------------------------------------------
            PublishedTreeView? tree = PublishedTreeProbe.For(root);
            var findings = new List<ValidationFinding>();
            var files = new List<SubmissionFile>();

            foreach (string path in SubmissionManifestBuilder.CollectReferencedFiles(SubmissionRowsBoard.ToBoard(edited)))
            {
                if (BoardEditFlow.TryFindInBeta(root, tree, path, out SubmissionFile? file))
                {
                    files.Add(file!);
                    continue;
                }

                findings.Add(new ValidationFinding
                {
                    Severity = ValidationSeverity.Error,
                    Code = "board_edit.file_unknown",
                    Subject = path,
                    Message = $"[{path}] is not in BETA. A change made on the Boards screen can only use files BETA already holds."
                });
            }

            if (findings.Count > 0)
                return (null, BoardEditOutcome.Refused("The change cannot be sent:", findings));

            string[] parts = boardId!.Split('/');

            var manifest = new SubmissionManifest
            {
                FormatVersion = SubmissionFormat.CurrentVersion,
                BoardId = boardId,
                Manufacturer = parts[0],
                Hardware = parts[1],
                Board = parts[2],

                // What BETA holds now - the board this edit was made against, and what a contributor's
                // draft of it would carry (PublishMerge.RevisionOf).
                BaseRevision = PublishMerge.RevisionOf(beta.Value.Board),
                Summary = request.Summary?.Trim() ?? string.Empty,
                CreatedUtc = nowUtc,
                Rows = edited,
                Files = files
            };

            // ---- 6. Removals -------------------------------------------------------------------
            //
            // The approval's own computation, against BETA's board as just read - the board the
            // approval will read too, unless something publishes in between (then it refuses, and the
            // submission waits in the queue - see the header).
            FileRemovalPreview removals = ApprovePublishFlow.PreviewRemovals(root, manifest, beta.Value.Board, nowUtc);

            return (new Prepared(root, tree, manifest, removals), null);
        }

        // ###########################################################################################
        // BETA's board as rows, with the calibrations its highlight file carries - or null when BETA
        // holds nothing of the board. A board that cannot be READ throws (PublishedBoardReader):
        // "not in BETA" would be a false claim about it.
        // ###########################################################################################
        private static async Task<(BoardData Board, SubmissionRows Rows)?> ReadBetaAsync(
            string root,
            string boardId,
            PublishedBoardReader boards,
            CancellationToken cancellationToken)
        {
            PublishedBoardLocation location = PublishedBoardLocator.LocateBoard(root, boardId);

            if (!location.Exists)
                return null;

            BoardData? board = await boards.TryReadBoardAsync(root, boardId, cancellationToken).ConfigureAwait(false);

            if (board is null)
                return null;

            SubmissionRows rows = SubmissionRowsBoard.FromBoard(board);
            rows.KiCadCalibrations = DraftBoardSource.CollectCalibrations(location.WorkbookPath, board).ToList();

            return (board, rows);
        }

        // ###########################################################################################
        // Null when this account may change the board; otherwise why not, for the screen.
        //
        // *** NOT WHILE THE BOARD WAITS IN BETA (owner decision, 2026-10-03). *** One submission in
        // BETA per board, as ApprovePublishFlow's step 3b - so a maintainer's second change waits
        // until the first has gone to stable, been pushed back or rejected. A board copied to stable
        // by hand is recorded as there the next time BETA > Stable's list is read, which the
        // Maintainer tab does every minute (ProductionPromotionFlow.ListEntriesAsync).
        //
        // *** NOT WHILE ITS LAST CHANGE OF THE BOARD STILL WAITS. *** A change that could not be
        // published stays in the queue (see the header), and a newer submission from the same account
        // REPLACES the older pending one (SubmissionReplacementRules) - so sending another would
        // quietly withdraw it. The table opens read-only instead, naming the submission to decide.
        //
        // The board's detail asks it too (BoardOverviewFlow.DetailAsync, code review 2026-10-09), so
        // the Boards screen says why above every view - one rule, never a second copy.
        // ###########################################################################################
        internal static async Task<string?> WhyNotEditableAsync(
            ReviewAccess? access,
            string boardId,
            ISubmissionStore store,
            bool oneSubmissionInBeta,
            CancellationToken cancellationToken)
        {
            if (!ReviewAuthority.CanReview(access, boardId))
                return BoardEditFlow.NotYoursMessage;

            if (await store.IsBoardAcceptingAsync(boardId, cancellationToken).ConfigureAwait(false) == false)
                return BoardEditFlow.ClosedMessage;

            if (oneSubmissionInBeta &&
                ProductionPromotionRules.IsAwaitingProduction(await store.FindBoardAsync(boardId, cancellationToken).ConfigureAwait(false)))
            {
                return OneSubmissionInBeta.NoChangeMessage(boardId);
            }

            SubmissionRecord? waiting = (await store.GetPendingForBoardAsync(boardId, cancellationToken).ConfigureAwait(false))
                .FirstOrDefault(record => record.AccountId == access!.Account.Id);

            // The words are CRT.Data's: CRT's Maintainer tab says the same when it cannot read the
            // table again after such a change (code review, 2026-10-10).
            return waiting is null ? null : BoardEditWording.AlreadyWaitingMessage(waiting.Id);
        }

        // ###########################################################################################
        // The file at `path` as BETA holds it - its hash from the tree's cache and its size from
        // disk - when the path is inside the tree, reached through no link, and there. The create
        // imports it from there and hashes what it copies, which is the real check.
        // ###########################################################################################
        private static bool TryFindInBeta(string root, PublishedTreeView? tree, string path, out SubmissionFile? file)
        {
            file = null;

            string? hash = tree?.HashOf(path);

            if (hash is null ||
                !SubmissionPathRules.TryResolve(root, path, out string resolved, out _) ||
                PublishPathSafety.FindLinkOnPath(root, resolved, PublishPathSafety.IsLink) is not null ||
                !File.Exists(resolved))
            {
                return false;
            }

            file = new SubmissionFile { Path = path, Sha256 = hash, SizeBytes = new FileInfo(resolved).Length };
            return true;
        }

        // ###########################################################################################
        // BETA's board as the table was opened on it, in one string: the SHA-256 of its rows in the
        // review API's own JSON. Computed only here, on both reads, so the same board gives the same
        // fingerprint whatever the JSON looks like; the client never works it out, only sends it back.
        // ###########################################################################################
        public static string Fingerprint(SubmissionRows rows)
        {
            ArgumentNullException.ThrowIfNull(rows);

            byte[] json = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(rows, ReviewApiContract.WireSettings));

            return Convert.ToHexStringLower(SHA256.HashData(json));
        }

        // *** NOT "its first submission is under the queue" (owner report, 2026-10-04). *** It said so
        // for every board BETA lacks - also one whose only submission was turned down long ago. Where
        // a board is - its submission, BETA, the stable source - is the Maintainer tab's stage line.
        public const string NotInBetaMessage =
            "Nothing of this board is in BETA, so there is no board to show.";

        public const string NotInStableMessage =
            "Nothing of this board is in the stable source, so there is no board to show.";

        public const string NoStableSourceMessage =
            "This server has no stable source to look in.";

        public const string StableReadOnlyMessage =
            "This is the stable source, as everybody using CRT gets it. Only a publish from BETA changes it - make a change on BETA's table.";

        public const string NotYoursMessage =
            "Only this board's maintainers and the administrator can change it. You can look, but not save.";

        public const string ClosedMessage =
            "This board is closed to contributions at the moment, so no change can be made to it.";

        public const string ChangedSinceMessage =
            "This board's data in BETA changed after you opened the table, so your change was not published - it would have undone that change. Open the board again, then make your change.";

        public const string NothingChangedMessage =
            "The table holds no change to publish.";

        public const string NoReasonMessage =
            "Give a reason for the change - it goes with it into BETA, " + MaintainerScreenWording.BetaQueueQuoted + " and the board's history.";

        public const string RemovalsChangedMessage =
            "The files this change would remove from BETA have changed since you were shown them - another publish has " +
            "changed what uses them. Save again to see the list as it is now.";
    }

    // What reading a board's table answered, or why not. The flags map onto status codes.
    public sealed record BoardTableOutcome(BoardTableAnswer? Answer, string Error, bool IsForbidden = false, bool IsNotFound = false)
    {
        public static BoardTableOutcome Read(BoardTableAnswer answer) => new(answer, string.Empty);

        public static BoardTableOutcome Forbidden(string error) => new(null, error, IsForbidden: true);

        public static BoardTableOutcome NotFound(string error) => new(null, error, IsNotFound: true);
    }

    // ###########################################################################################
    // What checking or sending an edit did, or why not. The flags map onto status codes in the
    // endpoint.
    //
    //   Checked           - the check passed; Removals is what a publish would remove.
    //   Published         - in BETA at Revision; Removals is what it removed.
    //   SavedNotPublished - made into a submission that the approval then did not publish
    //                       (NotPublishedReason); it waits under Contributor Submissions.
    // ###########################################################################################
    public sealed record BoardEditOutcome(
        bool IsAccepted,
        long SubmissionId,
        string Error,
        IReadOnlyList<ValidationFinding> Findings,
        bool IsForbidden = false,
        bool IsNotFound = false,
        bool IsConflict = false)
    {
        public bool IsPublished { get; init; }

        public string? Revision { get; init; }

        public IReadOnlyList<string> Removals { get; init; } = [];

        public string? NotPublishedReason { get; init; }

        public static BoardEditOutcome Checked(IReadOnlyList<string> removals) =>
            new(true, 0, string.Empty, []) { Removals = removals };

        public static BoardEditOutcome Published(
            long submissionId, string revision, IReadOnlyList<string> removed, IReadOnlyList<ValidationFinding> warnings) =>
            new(true, submissionId, string.Empty, warnings) { IsPublished = true, Revision = revision, Removals = removed };

        public static BoardEditOutcome SavedNotPublished(long submissionId, string reason, IReadOnlyList<ValidationFinding> warnings) =>
            new(true, submissionId, string.Empty, warnings) { NotPublishedReason = reason };

        public static BoardEditOutcome Refused(string error, IReadOnlyList<ValidationFinding> findings) =>
            new(false, 0, error, findings);

        public static BoardEditOutcome Forbidden(string error) => new(false, 0, error, [], IsForbidden: true);

        public static BoardEditOutcome NotFound(string error) => new(false, 0, error, [], IsNotFound: true);

        public static BoardEditOutcome Conflict(string error) => new(false, 0, error, [], IsConflict: true);

        // The refusal as ONE sentence - the Maintainer tab shows a refusal's `error` and nothing
        // else - the same shape AmendOutcome.FullError gives.
        public string FullError =>
            string.Join(" ", new[] { this.Error }
                .Concat(this.Findings.Where(finding => finding.Severity == ValidationSeverity.Error).Select(finding => finding.Message))
                .Where(part => !string.IsNullOrWhiteSpace(part)));
    }
}
