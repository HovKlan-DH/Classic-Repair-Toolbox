using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A MAINTAINER'S AMENDMENT to a submission (owner request, 2026-09-25): the review
    // application's table editor lets a maintainer correct a submission's rows before publishing it -
    // "The maintainer should be able to also edit whatever, if he chooses to publish it afterwards."
    //
    // *** THE ORDER OF THE CHECKS IS THE DESIGN, as in ApprovePublishFlow. *** Cheap refusals
    // first, and nothing is stored until every one has passed. Everything from the first read to
    // the store runs under the PublishLock, so an amendment can never land in the middle of a
    // publish (code review, 2026-09-25):
    //
    //   1. AUTHORITY - may this account review anything, and THIS submission's board?
    //   2. STATE     - is it still undecided (pending, or waiting for its second approval)?
    //   3. VERSION   - is it still at the amendment the maintainer opened? Two maintainers editing at
    //                  once must not silently overwrite each other; the second is told to reload.
    //                  Steps 2 and 3 are checked AGAIN by the store, inside the transaction that
    //                  writes - the answer here is only the early one.
    //   4. ROWS      - only the table's nine sheets are taken from the request
    //                  (SubmissionRowsBoard.WithTableSections); highlights, calibrations and the
    //                  revision date always stay as submitted.
    //   5. FILES     - rebuilt from what the rows cite. A file the submission carries is kept; a
    //                  file already PUBLISHED at that path is taken from the published tree (the
    //                  server's own copy, imported like a new submission's unchanged files); a
    //                  row citing anything else is refused - a maintainer edits rows, and cannot
    //                  bring in a file nobody has sent.
    //   6. VALIDATE  - the same path, file and row rules a new submission passes.
    //   7. STORE     - the store keeps what is replaced (the contributor's original first), clears
    //                  the approvals already given - they were given to other content - and
    //                  re-decides whether shared files change, so the two-person rule follows the
    //                  content actually being published.
    //
    // An audit row records who amended what. The contributor is told in the decision mail and in
    // CRT's "My submissions" (the status answer carries `amendedByMaintainer`).
    // ###########################################################################################
    public static class AmendSubmissionFlow
    {
        public const string AmendedAction = "submission.amended";

        public static async Task<AmendOutcome> AmendAsync(
            ReviewAccess? access,
            long submissionId,
            int expectedVersion,
            SubmissionRows? rows,
            string? dataTreeRoot,
            ISubmissionStore store,
            IAccountStore accounts,
            BlobStore blobs,
            PublishLock publishLock,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(accounts);
            ArgumentNullException.ThrowIfNull(blobs);
            ArgumentNullException.ThrowIfNull(publishLock);

            // ---- 1. Authority ------------------------------------------------------------------
            if (!ReviewAuthority.CanReviewAnything(access))
                return AmendOutcome.Forbidden(ReviewAuthority.DescribeRefusal(access, null));

            // *** UNDER THE PUBLISH LOCK, from the first read to the store (code review,
            // 2026-09-25). *** ApprovePublishFlow holds it from reading the submission to recording
            // it published. Without it an amendment could land in between: the tree got the rows as
            // they were, the database the amended ones, and the contributor was mailed that a
            // maintainer's change had been published. Now it waits for the publish, and then finds the
            // submission decided.
            using IDisposable gate = await publishLock.EnterAsync(cancellationToken).ConfigureAwait(false);

            SubmissionRecord? record = await store.FindAsync(submissionId, cancellationToken).ConfigureAwait(false);

            if (record is null)
                return AmendOutcome.NotFound();

            if (!ReviewAuthority.CanReview(access, record))
                return AmendOutcome.Forbidden(ReviewAuthority.DescribeRefusal(access, record));

            // ---- 2. State ----------------------------------------------------------------------
            if (!SubmissionState.CanBeAmended(record.State))
                return AmendOutcome.Conflict(AmendSubmissionFlow.NotAmendableMessage(record.State));

            // ---- 3. Version --------------------------------------------------------------------
            SubmissionAmendment? latest = await store.GetLatestAmendmentAsync(submissionId, cancellationToken).ConfigureAwait(false);
            int version = latest?.Version ?? 0;

            if (version != expectedVersion)
                return AmendOutcome.Conflict(AmendSubmissionFlow.ChangedSinceMessage(latest));

            if (rows is null)
                return AmendOutcome.Refused("The request carried no rows.", []);

            if (string.IsNullOrWhiteSpace(dataTreeRoot))
                return AmendOutcome.Refused("The server has no data tree configured.", []);

            SubmissionManifest? current = await store.LoadPayloadAsync(submissionId, cancellationToken).ConfigureAwait(false);

            if (current is null)
                return AmendOutcome.Refused("This submission's contents could not be loaded, so it cannot be changed.", []);

            // ---- 4. Rows -----------------------------------------------------------------------
            SubmissionRows amendedRows = SubmissionRowsBoard.WithTableSections(current.Rows ?? new SubmissionRows(), rows);

            // ---- 5. Files ----------------------------------------------------------------------
            PublishedTreeView? tree = PublishedTreeProbe.For(dataTreeRoot);
            var findings = new List<ValidationFinding>();
            var files = new List<SubmissionFile>();
            var imports = new List<(SubmissionFile File, string Source)>();

            foreach (string path in SubmissionManifestBuilder.CollectReferencedFiles(SubmissionRowsBoard.ToBoard(amendedRows)))
            {
                SubmissionFile? carried = current.Files.FirstOrDefault(file => string.Equals(file.Path, path, StringComparison.Ordinal));

                if (carried is not null)
                {
                    files.Add(carried);
                    continue;
                }

                string? hash = tree?.HashOf(path);

                if (hash is not null &&
                    SubmissionPathRules.TryResolve(dataTreeRoot, path, out string resolved, out _) &&
                    PublishPathSafety.FindLinkOnPath(dataTreeRoot, resolved, PublishPathSafety.IsLink) is null &&
                    File.Exists(resolved))
                {
                    var published = new SubmissionFile { Path = path, Sha256 = hash, SizeBytes = new FileInfo(resolved).Length };
                    files.Add(published);
                    imports.Add((published, resolved));
                    continue;
                }

                findings.Add(new ValidationFinding
                {
                    Severity = ValidationSeverity.Error,
                    Code = "amend.file_unknown",
                    Subject = path,
                    Message = $"[{path}] is neither in this submission nor in the published data. A change made " +
                        "in the Maintainer tab can only use files that are already there."
                });
            }

            // ###########################################################################################
            // *** THE SUBMISSION'S KiCad DATA SURVIVES AN AMENDMENT (2026-09-26). *** The loop above
            // rebuilds the file list from what the rows cite, and no row cites a KiCad file - so a
            // maintainer's table edit would silently strip the contributor's KiCad data from the
            // submission, and approving it would publish the board without its traces. Carried over
            // as they are: an amendment changes rows, never the KiCad folder.
            // ###########################################################################################
            foreach (SubmissionFile file in current.Files)
            {
                if (SubmissionKiCadFiles.IsSubmittable(current, file.Path) &&
                    !files.Any(kept => string.Equals(kept.Path, file.Path, StringComparison.Ordinal)))
                {
                    files.Add(file);
                }
            }

            var amended = new SubmissionManifest
            {
                FormatVersion = current.FormatVersion,
                BoardId = current.BoardId,
                Manufacturer = current.Manufacturer,
                Hardware = current.Hardware,
                Board = current.Board,
                BaseRevision = current.BaseRevision,
                Summary = current.Summary,
                ContactEmail = current.ContactEmail,
                CreatedUtc = current.CreatedUtc,
                Renames = current.Renames,
                Rows = amendedRows,
                Files = files
            };

            // ---- 6. Validate -------------------------------------------------------------------
            findings.AddRange(SubmissionPathRules.ValidateManifestPaths(amended, dataTreeRoot));
            findings.AddRange(SubmissionFileRules.ValidateManifestFiles(amended, tree));
            findings.AddRange(SubmissionValidator.Validate(amended, amended.Files.Select(file => file.Path).ToList()));

            if (!SubmissionValidator.CanBeQueued(findings))
                return AmendOutcome.Refused("The change cannot be saved:", findings);

            // ---- 7. Store ----------------------------------------------------------------------
            //
            // Inside the blob store's reference gate: a file imported here must not be collected
            // before the row naming it exists.
            //
            // The store checks the state and the version AGAIN, inside its transaction - steps 2
            // and 3 above are the cheap early answer, not the guard. A refusal there leaves any
            // file imported just now unreferenced in the blob store, where the collector finds it.
            AmendStoreResult stored;

            using (await blobs.EnterReferenceGateAsync(cancellationToken).ConfigureAwait(false))
            {
                foreach ((SubmissionFile file, string source) in imports)
                {
                    if (blobs.Contains(file.Sha256))
                        continue;

                    if (!await blobs.TryImportAsync(source, file.Sha256, cancellationToken).ConfigureAwait(false))
                    {
                        return AmendOutcome.Refused(
                            $"[{file.Path}] changed on the server while it was being read. Try saving again.", []);
                    }
                }

                stored = await store.AmendAsync(
                    submissionId,
                    expectedVersion,
                    amended,
                    SubmissionSharedFiles.TouchesSharedFiles(amended, tree),
                    access!.Account.Id,
                    ApprovePublishFlow.Label(access),
                    nowUtc,
                    cancellationToken).ConfigureAwait(false);
            }

            if (stored.Outcome == AmendStoreOutcome.NotAmendable)
            {
                SubmissionRecord? now = await store.FindAsync(submissionId, cancellationToken).ConfigureAwait(false);
                return AmendOutcome.Conflict(AmendSubmissionFlow.NotAmendableMessage(now?.State));
            }

            if (stored.Outcome == AmendStoreOutcome.VersionChanged)
            {
                SubmissionAmendment? theirs = await store.GetLatestAmendmentAsync(submissionId, cancellationToken).ConfigureAwait(false);
                return AmendOutcome.Conflict(AmendSubmissionFlow.ChangedSinceMessage(theirs));
            }

            int newVersion = stored.Version;

            // Warnings travel with the submission, as they do from creation.
            await store.SaveFindingsAsync(submissionId, findings, cancellationToken).ConfigureAwait(false);

            await accounts.WriteAuditAsync(
                new AuditEntry(
                    access.Account.Id,
                    access.Account.Email,
                    AmendSubmissionFlow.AmendedAction,
                    $"#{submissionId}",
                    $"{record.BoardId}: amendment {newVersion}, {amended.Files.Count} file(s)",
                    nowUtc),
                cancellationToken).ConfigureAwait(false);

            return AmendOutcome.Amended(newVersion, findings);
        }

        internal static string NotAmendableMessage(string? state) =>
            state == SubmissionState.Uploading
                ? "This submission is still uploading and cannot be changed yet."
                : $"This submission has already been decided ({state}) and cannot be changed.";

        internal static string ChangedSinceMessage(SubmissionAmendment? latest) =>
            latest is null
                ? "This submission has changed since you opened it. Open it again, then make your change."
                : $"{latest.By} changed this submission after you opened it. Open it again to see their change, then make yours.";
    }

    // What an amendment did, or why not. The flags map onto status codes in the endpoint.
    public sealed record AmendOutcome(
        bool IsAmended,
        int Version,
        string Error,
        IReadOnlyList<ValidationFinding> Findings,
        bool IsForbidden = false,
        bool IsNotFound = false,
        bool IsConflict = false)
    {
        public static AmendOutcome Amended(int version, IReadOnlyList<ValidationFinding> findings) =>
            new(true, version, string.Empty, findings);

        public static AmendOutcome Refused(string error, IReadOnlyList<ValidationFinding> findings) =>
            new(false, 0, error, findings);

        public static AmendOutcome Forbidden(string error) => new(false, 0, error, [], IsForbidden: true);

        public static AmendOutcome NotFound() => new(false, 0, "No such submission.", [], IsNotFound: true);

        public static AmendOutcome Conflict(string error) => new(false, 0, error, [], IsConflict: true);

        // The refusal as ONE sentence for the maintainer: the error, then what each blocking finding
        // says - the Maintainer tab shows a refusal's `error` and nothing else.
        public string FullError =>
            string.Join(" ", new[] { this.Error }
                .Concat(this.Findings.Where(finding => finding.Severity == ValidationSeverity.Error).Select(finding => finding.Message))
                .Where(part => !string.IsNullOrWhiteSpace(part)));
    }
}
