using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;

namespace CRT.Server.Tests.Fakes
{
    // ###########################################################################################
    // An in-memory ISubmissionStore, so the submission flows can be tested with no database. The
    // same role FakeAccountStore plays for accounts.
    //
    // IT BEHAVES LIKE THE REAL SCHEMA WHERE THAT MATTERS. In particular MarkUploadedAsync marks
    // every file carrying the hash, not just one, because a blob can legitimately appear at
    // several paths and uploading it once must satisfy all of them - a fake that marked one row
    // would let a test pass while the real finalise step still reported files missing.
    // ###########################################################################################
    public sealed class FakeSubmissionStore : ISubmissionStore
    {
        private long thisNextId = 1;

        public Dictionary<long, SubmissionRecord> Submissions { get; } = [];

        public Dictionary<long, List<SubmissionFileRecord>> Files { get; } = [];

        public Dictionary<long, SubmissionManifest> Payloads { get; } = [];

        public Dictionary<long, List<ValidationFinding>> Findings { get; } = [];

        // ###########################################################################################
        // The `systems` rows a submission implies, keyed by system id.
        //
        // THIS EXISTS BECAUSE ITS ABSENCE HID A REAL BUG. submissions.system_id is NOT NULL with a
        // foreign key to systems(system_id), and nothing ever wrote to `systems` - so every real
        // submission would have failed against MariaDB. Every test passed, because this fake
        // silently accepted a submission for a system that did not exist and threw the name parts
        // away.
        //
        // A fake that is more permissive than the real store is a fake that certifies bugs. This
        // one now records what the real store would have to insert, so the tests below can assert
        // it happens.
        // ###########################################################################################
        public Dictionary<string, NewSubmission> Systems { get; } = new(StringComparer.Ordinal);

        // Every NewSubmission as it was created, keyed by id - the address and byte count the rate
        // limit reads back, exactly as the real store's created_ip and bytes_to_upload columns do.
        public Dictionary<long, NewSubmission> Created { get; } = [];

        // Systems the administrator has closed (is_accepting = 0).
        public HashSet<string> ClosedSystems { get; } = new(StringComparer.Ordinal);

        public Task<long> CreateAsync(NewSubmission submission, CancellationToken cancellationToken = default)
        {
            long id = this.thisNextId++;

            this.Created[id] = submission;

            // INSERT IGNORE in the real store: an existing system is left completely alone, so a
            // second submission cannot rewrite its origin or its created date.
            if (!this.Systems.ContainsKey(submission.SystemId))
                this.Systems[submission.SystemId] = submission;

            this.Submissions[id] = new SubmissionRecord(
                id,
                submission.SystemId,
                submission.AccountId,
                submission.ContactEmail,
                submission.UploadTokenHash,
                submission.BaseRevision,
                SubmissionState.Uploading,
                submission.Summary,
                submission.FormatVersion,
                submission.CreatedUtc,
                submission.ExpiresUtc,
                DecidedUtc: null,
                TouchesSharedFiles: submission.TouchesSharedFiles);

            long fileId = 1;

            this.Files[id] = submission.Files
                .Select(file => new SubmissionFileRecord(
                    fileId++, file.Path, file.Sha256, file.SizeBytes, IsUploaded: false))
                .ToList();

            return Task.FromResult(id);
        }

        public Task<SubmissionRecord?> FindAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            return Task.FromResult(
                this.Submissions.TryGetValue(submissionId, out SubmissionRecord? record) ? record : null);
        }

        public Task<IReadOnlyList<SubmissionFileRecord>> GetFilesAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            IReadOnlyList<SubmissionFileRecord> files = this.Files.TryGetValue(submissionId, out List<SubmissionFileRecord>? list)
                ? list
                : [];

            return Task.FromResult(files);
        }

        public Task MarkUploadedAsync(long submissionId, string sha256, CancellationToken cancellationToken = default)
        {
            if (!this.Files.TryGetValue(submissionId, out List<SubmissionFileRecord>? files))
                return Task.CompletedTask;

            // EVERY file with this hash - see the class header.
            for (int index = 0; index < files.Count; index++)
            {
                if (string.Equals(files[index].Sha256, sha256, StringComparison.Ordinal))
                    files[index] = files[index] with { IsUploaded = true };
            }

            return Task.CompletedTask;
        }

        public Task SetStateAsync(long submissionId, string state, DateTimeOffset whenUtc, CancellationToken cancellationToken = default)
        {
            if (this.Submissions.TryGetValue(submissionId, out SubmissionRecord? record))
                this.Submissions[submissionId] = record with { State = state, DecidedUtc = whenUtc };

            return Task.CompletedTask;
        }

        // ###########################################################################################
        // *** THE PAYLOAD IS RECONSTRUCTED, NOT HANDED BACK. ***
        //
        // This fake used to store the manifest OBJECT and return that same instance, so every field
        // survived the round trip for free. The real store does not work that way: it writes only
        // the rows and renames as JSON, and rebuilds everything else from the `submissions` and
        // `systems` tables.
        //
        // It rebuilt everything EXCEPT Manufacturer/Hardware/Board, which came back as empty
        // strings - so the finalise-time validation pass, which runs against the RELOADED manifest,
        // rejected every submission that ever reached it with "does not name the hardware", "does
        // not name the board" and a consequent id mismatch. The findings blamed the client, and
        // three live submissions were spent chasing it there.
        //
        // Every test passed throughout, because this fake never lost anything. So it now models the
        // real store's actual behaviour: keep the rows, rebuild the rest from the submission and
        // system records. A fake that is kinder than the thing it stands in for certifies bugs.
        // ###########################################################################################
        public Task SavePayloadAsync(long submissionId, SubmissionManifest manifest, CancellationToken cancellationToken = default)
        {
            this.Payloads[submissionId] = manifest;
            return Task.CompletedTask;
        }

        public Task<SubmissionManifest?> LoadPayloadAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            if (!this.Payloads.TryGetValue(submissionId, out SubmissionManifest? stored))
                return Task.FromResult<SubmissionManifest?>(null);

            if (!this.Submissions.TryGetValue(submissionId, out SubmissionRecord? record))
                return Task.FromResult<SubmissionManifest?>(null);

            // The three name parts come from the SYSTEM row, exactly as the real store's join does.
            this.Systems.TryGetValue(record.SystemId, out NewSubmission? system);

            var rebuilt = new SubmissionManifest
            {
                FormatVersion = stored.FormatVersion,
                SystemId = record.SystemId,
                Manufacturer = system?.Manufacturer ?? string.Empty,
                Hardware = system?.Hardware ?? string.Empty,
                Board = system?.Board ?? string.Empty,
                BaseRevision = record.BaseRevision,
                Summary = record.Summary ?? string.Empty,
                ContactEmail = record.ContactEmail ?? string.Empty,
                CreatedUtc = record.CreatedUtc,
                Rows = stored.Rows,
                Renames = stored.Renames,
                Files = (this.Files.TryGetValue(submissionId, out List<SubmissionFileRecord>? files) ? files : [])
                    .Select(file => new SubmissionFile
                    {
                        Path = file.Path,
                        Sha256 = file.Sha256,
                        SizeBytes = file.SizeBytes
                    })
                    .ToList()
            };

            return Task.FromResult<SubmissionManifest?>(rebuilt);
        }

        public Task SaveFindingsAsync(long submissionId, IReadOnlyList<ValidationFinding> findings, CancellationToken cancellationToken = default)
        {
            this.Findings[submissionId] = findings.ToList();
            return Task.CompletedTask;
        }

        public Task<IReadOnlyList<ValidationFinding>> GetFindingsAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            IReadOnlyList<ValidationFinding> findings = this.Findings.TryGetValue(submissionId, out List<ValidationFinding>? list)
                ? list
                : [];

            return Task.FromResult(findings);
        }

        public Task<IReadOnlyList<long>> GetExpiredUploadsAsync(DateTimeOffset now, CancellationToken cancellationToken = default)
        {
            IReadOnlyList<long> expired = this.Submissions.Values
                .Where(record => record.State == SubmissionState.Uploading)
                .Where(record => record.ExpiresUtc is not null && record.ExpiresUtc <= now)
                .Select(record => record.Id)
                .ToList();

            return Task.FromResult(expired);
        }

        public Task<IReadOnlyList<SubmissionRecord>> GetForAccountAsync(long accountId, int limit, CancellationToken cancellationToken = default)
        {
            IReadOnlyList<SubmissionRecord> records = this.Submissions.Values
                .Where(record => record.AccountId == accountId)
                .OrderByDescending(record => record.Id)
                .Take(limit)
                .ToList();

            return Task.FromResult(records);
        }

        // ###########################################################################################
        // The review queue. Mirrors the real query exactly: only 'pending', oldest FIRST.
        //
        // The ordering is modelled rather than left to dictionary order on purpose - it is a real
        // behavioural decision (a work queue is worked from the front, so the longest-waiting
        // submission leads), and a fake that returned insertion order would let a test pass while
        // the real store ordered differently.
        // ###########################################################################################
        public Task<IReadOnlyList<SubmissionRecord>> GetQueueAsync(int limit, CancellationToken cancellationToken = default)
        {
            // 'approved' too: a shared-file change with one of its two approvals is still waiting.
            IReadOnlyList<SubmissionRecord> records = this.Submissions.Values
                .Where(record => record.State == SubmissionState.Pending || record.State == SubmissionState.Approved)
                .OrderBy(record => record.Id)
                .Take(Math.Clamp(limit, 1, 200))
                .ToList();

            return Task.FromResult(records);
        }

        // ###########################################################################################
        // Records a publish. Modelled on the REAL store's behaviour rather than on convenience -
        // see this class's own header on why a permissive fake certifies bugs.
        //
        // In particular it REFUSES a system that was never registered, because the real statement
        // is an UPDATE against a row CreateAsync inserts. A fake that happily invented the row
        // would hide a caller publishing a system with no `systems` entry, which against MariaDB
        // updates nothing at all and silently leaves current_revision NULL.
        // ###########################################################################################
        public Task SetSystemPublishedAsync(
            string systemId,
            string revision,
            string contentHash,
            DateTimeOffset publishedUtc,
            CancellationToken cancellationToken = default)
        {
            if (!this.Systems.ContainsKey(systemId))
            {
                throw new InvalidOperationException(
                    $"No `systems` row exists for [{systemId}], so a publish could not be recorded.");
            }

            this.PublishedSystems[systemId] = new PublishedSystemRow(revision, contentHash, publishedUtc);

            return Task.CompletedTask;
        }

        // What SetSystemPublishedAsync wrote, so tests can assert the revision and content hash
        // actually reached the systems row.
        public Dictionary<string, PublishedSystemRow> PublishedSystems { get; } = new(StringComparer.Ordinal);

        // ###########################################################################################
        // Records a decision. STRICTER than convenient, for the same reason the publish method is:
        // the real statement is an UPDATE against an existing row, so a fake that invented one
        // would hide a caller deciding a submission that does not exist - which against MariaDB
        // updates nothing at all and reports success.
        // ###########################################################################################
        public Task SetDecisionAsync(
            long submissionId,
            string state,
            long decidedByAccountId,
            string? comment,
            DateTimeOffset decidedUtc,
            CancellationToken cancellationToken = default)
        {
            if (!this.Submissions.TryGetValue(submissionId, out SubmissionRecord? row))
            {
                throw new InvalidOperationException(
                    $"No submission [{submissionId}] exists, so a decision could not be recorded.");
            }

            // The COMMENT lands on the row too, not only in the audit dictionary - the contributor
            // reads it back through FindAsync, exactly as the real store's SELECT returns it.
            this.Submissions[submissionId] = row with
            {
                State = state,
                DecidedUtc = decidedUtc,
                DecisionComment = string.IsNullOrWhiteSpace(comment) ? null : comment
            };

            this.Decisions[submissionId] =
                new RecordedDecision(state, decidedByAccountId, comment, decidedUtc);

            return Task.CompletedTask;
        }

        // What SetDecisionAsync wrote, so tests can assert WHO decided and WHY reached the row -
        // not just that the state moved. The comment is the contributor's entire feedback channel,
        // so a test that only checked the state would pass while it was being dropped.
        public Dictionary<long, RecordedDecision> Decisions { get; } = new();

        // The same WHERE as the real query: this address, strictly after `since`.
        public Task<IReadOnlyList<RecentSubmission>> GetRecentSubmissionsFromAddressAsync(
            string ipAddress,
            DateTimeOffset since,
            CancellationToken cancellationToken = default)
        {
            IReadOnlyList<RecentSubmission> recent = this.Created.Values
                .Where(created => string.Equals(created.CreatedIp, ipAddress, StringComparison.Ordinal))
                .Where(created => created.CreatedUtc > since)
                .Select(created => new RecentSubmission(created.CreatedUtc, created.BytesToUpload))
                .ToList();

            return Task.FromResult(recent);
        }

        // Null for a system with no row, like the real SELECT finding nothing.
        public Task<IReadOnlyList<SystemRecord>> ListSystemsAsync(CancellationToken cancellationToken = default)
        {
            IReadOnlyList<SystemRecord> systems = this.Systems.Values
                .Select(system => this.ToSystemRecord(system))
                .OrderBy(system => system.SystemId, StringComparer.Ordinal)
                .ToList();

            return Task.FromResult(systems);
        }

        // The row as the real SELECT assembles it: the BETA half from PublishedSystems, the
        // Production half from ProductionSystems.
        private SystemRecord ToSystemRecord(NewSubmission system)
        {
            this.PublishedSystems.TryGetValue(system.SystemId, out PublishedSystemRow? beta);
            this.ProductionSystems.TryGetValue(system.SystemId, out PublishedSystemRow? production);

            return new SystemRecord(
                system.SystemId,
                system.Manufacturer,
                system.Hardware,
                system.Board,
                beta?.Revision,
                !this.ClosedSystems.Contains(system.SystemId),
                beta?.ContentHash,
                production?.Revision,
                production?.ContentHash,
                production?.PublishedUtc);
        }

        // What SetSystemInProductionAsync wrote.
        public Dictionary<string, PublishedSystemRow> ProductionSystems { get; } = new(StringComparer.Ordinal);

        public Task<SystemRecord?> FindSystemAsync(string systemId, CancellationToken cancellationToken = default)
        {
            return Task.FromResult(
                this.Systems.TryGetValue(systemId, out NewSubmission? system) ? this.ToSystemRecord(system) : null);
        }

        // An UPDATE in the real store, so - like SetSystemPublishedAsync - a system with no row is
        // refused rather than invented.
        public Task SetSystemInProductionAsync(
            string systemId,
            string? revision,
            string? contentHash,
            DateTimeOffset publishedUtc,
            CancellationToken cancellationToken = default)
        {
            if (!this.Systems.ContainsKey(systemId))
            {
                throw new InvalidOperationException(
                    $"No `systems` row exists for [{systemId}], so a production publish could not be recorded.");
            }

            this.ProductionSystems[systemId] = new PublishedSystemRow(revision ?? string.Empty, contentHash ?? string.Empty, publishedUtc);

            return Task.CompletedTask;
        }

        // ---- Approvals (migration 0008). First approval in a role wins, as INSERT IGNORE does.

        public Dictionary<long, List<GivenApproval>> Approvals { get; } = [];

        // ---- Amendments (migration 0009): what each one replaced, as the real table keeps it.

        public List<(long SubmissionId, int Version, SubmissionManifest Replaced, string By, DateTimeOffset AtUtc)> Amendments { get; } = [];

        // Runs at the start of AmendAsync, so a test can land another change in the gap between
        // the flow's own checks and the store's transaction - the race the store must close.
        public Func<Task>? BeforeAmend { get; set; }

        public async Task<AmendStoreResult> AmendAsync(
            long submissionId,
            int expectedVersion,
            SubmissionManifest amended,
            bool touchesSharedFiles,
            long? accountId,
            string accountLabel,
            DateTimeOffset amendedUtc,
            CancellationToken cancellationToken = default)
        {
            if (this.BeforeAmend is { } before)
                await before();

            // The real store's in-transaction checks: still amendable, still at the version opened.
            if (!SubmissionState.CanBeAmended(this.Submissions[submissionId].State))
                return new AmendStoreResult(AmendStoreOutcome.NotAmendable, 0);

            int latest = this.Amendments.Count(amendment => amendment.SubmissionId == submissionId);

            if (latest != expectedVersion)
                return new AmendStoreResult(AmendStoreOutcome.VersionChanged, latest);

            SubmissionManifest replaced = (await this.LoadPayloadAsync(submissionId, cancellationToken))!;
            int version = latest + 1;

            this.Amendments.Add((submissionId, version, replaced, accountLabel, amendedUtc));
            this.Payloads[submissionId] = amended;

            long fileId = 1;
            this.Files[submissionId] = amended.Files
                .Select(file => new SubmissionFileRecord(fileId++, file.Path, file.Sha256, file.SizeBytes, IsUploaded: true))
                .ToList();

            this.Approvals.Remove(submissionId);

            SubmissionRecord record = this.Submissions[submissionId];
            this.Submissions[submissionId] = record with
            {
                TouchesSharedFiles = touchesSharedFiles,
                State = record.State == SubmissionState.Approved ? SubmissionState.Pending : record.State
            };

            return new AmendStoreResult(AmendStoreOutcome.Amended, version);
        }

        public Task<SubmissionAmendment?> GetLatestAmendmentAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            var latest = this.Amendments
                .Where(amendment => amendment.SubmissionId == submissionId)
                .OrderByDescending(amendment => amendment.Version)
                .Select(amendment => new SubmissionAmendment(amendment.Version, amendment.By, amendment.AtUtc))
                .FirstOrDefault();

            return Task.FromResult(latest);
        }

        public Task MarkTouchesSharedFilesAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            if (this.Submissions.TryGetValue(submissionId, out SubmissionRecord? record))
                this.Submissions[submissionId] = record with { TouchesSharedFiles = true };

            return Task.CompletedTask;
        }

        public Dictionary<(string SystemId, string Hash), List<GivenApproval>> ProductionApprovals { get; } = [];

        public Task<IReadOnlyList<GivenApproval>> GetApprovalsAsync(long submissionId, CancellationToken cancellationToken = default)
        {
            IReadOnlyList<GivenApproval> approvals = this.Approvals.TryGetValue(submissionId, out List<GivenApproval>? list)
                ? list.ToList()
                : [];

            return Task.FromResult(approvals);
        }

        public Task AddApprovalAsync(
            long submissionId,
            ApproverRole role,
            long accountId,
            string accountLabel,
            DateTimeOffset approvedUtc,
            CancellationToken cancellationToken = default)
        {
            if (!this.Approvals.TryGetValue(submissionId, out List<GivenApproval>? list))
                this.Approvals[submissionId] = list = [];

            if (list.All(approval => approval.Role != role))
                list.Add(new GivenApproval(role, accountLabel, approvedUtc, accountId));

            return Task.CompletedTask;
        }

        public Task<IReadOnlyList<GivenApproval>> GetProductionApprovalsAsync(
            string systemId,
            string betaContentHash,
            CancellationToken cancellationToken = default)
        {
            IReadOnlyList<GivenApproval> approvals = this.ProductionApprovals.TryGetValue((systemId, betaContentHash), out List<GivenApproval>? list)
                ? list.ToList()
                : [];

            return Task.FromResult(approvals);
        }

        public Task AddProductionApprovalAsync(
            string systemId,
            string betaContentHash,
            ApproverRole role,
            long accountId,
            string accountLabel,
            DateTimeOffset approvedUtc,
            CancellationToken cancellationToken = default)
        {
            if (!this.ProductionApprovals.TryGetValue((systemId, betaContentHash), out List<GivenApproval>? list))
                this.ProductionApprovals[(systemId, betaContentHash)] = list = [];

            if (list.All(approval => approval.Role != role))
                list.Add(new GivenApproval(role, accountLabel, approvedUtc, accountId));

            return Task.CompletedTask;
        }

        public Task<IReadOnlyList<SubmissionRecord>> GetMergedSubmissionsAsync(
            string systemId,
            DateTimeOffset? decidedAfter,
            DateTimeOffset decidedUpTo,
            CancellationToken cancellationToken = default)
        {
            IReadOnlyList<SubmissionRecord> records = this.Submissions.Values
                .Where(record => string.Equals(record.SystemId, systemId, StringComparison.Ordinal))
                .Where(record => record.State == SubmissionState.Merged)
                .Where(record => record.DecidedUtc is not null && record.DecidedUtc <= decidedUpTo)
                .Where(record => decidedAfter is null || record.DecidedUtc > decidedAfter)
                .OrderBy(record => record.Id)
                .ToList();

            return Task.FromResult(records);
        }

        // INSERT IGNORE, like CreateAsync: an existing row is left completely alone.
        public Task EnsureSystemAsync(
            string systemId,
            string manufacturer,
            string hardware,
            string board,
            string origin,
            DateTimeOffset createdUtc,
            CancellationToken cancellationToken = default)
        {
            if (!this.Systems.ContainsKey(systemId))
            {
                this.Systems[systemId] = new NewSubmission(
                    systemId, manufacturer, hardware, board,
                    null, null, null, string.Empty, string.Empty, string.Empty,
                    0, [], createdUtc, createdUtc);
            }

            return Task.CompletedTask;
        }

        public Task<bool?> IsSystemAcceptingAsync(string systemId, CancellationToken cancellationToken = default)
        {
            bool? accepting = this.Systems.ContainsKey(systemId)
                ? !this.ClosedSystems.Contains(systemId)
                : null;

            return Task.FromResult(accepting);
        }

        // Driven from the SAME state list the real query binds, so the fake cannot keep a
        // different set of blobs alive than MariaDB would.
        public Task<IReadOnlySet<string>> GetLiveBlobHashesAsync(CancellationToken cancellationToken = default)
        {
            IReadOnlySet<string> live = this.Submissions.Values
                .Where(record => SubmissionCollectionStates.Live.Contains(record.State))
                .SelectMany(record => this.Files.TryGetValue(record.Id, out List<SubmissionFileRecord>? files) ? files : [])
                .Select(file => file.Sha256)
                .ToHashSet(StringComparer.Ordinal);

            return Task.FromResult(live);
        }

        // Removes the payload AND the file list; keeps the submission row and its findings, as the
        // real two-statement transaction does.
        public Task<int> DeleteRetiredPayloadsAsync(DateTimeOffset decidedBefore, CancellationToken cancellationToken = default)
        {
            List<long> retired = this.Submissions.Values
                .Where(record => SubmissionCollectionStates.Retired.Contains(record.State))
                .Where(record => (record.DecidedUtc ?? record.CreatedUtc) < decidedBefore)
                .Select(record => record.Id)
                .ToList();

            int cleared = 0;

            foreach (long id in retired)
            {
                if (this.Payloads.Remove(id))
                    cleared++;

                this.Files.Remove(id);
            }

            return Task.FromResult(cleared);
        }
    }

    public sealed record PublishedSystemRow(string Revision, string ContentHash, DateTimeOffset PublishedUtc);

    public sealed record RecordedDecision(
        string State,
        long DecidedByAccountId,
        string? Comment,
        DateTimeOffset DecidedUtc);
}
