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

        // ###########################################################################################
        // A BETA rollback's bookkeeping (2026-09-27). Like the real store: only a row still MERGED
        // moves, its approvals go, and a system with no row is refused rather than invented.
        //
        // FailRollbackRecord makes it throw BEFORE changing anything - the real store's transaction
        // rolls back whole - so a test can prove the flow survives a failed record and that pushing
        // back again finishes the job.
        // ###########################################################################################
        public Task RecordRollbackAsync(
            string systemId,
            IReadOnlyList<long> returningSubmissionIds,
            long decidedByAccountId,
            string comment,
            string? betaRevision,
            string? betaContentHash,
            DateTimeOffset decidedUtc,
            bool reject = false,
            CancellationToken cancellationToken = default)
        {
            if (this.FailRollbackRecord)
                throw new InvalidOperationException("The database is not reachable.");

            if (!this.Systems.ContainsKey(systemId))
            {
                throw new InvalidOperationException(
                    $"No `systems` row exists for [{systemId}], so a rollback could not be recorded.");
            }

            foreach (long id in returningSubmissionIds)
            {
                if (!this.Submissions.TryGetValue(id, out SubmissionRecord? row) || row.State != SubmissionState.Merged)
                    continue;

                string state = reject ? SubmissionState.Rejected : SubmissionState.Pending;

                this.Submissions[id] = row with
                {
                    State = state,
                    DecidedUtc = decidedUtc,
                    DecisionComment = comment
                };

                // submission_beta_returns (migration 0015): returned to the queue, at this instant.
                if (!reject)
                    this.BetaReturns[id] = decidedUtc;

                this.Decisions[id] = new RecordedDecision(state, decidedByAccountId, comment, decidedUtc);
                this.Approvals.Remove(id);
            }

            this.BetaStates[systemId] = (betaRevision, betaContentHash);

            // The real store writes `systems.current_revision` / `content_hash`, which FindSystemAsync
            // reads back - so the fake moves the same row, or a caller re-reading the system would
            // still see the old BETA state.
            if (betaRevision is null && betaContentHash is null)
                this.PublishedSystems.Remove(systemId);
            else
                this.PublishedSystems[systemId] = new PublishedSystemRow(betaRevision ?? string.Empty, betaContentHash ?? string.Empty, default);

            return Task.CompletedTask;
        }

        public bool FailRollbackRecord { get; set; }

        // What RecordRollbackAsync wrote for each system, so a test can assert the recorded BETA
        // state followed the tree rather than still naming data that was removed.
        public Dictionary<string, (string? Revision, string? ContentHash)> BetaStates { get; } = new(StringComparer.Ordinal);

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
            this.ListSystemsCalls++;

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

        // How many times each way of reading systems was asked - for a test that a caller reads the
        // list ONCE rather than a system at a time.
        public int FindSystemCalls { get; private set; }

        public int ListSystemsCalls { get; private set; }

        public Task<SystemRecord?> FindSystemAsync(string systemId, CancellationToken cancellationToken = default)
        {
            this.FindSystemCalls++;

            return Task.FromResult(
                this.Systems.TryGetValue(systemId, out NewSubmission? system) ? this.ToSystemRecord(system) : null);
        }

        // ###########################################################################################
        // The real DELETE and its ON DELETE CASCADE: the system's row and everything the schema
        // hangs off it - every submission of it, whatever its state, with all of their rows. A
        // submission of ANOTHER system is untouched, which is what a test of a delete must be able
        // to see. (The maintainer pool and invitations live in FakeAccountStore, as their tables'
        // owner does; the cascade empties them in MariaDB.)
        // ###########################################################################################
        public Task<IReadOnlyList<long>> DeleteSystemAsync(string systemId, CancellationToken cancellationToken = default)
        {
            this.Systems.Remove(systemId);

            List<long> deleted = this.Submissions.Values
                .Where(record => string.Equals(record.SystemId, systemId, StringComparison.Ordinal))
                .Select(record => record.Id)
                .ToList();

            foreach (long id in deleted)
            {
                this.Submissions.Remove(id);
                this.Files.Remove(id);
                this.Payloads.Remove(id);
                this.Findings.Remove(id);
                this.Created.Remove(id);
                this.Decisions.Remove(id);
                this.Approvals.Remove(id);
                this.BetaReturns.Remove(id);
                this.DraftDiscards.Remove(id);
                this.Changes.Remove(id);
                this.Amendments.RemoveAll(amendment => amendment.SubmissionId == id);
            }

            this.ClosedSystems.Remove(systemId);
            this.BetaStates.Remove(systemId);
            this.PublishedSystems.Remove(systemId);
            this.ProductionSystems.Remove(systemId);
            this.Placements.Remove(systemId);

            foreach ((string SystemId, string Hash) key in this.ProductionApprovals.Keys.Where(key => key.SystemId == systemId).ToList())
                this.ProductionApprovals.Remove(key);

            this.AfterSystemDeleted?.Invoke();

            return Task.FromResult<IReadOnlyList<long>>(deleted);
        }

        // Run once a system's row is deleted - a test cancels the request there, as a client giving
        // up after the delete would.
        public Action? AfterSystemDeleted { get; set; }

        // ---- Placements (migration 0011): an UPDATE of the system's row, so a system with no row
        // is refused (false) rather than invented - as the real store's matched-rows count says.
        public Dictionary<string, SystemPlacement> Placements { get; } = new(StringComparer.Ordinal);

        public Task<SystemPlacement?> GetPlacementAsync(string systemId, CancellationToken cancellationToken = default) =>
            Task.FromResult(this.Placements.TryGetValue(systemId, out SystemPlacement? placement) ? placement : null);

        public Task<bool> SetPlacementAsync(
            string systemId,
            SystemPlacement placement,
            long setByAccountId,
            DateTimeOffset setUtc,
            CancellationToken cancellationToken = default)
        {
            if (!this.Systems.ContainsKey(systemId))
                return Task.FromResult(false);

            this.Placements[systemId] = placement;
            return Task.FromResult(true);
        }

        // submission_notes (migration 0019): read off what each submission was created with, joined
        // to its state NOW - as the real store's JOIN does - newest first.
        public Task<IReadOnlyList<SubmissionNotes>> GetHardwareNotesAsync(string systemId, CancellationToken cancellationToken = default) =>
            Task.FromResult<IReadOnlyList<SubmissionNotes>>(this.Created
                .Where(pair => string.Equals(pair.Value.SystemId, systemId, StringComparison.Ordinal) &&
                    !string.IsNullOrWhiteSpace(pair.Value.HardwareNotes) &&
                    this.Submissions.ContainsKey(pair.Key))
                .Select(pair => new SubmissionNotes(pair.Key, this.Submissions[pair.Key].State, this.Submissions[pair.Key].CreatedUtc, pair.Value.HardwareNotes!))
                .OrderByDescending(notes => notes.CreatedUtc)
                .ThenByDescending(notes => notes.SubmissionId)
                .ToList());

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

        // Pending only, this system only (exact, as the BINARY column compares), oldest first.
        public Task<IReadOnlyList<SubmissionRecord>> GetPendingForSystemAsync(string systemId, CancellationToken cancellationToken = default)
        {
            IReadOnlyList<SubmissionRecord> records = this.Submissions.Values
                .Where(record => record.State == SubmissionState.Pending && string.Equals(record.SystemId, systemId, StringComparison.Ordinal))
                .OrderBy(record => record.Id)
                .ToList();

            return Task.FromResult(records);
        }

        // The real query: this system exactly, never 'uploading' or 'abandoned', newest first, at
        // most `limit`. A maintainer decided it when SetDecisionAsync recorded it, as below.
        public Task<IReadOnlyList<SystemSubmissionRecord>> GetSubmissionsForSystemAsync(
            string systemId,
            int limit,
            CancellationToken cancellationToken = default)
        {
            IReadOnlyList<SystemSubmissionRecord> records = this.Submissions.Values
                .Where(record => string.Equals(record.SystemId, systemId, StringComparison.Ordinal))
                .Where(record => record.State is not (SubmissionState.Uploading or SubmissionState.Abandoned))
                .OrderByDescending(record => record.Id)
                .Take(Math.Clamp(limit, 1, 1000))
                .Select(record => new SystemSubmissionRecord(
                    record,
                    this.Decisions.ContainsKey(record.Id),
                    this.Decisions.TryGetValue(record.Id, out var decision) ? decision.DecidedByAccountId : null))
                .ToList();

            return Task.FromResult(records);
        }

        // The real query's two shapes: by account, or by email (trimmed, any case) among the
        // submissions sent without one. A maintainer decided it when SetDecisionAsync recorded it -
        // the automatic checks' rejection goes through SetStateAsync, as in the real store.
        public Task<IReadOnlyList<ContributorSubmission>> GetContributorSubmissionsAsync(
            long? accountId,
            string? contactEmail,
            CancellationToken cancellationToken = default)
        {
            string email = contactEmail?.Trim() ?? string.Empty;

            IReadOnlyList<ContributorSubmission> submissions = this.Submissions.Values
                .Where(record => accountId is not null
                    ? record.AccountId == accountId
                    : record.AccountId is null && email.Length > 0 &&
                      string.Equals(record.ContactEmail?.Trim(), email, StringComparison.OrdinalIgnoreCase))
                .Select(record => new ContributorSubmission(
                    record.Id,
                    record.State,
                    this.Decisions.ContainsKey(record.Id),
                    record.SystemId,
                    record.Summary,
                    record.CreatedUtc,
                    record.DecidedUtc,
                    record.DecisionComment))
                .ToList();

            return Task.FromResult(submissions);
        }

        // The real store's two checks - still pending, never amended - and its write: withdrawn,
        // decided now by nobody, with the comment on the row for the contributor to read.
        public Task<bool> WithdrawReplacedAsync(
            long submissionId,
            string comment,
            DateTimeOffset whenUtc,
            CancellationToken cancellationToken = default)
        {
            if (!this.Submissions.TryGetValue(submissionId, out SubmissionRecord? row) ||
                row.State != SubmissionState.Pending ||
                this.Amendments.Any(amendment => amendment.SubmissionId == submissionId))
            {
                return Task.FromResult(false);
            }

            this.Submissions[submissionId] = row with
            {
                State = SubmissionState.Withdrawn,
                DecidedUtc = whenUtc,
                DecisionComment = comment
            };

            return Task.FromResult(true);
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

        public Task SetTouchesSharedFilesAsync(long submissionId, bool touchesSharedFiles, CancellationToken cancellationToken = default)
        {
            if (this.Submissions.TryGetValue(submissionId, out SubmissionRecord? record))
                this.Submissions[submissionId] = record with { TouchesSharedFiles = touchesSharedFiles };

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

        // How many times each of the list's reads was asked - so a test can pin that the Beta > Prod
        // list asks a fixed number of times however many systems wait (code review, 2026-09-29).
        public int ProductionApprovalReads { get; private set; }

        public int MergedSubmissionReads { get; private set; }

        public int DraftDiscardReads { get; private set; }

        public Task<IReadOnlyDictionary<string, IReadOnlyList<GivenApproval>>> GetProductionApprovalsForAsync(
            IReadOnlyCollection<(string SystemId, string BetaContentHash)> states,
            CancellationToken cancellationToken = default)
        {
            this.ProductionApprovalReads++;

            IReadOnlyDictionary<string, IReadOnlyList<GivenApproval>> found = states
                .Distinct()
                .Where(this.ProductionApprovals.ContainsKey)
                .ToDictionary(state => state.SystemId, state => (IReadOnlyList<GivenApproval>)this.ProductionApprovals[state].ToList(), StringComparer.Ordinal);

            return Task.FromResult(found);
        }

        public Task<IReadOnlySet<string>> GetSystemsCarryingDiscardedDraftsAsync(
            IReadOnlyCollection<(string SystemId, DateTimeOffset? DecidedAfter)> windows,
            DateTimeOffset decidedUpTo,
            CancellationToken cancellationToken = default)
        {
            this.DraftDiscardReads++;

            IReadOnlySet<string> systems = windows
                .Where(window => this.Submissions.Values.Any(record =>
                    string.Equals(record.SystemId, window.SystemId, StringComparison.Ordinal) &&
                    record.State == SubmissionState.Merged &&
                    record.DecidedUtc is not null && record.DecidedUtc <= decidedUpTo &&
                    (window.DecidedAfter is null || record.DecidedUtc > window.DecidedAfter) &&
                    this.DraftDiscards.ContainsKey(record.Id)))
                .Select(window => window.SystemId)
                .ToHashSet(StringComparer.Ordinal);

            return Task.FromResult(systems);
        }

        public Task<IReadOnlyList<GivenApproval>> GetProductionApprovalsAsync(
            string systemId,
            string betaContentHash,
            CancellationToken cancellationToken = default)
        {
            this.ProductionApprovalReads++;

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
            this.MergedSubmissionReads++;

            IReadOnlyList<SubmissionRecord> records = this.Submissions.Values
                .Where(record => string.Equals(record.SystemId, systemId, StringComparison.Ordinal))
                .Where(record => record.State == SubmissionState.Merged)
                .Where(record => record.DecidedUtc is not null && record.DecidedUtc <= decidedUpTo)
                .Where(record => decidedAfter is null || record.DecidedUtc > decidedAfter)
                .OrderBy(record => record.Id)
                .ToList();

            return Task.FromResult(records);
        }

        // submission_beta_returns (migration 0015): when a rollback last returned each submission.
        public Dictionary<long, DateTimeOffset> BetaReturns { get; } = [];

        public Task<IReadOnlyDictionary<long, DateTimeOffset>> GetBetaReturnsAsync(
            IReadOnlyCollection<long> submissionIds,
            CancellationToken cancellationToken = default)
        {
            IReadOnlyDictionary<long, DateTimeOffset> found = submissionIds
                .Distinct()
                .Where(this.BetaReturns.ContainsKey)
                .ToDictionary(id => id, id => this.BetaReturns[id]);

            return Task.FromResult(found);
        }

        // submission_changes (migration 0017): what each publish changed; a later one replaces it.
        public Dictionary<long, SubmissionChanges> Changes { get; } = [];

        public Task SetChangesAsync(
            long submissionId,
            SubmissionChanges changes,
            DateTimeOffset recordedUtc,
            CancellationToken cancellationToken = default)
        {
            this.Changes[submissionId] = changes;
            return Task.CompletedTask;
        }

        public Task<IReadOnlyDictionary<long, SubmissionChanges>> GetChangesAsync(
            IReadOnlyCollection<long> submissionIds,
            CancellationToken cancellationToken = default)
        {
            IReadOnlyDictionary<long, SubmissionChanges> found = submissionIds
                .Distinct()
                .Where(this.Changes.ContainsKey)
                .ToDictionary(id => id, id => this.Changes[id]);

            return Task.FromResult(found);
        }

        // submission_draft_discards (migration 0014): the first time per submission is kept.
        public Dictionary<long, DateTimeOffset> DraftDiscards { get; } = [];

        public Task<bool> RecordDraftDiscardedAsync(
            long submissionId,
            DateTimeOffset discardedUtc,
            CancellationToken cancellationToken = default) =>
            Task.FromResult(this.DraftDiscards.TryAdd(submissionId, discardedUtc));

        public Task<IReadOnlyDictionary<long, DateTimeOffset>> GetDraftDiscardsAsync(
            IReadOnlyCollection<long> submissionIds,
            CancellationToken cancellationToken = default)
        {
            this.DraftDiscardReads++;

            IReadOnlyDictionary<long, DateTimeOffset> found = submissionIds
                .Distinct()
                .Where(this.DraftDiscards.ContainsKey)
                .ToDictionary(id => id, id => this.DraftDiscards[id]);

            return Task.FromResult(found);
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
