using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Create, negotiate, upload and finalise - the three-step transport from
    // NewContributeStrategy.md, expressed as decisions.
    //
    // Every method takes its collaborators as arguments (the store, the blob store, the clock), so
    // the whole file is testable against fakes with no database, no filesystem and no network.
    // The endpoints are a rim over these, the same split AccountFlows and AccountEndpoints use.
    //
    // THE FOUR PROPERTIES THIS FILE GUARANTEES:
    //
    // 1. NOTHING IS PUBLISHED HERE. A finalised submission is QUEUED. Promotion to the production
    //    tree stays a manual act by the maintainer - Phase 3 step 3's filesystem permissions make
    //    it impossible for this service to do otherwise, and this code must not pretend otherwise
    //    either.
    //
    // 2. VALIDATION HAPPENS BEFORE QUEUEING, NEVER AFTER. A submission with any error is rejected
    //    automatically with its reasons recorded, and never reaches a human. That is the
    //    highest-leverage rule in the plan.
    //
    // 3. A SUBMISSION IS ONLY EVER TOUCHED BY WHOEVER CREATED IT, proved by a CAPABILITY TOKEN.
    //    Contributing requires no account (see NewContributeStrategy.md, "CONTRIBUTING NEEDS NO
    //    ACCOUNT"), so there is no identity to check against - and a submission id is a small
    //    integer anyone can guess. Creating a submission therefore returns a random 256-bit token
    //    that every later call must present. Only its hash is stored. This is the authorisation
    //    boundary of the whole phase.
    //
    // 4. THE MANIFEST IS VALIDATED BEFORE ANY UPLOAD IS ACCEPTED. A contributor should be told
    //    their data is wrong before spending an hour uploading, not after.
    // ###########################################################################################
    public static class SubmissionFlows
    {
        // How long a client has to finish uploading before the submission is abandoned and its
        // partial blobs collected. Generous, because a large system over a domestic connection is
        // genuinely slow, and the cost of being wrong is a contributor losing their work.
        public static readonly TimeSpan UploadWindow = TimeSpan.FromHours(24);

        // ###########################################################################################
        // Step 1 and 2 together: accept a manifest, validate it, and answer with the hashes the
        // server does not already hold.
        //
        // VALIDATION RUNS FIRST, BEFORE A SUBMISSION ROW EXISTS. A manifest that cannot be
        // published is refused here, so a broken client cannot fill the table with rows that will
        // never be finalised - and the contributor learns immediately rather than after uploading.
        // ###########################################################################################
        public static async Task<SubmissionCreationOutcome> CreateAsync(
            SubmissionManifest manifest,
            Submitter submitter,
            string systemFolder,
            ISubmissionStore store,
            BlobStore blobs,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(manifest);
            ArgumentNullException.ThrowIfNull(submitter);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(blobs);

            var findings = new List<ValidationFinding>();

            // A contact address is the ONE thing a contributor must give. Not a credential -
            // nothing signs in with it - but without it a reviewer cannot say "accepted" or "this
            // needs changing", and a submission nobody can reply to can only be taken or dropped.
            if (!submitter.HasContact)
            {
                findings.Add(new ValidationFinding
                {
                    Severity = ValidationSeverity.Error,
                    Code = "contact.missing",
                    Subject = string.Empty,
                    Message =
                        "An email address is needed so you can be told whether your contribution " +
                        "was accepted. It is used for that and nothing else - there is no account " +
                        "to create and no password to choose."
                });
            }
            else if (submitter.AccountId is null && !AccountRules.IsPlausibleEmail(submitter.ContactEmail))
            {
                findings.Add(new ValidationFinding
                {
                    Severity = ValidationSeverity.Error,
                    Code = "contact.malformed",
                    Subject = submitter.ContactEmail ?? string.Empty,
                    Message = "That does not look like an email address."
                });
            }

            if (findings.Count > 0)
                return SubmissionCreationOutcome.Refused(findings);

            // Paths first: these are the untrusted values that decide where files get written, and
            // a manifest that fails here is refused before anything else looks at it.
            findings.AddRange(SubmissionPathRules.ValidateManifestPaths(manifest, systemFolder));

            // Then the rows, against the paths the manifest itself supplies.
            findings.AddRange(SubmissionValidator.Validate(
                manifest, manifest.Files.Select(file => file.Path).ToList()));

            if (!SubmissionValidator.CanBeQueued(findings))
                return SubmissionCreationOutcome.Refused(findings);

            // THE CAPABILITY TOKEN. With no account, a submission id is a small integer anyone can
            // guess, so it cannot be what proves ownership. This random 256-bit value is returned
            // once and required on every later call for this submission - upload and finalise -
            // which is what stops a passer-by uploading into, or finalising, somebody else's work.
            //
            // Only its HASH is stored, the same rule as every other token in this service: a
            // database read must not yield the ability to act on submissions in flight.
            string uploadToken = SecureToken.Create();

            long submissionId = await store.CreateAsync(
                new NewSubmission(
                    manifest.SystemId,

                    // The three parts travel alongside the id so the store can create the
                    // `systems` row a first submission implies, without splitting the id back
                    // apart - see NewSubmission's own header.
                    manifest.Manufacturer,
                    manifest.Hardware,
                    manifest.Board,
                    submitter.AccountId,
                    submitter.ContactEmail,
                    submitter.IpAddress,
                    SecureToken.Hash(uploadToken),
                    manifest.BaseRevision,
                    manifest.Summary,
                    manifest.FormatVersion,
                    manifest.Files,
                    now,
                    now + SubmissionFlows.UploadWindow),
                cancellationToken);

            await store.SavePayloadAsync(submissionId, manifest, cancellationToken);

            // Warnings are kept even on a submission that is proceeding: the reviewer needs to see
            // what was flagged at the time it was accepted.
            if (findings.Count > 0)
                await store.SaveFindingsAsync(submissionId, findings, cancellationToken);

            // Now the negotiation. A hash the store already holds is marked uploaded immediately -
            // THIS is what makes a typo fix cost no upload at all, and what stops a shared image
            // being re-sent for every board that references it.
            var missing = new List<string>();
            var alreadyHeld = new HashSet<string>(StringComparer.Ordinal);
            long bytesToUpload = 0;

            foreach (SubmissionFile file in manifest.Files)
            {
                if (alreadyHeld.Contains(file.Sha256) || missing.Contains(file.Sha256))
                    continue;

                if (blobs.Contains(file.Sha256))
                {
                    alreadyHeld.Add(file.Sha256);
                    await store.MarkUploadedAsync(submissionId, file.Sha256, cancellationToken);
                }
                else
                {
                    missing.Add(file.Sha256);
                    bytesToUpload += file.SizeBytes;
                }
            }

            return SubmissionCreationOutcome.Accepted(new HashNegotiationResponse
            {
                SubmissionId = submissionId,
                // The capability token that proves ownership of this submission for the upload and
                // finalise steps - see SubmissionFlows' header. Returned once, here, and never
                // again: the server stores only its hash.
                UploadToken = uploadToken,
                MissingHashes = missing,
                AlreadyHeldCount = alreadyHeld.Count,
                TotalBytesToUpload = bytesToUpload
            });
        }

        // ###########################################################################################
        // Step 3: accept a chunk of a blob.
        //
        // OWNERSHIP IS CHECKED HERE, not by the endpoint. A submission id is a small integer, so
        // without this anyone could upload into anyone else's submission - which would let them
        // replace a file in a submission somebody else is about to have reviewed and merged.
        // ###########################################################################################
        public static async Task<BlobUploadOutcome> UploadChunkAsync(
            long submissionId,
            string hash,
            long offset,
            Stream content,
            string? uploadToken,
            ISubmissionStore store,
            BlobStore blobs,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(blobs);
            ArgumentNullException.ThrowIfNull(content);

            SubmissionRecord? submission = await store.FindAsync(submissionId, cancellationToken);

            if (!SubmissionFlows.OwnsSubmission(submission, uploadToken))
                return BlobUploadOutcome.NotFound();

            if (submission.State != SubmissionState.Uploading)
                return BlobUploadOutcome.WrongState(submission.State);

            if (submission.ExpiresUtc is not null && submission.ExpiresUtc <= now)
                return BlobUploadOutcome.Expired();

            // The hash must be one this submission actually asked for. Otherwise a contributor
            // could use the upload endpoint to put arbitrary content into the shared blob store
            // under a hash of their choosing - content that a LATER submission could then
            // reference without uploading it.
            IReadOnlyList<SubmissionFileRecord> files = await store.GetFilesAsync(submissionId, cancellationToken);

            if (!files.Any(file => string.Equals(file.Sha256, hash, StringComparison.Ordinal)))
                return BlobUploadOutcome.UnexpectedHash();

            BlobChunkResult result = await blobs.AppendChunkAsync(
                submissionId, hash, offset, content, cancellationToken);

            if (!result.IsAccepted)
                return BlobUploadOutcome.ChunkRejected(result.ResumeFrom, result.Error ?? "The chunk was refused.");

            // The client says it has sent everything by asking to complete. Verification of the
            // hash happens inside TryCompleteAsync and is what makes the store trustworthy.
            long expectedSize = files
                .First(file => string.Equals(file.Sha256, hash, StringComparison.Ordinal))
                .SizeBytes;

            if (result.BytesOnDisk < expectedSize)
                return BlobUploadOutcome.Partial(result.BytesOnDisk);

            if (!await blobs.TryCompleteAsync(submissionId, hash, cancellationToken))
            {
                return BlobUploadOutcome.HashMismatch(
                    "The uploaded bytes do not match the hash the submission declared. " +
                    "The partial upload has been discarded; send it again.");
            }

            await store.MarkUploadedAsync(submissionId, hash, cancellationToken);

            return BlobUploadOutcome.Completed(result.BytesOnDisk);
        }

        // ###########################################################################################
        // Finalise: every blob has arrived, so validate once more and queue.
        //
        // WHY VALIDATE AGAIN. The manifest was checked at creation, but the FILES were not there
        // yet - only their declared names and hashes. Nothing about the bytes has been verified
        // against the rows until now. Re-running is cheap and closes the window where a manifest
        // was accepted and the content turned out to be something else.
        // ###########################################################################################
        public static async Task<SubmissionResult> FinaliseAsync(
            long submissionId,
            string? uploadToken,
            ISubmissionStore store,
            BlobStore blobs,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(blobs);

            SubmissionRecord? submission = await store.FindAsync(submissionId, cancellationToken);

            if (!SubmissionFlows.OwnsSubmission(submission, uploadToken))
            {
                return new SubmissionResult
                {
                    SubmissionId = submissionId,
                    IsAccepted = false,
                    State = string.Empty,
                    Findings =
                    [
                        new ValidationFinding
                        {
                            Severity = ValidationSeverity.Error,
                            Code = "submission.not_found",
                            Subject = string.Empty,
                            Message = "That submission does not exist."
                        }
                    ]
                };
            }

            if (submission.State != SubmissionState.Uploading)
            {
                // Already finalised. Answering with the current state rather than an error means a
                // client that retried after a dropped connection gets a sensible answer instead of
                // a failure for something that actually worked.
                return new SubmissionResult
                {
                    SubmissionId = submissionId,
                    IsAccepted = submission.State == SubmissionState.Pending,
                    State = submission.State,
                    Findings = (await store.GetFindingsAsync(submissionId, cancellationToken)).ToList()
                };
            }

            IReadOnlyList<SubmissionFileRecord> files = await store.GetFilesAsync(submissionId, cancellationToken);

            List<SubmissionFileRecord> missing = files.Where(file => !file.IsUploaded).ToList();

            if (missing.Count > 0)
            {
                return new SubmissionResult
                {
                    SubmissionId = submissionId,
                    IsAccepted = false,
                    State = submission.State,
                    Findings = missing
                        .Take(20)
                        .Select(file => new ValidationFinding
                        {
                            Severity = ValidationSeverity.Error,
                            Code = "file.not_uploaded",
                            Subject = file.Path,
                            Message = $"The file [{file.Path}] has not finished uploading."
                        })
                        .ToList()
                };
            }

            SubmissionManifest? manifest = await store.LoadPayloadAsync(submissionId, cancellationToken);

            if (manifest is null)
            {
                return new SubmissionResult
                {
                    SubmissionId = submissionId,
                    IsAccepted = false,
                    State = submission.State,
                    Findings =
                    [
                        new ValidationFinding
                        {
                            Severity = ValidationSeverity.Error,
                            Code = "submission.payload_missing",
                            Subject = string.Empty,
                            Message = "The submission's contents could not be read. Submit again."
                        }
                    ]
                };
            }

            IReadOnlyList<ValidationFinding> findings = SubmissionValidator.Validate(
                manifest, files.Select(file => file.Path).ToList());

            await store.SaveFindingsAsync(submissionId, findings, cancellationToken);

            bool accepted = SubmissionValidator.CanBeQueued(findings);

            string state = accepted ? SubmissionState.Pending : SubmissionState.Rejected;

            await store.SetStateAsync(submissionId, state, now, cancellationToken);

            // Partial uploads are cleared either way: a rejected submission's bytes are no longer
            // needed, and a queued one's have already been moved into the content-addressed store.
            blobs.ClearPartials(submissionId);

            return new SubmissionResult
            {
                SubmissionId = submissionId,
                IsAccepted = accepted,
                State = state,
                Findings = findings.ToList()
            };
        }

        // ###########################################################################################
        // Does the caller hold the capability token for this submission?
        //
        // FIXED-TIME COMPARISON, because an early-exit string compare on a secret leaks through
        // timing how many leading characters matched - which turns guessing a 256-bit token into a
        // per-character search. SecureToken.HashesEqual is the one place that comparison lives.
        //
        // A missing submission and a wrong token answer the same way, so a caller cannot learn
        // which submission ids exist by walking integers.
        // ###########################################################################################
        private static bool OwnsSubmission(
            [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] SubmissionRecord? submission,
            string? uploadToken)
        {
            if (submission is null || string.IsNullOrWhiteSpace(uploadToken))
                return false;

            return SecureToken.HashesEqual(
                submission.UploadTokenHash, SecureToken.Hash(uploadToken));
        }

        // ###########################################################################################
        // Garbage-collects submissions whose upload window passed without being finalised.
        //
        // Blobs from abandoned submissions must be collected or the disk fills quietly - named as a
        // trap in NewContributeStrategy.md. Only PARTIAL uploads are removed; completed blobs are
        // content-addressed and may be referenced by other systems, so reclaiming those is a
        // separate sweep against what the published tree actually references.
        // ###########################################################################################
        public static async Task<int> CollectAbandonedAsync(
            ISubmissionStore store,
            BlobStore blobs,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(blobs);

            IReadOnlyList<long> expired = await store.GetExpiredUploadsAsync(now, cancellationToken);

            foreach (long submissionId in expired)
            {
                blobs.ClearPartials(submissionId);
                await store.SetStateAsync(submissionId, SubmissionState.Abandoned, now, cancellationToken);
            }

            return expired.Count;
        }
    }

    // ###########################################################################################
    // Who is submitting.
    //
    // CONTRIBUTING REQUIRES NO ACCOUNT. AccountId is null for the ordinary case - somebody who
    // opened CRT, fixed a typo and pressed Submit - and set only when a signed-in maintainer
    // submits to a system they maintain.
    //
    // ContactEmail is what a reviewer replies to. It is NOT a credential and NOT an identity:
    // nothing signs in with it, and two submissions from the same address are two unrelated
    // submissions. It is required precisely because a contribution nobody can reply to can only be
    // taken or dropped, never improved.
    //
    // A signed-in maintainer needs no separate address - their account carries one - which is why
    // HasContact is satisfied by either.
    // ###########################################################################################
    public sealed record Submitter(long? AccountId, string? ContactEmail, string? IpAddress)
    {
        public bool HasContact =>
            this.AccountId is not null || !string.IsNullOrWhiteSpace(this.ContactEmail);

        public static Submitter Anonymous(string? contactEmail, string? ipAddress) =>
            new(null, contactEmail, ipAddress);

        public static Submitter SignedIn(long accountId, string? ipAddress) =>
            new(accountId, null, ipAddress);
    }

    // ###########################################################################################
    // Outcomes. Coarse where revealing more would leak, detailed where the client has to act.
    // ###########################################################################################
    public sealed record SubmissionCreationOutcome(
        bool IsAccepted,
        HashNegotiationResponse? Negotiation,
        IReadOnlyList<ValidationFinding> Findings)
    {
        public static SubmissionCreationOutcome Accepted(HashNegotiationResponse negotiation) =>
            new(true, negotiation, Array.Empty<ValidationFinding>());

        public static SubmissionCreationOutcome Refused(IReadOnlyList<ValidationFinding> findings) =>
            new(false, null, findings);
    }

    // ###########################################################################################
    // The outcome of one chunk.
    //
    // ResumeFrom is what a client needs after any failure that is not fatal: it says where to
    // continue from rather than leaving it to restart a 300 MB upload.
    // ###########################################################################################
    public sealed record BlobUploadOutcome(
        BlobUploadStatus Status,
        long ResumeFrom,
        string? Error)
    {
        public static BlobUploadOutcome Completed(long bytes) => new(BlobUploadStatus.Completed, bytes, null);

        public static BlobUploadOutcome Partial(long bytes) => new(BlobUploadStatus.Partial, bytes, null);

        public static BlobUploadOutcome NotFound() => new(BlobUploadStatus.NotFound, 0, null);

        public static BlobUploadOutcome Expired() =>
            new(BlobUploadStatus.Expired, 0, "The upload window for this submission has passed. Submit again.");

        public static BlobUploadOutcome WrongState(string state) =>
            new(BlobUploadStatus.WrongState, 0, $"This submission is {state} and is no longer accepting uploads.");

        public static BlobUploadOutcome UnexpectedHash() =>
            new(BlobUploadStatus.UnexpectedHash, 0, "That file is not part of this submission.");

        public static BlobUploadOutcome ChunkRejected(long resumeFrom, string error) =>
            new(BlobUploadStatus.ChunkRejected, resumeFrom, error);

        public static BlobUploadOutcome HashMismatch(string error) =>
            new(BlobUploadStatus.HashMismatch, 0, error);
    }

    public enum BlobUploadStatus
    {
        Completed,
        Partial,
        NotFound,
        Expired,
        WrongState,
        UnexpectedHash,
        ChunkRejected,
        HashMismatch
    }
}
