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
    //    tree stays a manual act by the project owner - Phase 3 step 3's filesystem permissions make
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
        //
        // `tree` is the published tree as it stands (security review, 2026-09-25): it is what lets a
        // submission cite another board's file UNCHANGED, and what catches a path differing from a
        // published one only by capitalisation. Null means it could not be consulted, and then
        // every foreign file is refused - see SubmissionFileRules.ValidateManifestFiles.
        //
        // It is also what lets a file published UNCHANGED be taken from this server's own disk
        // rather than uploaded again (2026-09-25) - see the import step below. `containmentRoot`
        // is the data-tree root every submitted path is contained to, and the same root the tree
        // was read from; the endpoint passes DataTreeRoot for both.
        // ###########################################################################################
        public static async Task<SubmissionCreationOutcome> CreateAsync(
            SubmissionManifest manifest,
            Submitter submitter,
            string containmentRoot,
            ISubmissionStore store,
            BlobStore blobs,
            DateTimeOffset now,
            CancellationToken cancellationToken = default,
            PublishedTreeView? tree = null)
        {
            ArgumentNullException.ThrowIfNull(manifest);
            ArgumentNullException.ThrowIfNull(submitter);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(blobs);

            var findings = new List<ValidationFinding>();

            // A contact address is the ONE thing a contributor must give. Not a credential -
            // nothing signs in with it - but without it a maintainer cannot say "accepted" or "this
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

            // ---- The per-address limit, BEFORE any real work (security review, 2026-09-25) -----
            //
            // Checked first on COUNT alone, so a sender over the limit costs one indexed query
            // rather than a full validation of a large manifest. The byte half is checked again
            // below, once the server knows how much it would actually be asked to store.
            IReadOnlyList<RecentSubmission> recent = await SubmissionFlows.RecentFromAsync(
                submitter, store, now, cancellationToken);

            RateLimitVerdict early = SubmissionRateLimitPolicy.Check(recent, 0, now);

            if (!early.IsAllowed)
                return SubmissionCreationOutcome.RateLimited(early.RetryAfter);

            // Paths first: these are the untrusted values that decide where files get written, and
            // a manifest that fails here is refused before anything else looks at it.
            findings.AddRange(SubmissionPathRules.ValidateManifestPaths(manifest, containmentRoot));

            // Then WHICH files and WHERE: allowed types, used by a row, and nothing belonging to
            // another board unless it is left exactly as published. See SubmissionFileRules.
            findings.AddRange(SubmissionFileRules.ValidateManifestFiles(manifest, tree));

            // Then the rows, against the paths the manifest itself supplies.
            findings.AddRange(SubmissionValidator.Validate(
                manifest, manifest.Files.Select(file => file.Path).ToList()));

            if (!SubmissionValidator.CanBeQueued(findings))
                return SubmissionCreationOutcome.Refused(findings);

            // A system the administrator has closed to contributions (systems.is_accepting = 0).
            // A board with no row yet is open: that is every new system and every shipped board
            // nobody has submitted to.
            if (await store.IsSystemAcceptingAsync(manifest.SystemId, cancellationToken) == false)
            {
                return SubmissionCreationOutcome.Refused(
                [
                    new ValidationFinding
                    {
                        Severity = ValidationSeverity.Error,
                        Code = "system.closed",
                        Subject = manifest.SystemId,
                        Message = "This board is not accepting contributions at the moment."
                    }
                ]);
            }

            // ###########################################################################################
            // *** THE REFERENCE GATE COVERS ONLY THE LAST STEP (code review, 2026-09-25). ***
            //
            // The gate is ONE lock for the whole service, shared with the blob collector. This
            // method used to hold it from here to the end - through importing a whole board from the
            // published tree (up to ~120 MB copied and hashed) and every database write - so one
            // first-time submission made every other contributor's create, and the collector, wait.
            //
            // Now what the server holds, what it can take from the tree, the room and budget checks
            // and the imports all happen OUTSIDE it. Inside it: every blob about to be called "held"
            // is confirmed still there - the collector may have run meanwhile, and anything it took
            // is asked for instead - and the submission row is created, which makes every one of its
            // hashes live to the collector. The rest follows outside. See BlobStore's header.
            // ###########################################################################################

            // Which blobs the server lacks, BEFORE a row exists, so the storage and byte checks
            // below can refuse without leaving anything behind.
            var missing = new List<string>();
            var alreadyHeld = new HashSet<string>(StringComparer.Ordinal);
            var importable = new Dictionary<string, (string Source, long SizeBytes)>(StringComparer.Ordinal);
            long bytesToUpload = 0;
            long bytesToImport = 0;

            foreach (SubmissionFile file in manifest.Files)
            {
                if (alreadyHeld.Contains(file.Sha256) || missing.Contains(file.Sha256) || importable.ContainsKey(file.Sha256))
                    continue;

                if (blobs.Contains(file.Sha256))
                {
                    alreadyHeld.Add(file.Sha256);
                }
                else if (SubmissionFlows.TryFindPublishedCopy(file, containmentRoot, tree, out string source))
                {
                    importable[file.Sha256] = (source, file.SizeBytes);
                    bytesToImport += file.SizeBytes;
                }
                else
                {
                    missing.Add(file.Sha256);
                    bytesToUpload += file.SizeBytes;
                }
            }

            // An import writes to this disk as surely as an upload does.
            if (!blobs.HasRoomFor(bytesToUpload + bytesToImport))
                return SubmissionCreationOutcome.NoRoom();

            // The sender's budget counts what THEY send. An import costs them nothing, and what it
            // stores is bounded by the published tree itself - a sender cannot make it bigger.
            RateLimitVerdict budget = SubmissionRateLimitPolicy.Check(recent, bytesToUpload, now);

            if (!budget.IsAllowed)
                return SubmissionCreationOutcome.RateLimited(budget.RetryAfter);

            // ###########################################################################################
            // *** THE SERVER'S OWN COPY OF EVERY FILE PUBLISHED UNCHANGED (2026-09-25). ***
            //
            // The blob store only ever held what somebody had UPLOADED, and the published tree was
            // never in it - so the first submission to any board uploaded all of it, whatever had
            // changed. Correcting one component's text on a shipped board sent 1,212 files and
            // 121 MB (reported), every byte of which was already on this disk.
            //
            // Taken INTO the store, verified, rather than read from the tree later: the maintainer
            // reads a submission's files from the store, finalise checks them there, and the
            // publish copies them from there. A file only NAMED as "in the tree" could be replaced
            // there by another board's publish before this one was reviewed, and the submission
            // would then name bytes that exist nowhere. So nothing downstream changes, and the
            // store holds exactly what an upload would have put there.
            //
            // OUTSIDE the reference gate. The collector may take an imported blob before the row
            // below records it; the check inside the gate then finds it gone and asks for it.
            // ###########################################################################################
            var imported = new HashSet<string>(StringComparer.Ordinal);
            bool uploadGrew = false;

            foreach ((string hash, (string source, long sizeBytes)) in importable)
            {
                bool wasHeld = blobs.Contains(hash);

                if (await blobs.TryImportAsync(source, hash, cancellationToken))
                {
                    alreadyHeld.Add(hash);

                    // Only what THIS create put there is its to take back on a refusal below.
                    if (!wasHeld)
                        imported.Add(hash);
                }
                else
                {
                    // Changed on disk since it was hashed, or unreadable. Ask for it instead.
                    missing.Add(hash);
                    bytesToUpload += sizeBytes;
                    uploadGrew = true;
                }
            }

            // THE CAPABILITY TOKEN. With no account, a submission id is a small integer anyone can
            // guess, so it cannot be what proves ownership. This random 256-bit value is returned
            // once and required on every later call for this submission - upload and finalise -
            // which is what stops a passer-by uploading into, or finalising, somebody else's work.
            //
            // Only its HASH is stored, the same rule as every other token in this service: a
            // database read must not yield the ability to act on submissions in flight.
            string uploadToken = SecureToken.Create();

            // Decided now, from the manifest against the published tree, and stored on the row: a
            // submission changing a shared file needs two approvals (ApprovalRules). With no view
            // of the tree this is TRUE, the safe side. Before the gate, because it may hash files.
            bool touchesSharedFiles = SubmissionSharedFiles.TouchesSharedFiles(manifest, tree);

            long submissionId;

            using (await blobs.EnterReferenceGateAsync(cancellationToken))
            {
                // Every blob about to be called "held" is still there - or it is asked for.
                foreach (string hash in alreadyHeld.ToList())
                {
                    if (blobs.Contains(hash))
                        continue;

                    alreadyHeld.Remove(hash);
                    imported.Remove(hash);
                    missing.Add(hash);
                    bytesToUpload += manifest.Files.First(file => file.Sha256 == hash).SizeBytes;
                    uploadGrew = true;
                }

                // Rare - only a failed import or a collected blob moves bytes here - but those ARE
                // sent by the contributor, so they are held to the same budget. A refusal takes back
                // what this create imported: nothing else can reference it yet except a submission
                // recorded meanwhile, which the live set shows.
                if (uploadGrew && bytesToUpload > 0)
                {
                    RateLimitVerdict afterImports = SubmissionRateLimitPolicy.Check(recent, bytesToUpload, now);

                    if (!afterImports.IsAllowed)
                    {
                        await SubmissionFlows.TakeBackImportsAsync(imported, store, blobs, cancellationToken);
                        return SubmissionCreationOutcome.RateLimited(afterImports.RetryAfter);
                    }
                }

                submissionId = await store.CreateAsync(
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
                        now + SubmissionFlows.UploadWindow,
                        bytesToUpload,

                        // Decided before the gate (it may hash files) - see touchesSharedFiles.
                        touchesSharedFiles,

                        // A new system's notes (2026-10-05), for the placement to start with.
                        // Their length was held to the notes column's by SubmissionValidator.
                        HardwareNotes: string.IsNullOrWhiteSpace(manifest.HardwareNotes) ? null : manifest.HardwareNotes.Trim()),
                    cancellationToken);
            }

            // Outside the gate: the row exists, so every hash it names is live and the collector
            // leaves it alone.
            await store.SavePayloadAsync(submissionId, manifest, cancellationToken);

            // Warnings are kept even on a submission that is proceeding: the maintainer needs to see
            // what was flagged at the time it was accepted.
            if (findings.Count > 0)
                await store.SaveFindingsAsync(submissionId, SubmissionFlows.FitForStorage(findings), cancellationToken);

            // Now the negotiation. A hash the store already holds is marked uploaded immediately -
            // THIS is what makes a typo fix cost no upload at all, and what stops a shared image
            // being re-sent for every board that references it.
            foreach (string hash in alreadyHeld)
                await store.MarkUploadedAsync(submissionId, hash, cancellationToken);

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

            if (result.IsNoRoom)
                return BlobUploadOutcome.NoRoom(result.Error ?? "The server is short of storage space.");

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
        // How much of one blob has arrived, for a client resuming an upload - or null, which the
        // endpoint answers as 404, for a caller not holding the token OR asking about a hash this
        // submission never listed.
        //
        // *** THE HASH IS SCOPED TO THE SUBMISSION (security review, 2026-09-25). *** This used to
        // check the token and then answer about ANY hash, so anyone holding any upload token - which
        // anyone can get, by submitting - could ask whether the shared store holds any file at all:
        // an oracle for "has somebody submitted this exact image". The upload step already refused
        // a hash the submission did not list; the question about it now refuses the same way.
        // ###########################################################################################
        public static async Task<UploadState?> GetUploadStateAsync(
            long submissionId,
            string hash,
            string? uploadToken,
            ISubmissionStore store,
            BlobStore blobs,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(blobs);

            SubmissionRecord? submission = await store.FindAsync(submissionId, cancellationToken);

            if (!SubmissionFlows.OwnsSubmission(submission, uploadToken))
                return null;

            IReadOnlyList<SubmissionFileRecord> files = await store.GetFilesAsync(submissionId, cancellationToken);

            if (!files.Any(file => string.Equals(file.Sha256, hash, StringComparison.Ordinal)))
                return null;

            return new UploadState(blobs.GetUploadedLength(submissionId, hash), blobs.Contains(hash));
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

            var findings = SubmissionValidator.Validate(
                manifest, files.Select(file => file.Path).ToList()).ToList();

            // ###########################################################################################
            // *** THE BYTES ARE CHECKED AGAINST THEIR NAMES (security review, 2026-09-25). ***
            //
            // Only now are all of them here. A ".png" whose bytes are an executable, or a ".txt"
            // carrying NUL bytes, is refused before any maintainer sees it - and before it could ever
            // reach every user's disk under a name that invites opening it. See
            // SubmissionContentRules for what is and is not checked.
            // ###########################################################################################
            findings.AddRange(await SubmissionFlows.CheckContentAsync(files, blobs, cancellationToken));

            await store.SaveFindingsAsync(submissionId, SubmissionFlows.FitForStorage(findings), cancellationToken);

            bool accepted = SubmissionValidator.CanBeQueued(findings);

            string state = accepted ? SubmissionState.Pending : SubmissionState.Rejected;

            await store.SetStateAsync(submissionId, state, now, cancellationToken);

            // Queued: it replaces the same contributor's older, untouched submissions of this
            // system (owner decision, 2026-09-26) - see SubmissionReplacementRules. Only once it is
            // queued: an upload that never finishes must not have taken the older one's place.
            if (accepted)
                await SubmissionFlows.WithdrawReplacedAsync(submission with { State = SubmissionState.Pending }, store, now, cancellationToken);

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
        // Withdraws what a newly queued submission replaces, and returns the ids withdrawn. The
        // rules pick the candidates; the store withdraws each only if it is still untouched when
        // it gets there, so one a maintainer amended meanwhile stays.
        // ###########################################################################################
        public static async Task<IReadOnlyList<long>> WithdrawReplacedAsync(
            SubmissionRecord arrived,
            ISubmissionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(arrived);
            ArgumentNullException.ThrowIfNull(store);

            IReadOnlyList<SubmissionRecord> waiting = await store.GetPendingForSystemAsync(arrived.SystemId, cancellationToken);
            var withdrawn = new List<long>();

            foreach (SubmissionRecord older in SubmissionReplacementRules.ReplacedBy(arrived, waiting))
            {
                if (await store.WithdrawReplacedAsync(older.Id, SubmissionReplacementRules.ReplacedComment, now, cancellationToken))
                    withdrawn.Add(older.Id);
            }

            return withdrawn;
        }

        // ###########################################################################################
        // Where this server's own copy of a submitted file is, when the published tree holds it
        // at the SAME path with the SAME hash - the file the contributor did not change.
        //
        // SAME PATH ONLY. The same bytes published under some other name would be found only by
        // hashing the whole tree, and the case that costs contributors real uploads is simpler:
        // every untouched file of the board they edited.
        //
        // The hash comes from the tree's cache and is only a reason to TRY; the import hashes what
        // it copies and is the real check. The path goes through the same containment rule and the
        // same link refusal the publish uses, so a link inside the tree cannot make an import read
        // from somewhere else.
        // ###########################################################################################
        private static bool TryFindPublishedCopy(
            SubmissionFile file,
            string containmentRoot,
            PublishedTreeView? tree,
            out string source)
        {
            source = string.Empty;

            if (tree is null)
                return false;

            if (!string.Equals(tree.HashOf(file.Path), file.Sha256, StringComparison.Ordinal))
                return false;

            if (!SubmissionPathRules.TryResolve(containmentRoot, file.Path, out string resolved, out _))
                return false;

            if (PublishPathSafety.FindLinkOnPath(containmentRoot, resolved, PublishPathSafety.IsLink) is not null)
                return false;

            source = resolved;
            return true;
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
        // content-addressed and may be referenced by other submissions, so reclaiming those is
        // CollectUnreferencedBlobsAsync below.
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

        // ###########################################################################################
        // How long a submission that ended without publishing keeps its stored rows and file list
        // before CollectRetiredAsync removes them (security review, 2026-09-25). Long enough that a
        // contributor or maintainer asking "what was in that one?" a few weeks later can be answered
        // from the database; short enough that an anonymous sender cannot pile them up for ever.
        // ###########################################################################################
        public static readonly TimeSpan RetiredRetention = TimeSpan.FromDays(30);

        // ###########################################################################################
        // Clears the bulk of submissions that ended without publishing - see
        // ISubmissionStore.DeleteRetiredPayloadsAsync for exactly what goes and what stays.
        // ###########################################################################################
        public static Task<int> CollectRetiredAsync(
            ISubmissionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);

            return store.DeleteRetiredPayloadsAsync(now - SubmissionFlows.RetiredRetention, cancellationToken);
        }

        // ###########################################################################################
        // Takes back blobs a create imported from the published tree and then did not use (a
        // refusal after the imports). Inside the reference gate, and only those no live submission
        // names - another create may have been told "held" about one and recorded it meanwhile.
        // Whatever this misses, the hourly collector takes.
        // ###########################################################################################
        private static async Task TakeBackImportsAsync(
            IReadOnlyCollection<string> imported,
            ISubmissionStore store,
            BlobStore blobs,
            CancellationToken cancellationToken)
        {
            if (imported.Count == 0)
                return;

            IReadOnlySet<string> live = await store.GetLiveBlobHashesAsync(cancellationToken);

            foreach (string hash in imported)
            {
                if (!live.Contains(hash))
                    blobs.TryDeleteBlob(hash);
            }
        }

        // ###########################################################################################
        // Deletes every completed blob no live submission needs (security review, 2026-09-25).
        //
        // *** UNTIL NOW NOTHING REMOVED A COMPLETED BLOB, EVER. *** The abandoned-upload sweep only
        // cleared PARTIAL uploads, so every file anyone finished uploading stayed on disk for good -
        // including every file of every rejected or abandoned submission from anybody at all.
        //
        // "Live" is ISubmissionStore.GetLiveBlobHashesAsync: uploading, pending, approved and
        // merged. Held inside the reference gate, so a submission being told "already held" at the
        // same moment cannot have its blob deleted from under it - see BlobStore.
        // ###########################################################################################
        public static async Task<int> CollectUnreferencedBlobsAsync(
            ISubmissionStore store,
            BlobStore blobs,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(blobs);

            using IDisposable gate = await blobs.EnterReferenceGateAsync(cancellationToken);

            IReadOnlySet<string> live = await store.GetLiveBlobHashesAsync(cancellationToken);

            int deleted = 0;

            foreach (string hash in blobs.EnumerateBlobHashes().ToList())
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (!live.Contains(hash) && blobs.TryDeleteBlob(hash))
                    deleted++;
            }

            return deleted;
        }

        // ###########################################################################################
        // Reads the first bytes of each submitted file's blob and checks them against the name
        // it will be published under. One finding per distinct path.
        // ###########################################################################################
        private static async Task<IReadOnlyList<ValidationFinding>> CheckContentAsync(
            IReadOnlyList<SubmissionFileRecord> files,
            BlobStore blobs,
            CancellationToken cancellationToken)
        {
            var findings = new List<ValidationFinding>();
            var seen = new HashSet<string>(StringComparer.Ordinal);

            foreach (SubmissionFileRecord file in files)
            {
                if (!seen.Add(file.Path))
                    continue;

                BlobVerification blob = await blobs.VerifyAsync(
                    file.Sha256, SubmissionContentRules.SampleBytes, cancellationToken);

                if (!blob.IsVerified)
                {
                    findings.Add(new ValidationFinding
                    {
                        Severity = ValidationSeverity.Error,
                        Code = "file.not_uploaded",
                        Subject = file.Path,
                        Message = $"The file [{file.Path}] is not on the server intact. Submit again."
                    });

                    continue;
                }

                if (!SubmissionContentRules.Matches(file.Path, blob.Head, out string reason))
                {
                    findings.Add(new ValidationFinding
                    {
                        Severity = ValidationSeverity.Error,
                        Code = "file.content_mismatch",
                        Subject = file.Path,
                        Message = $"[{file.Path}] cannot be submitted. {reason}"
                    });
                }
            }

            return findings;
        }

        // ###########################################################################################
        // What an address has submitted inside the rate-limit window, or nothing at all for a
        // caller the policy does not limit - a trusted account, or a request with no address,
        // which in practice is a test (see AccountFlows.MaySendMailAsync for the same reasoning).
        // ###########################################################################################
        private static async Task<IReadOnlyList<RecentSubmission>> RecentFromAsync(
            Submitter submitter,
            ISubmissionStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (submitter.IsTrusted || string.IsNullOrWhiteSpace(submitter.IpAddress))
                return [];

            return await store.GetRecentSubmissionsFromAddressAsync(
                submitter.IpAddress, now - SubmissionRateLimitPolicy.Window, cancellationToken);
        }

        // ###########################################################################################
        // Findings cut to the columns they are stored in (security review, 2026-09-25).
        //
        // A finding's Subject is often a board label or a path straight from the contributor's
        // rows, and nothing bounded it: a label of ten thousand characters produced a finding whose
        // INSERT failed against submission_findings.subject VARCHAR(500) - AFTER the submission row
        // had been written, answering a 500 and stranding a half-made submission. The message is
        // TEXT and needs no cut; the code is ours and always short, but is bounded anyway.
        // ###########################################################################################
        internal const int MaximumFindingSubjectLength = 500;

        internal const int MaximumFindingCodeLength = 64;

        internal static IReadOnlyList<ValidationFinding> FitForStorage(IReadOnlyList<ValidationFinding> findings)
        {
            ArgumentNullException.ThrowIfNull(findings);

            return findings
                .Select(finding => new ValidationFinding
                {
                    Severity = finding.Severity,
                    Code = SubmissionFlows.Cut(finding.Code, SubmissionFlows.MaximumFindingCodeLength),
                    Subject = SubmissionFlows.Cut(finding.Subject, SubmissionFlows.MaximumFindingSubjectLength),
                    Message = finding.Message ?? string.Empty
                })
                .ToList();
        }

        private static string Cut(string? value, int maximum)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            return value.Length <= maximum ? value : value[..(maximum - 3)] + "...";
        }
    }

    // ###########################################################################################
    // Who is submitting.
    //
    // CONTRIBUTING REQUIRES NO ACCOUNT. AccountId is null for the ordinary case - somebody who
    // opened CRT, fixed a typo and pressed Submit - and set only when a signed-in maintainer
    // submits to a system they maintain.
    //
    // ContactEmail is what a maintainer replies to. It is NOT a credential and NOT an identity:
    // nothing signs in with it, and two submissions from the same address are two unrelated
    // submissions. It is required precisely because a contribution nobody can reply to can only be
    // taken or dropped, never improved.
    //
    // A signed-in maintainer needs no separate address - their account carries one - which is why
    // HasContact is satisfied by either.
    //
    // IsTrusted is set ONLY for an account the database lists as a maintainer or administrator
    // (ReviewAuthority.CanReview), and exempts it from SubmissionRateLimitPolicy. Anybody may
    // register an ordinary account, so being signed in alone earns nothing.
    // ###########################################################################################
    public sealed record Submitter(long? AccountId, string? ContactEmail, string? IpAddress, bool IsTrusted = false)
    {
        public bool HasContact =>
            this.AccountId is not null || !string.IsNullOrWhiteSpace(this.ContactEmail);

        public static Submitter Anonymous(string? contactEmail, string? ipAddress) =>
            new(null, contactEmail, ipAddress);

        public static Submitter SignedIn(long accountId, string? ipAddress, bool isTrusted = false) =>
            new(accountId, null, ipAddress, isTrusted);
    }

    // ###########################################################################################
    // Outcomes. Coarse where revealing more would leak, detailed where the client has to act.
    // ###########################################################################################
    public sealed record SubmissionCreationOutcome(
        bool IsAccepted,
        HashNegotiationResponse? Negotiation,
        IReadOnlyList<ValidationFinding> Findings,
        bool IsRateLimited = false,
        TimeSpan RetryAfter = default,
        bool IsNoRoom = false)
    {
        public static SubmissionCreationOutcome Accepted(HashNegotiationResponse negotiation) =>
            new(true, negotiation, Array.Empty<ValidationFinding>());

        public static SubmissionCreationOutcome Refused(IReadOnlyList<ValidationFinding> findings) =>
            new(false, null, findings);

        // ###########################################################################################
        // Too many submissions, or too many bytes, from this address in the last day. Carries a
        // finding as well as a flag, because the client shows findings - a bare 429 would reach
        // the contributor as "the server refused the submission" and nothing more.
        // ###########################################################################################
        public static SubmissionCreationOutcome RateLimited(TimeSpan retryAfter) =>
            new(
                false,
                null,
                [
                    new ValidationFinding
                    {
                        Severity = ValidationSeverity.Error,
                        Code = "submission.rate_limited",
                        Subject = string.Empty,
                        Message =
                            "A lot has been submitted from your internet connection today. Try again in " +
                            $"about {Math.Max(1, (int)Math.Ceiling(retryAfter.TotalHours))} hour(s) - your draft is kept."
                    }
                ],
                IsRateLimited: true,
                RetryAfter: retryAfter);

        // The server's disk is below its reserve. Nothing about the submission is wrong.
        public static SubmissionCreationOutcome NoRoom() =>
            new(
                false,
                null,
                [
                    new ValidationFinding
                    {
                        Severity = ValidationSeverity.Error,
                        Code = "server.storage_full",
                        Subject = string.Empty,
                        Message =
                            "The server is short of storage space and cannot take submissions right now. " +
                            "Nothing is wrong with yours - try again later; your draft is kept."
                    }
                ],
                IsNoRoom: true);
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

        public static BlobUploadOutcome NoRoom(string error) =>
            new(BlobUploadStatus.NoRoom, 0, error);
    }

    // How much of one blob has arrived, and whether the store now holds it complete.
    public sealed record UploadState(long Uploaded, bool Complete);

    public enum BlobUploadStatus
    {
        Completed,
        Partial,
        NotFound,
        Expired,
        WrongState,
        UnexpectedHash,
        ChunkRejected,
        HashMismatch,
        NoRoom
    }
}
