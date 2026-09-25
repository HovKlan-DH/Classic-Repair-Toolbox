using System.Security.Cryptography;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Stores and retrieves submission blobs on disk. The I/O half; BlobStorePaths decides WHERE
    // and is pure.
    //
    // RESUMABLE, because the strategy document requires it: "a contributor whose 300 MB upload
    // dies at 90% and must start again will not start again." Resumption works because a partial
    // file's LENGTH is the offset - there is no separate bookkeeping to get out of step with the
    // bytes on disk. A client asks how much arrived, then sends from there.
    //
    // *** THE HASH IS VERIFIED BEFORE A BLOB IS ACCEPTED, ALWAYS. *** The client names the hash it
    // claims to be uploading; the server computes the hash of what actually arrived and refuses a
    // mismatch. Without that, "content-addressed" is a lie a caller can tell: uploading arbitrary
    // bytes under the hash of an innocent file would let a submission reference something the
    // server believes it has already vetted. This is the single most important line in the file.
    //
    // *** AND AGAIN WHEN IT IS COPIED OUT (security review, 2026-09-25). *** A blob used to be
    // verified once, on arrival, and copied into the published tree on trust. Now every copy is
    // hashed as it is written and refused on a mismatch, so a blob that changed on disk - by a
    // race, a bad sector or a hand edit - never reaches every user's machine under the name of the
    // file a reviewer approved.
    //
    // A blob is written to its partial path, verified, and only then MOVED into the store - so a
    // file in the store has always been verified, and an interrupted upload can never be mistaken
    // for a complete one.
    // ###########################################################################################
    public sealed class BlobStore
    {
        // How many locks the partial uploads share. Striped rather than one per upload, so the set
        // is fixed and never needs cleaning up; two unrelated uploads sharing a stripe merely wait
        // for each other for the length of one chunk.
        private const int LockStripes = 64;

        private readonly string thisBlobRoot;
        private readonly ILogger<BlobStore> thisLogger;
        private readonly Func<long?> thisFreeSpaceProbe;
        private readonly long thisMinimumFreeBytes;

        // ###########################################################################################
        // *** ONE PARTIAL UPLOAD IS TOUCHED BY ONE REQUEST AT A TIME (security review, 2026-09-25). ***
        //
        // Completing an upload hashed the partial file, closed it, and THEN moved it into the
        // store. A second request appending at the right offset in that gap added bytes to a file
        // that was about to enter the store under a hash those bytes no longer matched - and every
        // later submission citing that hash was told "already held" and published the poisoned
        // copy. Two appends at the same offset could also both pass the offset check and both
        // write. Taking this lock around the check-and-append and around the hash-and-move makes
        // each of those steps whole.
        // ###########################################################################################
        private readonly SemaphoreSlim[] thisUploadLocks =
            Enumerable.Range(0, BlobStore.LockStripes).Select(_ => new SemaphoreSlim(1, 1)).ToArray();

        // ###########################################################################################
        // *** THE REFERENCE GATE: garbage collection versus "you already have this" (2026-09-25). ***
        //
        // Creating a submission asks "does the store hold this hash?" and, when it does, tells the
        // client not to upload it. The collector deletes blobs no live submission references. If
        // the collector read its list of references just before a new submission recorded its
        // files, and deleted just after that submission was told "already held", the submission
        // would reach review naming a blob that no longer exists. Both hold this gate for the
        // whole of their check-and-act, so each sees the other's work completely or not at all.
        //
        // The create's check-and-act is SHORT (code review, 2026-09-25): it decides and imports
        // outside the gate, then inside it confirms each "held" blob still exists and creates the
        // row that makes them live - milliseconds, rather than a whole board's import. See
        // SubmissionFlows.CreateAsync.
        // ###########################################################################################
        private readonly SemaphoreSlim thisReferenceGate = new(1, 1);

        public BlobStore(
            string blobRoot,
            ILogger<BlobStore> logger,
            Func<long?>? freeSpaceProbe = null,
            long minimumFreeBytes = 0)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(blobRoot);
            ArgumentNullException.ThrowIfNull(logger);

            this.thisBlobRoot = blobRoot;
            this.thisLogger = logger;

            // No probe means "unknown", which is treated as "room" - a store that could not measure
            // its disk must not refuse every upload, and the real service always passes one.
            this.thisFreeSpaceProbe = freeSpaceProbe ?? (() => null);
            this.thisMinimumFreeBytes = Math.Max(0, minimumFreeBytes);
        }

        // ###########################################################################################
        // Does the store already hold this blob? This is what makes a typo fix upload nothing.
        // ###########################################################################################
        public bool Contains(string hash)
        {
            return BlobStorePaths.TryGetBlobPath(this.thisBlobRoot, hash, out string path)
                && File.Exists(path);
        }

        // ###########################################################################################
        // Would storing this many more bytes still leave the configured reserve free?
        //
        // *** THE DISK IS SHARED WITH THE WEB SITE AND THE DATABASE (security review, 2026-09-25). ***
        // Contributing needs no account, so the per-address limits are all that bound one sender -
        // and a sender with many addresses is bounded by nothing but the disk. Refusing new work
        // while the disk is below its reserve turns "the whole site fell over" into "contributions
        // paused", which is the failure worth having.
        // ###########################################################################################
        public bool HasRoomFor(long bytes)
        {
            long? free = this.thisFreeSpaceProbe();

            if (free is null)
                return true;

            return free.Value - Math.Max(0, bytes) >= this.thisMinimumFreeBytes;
        }

        // ###########################################################################################
        // Enters the reference gate - see its field's header. Dispose the result to leave it.
        // ###########################################################################################
        public async Task<IDisposable> EnterReferenceGateAsync(CancellationToken cancellationToken = default)
        {
            await this.thisReferenceGate.WaitAsync(cancellationToken).ConfigureAwait(false);

            return new Releaser(this.thisReferenceGate);
        }

        // ###########################################################################################
        // How many bytes of this blob have already arrived for this submission.
        //
        // Zero means "start from the beginning", which is also the answer when nothing has been
        // uploaded yet - so a client needs no special case for a first attempt.
        // ###########################################################################################
        public long GetUploadedLength(long submissionId, string hash)
        {
            if (!BlobStorePaths.TryGetPartialPath(this.thisBlobRoot, submissionId, hash, out string path))
                return 0;

            var file = new FileInfo(path);

            return file.Exists ? file.Length : 0;
        }

        // ###########################################################################################
        // Appends a chunk to a partial upload.
        //
        // offset is where the client believes it is resuming from, and it MUST match what is
        // already on disk. A mismatch means the two sides disagree about how much arrived, and
        // appending anyway would silently corrupt the blob - it would fail the hash check later,
        // but only after the whole upload had completed, which is exactly the outcome resumption
        // exists to avoid. Refusing immediately tells the client to re-ask for the offset.
        //
        // The size cap is enforced DURING the write, not after, so a caller cannot fill the disk by
        // streaming an unbounded body.
        // ###########################################################################################
        public async Task<BlobChunkResult> AppendChunkAsync(
            long submissionId,
            string hash,
            long offset,
            Stream content,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(content);

            if (!BlobStorePaths.TryGetPartialPath(this.thisBlobRoot, submissionId, hash, out string partialPath))
                return BlobChunkResult.Rejected("The blob hash or submission is not valid.");

            if (!this.HasRoomFor(0))
            {
                this.thisLogger.LogWarning(
                    "Refused an upload chunk for submission {SubmissionId}: the disk is below its free-space reserve.",
                    submissionId);

                return BlobChunkResult.NoRoom(
                    "The server is short of storage space and cannot accept uploads right now. Try again later.");
            }

            SemaphoreSlim uploadLock = this.LockFor(submissionId, hash);

            await uploadLock.WaitAsync(cancellationToken).ConfigureAwait(false);

            try
            {
                long existing = this.GetUploadedLength(submissionId, hash);

                if (offset != existing)
                {
                    return BlobChunkResult.OffsetMismatch(
                        existing,
                        $"The upload resumes at {existing} bytes, but the request began at {offset}.");
                }

                Directory.CreateDirectory(Path.GetDirectoryName(partialPath)!);

                long written = existing;

                // Append mode, so a resumed upload continues the same file rather than truncating it.
                await using (var target = new FileStream(
                    partialPath, FileMode.Append, FileAccess.Write, FileShare.None, 81920, useAsync: true))
                {
                    byte[] buffer = new byte[81920];
                    int read;

                    while ((read = await content.ReadAsync(buffer, cancellationToken)) > 0)
                    {
                        written += read;

                        // Checked as it streams: a client that lies about its size in the manifest
                        // must not be able to fill the disk before anyone notices.
                        if (written > SubmissionFormat.MaximumBlobBytes)
                        {
                            this.thisLogger.LogWarning(
                                "Submission {SubmissionId} exceeded the blob size limit while uploading {Hash}.",
                                submissionId, hash);

                            return BlobChunkResult.Rejected(
                                $"The file exceeds the maximum of {SubmissionFormat.MaximumBlobBytes} bytes.");
                        }

                        await target.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
                    }
                }

                return BlobChunkResult.Accepted(written);
            }
            finally
            {
                uploadLock.Release();
            }
        }

        // ###########################################################################################
        // Verifies a completed upload and moves it into the store.
        //
        // THE VERIFICATION IS THE POINT - see the class header. The hash is computed from the bytes
        // on disk, not from anything the client said, and a mismatch deletes the partial file so a
        // retry starts clean rather than resuming a corrupt one.
        //
        // Under the same lock as AppendChunkAsync, from the first byte hashed to the move - so the
        // bytes that were verified are the bytes that enter the store. See thisUploadLocks.
        // ###########################################################################################
        public async Task<bool> TryCompleteAsync(
            long submissionId,
            string hash,
            CancellationToken cancellationToken = default)
        {
            if (!BlobStorePaths.TryGetPartialPath(this.thisBlobRoot, submissionId, hash, out string partialPath))
                return false;

            if (!BlobStorePaths.TryGetBlobPath(this.thisBlobRoot, hash, out string finalPath))
                return false;

            SemaphoreSlim uploadLock = this.LockFor(submissionId, hash);

            await uploadLock.WaitAsync(cancellationToken).ConfigureAwait(false);

            try
            {
                if (!File.Exists(partialPath))
                    return false;

                string actualHash;

                await using (var stream = new FileStream(
                    partialPath, FileMode.Open, FileAccess.Read, FileShare.None, 81920, useAsync: true))
                {
                    byte[] computed = await SHA256.HashDataAsync(stream, cancellationToken);
                    actualHash = Convert.ToHexStringLower(computed);
                }

                if (!string.Equals(actualHash, hash, StringComparison.Ordinal))
                {
                    this.thisLogger.LogWarning(
                        "Submission {SubmissionId} uploaded a blob claiming hash {Claimed} but the content " +
                        "hashes to {Actual}. The partial upload has been discarded.",
                        submissionId, hash, actualHash);

                    VerifiedFileCopy.TryDelete(partialPath);

                    return false;
                }

                Directory.CreateDirectory(Path.GetDirectoryName(finalPath)!);

                // Another submission may have completed the same blob while this one was uploading -
                // legitimate and common, since the hash is the identity. The bytes are identical by
                // definition, so the existing file wins and this copy is discarded.
                if (File.Exists(finalPath))
                {
                    VerifiedFileCopy.TryDelete(partialPath);
                    return true;
                }

                try
                {
                    File.Move(partialPath, finalPath);
                }
                catch (IOException)
                {
                    // Lost a race with another submission moving the same verified blob into place.
                    // Identical bytes, so this is success, not failure.
                    if (File.Exists(finalPath))
                    {
                        VerifiedFileCopy.TryDelete(partialPath);
                        return true;
                    }

                    throw;
                }

                return true;
            }
            finally
            {
                uploadLock.Release();
            }
        }

        // ###########################################################################################
        // Clears every partial upload for a submission - called when it is finalised or abandoned.
        //
        // Blobs from abandoned submissions must be garbage-collected or the disk fills quietly;
        // one folder per submission is what makes that a directory delete rather than a query.
        //
        // Completed blobs are NOT removed here: they are content-addressed and may be referenced by
        // other submissions. SubmissionFlows.CollectUnreferencedBlobsAsync reclaims those.
        // ###########################################################################################
        public void ClearPartials(long submissionId)
        {
            if (!BlobStorePaths.TryGetPartialFolder(this.thisBlobRoot, submissionId, out string folder))
                return;

            try
            {
                if (Directory.Exists(folder))
                    Directory.Delete(folder, recursive: true);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Leaving a partial folder behind wastes disk but breaks nothing, and throwing
                // here would fail a submission that has otherwise succeeded.
                this.thisLogger.LogWarning(
                    ex, "Could not clear partial uploads for submission {SubmissionId}.", submissionId);
            }
        }

        // ###########################################################################################
        // Where a completed blob's bytes are, for a caller that wants to READ them.
        //
        // Returns the PATH rather than a stream so the caller can hand it to Results.File, which
        // streams it and answers range requests without this class reimplementing either.
        //
        // *** THIS DOES NOT DECIDE WHETHER THE CALLER MAY HAVE IT. *** It answers "do these bytes
        // exist", nothing more. The store is shared across every submission and content-addressed,
        // so anyone holding a hash could otherwise read anyone's upload. ReviewAssetLocator is
        // where the scope check lives, and every caller must go through it first.
        // ###########################################################################################
        public bool TryGetReadablePath(string hash, out string path)
        {
            path = string.Empty;

            if (!BlobStorePaths.TryGetBlobPath(this.thisBlobRoot, hash, out string candidate))
                return false;

            if (!File.Exists(candidate))
                return false;

            path = candidate;
            return true;
        }

        // ###########################################################################################
        // Hashes a stored blob in full and hands back its first `headBytes` bytes (security review,
        // 2026-09-25).
        //
        // One read serves both checks a caller makes before trusting a blob: that it still IS the
        // hash it is filed under, and that its opening bytes match the file type it will be
        // published as (SubmissionContentRules). Reading it twice would double the work on the
        // largest board scans for nothing.
        // ###########################################################################################
        public async Task<BlobVerification> VerifyAsync(
            string hash,
            int headBytes,
            CancellationToken cancellationToken = default)
        {
            if (!BlobStorePaths.TryGetBlobPath(this.thisBlobRoot, hash, out string path) || !File.Exists(path))
                return BlobVerification.Missing;

            byte[] head = new byte[Math.Max(0, headBytes)];
            int headLength = 0;

            using var hasher = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);

            await using (var stream = new FileStream(
                path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true))
            {
                byte[] buffer = new byte[81920];
                int read;

                while ((read = await stream.ReadAsync(buffer, cancellationToken)) > 0)
                {
                    hasher.AppendData(buffer, 0, read);

                    int take = Math.Min(read, head.Length - headLength);

                    if (take > 0)
                    {
                        Array.Copy(buffer, 0, head, headLength, take);
                        headLength += take;
                    }
                }
            }

            string actual = Convert.ToHexStringLower(hasher.GetHashAndReset());

            return string.Equals(actual, hash, StringComparison.Ordinal)
                ? BlobVerification.Verified(head[..headLength])
                : BlobVerification.Mismatch;
        }

        // ###########################################################################################
        // Copies a stored blob out to a path in the published tree - VERIFIED, and never left
        // half-written.
        //
        // The DESTINATION has already been through SubmissionPathRules - this method does not
        // re-derive it, because a second derivation is a second chance to get it wrong. It takes
        // the resolved path its caller validated.
        //
        // *** WRITTEN BESIDE THE TARGET, HASHED, THEN MOVED OVER IT (security review, 2026-09-25). ***
        // The copy used to stream straight into the destination, so a failure part way left a
        // truncated file in the published tree, and nothing checked the bytes at all. Now the copy
        // goes to a hidden temporary file next to the target (a dot-name, so the checksum manifest
        // skips it if the process dies), its hash is compared with the blob's name, and only a
        // match replaces the real file - in one rename.
        //
        // A rename needs write permission on the FOLDER, not on the old file, so a board image
        // whose group was lost to a network-share copy no longer stops a publish.
        // ###########################################################################################
        public async Task<BlobCopyResult> TryCopyToAsync(
            string hash,
            string resolvedDestinationPath,
            CancellationToken cancellationToken = default)
        {
            if (!BlobStorePaths.TryGetBlobPath(this.thisBlobRoot, hash, out string source))
                return BlobCopyResult.Missing;

            // The copy itself is VerifiedFileCopy's - the same rule the production promotion
            // copies with, so the two cannot come to verify differently.
            VerifiedCopyResult result = await VerifiedFileCopy
                .CopyAsync(source, hash, resolvedDestinationPath, cancellationToken: cancellationToken)
                .ConfigureAwait(false);

            if (result == VerifiedCopyResult.HashMismatch)
            {
                this.thisLogger.LogError(
                    "The stored blob {Hash} no longer hashes to its name. It was NOT copied to {Destination}.",
                    hash, resolvedDestinationPath);
            }

            return result switch
            {
                VerifiedCopyResult.Copied => BlobCopyResult.Copied,
                VerifiedCopyResult.HashMismatch => BlobCopyResult.HashMismatch,
                _ => BlobCopyResult.Missing
            };
        }

        // ###########################################################################################
        // Takes a file this server ALREADY HAS into the store under its hash - the published copy
        // of a file a submission cites unchanged (2026-09-25). See SubmissionFlows.CreateAsync.
        //
        // *** ON EXACTLY THE TERMS AN UPLOAD GETS. *** Written to a temporary file beside where the
        // blob belongs, hashed as it is written, and moved in only on a match - so "a file in the
        // store has always been verified" holds for an import too. The caller chose to import
        // because a CACHED hash said the file was unchanged, and a cache can be stale; a blob filed
        // under a hash its bytes do not match would be "already held" for every later submission
        // citing it, and fail every one of their finalises.
        //
        // COPIED, NEVER MOVED OR LINKED. The source is what every client syncs, and a hard link
        // would let an in-place edit of the published file change a blob that is supposed to be
        // immutable.
        //
        // The SOURCE has already been through SubmissionPathRules and the link check, as for
        // TryCopyToAsync's destination - this method takes the resolved path its caller validated.
        // False for anything short of a verified blob in the store; the caller then asks the
        // contributor to upload the file, which is always correct.
        // ###########################################################################################
        public async Task<bool> TryImportAsync(
            string resolvedSourcePath,
            string hash,
            CancellationToken cancellationToken = default)
        {
            if (!BlobStorePaths.TryGetBlobPath(this.thisBlobRoot, hash, out string finalPath))
                return false;

            if (File.Exists(finalPath))
                return true;

            try
            {
                // VerifiedFileCopy's loop - the one every other verified copy uses - under the same
                // cap an upload streams under. Nothing oversized should be published, but an import
                // must not be the one way round the limit. The temporary file's dot-name is never a
                // hash, so EnumerateBlobHashes never takes it for a blob if the process dies part way.
                VerifiedCopyResult result = await VerifiedFileCopy
                    .CopyAsync(
                        resolvedSourcePath,
                        hash,
                        finalPath,
                        maximumBytes: SubmissionFormat.MaximumBlobBytes,
                        temporaryTag: "crt-import",
                        cancellationToken: cancellationToken)
                    .ConfigureAwait(false);

                if (result == VerifiedCopyResult.HashMismatch)
                {
                    this.thisLogger.LogWarning(
                        "The published file {Source} was expected to hash to {Expected} but does not. " +
                        "It was not taken into the blob store; the contributor will be asked to upload it.",
                        resolvedSourcePath, hash);
                }

                return result == VerifiedCopyResult.Copied;
            }
            catch (IOException) when (File.Exists(finalPath))
            {
                // An upload of the same bytes completed meanwhile. Identical by definition, so this
                // is success - the same rule TryCompleteAsync applies.
                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                this.thisLogger.LogWarning(
                    ex, "Could not take the published file {Source} into the blob store.", resolvedSourcePath);

                return false;
            }
        }

        // ###########################################################################################
        // Every hash the store holds, for the garbage collector. A file whose name is not a
        // well-formed hash is not a blob and is left alone - it is not this class's to delete.
        // ###########################################################################################
        public IEnumerable<string> EnumerateBlobHashes()
        {
            string root = Path.Combine(this.thisBlobRoot, BlobStorePaths.BlobFolderName);

            if (!Directory.Exists(root))
                yield break;

            foreach (string path in Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories))
            {
                string name = Path.GetFileName(path);

                // Only a blob filed where its own name says it belongs - "ab/cd/abcd...". A stray
                // file elsewhere is not something this store put there.
                if (SubmissionPathRules.IsValidHash(name) &&
                    BlobStorePaths.TryGetBlobPath(this.thisBlobRoot, name, out string expected) &&
                    string.Equals(Path.GetFullPath(expected), Path.GetFullPath(path), StringComparison.Ordinal))
                {
                    yield return name;
                }
            }
        }

        // Deletes one completed blob. False when it was not there or could not be removed.
        public bool TryDeleteBlob(string hash)
        {
            if (!BlobStorePaths.TryGetBlobPath(this.thisBlobRoot, hash, out string path) || !File.Exists(path))
                return false;

            try
            {
                File.Delete(path);
                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                this.thisLogger.LogWarning(ex, "Could not delete the unreferenced blob {Hash}.", hash);
                return false;
            }
        }

        // The lock guarding one partial upload. Internal so a test can hold it and prove that an
        // append or a completion really waits for it.
        internal SemaphoreSlim LockFor(long submissionId, string hash)
        {
            int index = (int)((uint)HashCode.Combine(submissionId, hash) % BlobStore.LockStripes);

            return this.thisUploadLocks[index];
        }

        private sealed class Releaser(SemaphoreSlim gate) : IDisposable
        {
            private SemaphoreSlim? thisGate = gate;

            public void Dispose() => Interlocked.Exchange(ref this.thisGate, null)?.Release();
        }
    }

    // ###########################################################################################
    // The outcome of appending a chunk.
    //
    // ResumeFrom is meaningful only on an offset mismatch, where it tells the client where to
    // actually resume rather than leaving it to guess or restart.
    // ###########################################################################################
    public sealed record BlobChunkResult(bool IsAccepted, long BytesOnDisk, long ResumeFrom, string? Error, bool IsNoRoom = false)
    {
        public static BlobChunkResult Accepted(long bytesOnDisk) =>
            new(true, bytesOnDisk, bytesOnDisk, null);

        public static BlobChunkResult OffsetMismatch(long resumeFrom, string error) =>
            new(false, resumeFrom, resumeFrom, error);

        public static BlobChunkResult Rejected(string error) =>
            new(false, 0, 0, error);

        public static BlobChunkResult NoRoom(string error) =>
            new(false, 0, 0, error, IsNoRoom: true);
    }

    public enum BlobCopyResult
    {
        Copied,
        Missing,
        HashMismatch
    }

    // ###########################################################################################
    // What VerifyAsync found. Head is the blob's opening bytes, present only when it verified.
    // ###########################################################################################
    public sealed record BlobVerification(bool Exists, bool HashMatches, byte[] Head)
    {
        public static BlobVerification Missing { get; } = new(false, false, []);

        public static BlobVerification Mismatch { get; } = new(true, false, []);

        public static BlobVerification Verified(byte[] head) => new(true, true, head);

        public bool IsVerified => this.Exists && this.HashMatches;
    }
}
