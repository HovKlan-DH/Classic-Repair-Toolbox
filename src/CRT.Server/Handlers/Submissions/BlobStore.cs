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
    // A blob is written to its partial path, verified, and only then MOVED into the store - so a
    // file in the store has always been verified, and an interrupted upload can never be mistaken
    // for a complete one.
    // ###########################################################################################
    public sealed class BlobStore
    {
        private readonly string thisBlobRoot;
        private readonly ILogger<BlobStore> thisLogger;

        public BlobStore(string blobRoot, ILogger<BlobStore> logger)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(blobRoot);
            ArgumentNullException.ThrowIfNull(logger);

            this.thisBlobRoot = blobRoot;
            this.thisLogger = logger;
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

        // ###########################################################################################
        // Verifies a completed upload and moves it into the store.
        //
        // THE VERIFICATION IS THE POINT - see the class header. The hash is computed from the bytes
        // on disk, not from anything the client said, and a mismatch deletes the partial file so a
        // retry starts clean rather than resuming a corrupt one.
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

            if (!File.Exists(partialPath))
                return false;

            string actualHash;

            await using (var stream = new FileStream(
                partialPath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true))
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

                BlobStore.TryDelete(partialPath);

                return false;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(finalPath)!);

            // Another submission may have completed the same blob while this one was uploading -
            // legitimate and common, since the hash is the identity. The bytes are identical by
            // definition, so the existing file wins and this copy is discarded.
            if (File.Exists(finalPath))
            {
                BlobStore.TryDelete(partialPath);
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
                    BlobStore.TryDelete(partialPath);
                    return true;
                }

                throw;
            }

            return true;
        }

        // ###########################################################################################
        // Clears every partial upload for a submission - called when it is finalised or abandoned.
        //
        // Blobs from abandoned submissions must be garbage-collected or the disk fills quietly;
        // one folder per submission is what makes that a directory delete rather than a query.
        //
        // Completed blobs are NOT removed: they are content-addressed and may be referenced by
        // other systems. Reclaiming those is a separate sweep against what the published tree
        // actually references.
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
        // Copies a stored blob out to a path in the published tree.
        //
        // The DESTINATION has already been through SubmissionPathRules - this method does not
        // re-derive it, because a second derivation is a second chance to get it wrong. It takes
        // the resolved path its caller validated.
        // ###########################################################################################
        public async Task<bool> TryCopyToAsync(
            string hash,
            string resolvedDestinationPath,
            CancellationToken cancellationToken = default)
        {
            if (!BlobStorePaths.TryGetBlobPath(this.thisBlobRoot, hash, out string source))
                return false;

            if (!File.Exists(source))
                return false;

            Directory.CreateDirectory(Path.GetDirectoryName(resolvedDestinationPath)!);

            await using FileStream input = new(source, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true);
            await using FileStream output = new(resolvedDestinationPath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true);

            await input.CopyToAsync(output, cancellationToken);

            return true;
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Nothing useful to do; a stray partial file is harmless and will be cleared with
                // the submission's folder.
            }
        }
    }

    // ###########################################################################################
    // The outcome of appending a chunk.
    //
    // ResumeFrom is meaningful only on an offset mismatch, where it tells the client where to
    // actually resume rather than leaving it to guess or restart.
    // ###########################################################################################
    public sealed record BlobChunkResult(bool IsAccepted, long BytesOnDisk, long ResumeFrom, string? Error)
    {
        public static BlobChunkResult Accepted(long bytesOnDisk) =>
            new(true, bytesOnDisk, bytesOnDisk, null);

        public static BlobChunkResult OffsetMismatch(long resumeFrom, string error) =>
            new(false, resumeFrom, resumeFrom, error);

        public static BlobChunkResult Rejected(string error) =>
            new(false, 0, 0, error);
    }
}
