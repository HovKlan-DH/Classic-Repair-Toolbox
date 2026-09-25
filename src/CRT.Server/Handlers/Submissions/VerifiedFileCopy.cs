using System.Security.Cryptography;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Copies one file - VERIFIED, and never left half-written (security review, 2026-09-25; shared
    // with the production promotion and the blob store's import since).
    //
    // The copy goes to a hidden temporary file beside the target (a dot-name, so the checksum
    // manifest skips it, a promotion never carries it and the blob store never takes it for a blob
    // if the process dies), is hashed as it is written, and only a match replaces the real file -
    // in one rename. A rename needs write permission on the FOLDER, not on the old file.
    //
    // Three callers, one rule: BlobStore.TryCopyToAsync (blob -> BETA, a publish),
    // ProductionPromoter (BETA -> Production) and BlobStore.TryImportAsync (a published file -> the
    // blob store). The import was a hand-kept second copy of this loop until the code review of
    // 2026-09-25; a fix to one would not have reached the other. Every DESTINATION has already been
    // through SubmissionPathRules and the link check; this method does not re-derive it.
    // ###########################################################################################
    public static class VerifiedFileCopy
    {
        // The temporary file's tag for a publish or a promotion - the default.
        public const string PublishTag = "crt-publish";

        // `maximumBytes`: stop, write nothing, and answer TooLarge once the source passes it - the
        // import's rule, so an import cannot be the one way round the upload limit. Null is no cap.
        // `temporaryTag`: names the temporary file, so a stray one says which operation left it.
        public static async Task<VerifiedCopyResult> CopyAsync(
            string sourcePath,
            string expectedSha256,
            string resolvedDestinationPath,
            long? maximumBytes = null,
            string temporaryTag = VerifiedFileCopy.PublishTag,
            CancellationToken cancellationToken = default)
        {
            if (!File.Exists(sourcePath))
                return VerifiedCopyResult.Missing;

            string directory = Path.GetDirectoryName(resolvedDestinationPath)!;

            Directory.CreateDirectory(directory);

            string temporary = Path.Combine(
                directory,
                $".{Path.GetFileName(resolvedDestinationPath)}.{temporaryTag}-{Guid.NewGuid():N}.tmp");

            try
            {
                string actual;

                await using (FileStream input = new(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true))
                await using (FileStream output = new(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, useAsync: true))
                {
                    using var hasher = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);

                    byte[] buffer = new byte[81920];
                    long written = 0;
                    int read;

                    while ((read = await input.ReadAsync(buffer, cancellationToken)) > 0)
                    {
                        written += read;

                        if (written > maximumBytes)
                            return VerifiedCopyResult.TooLarge;

                        hasher.AppendData(buffer, 0, read);
                        await output.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
                    }

                    actual = Convert.ToHexStringLower(hasher.GetHashAndReset());
                }

                if (!string.Equals(actual, expectedSha256, StringComparison.Ordinal))
                    return VerifiedCopyResult.HashMismatch;

                File.Move(temporary, resolvedDestinationPath, overwrite: true);

                return VerifiedCopyResult.Copied;
            }
            finally
            {
                VerifiedFileCopy.TryDelete(temporary);
            }
        }

        // Deletes a file if it is there. A stray temporary or partial file is harmless - a dot-name
        // is never published, promoted or taken for a blob - so a failure is not worth reporting.
        // Shared with BlobStore, which had its own copy of this.
        internal static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // See above.
            }
        }
    }

    public enum VerifiedCopyResult
    {
        Copied,
        Missing,
        HashMismatch,

        // The source passed the caller's `maximumBytes`; nothing was written.
        TooLarge
    }
}
