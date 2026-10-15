using System;
using System.Buffers;
using System.IO;
using System.Security.Cryptography;
using System.Threading.Tasks;

namespace Handlers.OnlineHandling
{
    // ###########################################################################################
    // A download's body copied to its file a chunk at a time, hashed as it goes (code review,
    // 2026-10-13). It used to be read whole into memory (ReadAsByteArrayAsync - 44 MB for the
    // biggest KiCad board) under a total time budget; now every chunk that arrives tells the
    // watchdog there is progress, and the SHA-256 is ready the moment the last byte is written.
    // OnlineServices.DownloadFileAsync checks it against the manifest before the file is moved
    // into place - this never decides whether a file is good.
    // ###########################################################################################
    internal static class StreamingDownload
    {
        internal const int ChunkSize = 81920;

        // ###########################################################################################
        // Copies `source` to `target` and returns the lowercase hex SHA-256 of what was copied.
        // Throws OperationCanceledException when the watchdog's token is cancelled - a stall, or
        // the caller giving up (watchdog.Stalled says which).
        // ###########################################################################################
        internal static async Task<string> CopyAndHashAsync(Stream source, Stream target, NoProgressWatchdog watchdog)
        {
            using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            byte[] buffer = ArrayPool<byte>.Shared.Rent(ChunkSize);

            try
            {
                int read;

                while ((read = await source.ReadAsync(buffer.AsMemory(0, ChunkSize), watchdog.Token).ConfigureAwait(false)) > 0)
                {
                    watchdog.Progress();
                    hash.AppendData(buffer, 0, read);
                    await target.WriteAsync(buffer.AsMemory(0, read), watchdog.Token).ConfigureAwait(false);
                }

                return Convert.ToHexStringLower(hash.GetHashAndReset());
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }
    }
}
