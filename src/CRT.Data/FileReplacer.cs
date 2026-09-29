using System;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WRITES A FILE BY REPLACING IT, never by opening the old one for writing (owner report,
    // 2026-09-28).
    //
    // *** WHY: A FILE COPIED IN BY HAND. *** The data trees are group-writable folders
    // (DEPLOYMENT.md step 3), but a file copied into them as root arrives owned by root with no
    // group write - so the service may DELETE or REPLACE it (that is the folder's permission) but
    // not OPEN it for writing (that is the file's). A publish that wrote the board's highlight
    // file in place was refused on exactly that after production's files had been copied into
    // BETA by hand, half-way through the board: the workbook beside it, which EPPlus happens to
    // delete and recreate, had already been written.
    //
    // So the new content goes into a temporary file BESIDE the target and is renamed over it. A
    // rename needs only the folder, so who owns the old file stops mattering - and a reader never
    // sees half a file, which a CRT syncing mid-write otherwise could.
    //
    // The temporary is a DOT-name: nothing publishes, promotes or lists one (the checksum
    // manifest skips dot-segments), so one left by a crash is litter, never data. Removed on
    // every path out of here.
    //
    // Used by every writer of a board file that is not already rename-based - BoardSidecarWriter
    // and BoardWorkbookWriter; blobs and copies go through the server's VerifiedFileCopy, which
    // works the same way.
    // ###########################################################################################
    public static class FileReplacer
    {
        // The content as the new `path`, UTF-8 without a byte order mark - what File.WriteAllText
        // writes, so a file replaced here is byte-identical to one written the old way.
        public static void WriteAllText(string path, string contents)
        {
            FileReplacer.Replace(path, temporary => File.WriteAllText(temporary, contents));
        }

        // ###########################################################################################
        // `write` produces the new content at the temporary path it is given; that file then
        // replaces `path`. The folder must exist. Throws what the write or the rename throws -
        // IOException or UnauthorizedAccessException for a folder that may not be written - with
        // the target left as it was.
        // ###########################################################################################
        public static void Replace(string path, Action<string> write)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(path);
            ArgumentNullException.ThrowIfNull(write);

            string target = Path.GetFullPath(path);
            string temporary = FileReplacer.TemporaryPathFor(target);

            try
            {
                write(temporary);
                File.Move(temporary, target, overwrite: true);
            }
            finally
            {
                FileReplacer.TryDelete(temporary);
            }
        }

        // Beside the target, so the rename never crosses a file system, and unique, so two writers
        // of one file cannot share a temporary.
        internal static string TemporaryPathFor(string target) =>
            Path.Combine(
                Path.GetDirectoryName(target) ?? string.Empty,
                $".{Path.GetFileName(target)}.{Guid.NewGuid():N}.writing.tmp");

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // A stray dot-named temporary is harmless - see the header.
            }
        }
    }
}
