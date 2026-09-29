using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // BUILDS dataChecksums.json - THE FILE EVERY CRT CLIENT SYNCS AGAINST
    // (owner request, 2026-09-23).
    //
    // *** WITHOUT THIS, PUBLISHING IS INVISIBLE TO EVERY USER. *** CRT decides what to download by
    // comparing this manifest's checksums against what it already holds. ApprovePublishFlow wrote
    // the board into the data tree and never touched the manifest, so the manifest went on
    // advertising the OLD checksum - and every client, correctly, concluded there was nothing to
    // fetch. Reported after the first real publish: the board on the server was right, the
    // project owner's sync completed successfully, and the change never arrived.
    //
    // The manifest sits ONE FOLDER UP from the data root (beside `Data/`, not inside it), so it is
    // configured separately and cannot be derived - see ServerOptions.ManifestPath.
    //
    // *** THIS IS A PORT OF THE REVIEW TOOL'S PHP GENERATOR, NOT THE SERVER'S. *** There are two
    // and they disagree. `app-data-BETA/dataGenerate.php` is a plain recursive scan: no junk
    // filtering, no explicit sort, and a non-atomic write. `review/function_file-ops.php`'s
    // regenerateDataChecksumsManifest filters dot-files and OS junk, sorts by path, forces
    // lowercase and writes atomically. The stricter one is the one worth keeping, and its four
    // rules are each load-bearing:
    //
    //   - JUNK IS SKIPPED. Thumbs.db, desktop.ini, anything dot-prefixed and any leftover
    //     .tmp_NNN from an interrupted write. A client that syncs those downloads files it can
    //     never use, and a .tmp_ entry names a file that is about to vanish.
    //   - SORTED BY PATH. The live manifest is only sorted today because scandir happened to
    //     return it that way. A stable order means two regenerations of unchanged data produce an
    //     identical file, which is what makes "did anything actually change" answerable.
    //   - LOWERCASE HEX. A hash differing only in case compares unequal on the client, which
    //     reads as every file being permanently out of date.
    //   - WRITTEN ATOMICALLY. A client fetching a half-written manifest aborts its sync. That is
    //     safe (OnlineServices returns null and reports a failure) but it is a real interruption,
    //     and temp-plus-move costs nothing to avoid it.
    //
    // PURE except for reading the tree and writing one file. The SCAN is separated from the WRITE
    // so the rules above are unit tested without a filesystem write.
    // ###########################################################################################
    public static class DataChecksumManifest
    {
        // ###########################################################################################
        // One row, matching CRT.App's DataFileEntry exactly - the same three property names in the
        // same lowercase spelling.
        //
        // *** THIS IS A WIRE CONTRACT WITH EVERY INSTALLED COPY OF CRT, including ones that will
        // never be updated. *** A renamed property here is not a refactor; it is a manifest that
        // older clients silently read as empty. The names are fixed by JsonPropertyName rather
        // than left to a serializer policy so that changing a C# property name cannot change them.
        // ###########################################################################################
        public sealed class Entry
        {
            [JsonPropertyName("file")] public string File { get; init; } = string.Empty;

            [JsonPropertyName("checksum")] public string Checksum { get; init; } = string.Empty;

            [JsonPropertyName("url")] public string Url { get; init; } = string.Empty;
        }

        // Indented, because this file is read by humans diagnosing a sync as often as by the app.
        // Slash escaping is left at the default: the PHP generator escapes them ("https:\/\/") and
        // System.Text.Json does not, which is a difference in the BYTES and not in the JSON - both
        // parse identically, and the client deserialises rather than string-matching.
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true
        };

        // ###########################################################################################
        // Every file under the data root, as manifest rows.
        //
        // `publicBaseUrl` is the public address of the DATA ROOT itself
        // (https://.../app-data-BETA/Data), and each path segment is escaped individually - a space
        // becomes %20 - so a board folder with spaces in its name resolves. Escaping the whole
        // relative path in one go would turn its slashes into %2F and every URL would 404.
        //
        // Throws nothing for an absent root: an empty list is returned, and the caller decides
        // whether that is a failure. Writing an empty manifest is refused in Write below, which is
        // where that judgement belongs.
        // ###########################################################################################
        public static IReadOnlyList<Entry> Scan(string dataRoot, string publicBaseUrl)
        {
            if (string.IsNullOrWhiteSpace(dataRoot) || !Directory.Exists(dataRoot))
            {
                return [];
            }

            string root = Path.GetFullPath(dataRoot).TrimEnd(Path.DirectorySeparatorChar);
            string baseUrl = (publicBaseUrl ?? string.Empty).TrimEnd('/');

            var entries = new List<Entry>();

            foreach (string path in Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories))
            {
                string relative = Path.GetRelativePath(root, path).Replace('\\', '/');

                if (!DataChecksumManifest.IsSyncable(relative))
                {
                    continue;
                }

                entries.Add(new Entry
                {
                    File = relative,
                    Checksum = DataChecksumManifest.HashOf(path),
                    Url = baseUrl + "/" + string.Join(
                        "/",
                        relative.Split('/').Select(Uri.EscapeDataString)),
                });
            }

            // Ordinal, matching the PHP's strcmp and the case-sensitive tree the server runs on.
            return entries
                .OrderBy(entry => entry.File, StringComparer.Ordinal)
                .ToList();
        }

        // ###########################################################################################
        // Does this relative path belong in the sync set?
        //
        // Checked SEGMENT BY SEGMENT, not against the whole path, so a dot-folder excludes
        // everything beneath it - ".git/config" is skipped because of the ".git", not because the
        // file itself is dotted. That is what the PHP does and it is the behaviour that matters:
        // a dot-directory's contents are never ordinary data.
        // ###########################################################################################
        public static bool IsSyncable(string? relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
            {
                return false;
            }

            foreach (string segment in relativePath.Replace('\\', '/').Split('/'))
            {
                if (segment.Length == 0 || segment[0] == '.')
                {
                    return false;
                }

                if (string.Equals(segment, "Thumbs.db", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(segment, "desktop.ini", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                // A leftover ".tmp_1758648000" from an interrupted write. Listing one publishes a
                // checksum for a file that is about to disappear.
                if (segment.Contains(".tmp_", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }

        // ###########################################################################################
        // Scans and writes the manifest, atomically.
        //
        // *** IT REFUSES TO WRITE AN EMPTY MANIFEST, and that guard is the important one here. ***
        // An empty result means the data root was wrong, unreadable, or momentarily unmounted - and
        // writing it would tell every client that the entire data tree had been deleted. Keeping
        // the previous manifest is always the better failure: clients simply stay on what they
        // have. (A genuinely empty tree is not a case worth supporting; there is no such server.)
        //
        // Returns the number of entries written, or -1 when nothing was written. The caller logs.
        //
        // ###########################################################################################
        // *** IT THROWS NOTHING, AND THE SCAN IS INSIDE THE GUARD - which it was not at first, and
        // that shipped a 500 (owner report, 2026-09-23). ***
        //
        // The first version wrapped only the WRITE, while the scan ran outside it, under a comment
        // claiming the method could not throw. Walking a real tree of ~11,000 files can throw
        // plenty that has nothing to do with writing: a folder deleted mid-enumeration
        // (DirectoryNotFoundException), a path the service cannot traverse, a file locked while it
        // is being hashed. Every one of those escaped into the endpoint and answered 500 - AFTER
        // the board had already been published, so the maintainer was told the publish failed when
        // it had in fact succeeded. That is the exact failure the "must not fail the publish"
        // reasoning existed to prevent, defeated by putting the try in the wrong place.
        //
        // Catching Exception rather than a list of types is deliberate here, and is the same
        // judgement SubmissionNotifier makes: this runs after an irreversible operation, so the
        // only acceptable outcome for ANY fault is a logged warning and a stale manifest.
        //
        // *** ONE REBUILD AT A TIME, EACH WITH ITS OWN TEMPORARY FILE (code review, 2026-09-27). ***
        // Several server paths rebuild the BETA manifest - a publish, a rollback, the unused files
        // screen, and saving a new system's place in the lists, which runs OUTSIDE the publish lock.
        // The temporary file used to be named by the Unix second, so two rebuilds in the same second
        // wrote the SAME file: the writes interleaved, or one File.Move found its file already moved.
        // Now each carries a random name, and WriteGate makes a whole scan-and-write wait for the
        // previous one - so the rebuild that finishes last also scanned last, and the manifest never
        // ends up describing an older tree than a rebuild that had already finished.
        // ###########################################################################################
        private static readonly object WriteGate = new();

        public static int Write(string dataRoot, string publicBaseUrl, string manifestPath)
        {
            if (string.IsNullOrWhiteSpace(manifestPath))
            {
                return -1;
            }

            lock (DataChecksumManifest.WriteGate)
            {
                return DataChecksumManifest.WriteOnce(dataRoot, publicBaseUrl, manifestPath);
            }
        }

        private static int WriteOnce(string dataRoot, string publicBaseUrl, string manifestPath)
        {
            // ".tmp_" is what Scan skips as junk, so a temporary file left by a crash is never listed.
            string temporary = manifestPath + ".tmp_" + Guid.NewGuid().ToString("N");

            try
            {
                IReadOnlyList<Entry> entries = DataChecksumManifest.Scan(dataRoot, publicBaseUrl);

                if (entries.Count == 0)
                {
                    return -1;
                }

                string json = JsonSerializer.Serialize(entries, DataChecksumManifest.JsonOptions);

                string? directory = Path.GetDirectoryName(manifestPath);

                if (!string.IsNullOrWhiteSpace(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllText(temporary, json);
                File.Move(temporary, manifestPath, overwrite: true);

                return entries.Count;
            }
            catch (Exception ex)
            {
                // Deliberately every exception - see the header. This runs after an irreversible
                // publish, so there is no fault whose right answer is "throw at the maintainer".
                // Logged with the full exception rather than just its message, because a stale
                // manifest is diagnosed from this line and the type is usually the whole answer.
                CrtLog.Warning($"Could not write the checksum manifest [{manifestPath}] - [{ex}]");

                DataChecksumManifest.TryDelete(temporary);

                return -1;
            }
        }

        private static string HashOf(string path)
        {
            using FileStream stream = new(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                bufferSize: 1 << 16);

            return Convert.ToHexStringLower(SHA256.HashData(stream));
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // A stranded temp file is untidy, not harmful - and the real failure has already
                // been reported.
                CrtLog.Warning($"Could not remove the temporary manifest file [{path}] - [{ex.Message}]");
            }
        }
    }
}
