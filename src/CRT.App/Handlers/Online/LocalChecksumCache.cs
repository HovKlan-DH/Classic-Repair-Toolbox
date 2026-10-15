using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Handlers.OnlineHandling
{
    // ###########################################################################################
    // The SHA-256 of every local data file, remembered with the file's STAMP - its length, last
    // write time and creation time - so a launch hashes only the files that changed (code review,
    // 2026-10-13: every launch read the whole 1.1 GB data tree to hash it, partly on the UI thread).
    //
    // *** IT MUST NEVER MAKE CRT TRUST A FILE IT WOULD OTHERWISE DOWNLOAD AGAIN. *** So:
    //   - a hash answers only while the file's stamp is EXACTLY the one it was taken with; any
    //     write, a replaced file or a copy over it moves the stamp, and the file is hashed again;
    //   - a file written within RacyWindow of being hashed is not remembered at all: a second write
    //     inside the file system's timestamp step would leave the stamp unchanged (git's "racy
    //     clean" problem) - it is simply hashed again next time;
    //   - a cache file that is missing, cannot be read or parsed, or was made for another data
    //     folder is an EMPTY cache: everything is hashed, as before it existed.
    //
    // Kept in CRT's AppData folder beside "Data" (AppConfig.DataChecksumCacheFileName), written
    // through AtomicJsonFile. The data folder it is for is inside it, so a "--data-root=" run
    // never takes another folder's hashes - and never writes a file into a folder of its own.
    // ###########################################################################################
    internal sealed class LocalChecksumCache
    {
        // A file written this close to its hashing is not remembered (see the header).
        internal static readonly TimeSpan RacyWindow = TimeSpan.FromSeconds(3);

        private readonly object thisSync = new();
        private readonly string thisCachePath;
        private readonly string thisDataRoot;
        private readonly Dictionary<string, Entry> thisEntries;
        private bool thisChanged;

        private LocalChecksumCache(string cachePath, string dataRoot, Dictionary<string, Entry> entries)
        {
            this.thisCachePath = cachePath;
            this.thisDataRoot = dataRoot;
            this.thisEntries = entries;
        }

        internal int Count
        {
            get
            {
                lock (this.thisSync)
                {
                    return this.thisEntries.Count;
                }
            }
        }

        internal readonly record struct FileStamp(long Length, long LastWriteUtcTicks, long CreationUtcTicks)
        {
            internal static FileStamp Of(FileInfo file) =>
                new(file.Length, file.LastWriteTimeUtc.Ticks, file.CreationTimeUtc.Ticks);
        }

        // ###########################################################################################
        // The cache in `cachePath` for the data folder `dataRoot` - empty when there is none, it
        // cannot be read, or it was made for another data folder. Never throws.
        // ###########################################################################################
        internal static LocalChecksumCache Load(string cachePath, string dataRoot)
        {
            string root = LocalChecksumCache.NormaliseRoot(dataRoot);
            var entries = new Dictionary<string, Entry>(StringComparer.Ordinal);

            try
            {
                if (!string.IsNullOrWhiteSpace(cachePath) && File.Exists(cachePath))
                {
                    var stored = JsonSerializer.Deserialize<StoredCache>(File.ReadAllText(cachePath));

                    if (stored?.Files != null && string.Equals(stored.DataRoot, root, StringComparison.Ordinal))
                    {
                        foreach (var (path, entry) in stored.Files)
                        {
                            if (path != null && entry != null && LocalChecksumCache.IsHash(entry.Sha256))
                            {
                                entries[path] = entry;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Local checksum cache could not be read - every data file is hashed - [{ex.Message}]");
                entries.Clear();
            }

            return new LocalChecksumCache(cachePath, root, entries);
        }

        // The remembered hash of the file at `relativePath`, if its stamp is still `stamp`.
        internal string? TryGet(string relativePath, FileStamp stamp)
        {
            lock (this.thisSync)
            {
                return this.thisEntries.TryGetValue(relativePath, out var entry) &&
                       entry.Length == stamp.Length &&
                       entry.LastWriteUtcTicks == stamp.LastWriteUtcTicks &&
                       entry.CreationUtcTicks == stamp.CreationUtcTicks
                    ? entry.Sha256
                    : null;
            }
        }

        // ###########################################################################################
        // Remembers `sha256` for the file as it was when hashed at `hashedAtUtc` - unless it was
        // written within RacyWindow of that, when whatever was remembered for it is forgotten.
        // ###########################################################################################
        internal void Remember(string relativePath, FileStamp stamp, string sha256, DateTime hashedAtUtc)
        {
            lock (this.thisSync)
            {
                bool racy = hashedAtUtc.Ticks - stamp.LastWriteUtcTicks < RacyWindow.Ticks;

                if (racy || !LocalChecksumCache.IsHash(sha256))
                {
                    this.thisChanged |= this.thisEntries.Remove(relativePath);
                    return;
                }

                this.thisEntries[relativePath] = new Entry
                {
                    Length = stamp.Length,
                    LastWriteUtcTicks = stamp.LastWriteUtcTicks,
                    CreationUtcTicks = stamp.CreationUtcTicks,
                    Sha256 = sha256,
                };

                this.thisChanged = true;
            }
        }

        internal void Forget(string relativePath)
        {
            lock (this.thisSync)
            {
                this.thisChanged |= this.thisEntries.Remove(relativePath);
            }
        }

        // Writes the cache when anything changed since it was loaded or last saved.
        internal void Save()
        {
            StoredCache stored;

            lock (this.thisSync)
            {
                if (!this.thisChanged || string.IsNullOrWhiteSpace(this.thisCachePath))
                {
                    return;
                }

                stored = new StoredCache
                {
                    DataRoot = this.thisDataRoot,
                    Files = new Dictionary<string, Entry>(this.thisEntries, StringComparer.Ordinal),
                };

                this.thisChanged = false;
            }

            AtomicJsonFile.Write(this.thisCachePath, stored, new JsonSerializerOptions(), "local checksum cache");
        }

        private static string NormaliseRoot(string dataRoot) =>
            string.IsNullOrWhiteSpace(dataRoot)
                ? string.Empty
                : Path.TrimEndingDirectorySeparator(Path.GetFullPath(dataRoot));

        private static bool IsHash(string? value)
        {
            if (value == null || value.Length != 64)
            {
                return false;
            }

            foreach (char c in value)
            {
                if (!((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f')))
                {
                    return false;
                }
            }

            return true;
        }

        internal sealed class Entry
        {
            [JsonPropertyName("length")] public long Length { get; set; }
            [JsonPropertyName("lastWriteUtcTicks")] public long LastWriteUtcTicks { get; set; }
            [JsonPropertyName("creationUtcTicks")] public long CreationUtcTicks { get; set; }
            [JsonPropertyName("sha256")] public string Sha256 { get; set; } = string.Empty;
        }

        internal sealed class StoredCache
        {
            [JsonPropertyName("dataRoot")] public string DataRoot { get; set; } = string.Empty;
            [JsonPropertyName("files")] public Dictionary<string, Entry>? Files { get; set; }
        }
    }
}
