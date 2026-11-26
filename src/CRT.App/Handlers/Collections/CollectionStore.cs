using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using Handlers.DataHandling;

namespace Handlers.Collections
{
    // ###########################################################################################
    // THE OSCILLOSCOPE COLLECTIONS ON THIS COMPUTER (2026-11-26) - one folder each under the root (AppConfig.CollectionsFolder
    // for CRT, a folder of its own for a test), named by its local id: "collection.json" and its captures as "<sha256>.png".
    //
    // A capture is kept once a collection - the same picture added twice is one. Every path is made from a local id of the
    // right shape and a checksum, never from anything else, so nothing leads out of the root.
    // ###########################################################################################
    public sealed class CollectionStore
    {
        public const string FileName = "collection.json";

        private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

        private readonly string thisRoot;

        public CollectionStore(string root)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(root);
            this.thisRoot = Path.GetFullPath(root);
        }

        public string Root => this.thisRoot;

        // A local id: a GUID without dashes, lower case - the folder's name.
        public static bool IsLocalId(string? id) =>
            id is { Length: 32 } && id.All(c => c is (>= '0' and <= '9') or (>= 'a' and <= 'f'));

        // ###########################################################################################
        // Every collection kept, newest first - one whose file cannot be read is left out (and said once in the log).
        // ###########################################################################################
        public IReadOnlyList<LocalCollection> All()
        {
            if (!Directory.Exists(this.thisRoot))
                return [];

            var collections = new List<LocalCollection>();

            foreach (string folder in Directory.EnumerateDirectories(this.thisRoot))
            {
                if (CollectionStore.IsLocalId(Path.GetFileName(folder)) && this.Read(Path.GetFileName(folder)) is { } collection)
                    collections.Add(collection);
            }

            return collections.OrderByDescending(collection => collection.CreatedUtc).ToList();
        }

        // The board's collections, newest first.
        public IReadOnlyList<LocalCollection> OfBoard(string hardwareName, string boardName) =>
            this.All()
                .Where(collection =>
                    string.Equals(collection.HardwareName, hardwareName, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(collection.BoardName, boardName, StringComparison.OrdinalIgnoreCase))
                .ToList();

        public LocalCollection? Find(string? localId) => CollectionStore.IsLocalId(localId) ? this.Read(localId!) : null;

        // ###########################################################################################
        // A new, empty collection of the board - kept at once.
        // ###########################################################################################
        public LocalCollection? Create(string hardwareName, string boardName, string? region, string? title, DateTimeOffset now)
        {
            var collection = new LocalCollection
            {
                LocalId = Guid.NewGuid().ToString("N"),
                HardwareName = hardwareName,
                BoardName = boardName,
                Region = string.IsNullOrWhiteSpace(region) ? null : region,
                Title = string.IsNullOrWhiteSpace(title) ? null : title,
                CreatedUtc = now
            };

            return this.Save(collection) ? collection : null;
        }

        public bool Save(LocalCollection collection)
        {
            ArgumentNullException.ThrowIfNull(collection);

            if (!CollectionStore.IsLocalId(collection.LocalId))
                return false;

            Directory.CreateDirectory(this.FolderOf(collection.LocalId));

            return AtomicJsonFile.Write(
                Path.Combine(this.FolderOf(collection.LocalId), CollectionStore.FileName),
                collection,
                CollectionStore.JsonOptions,
                $"oscilloscope collection [{collection.Title ?? collection.LocalId}]");
        }

        // ###########################################################################################
        // A capture added: its PNG kept under its checksum, the capture last in the list - none added twice. The collection as
        // it now is, or null when it could not be kept.
        // ###########################################################################################
        public LocalCollection? AddCapture(string localId, byte[] png, LocalCapture capture)
        {
            ArgumentNullException.ThrowIfNull(png);
            ArgumentNullException.ThrowIfNull(capture);

            if (this.Find(localId) is not { } collection)
                return null;

            string sha256 = Convert.ToHexStringLower(SHA256.HashData(png));

            if (collection.Captures.Any(kept => kept.Sha256 == sha256))
                return collection;

            try
            {
                string path = this.CapturePath(localId, sha256);

                if (!File.Exists(path))
                    File.WriteAllBytes(path, png);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Could not keep a capture in oscilloscope collection [{collection.Title ?? localId}]: {ex.Message}");
                return null;
            }

            LocalCollection changed = collection with { Captures = [.. collection.Captures, capture with { Sha256 = sha256 }] };

            return this.Save(changed) ? changed : null;
        }

        // A capture taken out - and its file with it.
        public LocalCollection? RemoveCapture(string localId, string sha256)
        {
            if (this.Find(localId) is not { } collection || !DataTreeLayout.IsSha256(sha256))
                return null;

            LocalCollection changed = collection with { Captures = collection.Captures.Where(kept => kept.Sha256 != sha256).ToList() };

            if (!this.Save(changed))
                return null;

            try
            {
                File.Delete(this.CapturePath(localId, sha256));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Could not delete a capture taken out of oscilloscope collection [{collection.Title ?? localId}]: {ex.Message}");
            }

            return changed;
        }

        public bool Delete(string localId)
        {
            if (!CollectionStore.IsLocalId(localId) || !Directory.Exists(this.FolderOf(localId)))
                return false;

            try
            {
                Directory.Delete(this.FolderOf(localId), recursive: true);
                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Could not delete oscilloscope collection [{localId}]: {ex.Message}");
                return false;
            }
        }

        public string CapturePath(string localId, string sha256)
        {
            if (!CollectionStore.IsLocalId(localId) || !DataTreeLayout.IsSha256(sha256))
                throw new ArgumentException("No local collection id or capture checksum.");

            return Path.Combine(this.FolderOf(localId), sha256 + ".png");
        }

        // A capture's bytes, or null when they are not there.
        public byte[]? ReadCapture(string localId, string sha256)
        {
            try
            {
                return File.ReadAllBytes(this.CapturePath(localId, sha256));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
            {
                return null;
            }
        }

        private string FolderOf(string localId) => Path.Combine(this.thisRoot, localId);

        private LocalCollection? Read(string localId)
        {
            string path = Path.Combine(this.FolderOf(localId), CollectionStore.FileName);

            if (!File.Exists(path))
                return null;

            try
            {
                LocalCollection? collection = JsonSerializer.Deserialize<LocalCollection>(File.ReadAllText(path), CollectionStore.JsonOptions);

                // The folder's name is the id, whatever the file says.
                return collection is null ? null : collection with { LocalId = localId };
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
            {
                UnreadableFiles.Report(path, $"Could not read oscilloscope collection [{path}]: {ex.Message}");
                return null;
            }
        }
    }
}
