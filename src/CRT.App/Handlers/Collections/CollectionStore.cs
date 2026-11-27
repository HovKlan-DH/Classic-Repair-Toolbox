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
        public static bool IsLocalId(string? id) => CollectionRules.IsClientKey(id);

        // ###########################################################################################
        // Every collection kept, newest first - one whose file cannot be read is left out (and said once in the log).
        // ###########################################################################################
        // ###########################################################################################
        // *** KEPT UNTIL THE STORE WRITES AGAIN (code review, 2026-11-27). *** Every component shown in a popup asks for the
        // board's collections - twice - and every collection's file was read and parsed on the UI thread each time. CRT makes
        // one store for the run, and every change of a collection goes through it (thisGeneration moves with each), so the list
        // read last stands until then. A file changed by hand is seen at the next launch.
        // ###########################################################################################
        private int thisGeneration;

        private (int Generation, IReadOnlyList<LocalCollection> Collections)? thisAll;

        public IReadOnlyList<LocalCollection> All()
        {
            lock (this.thisLock)
            {
                if (this.thisAll is { } kept && kept.Generation == this.thisGeneration)
                    return kept.Collections;

                var collections = new List<LocalCollection>();

                // A collection whose file is there and could not be read is no reason to keep the list - it may be held a
                // moment (code review, 2026-11-27: left out, it stayed left out until the next write).
                bool complete = true;

                if (Directory.Exists(this.thisRoot))
                {
                    foreach (string folder in Directory.EnumerateDirectories(this.thisRoot))
                    {
                        string localId = Path.GetFileName(folder);

                        if (!CollectionStore.IsLocalId(localId))
                            continue;

                        if (this.Read(localId) is { } collection)
                            collections.Add(collection);
                        else if (File.Exists(Path.Combine(folder, CollectionStore.FileName)))
                            complete = false;
                    }
                }

                IReadOnlyList<LocalCollection> sorted = collections.OrderByDescending(collection => collection.CreatedUtc).ToList();

                this.thisAll = complete ? (this.thisGeneration, sorted) : null;
                return sorted;
            }
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
        // A new board renamed ("Board: Name"): its collections take the new name, as everything else this machine keeps of it
        // does (Main.AfterBoardRenamed). A shared one's page follows at its next "Share". How many were renamed.
        // ###########################################################################################
        public int RenameBoard(string fromHardware, string fromBoard, string toHardware, string toBoard)
        {
            int renamed = 0;

            foreach (LocalCollection collection in this.OfBoard(fromHardware, fromBoard))
            {
                this.Update(collection.LocalId, current => current with { HardwareName = toHardware, BoardName = toBoard }, out bool kept);

                if (kept)
                    renamed++;
            }

            return renamed;
        }

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

        // ###########################################################################################
        // *** EVERY CHANGE OF A COLLECTION IS READ, MADE AND KEPT UNDER ONE LOCK (code review, 2026-11-27). *** A capture goes in
        // off the UI thread (the component info window) while the collection's window keeps its edits, the comments read and
        // what a share came to on the UI thread - each read the collection, changed it and saved it, and the later save threw
        // the other's change away: a capture added was gone from the list, its picture left in the folder. So the window
        // changes it through Update, and AddCapture and RemoveCapture hold the same lock.
        // ###########################################################################################
        private readonly object thisLock = new();

        public bool Save(LocalCollection collection)
        {
            ArgumentNullException.ThrowIfNull(collection);

            if (!CollectionStore.IsLocalId(collection.LocalId))
                return false;

            lock (this.thisLock)
            {
                // Whatever the write comes to, the list read before is no longer known to be the folder's.
                this.thisGeneration++;

                Directory.CreateDirectory(this.FolderOf(collection.LocalId));

                return AtomicJsonFile.Write(
                    Path.Combine(this.FolderOf(collection.LocalId), CollectionStore.FileName),
                    collection,
                    CollectionStore.JsonOptions,
                    $"oscilloscope collection [{collection.Title ?? collection.LocalId}]");
            }
        }

        // The collection as `change` makes it of the one kept NOW - kept, unless `change` hands back the very one it was given
        // (nothing to keep). Null when there is no such collection; `kept` false when the change could not be kept.
        public LocalCollection? Update(string localId, Func<LocalCollection, LocalCollection> change, out bool kept)
        {
            ArgumentNullException.ThrowIfNull(change);

            lock (this.thisLock)
            {
                kept = true;

                if (this.Find(localId) is not { } collection)
                    return null;

                LocalCollection changed = change(collection);

                if (ReferenceEquals(changed, collection))
                    return collection;

                kept = this.Save(changed);
                return kept ? changed : collection;
            }
        }

        // ###########################################################################################
        // A capture added: its PNG kept under its checksum, the capture last in the list - none added twice. What came of it
        // (CaptureAdded): added, or ALREADY IN - the very same screen, two ground pins at the same settings, is one capture, so
        // the popup offers no "Take out" for a capture it did not add (code review, 2026-11-27: it took the earlier one out) - or
        // refused by the rule the server shares by (CollectionPng: a screen too large is said now, not at "Share"), or not kept.
        // ###########################################################################################
        public CaptureAdded AddCapture(string localId, byte[] png, LocalCapture capture)
        {
            ArgumentNullException.ThrowIfNull(png);
            ArgumentNullException.ThrowIfNull(capture);

            // The checksum and the picture's check - inflating up to 12 MB of pixels - before the lock, which the UI thread takes
            // for every list and save (code review, 2026-11-27).
            string sha256 = Convert.ToHexStringLower(SHA256.HashData(png));
            CollectionPng.Result checkedPng = CollectionPng.Check(png);

            lock (this.thisLock)
            {
                return this.AddCaptureLocked(localId, png, sha256, checkedPng, capture);
            }
        }

        private CaptureAdded AddCaptureLocked(string localId, byte[] png, string sha256, CollectionPng.Result checkedPng, LocalCapture capture)
        {
            if (this.Find(localId) is not { } collection)
                return new CaptureAdded(CaptureAddedKind.NotKept, null);

            if (collection.Captures.Any(kept => kept.Sha256 == sha256))
                return new CaptureAdded(CaptureAddedKind.AlreadyIn, collection);

            if (!checkedPng.IsCapture)
                return new CaptureAdded(CaptureAddedKind.Refused, collection, checkedPng.Problem ?? CollectionPng.NotAPicture);

            string path = this.CapturePath(localId, sha256);
            bool written = false;

            try
            {
                if (!File.Exists(path))
                {
                    File.WriteAllBytes(path, png);
                    written = true;
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Could not keep a capture in oscilloscope collection [{collection.Title ?? localId}]: {ex.Message}");
                return new CaptureAdded(CaptureAddedKind.NotKept, collection);
            }

            LocalCollection changed = collection with { Captures = [.. collection.Captures, capture with { Sha256 = sha256 }] };

            if (this.Save(changed))
                return new CaptureAdded(CaptureAddedKind.Added, changed);

            // Not listed: the picture just written goes too, never left in the folder named by no capture (code review,
            // 2026-11-27).
            if (written)
                CollectionStore.DeleteQuietly(path);

            return new CaptureAdded(CaptureAddedKind.NotKept, collection);
        }

        private static void DeleteQuietly(string path)
        {
            try
            {
                File.Delete(path);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Could not delete a capture that was not kept [{path}]: {ex.Message}");
            }
        }

        // A capture taken out - and its file with it.
        public LocalCollection? RemoveCapture(string localId, string sha256)
        {
            lock (this.thisLock)
            {
                return this.RemoveCaptureLocked(localId, sha256);
            }
        }

        private LocalCollection? RemoveCaptureLocked(string localId, string sha256)
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

        // Under the store's lock as every change is (code review, 2026-11-27: a capture going in off the UI thread as the
        // collection was deleted made its folder again, the collection back with it).
        public bool Delete(string localId)
        {
            if (!CollectionStore.IsLocalId(localId))
                return false;

            lock (this.thisLock)
            {
                if (!Directory.Exists(this.FolderOf(localId)))
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
                finally
                {
                    // A delete that failed halfway - its file gone, a picture held - changed the folder all the same.
                    this.thisGeneration++;
                }
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

        // ###########################################################################################
        // *** THE COLLECTIONS ARE KEPT PER SERVICE (code review, 2026-11-27), as everything CRT remembers of a service is
        // (ContributionServer.FileNameFor): a collection's server id is one database's, so a pre-release's "Delete collection"
        // on a release's collection asked the test instance, was answered 404 and deleted it here, its live page left up. ***
        // The 4.1.0 pre-releases - the only CRTs that ever kept collections, all shared with the test instance - kept them in
        // the folder a release now uses; such a pre-release moves that folder to its own the first time, as long as it has none.
        // Never once 4.1.0 is released: a later pre-release would take the release's collections away from it.
        // ###########################################################################################
        public static bool TakesOverEarlierFolder(bool testServer, string appVersion) =>
            testServer && CrtVersion.TryParse(appVersion, out CrtVersion? version) && version! < CrtVersion.Parse("4.1.0");

        // ###########################################################################################
        // `earlier` moved to `root` - whole while there is no `root` yet, else EACH COLLECTION IT STILL HOLDS that `root` has not,
        // and the folder then taken away once empty (code review, 2026-11-27: a move refused once - a picture held open - and a
        // collection made since left the rest behind for good, test-instance ids in the folder a release reads). Tried at every
        // launch of such a pre-release; what cannot be moved is said in the log and tried again next time. True when anything
        // moved.
        // ###########################################################################################
        public static bool TakeOverEarlierFolder(string earlier, string root)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(earlier);
            ArgumentException.ThrowIfNullOrWhiteSpace(root);

            if (!Directory.Exists(earlier) || File.Exists(root))
                return false;

            try
            {
                if (!Directory.Exists(root))
                {
                    Directory.Move(earlier, root);
                    Logger.Info($"Moved the oscilloscope collections of this pre-release from [{earlier}] to [{root}]");
                    return true;
                }

                bool moved = false;

                foreach (string folder in Directory.EnumerateDirectories(earlier).ToList())
                {
                    string name = Path.GetFileName(folder);
                    string target = Path.Combine(root, name);

                    if (!CollectionStore.IsLocalId(name) || Directory.Exists(target))
                        continue;

                    try
                    {
                        Directory.Move(folder, target);
                        moved = true;
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                    {
                        Logger.Warning($"Could not move the oscilloscope collection [{folder}] to [{target}]: {ex.Message}");
                    }
                }

                if (!Directory.EnumerateFileSystemEntries(earlier).Any())
                    Directory.Delete(earlier);

                return moved;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Could not move the oscilloscope collections from [{earlier}] to [{root}]: {ex.Message}");
                return false;
            }
        }
    }

    public enum CaptureAddedKind
    {
        Added,
        AlreadyIn,
        Refused,
        NotKept
    }

    // What adding a capture came to - the collection as it now is (null when there is none), and why one was refused.
    public sealed record CaptureAdded(CaptureAddedKind Kind, LocalCollection? Collection, string? Problem = null);
}
