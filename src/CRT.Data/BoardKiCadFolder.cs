using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A CRT 4.0 BOARD'S KiCad PROJECT AS A FOLDER (2026-10-24).
    //
    // A board file lists its KiCad files by name ("KiCad data/Open128.kicad_pcb",
    // "KiCad data/Pages/vic.kicad_sch") and the store keeps their bytes under their checksum. KiCad's
    // own files point at each other by NAME - a root sheet names its pages "Pages/vic.kicad_sch" -
    // so the loaders need them side by side under their names. This copies them into a folder of
    // their own in CRT's cache, one per board file, and answers that folder's "KiCad data" for
    // KiCadRawFileScanner.
    //
    // A file is copied again only when its entry changed (the copy's own record, RecordFileName);
    // a file no entry names any more is removed. A DRAFT's new KiCad files, lying in
    // "<draft>/KiCad data" until the draft's next write takes them into its store (DraftFiles), are
    // copied too and win over a listed file of the same name. A listed file's bytes are looked for
    // in the board folder's own store first (a draft's), then in the data tree's.
    //
    // *** A NAME IS NEVER TRUSTED TO STAY INSIDE THE FOLDER. *** Only names under "KiCad data/"
    // that resolve inside the copy's folder are copied - a board file can come from anywhere a
    // draft folder can - and nothing is written or deleted through a folder that links elsewhere
    // (FolderLinks): the copy deletes every file it does not want.
    //
    // *** ONLY WHAT THE LOADERS READ, AND NOTHING TWICE (code review, 2026-10-25). *** The files of
    // the types KiCadRawFileScanner reads (ComponentListBuilder.IsSupportedKiCadRawFile) - a board's
    // 21 MB generated "KiCad-traces.json" was copied for nobody. The board file's list is read
    // through DraftFiles.ListedIn (parsed once while the file stays as it is) and the record written
    // only when it changed, so asking again with nothing changed touches no file. The first copy of
    // a large board can take seconds: callers ask off the UI thread.
    //
    // *** A COPY IS NAMED BY WHAT IT HOLDS, NEVER BY THE BOARD FILE (2026-11-16, performance
    // walkthrough). *** One folder per set of names and bytes (FolderNameOf): a draft lists the
    // published KiCad files as they are, so a board and its draft - the first change on a board
    // makes one - share ONE copy, and KiCadProjectLoader, which keeps its parse by path and write
    // time, ONE parse; named by the board file's path, the first change copied Open128's 45 MB PCB
    // again and parsed it for seconds. Other bytes are another folder.
    //
    // *** A COPY GOES AS SOON AS NO BOARD'S KiCad FILES ARE WHAT IT HOLDS (owner request, 2026-11-16:
    // "a temporary file should be deleted when it is not used anymore" - one nothing had used for 14
    // days went until then). *** Which board file each copy holds the KiCad files of is kept beside
    // them (OwnersFileName): a board file whose KiCad files change gets its new copy and lets go of
    // the old one, which goes at once unless another board file still holds it; a discarded draft lets
    // go of its own (LetGoOf); and the sweep CRT runs after launch lets go of every board file gone or
    // changed since, and takes away every copy no board file holds - and the folders named by path
    // that CRT 4.0's earlier pre-releases left. So the cache holds one copy of the KiCad files of each
    // board looked at, as they now are - kept between runs, as copying a large board's again at every
    // launch takes seconds. A copy that cannot be taken away now (a file of it open) goes at the next.
    // ###########################################################################################
    public static class BoardKiCadFolder
    {
        public const string RecordFileName = ".crt-kicad.json";

        // Which board file holds which copy: its full path, and its copy's folder name.
        public const string OwnersFileName = ".crt-kicad-owners.json";

        // Every folder named by content starts so.
        public const string FolderPrefix = "v2-";

        private static readonly ConcurrentDictionary<string, object> Locks = new(StringComparer.OrdinalIgnoreCase);

        // Around every read and write of the owners file - taken inside a copy's lock, never the other way round.
        private static readonly object OwnersGate = new();

        // ###########################################################################################
        // The folder holding the board's KiCad files under their names - "<copy>/KiCad data" - or ""
        // when the board names none. `boardFile` is a board.json (published or a draft's),
        // `dataRoot` the tree whose store holds the published bytes, `cacheRoot` where copies are
        // kept. Never throws: a board whose KiCad files cannot be copied draws no traces, as one
        // with none.
        // ###########################################################################################
        public static string For(string boardFile, string dataRoot, string cacheRoot)
        {
            if (string.IsNullOrWhiteSpace(boardFile) || string.IsNullOrWhiteSpace(cacheRoot))
                return string.Empty;

            string fullBoardFile = Path.GetFullPath(boardFile);
            string boardFolder = Path.GetDirectoryName(fullBoardFile) ?? string.Empty;
            string cache = Path.GetFullPath(cacheRoot);
            string copyFolder = cache;

            try
            {
                Dictionary<string, (string Source, string Token)> wanted = BoardKiCadFolder.Wanted(boardFile, boardFolder, dataRoot);

                if (wanted.Count == 0)
                {
                    BoardKiCadFolder.LetGoOf(fullBoardFile, cache);
                    return string.Empty;
                }

                string folderName = BoardKiCadFolder.FolderNameOf(wanted);
                copyFolder = Path.Combine(cache, folderName);

                // Held by this board file from now on - the copy it held before let go of (below).
                string? letGo = BoardKiCadFolder.Hold(cache, fullBoardFile, folderName);
                string kiCadFolder;

                lock (BoardKiCadFolder.Locks.GetOrAdd(copyFolder, _ => new object()))
                {
                    if (FolderLinks.AnyBetween(cache, Path.Combine(copyFolder, DataTreeUsage.KiCadFolderName, "x")))
                        throw new IOException("the copy's folder links elsewhere, so nothing is written or deleted there");

                    BoardKiCadFolder.Copy(cache, copyFolder, wanted);
                    kiCadFolder = Path.Combine(copyFolder, DataTreeUsage.KiCadFolderName);
                }

                // The board file's KiCad files changed: its copy before goes, unless another board file holds it.
                if (letGo is not null)
                    BoardKiCadFolder.TakeAwayUnheld(cache, letGo);

                return kiCadFolder;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
            {
                CrtLog.Warning($"The KiCad files of [{boardFile}] could not be copied to [{copyFolder}] - [{ex.Message}]");
                return string.Empty;
            }
        }

        // ###########################################################################################
        // A board file's copy let go of - a draft discarded - and taken away at once unless another board file holds it
        // (the class header). Nothing when the board file holds none. Never throws.
        // ###########################################################################################
        public static void LetGoOf(string boardFile, string cacheRoot)
        {
            if (string.IsNullOrWhiteSpace(boardFile) || string.IsNullOrWhiteSpace(cacheRoot))
                return;

            try
            {
                string cache = Path.GetFullPath(cacheRoot);
                string? letGo = null;

                lock (BoardKiCadFolder.OwnersGate)
                {
                    Dictionary<string, string> owners = BoardKiCadFolder.ReadOwners(cache);

                    if (owners.Remove(Path.GetFullPath(boardFile), out string? held))
                    {
                        BoardKiCadFolder.WriteOwners(cache, owners);
                        letGo = held;
                    }
                }

                if (letGo is not null)
                    BoardKiCadFolder.TakeAwayUnheld(cache, letGo);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
            {
                CrtLog.Warning($"The KiCad copy of [{boardFile}] could not be let go of - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // `boardFile` holds `folderName` from now on - written only when that is news. Answers the copy it held before when no
        // board file holds that any more, for the caller to take away; null otherwise.
        // ###########################################################################################
        private static string? Hold(string cache, string boardFile, string folderName)
        {
            lock (BoardKiCadFolder.OwnersGate)
            {
                Dictionary<string, string> owners = BoardKiCadFolder.ReadOwners(cache);

                if (owners.TryGetValue(boardFile, out string? before) && string.Equals(before, folderName, StringComparison.OrdinalIgnoreCase))
                    return null;

                owners[boardFile] = folderName;

                // Not written, the copy is still made and drawn from: the next launch's sweep finds it held by nobody and takes
                // it away, to be made again as its board is shown.
                try
                {
                    BoardKiCadFolder.WriteOwners(cache, owners);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    CrtLog.Info($"Which board holds a KiCad copy could not be written: [{cache}] - [{ex.Message}]");
                    return null;
                }

                return before is not null && !owners.Values.Contains(before, StringComparer.OrdinalIgnoreCase) ? before : null;
            }
        }

        // ###########################################################################################
        // The copy `folderName` taken away - unless a board file holds it by now (asked again under its own lock, so a board
        // asking for it meanwhile makes it again after), it links elsewhere, or a file of it is open (left for the next
        // launch's sweep). Whether it went.
        // ###########################################################################################
        private static bool TakeAwayUnheld(string cache, string folderName)
        {
            string folder = Path.Combine(cache, folderName);

            lock (BoardKiCadFolder.Locks.GetOrAdd(folder, _ => new object()))
            {
                lock (BoardKiCadFolder.OwnersGate)
                {
                    if (BoardKiCadFolder.ReadOwners(cache).Values.Contains(folderName, StringComparer.OrdinalIgnoreCase))
                        return false;
                }

                try
                {
                    if (!Directory.Exists(folder) || FolderLinks.AnyBetween(cache, Path.Combine(folder, DataTreeUsage.KiCadFolderName, "x")))
                        return false;

                    Directory.Delete(folder, recursive: true);
                    return true;
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    CrtLog.Info($"A KiCad copy nothing uses could not be taken away now, so the next launch will: [{folder}] - [{ex.Message}]");
                    return false;
                }
            }
        }

        // The owners file: each board file's full path and the folder name of its copy - empty when there is none, or it cannot
        // be read (the sweep then takes away every copy, each made again as its board is shown).
        private static Dictionary<string, string> ReadOwners(string cache)
        {
            string path = Path.Combine(cache, BoardKiCadFolder.OwnersFileName);

            try
            {
                if (File.Exists(path) &&
                    JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(path)) is { } stored)
                {
                    return new Dictionary<string, string>(stored, StringComparer.OrdinalIgnoreCase);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
            {
                CrtLog.Info($"The KiCad copies' owners could not be read: [{path}] - [{ex.Message}]");
            }

            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        // Written whole, swapped into place.
        private static void WriteOwners(string cache, Dictionary<string, string> owners)
        {
            Directory.CreateDirectory(cache);

            string path = Path.Combine(cache, BoardKiCadFolder.OwnersFileName);
            string temporary = path + ".crt-writing";

            File.WriteAllText(temporary, JsonSerializer.Serialize(owners));
            File.Move(temporary, path, overwrite: true);
        }

        // ###########################################################################################
        // THE SWEEP, after launch (the class header): every board file that is gone, or whose KiCad files are no longer what
        // its copy holds (a data update, a draft's import while CRT was closed), lets go of its copy; then every copy no board
        // file holds is taken away, and every folder named by a board file's path (an earlier pre-release's, 16 hex digits) -
        // never through a folder that links elsewhere, and nothing at the cache's top but CRT's folders. `dataRoot` is the
        // tree whose store holds the published bytes, as For takes it. Never throws: a copy that cannot be taken away now (a
        // file of it open) goes at the next launch. Answers how many went.
        // ###########################################################################################
        public static int Sweep(string cacheRoot, string dataRoot)
        {
            if (string.IsNullOrWhiteSpace(cacheRoot))
                return 0;

            int removed = 0;

            try
            {
                string cache = Path.GetFullPath(cacheRoot);

                if (!Directory.Exists(cache))
                    return 0;

                BoardKiCadFolder.LetGoOfWhatChanged(cache, dataRoot);

                foreach (string folder in Directory.EnumerateDirectories(cache).ToList())
                {
                    string name = Path.GetFileName(folder);

                    // Named by a board file's path: nothing copies into those any more.
                    if (BoardKiCadFolder.IsNamedByPath(name))
                    {
                        try
                        {
                            if (!FolderLinks.AnyBetween(cache, Path.Combine(folder, DataTreeUsage.KiCadFolderName, "x")))
                            {
                                Directory.Delete(folder, recursive: true);
                                removed++;
                            }
                        }
                        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                        {
                            CrtLog.Info($"An earlier KiCad copy could not be taken away now, so the next launch will: [{folder}] - [{ex.Message}]");
                        }

                        continue;
                    }

                    if (name.StartsWith(BoardKiCadFolder.FolderPrefix, StringComparison.Ordinal) && BoardKiCadFolder.TakeAwayUnheld(cache, name))
                        removed++;
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"The KiCad copies in [{cacheRoot}] could not be looked through - [{ex.Message}]");
            }

            return removed;
        }

        // ###########################################################################################
        // Every board file that is gone, or whose KiCad files are no longer what its copy holds, lets go of its copy. Worked
        // out OUTSIDE the owners' lock - each board file read - and let go of under it only while it still holds that copy
        // (a board shown meanwhile has its new copy already). A board file that cannot be read keeps its copy.
        // ###########################################################################################
        private static void LetGoOfWhatChanged(string cache, string dataRoot)
        {
            Dictionary<string, string> owners;

            lock (BoardKiCadFolder.OwnersGate)
                owners = BoardKiCadFolder.ReadOwners(cache);

            var stale = new List<(string BoardFile, string Folder)>();

            foreach ((string boardFile, string folder) in owners)
            {
                try
                {
                    if (!File.Exists(boardFile))
                    {
                        stale.Add((boardFile, folder));
                        continue;
                    }

                    Dictionary<string, (string Source, string Token)> wanted =
                        BoardKiCadFolder.Wanted(boardFile, Path.GetDirectoryName(boardFile) ?? string.Empty, dataRoot);

                    if (wanted.Count == 0 || !string.Equals(BoardKiCadFolder.FolderNameOf(wanted), folder, StringComparison.OrdinalIgnoreCase))
                        stale.Add((boardFile, folder));
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
                {
                    CrtLog.Info($"Whether the KiCad copy of [{boardFile}] is still its own could not be told - [{ex.Message}]");
                }
            }

            if (stale.Count == 0)
                return;

            lock (BoardKiCadFolder.OwnersGate)
            {
                Dictionary<string, string> now = BoardKiCadFolder.ReadOwners(cache);
                bool changed = false;

                foreach ((string boardFile, string folder) in stale)
                {
                    if (now.TryGetValue(boardFile, out string? held) && string.Equals(held, folder, StringComparison.OrdinalIgnoreCase))
                        changed |= now.Remove(boardFile);
                }

                if (changed)
                {
                    try
                    {
                        BoardKiCadFolder.WriteOwners(cache, now);
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                    {
                        CrtLog.Info($"Which board holds a KiCad copy could not be written: [{cache}] - [{ex.Message}]");
                    }
                }
            }
        }

        // ###########################################################################################
        // The KiCad folder of the board CRT identifies as `boardIdentity`, as it is SHOWN: a draft is
        // the whole board - its file lists the published KiCad files as well as its own - so with a
        // draft (`draftFolder` holding a board file) only the draft's, else the published board's.
        // ###########################################################################################
        public static string ForBoard(string dataRoot, string? draftFolder, string boardIdentity, string cacheRoot)
        {
            string draftBoardFile = string.IsNullOrWhiteSpace(draftFolder)
                ? string.Empty
                : Path.Combine(draftFolder, DataTreeLayout.BoardFileName);

            string boardFile = File.Exists(draftBoardFile)
                ? draftBoardFile
                : DraftBoardSource.PublishedPathOf(dataRoot, boardIdentity);

            return BoardKiCadFolder.For(boardFile, dataRoot, cacheRoot);
        }

        // ###########################################################################################
        // The KiCad files a draft brought itself, by name ("KiCad data/Pages/vic.kicad_sch"): those
        // its board file lists with other bytes than the published board's - or that the published
        // board does not have - and every new file lying in "<draft>/KiCad data", not yet taken in.
        // What the "KiCad data" window lists as imported, and what its Remove takes away; the
        // published board's own KiCad files, which a seeded draft lists too, are not among them.
        // ###########################################################################################
        public static List<string> ImportedNames(string draftBoardFile, string publishedBoardFile)
        {
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow> published =
                DraftFiles.ListedIn(publishedBoardFile) ?? new Dictionary<string, BoardDocument.StoredFileRow>();

            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (DraftFiles.ListedIn(draftBoardFile) is { } draft)
            {
                foreach (BoardDocument.StoredFileRow file in draft.Values.Where(DraftFiles.IsKiCad))
                {
                    if (!published.TryGetValue(file.Name, out BoardDocument.StoredFileRow? theirs) ||
                        !string.Equals(theirs.Sha256, file.Sha256, StringComparison.OrdinalIgnoreCase))
                    {
                        names.Add(file.Name);
                    }
                }
            }

            string draftFolder = Path.GetDirectoryName(Path.GetFullPath(draftBoardFile)) ?? string.Empty;

            foreach (string name in DraftFiles.ArrivedKiCadNames(draftFolder))
                names.Add(BoardFileNames.Normalise(name));

            return names.Order(StringComparer.OrdinalIgnoreCase).ToList();
        }

        // The names of the KiCad files a board file lists of the types the loaders read - what is
        // copied of a board with no draft, counted without copying anything.
        public static List<string> ListedNames(string boardFile) =>
            (DraftFiles.ListedIn(boardFile)?.Values ?? [])
                .Where(BoardKiCadFolder.IsRead)
                .Select(file => file.Name)
                .Order(StringComparer.OrdinalIgnoreCase)
                .ToList();

        // A listed file the loaders read: KiCad data of a type KiCadRawFileScanner takes.
        private static bool IsRead(BoardDocument.StoredFileRow file) =>
            DraftFiles.IsKiCad(file) && ComponentListBuilder.IsSupportedKiCadRawFile(file.Name);

        // ###########################################################################################
        // One folder per set of names and bytes: the same files under the same names - whichever
        // board file lists them - are the same folder, in any order and any capitals of a name.
        // ###########################################################################################
        internal static string FolderNameOf(IReadOnlyDictionary<string, (string Source, string Token)> wanted)
        {
            var lines = new StringBuilder();

            foreach ((string name, (string _, string token)) in wanted.OrderBy(pair => pair.Key.ToUpperInvariant(), StringComparer.Ordinal))
                lines.Append(name.ToUpperInvariant()).Append('\u001F').Append(token).Append('\n');

            byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(lines.ToString()));
            return BoardKiCadFolder.FolderPrefix + Convert.ToHexString(hash, 0, 8).ToLowerInvariant();
        }

        // An earlier pre-release's folder, named after a board file's path: 16 hex digits.
        private static bool IsNamedByPath(string name) =>
            name.Length == 16 && name.All(Uri.IsHexDigit);

        // ###########################################################################################
        // Every KiCad file the board has, by name: where its bytes are and a token that changes
        // when they do (the checksum of a listed file, size and time of a new one).
        // ###########################################################################################
        private static Dictionary<string, (string Source, string Token)> Wanted(string boardFile, string boardFolder, string dataRoot)
        {
            var wanted = new Dictionary<string, (string Source, string Token)>(StringComparer.OrdinalIgnoreCase);

            foreach (BoardDocument.StoredFileRow file in (DraftFiles.ListedIn(boardFile)?.Values ?? []).Where(BoardKiCadFolder.IsRead))
            {
                if (!DataTreeLayout.IsSha256(file.Sha256))
                    continue;

                string stored = DataTreeLayout.StoredFilePath(file.Sha256, file.Name).Replace('/', Path.DirectorySeparatorChar);
                string own = Path.Combine(boardFolder, stored);
                string tree = string.IsNullOrWhiteSpace(dataRoot) ? string.Empty : Path.Combine(dataRoot, stored);
                string source = File.Exists(own) ? own : File.Exists(tree) ? tree : string.Empty;

                if (source.Length > 0)
                    wanted[BoardFileNames.Normalise(file.Name)] = (source, file.Sha256.ToLowerInvariant());
            }

            foreach (string name in DraftFiles.ArrivedKiCadNames(boardFolder).Where(ComponentListBuilder.IsSupportedKiCadRawFile))
            {
                string path = Path.Combine(boardFolder, name.Replace('/', Path.DirectorySeparatorChar));
                var info = new FileInfo(path);
                wanted[BoardFileNames.Normalise(name)] = (path, $"{info.Length}:{info.LastWriteTimeUtc.Ticks}");
            }

            return wanted;
        }

        private static void Copy(string cache, string copyFolder, Dictionary<string, (string Source, string Token)> wanted)
        {
            string fullCopyFolder = Path.TrimEndingDirectorySeparator(copyFolder) + Path.DirectorySeparatorChar;
            string recordPath = Path.Combine(copyFolder, BoardKiCadFolder.RecordFileName);
            Dictionary<string, string> record = BoardKiCadFolder.ReadRecord(recordPath);
            var kept = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var keptPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            Directory.CreateDirectory(copyFolder);

            foreach ((string name, (string source, string token)) in wanted)
            {
                if (!name.StartsWith(DataTreeUsage.KiCadFolderName + "/", StringComparison.OrdinalIgnoreCase))
                    continue;

                string target = Path.GetFullPath(Path.Combine(copyFolder, name.Replace('/', Path.DirectorySeparatorChar)));

                if (!target.StartsWith(fullCopyFolder, StringComparison.OrdinalIgnoreCase) || FolderLinks.AnyBetween(cache, target))
                {
                    CrtLog.Warning($"A KiCad file name pointing outside its folder was not copied: [{name}]");
                    continue;
                }

                if (!(record.TryGetValue(name, out string? had) && had == token && File.Exists(target)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                    string temporary = target + ".crt-copying";
                    DateTime? replaced = File.Exists(target) ? File.GetLastWriteTimeUtc(target) : null;

                    File.Copy(source, temporary, overwrite: true);

                    // ###########################################################################################
                    // *** A COPY MADE AGAIN, A LATER WRITE TIME (2026-10-25). *** File.Copy keeps the source's
                    // time, and two sources written in one clock tick (a contributor's file and the published
                    // one) gave the copy the very stamp it had - so KiCadProjectLoader, which keeps its parse
                    // by path and write time, answered with the replaced file's parse. Since folders are named
                    // by what they hold, other bytes go to another folder; this stays for a copy made again
                    // over one whose record was lost, which a stamp of its own costs nothing.
                    // ###########################################################################################
                    if (replaced is { } previous && File.GetLastWriteTimeUtc(temporary) <= previous)
                        File.SetLastWriteTimeUtc(temporary, previous.AddTicks(1));

                    File.Move(temporary, target, overwrite: true);
                }

                kept[name] = token;
                keptPaths.Add(target);
            }

            string kiCadFolder = Path.Combine(copyFolder, DataTreeUsage.KiCadFolderName);

            if (Directory.Exists(kiCadFolder))
            {
                // Never through a folder that links elsewhere - only the copy's own files go.
                foreach (string path in Directory.EnumerateFiles(kiCadFolder, "*", SearchOption.AllDirectories).ToList())
                {
                    if (!keptPaths.Contains(Path.GetFullPath(path)) && !FolderLinks.AnyBetween(cache, path))
                        File.Delete(path);
                }

                foreach (string folder in Directory.EnumerateDirectories(kiCadFolder, "*", SearchOption.AllDirectories).OrderByDescending(path => path.Length).ToList())
                {
                    if (!FolderLinks.AnyBetween(cache, Path.Combine(folder, "x")) && !Directory.EnumerateFileSystemEntries(folder).Any())
                        Directory.Delete(folder);
                }
            }

            // Written only when it changed: asking again with nothing changed writes nothing.
            if (record.Count != kept.Count || kept.Any(pair => !record.TryGetValue(pair.Key, out string? had) || had != pair.Value))
                File.WriteAllText(recordPath, JsonSerializer.Serialize(kept));
        }

        private static Dictionary<string, string> ReadRecord(string recordPath)
        {
            try
            {
                if (File.Exists(recordPath) &&
                    JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(recordPath)) is { } stored)
                {
                    return new Dictionary<string, string>(stored, StringComparer.OrdinalIgnoreCase);
                }
            }
            catch (JsonException)
            {
                // A damaged record copies everything again.
            }

            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
    }
}
