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
    // ###########################################################################################
    public static class BoardKiCadFolder
    {
        public const string RecordFileName = ".crt-kicad.json";

        private static readonly ConcurrentDictionary<string, object> Locks = new(StringComparer.OrdinalIgnoreCase);

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

            string boardFolder = Path.GetDirectoryName(Path.GetFullPath(boardFile)) ?? string.Empty;
            string cache = Path.GetFullPath(cacheRoot);
            string copyFolder = Path.Combine(cache, BoardKiCadFolder.FolderNameFor(boardFile));

            lock (BoardKiCadFolder.Locks.GetOrAdd(copyFolder, _ => new object()))
            {
                try
                {
                    Dictionary<string, (string Source, string Token)> wanted = BoardKiCadFolder.Wanted(boardFile, boardFolder, dataRoot);

                    if (FolderLinks.AnyBetween(cache, Path.Combine(copyFolder, DataTreeUsage.KiCadFolderName, "x")))
                        throw new IOException("the copy's folder links elsewhere, so nothing is written or deleted there");

                    if (wanted.Count == 0)
                    {
                        if (Directory.Exists(copyFolder))
                            Directory.Delete(copyFolder, recursive: true);

                        return string.Empty;
                    }

                    BoardKiCadFolder.Copy(cache, copyFolder, wanted);

                    return Path.Combine(copyFolder, DataTreeUsage.KiCadFolderName);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
                {
                    CrtLog.Warning($"The KiCad files of [{boardFile}] could not be copied to [{copyFolder}] - [{ex.Message}]");
                    return string.Empty;
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

        // One folder per board file, named after its full path - stable across launches.
        internal static string FolderNameFor(string boardFile)
        {
            byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes(Path.GetFullPath(boardFile).ToUpperInvariant()));
            return Convert.ToHexString(hash, 0, 8).ToLowerInvariant();
        }

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
                    // *** NEW BYTES, A LATER WRITE TIME (2026-10-25). *** File.Copy keeps the source's time,
                    // and two sources written in one clock tick (a contributor's file and the published
                    // one) gave the copy the very stamp it had - so KiCadProjectLoader, which keeps its
                    // parse by path and write time, answered with the replaced file's parse.
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
