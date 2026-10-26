using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A CRT 4.0 DRAFT'S FILES - which ones its board file lists, and how a new one gets in
    // (2026-10-24).
    //
    // A draft's board file lists every file of the board by name and SHA-256 (BoardDocument.Files);
    // the bytes of those the 4.0 tree does not hold are in the draft's own store, "Files/" in its
    // folder (DataTreeLayout.StoredFilePath). Nothing outside this class has to keep that list:
    //
    //   - A NEW FILE ARRIVES IN THE DRAFT FOLDER UNDER ITS NAME - "<draft>/Component images/HotCPU.png",
    //     where every editor already copies a picked file (DraftFileResolver.BuildDraftFileDestination).
    //     The next write of the draft finds it there, stores its bytes under their checksum and lists
    //     it. Until then DraftFileResolver finds it where it lies.
    //   - THE LIST FOLLOWS THE ROWS. A name a row cites keeps its entry (or takes a new file's); a
    //     name no row cites any more is dropped. The KiCad files are listed whatever the rows say -
    //     the published board's as the draft has them, and every file under "<draft>/KiCad data".
    //   - BYTES NOTHING LISTS ARE REMOVED from the draft's store after each write - but those of a
    //     file the draft held in memory has seen listed stay while it is held, since its undo can
    //     bring back a row citing one (DraftFileMemory).
    //
    // A name a row cites with no file anywhere gets no entry: the board's checks say it is missing.
    // ###########################################################################################
    public static class DraftFiles
    {
        // ###########################################################################################
        // The files the board file at `boardFile` should list for `board`, given what it listed
        // before - new files in the folder COPIED into the store on the way, and named in
        // `arrived`.
        //
        // *** A NEW FILE LEAVES ITS PLACE ONLY ONCE THE BOARD FILE LISTS IT. *** The caller removes
        // the `arrived` files (RemoveArrived) after the board file is written - a write that fails
        // leaves the new file where it was, still found by its name, and the next write takes it in.
        // Moved straight into the store, a failed write would leave it unlisted, and the next
        // write's tidying would delete it.
        //
        // *** `remembered`: WHAT THE DRAFT HELD IN MEMORY HAS SEEN LISTED (code review, 2026-10-26). ***
        // Its history holds rows, not this list, so an undo can bring back a row citing a file an
        // earlier write dropped - the entry it had is taken again, for a cited name the list no
        // longer has (never a KiCad file: no row cites one, and one removed stays removed). Its
        // bytes are still in the store: the session keeps them (RemoveUnlisted's `alsoKeep`).
        // ###########################################################################################
        public static List<BoardDocument.StoredFileRow> ListFor(
            string boardFile,
            BoardData board,
            IReadOnlyList<BoardDocument.StoredFileRow> previous,
            out List<string> arrived,
            IReadOnlyCollection<BoardDocument.StoredFileRow>? remembered = null)
        {
            ArgumentNullException.ThrowIfNull(board);

            string folder = Path.GetDirectoryName(Path.GetFullPath(boardFile))!;
            var before = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);

            // By the name as Normalise spells it, as every name is looked up below - kept by its own
            // spelling ("KiCad data/./x"), it was never found and went at the next save, its bytes with it
            // (code review, 2026-10-26).
            foreach (BoardDocument.StoredFileRow file in previous ?? [])
                before.TryAdd(BoardFileNames.Normalise(file.Name), file);

            var earlier = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);

            foreach (BoardDocument.StoredFileRow file in remembered ?? [])
                earlier.TryAdd(BoardFileNames.Normalise(file.Name), file);

            var listed = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);
            var arrivedPaths = new List<string>();

            void List(string? name, bool cited)
            {
                string wanted = BoardFileNames.Normalise(name);

                if (wanted.Length == 0 || listed.ContainsKey(wanted))
                    return;

                if (DraftFiles.Arrived(folder, wanted, out string path) is BoardDocument.StoredFileRow taken)
                {
                    listed[wanted] = taken;
                    arrivedPaths.Add(path);
                }
                else if (before.TryGetValue(wanted, out BoardDocument.StoredFileRow? kept) ||
                         (cited && earlier.TryGetValue(wanted, out kept)))
                {
                    listed[wanted] = kept;
                }
            }

            foreach (string name in DraftFiles.CitedNames(board))
                List(name, cited: true);

            foreach (BoardDocument.StoredFileRow kiCad in before.Values.Where(DraftFiles.IsKiCad))
                List(kiCad.Name, cited: false);

            foreach (string name in DraftFiles.ArrivedKiCadNames(folder))
                List(name, cited: false);

            arrived = arrivedPaths;

            return listed.Values
                .OrderBy(file => file.Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(file => file.Name, StringComparer.Ordinal)
                .ToList();
        }

        // ###########################################################################################
        // Removes the new files ListFor took into the store from where they arrived, with the folders
        // that leaves empty - called once the board file listing them is written. A failure is
        // untidy, never lost work: the store holds the bytes, and a copy left behind is taken in
        // again (to the same stored bytes) by the next write.
        // ###########################################################################################
        public static void RemoveArrived(string boardFile, IReadOnlyList<string> arrived)
        {
            string folder = Path.GetDirectoryName(Path.GetFullPath(boardFile))!;

            foreach (string path in arrived ?? [])
            {
                try
                {
                    File.Delete(path);
                    DraftFiles.RemoveEmptyFolders(folder, Path.GetDirectoryName(path)!);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    CrtLog.Warning($"Could not tidy away [{path}] after taking it into the draft - [{ex.Message}]");
                }
            }
        }

        // Every file name the board's rows cite, in the order it shows them.
        public static IEnumerable<string> CitedNames(BoardData board) =>
            board.Schematics.Select(row => row.SchematicImageFile)
                .Concat(board.ComponentImages.Select(row => row.File))
                .Concat(board.ComponentLocalFiles.Select(row => row.File))
                .Concat(board.BoardLocalFiles.Select(row => row.File))
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Select(name => name.Trim());

        // Whether a listed file is KiCad data - SubmissionKiCadFiles.IsInKiCadFolder, the one rule.
        public static bool IsKiCad(BoardDocument.StoredFileRow file) =>
            SubmissionKiCadFiles.IsInKiCadFolder(file.Name);

        // ###########################################################################################
        // The KiCad files lying in "<draft>/KiCad data" that the draft's next write takes in, by name
        // ("KiCad data/Pages/vic.kicad_sch"), sorted - where the "KiCad data" window's import copies a
        // project. ListFor lists them, and a submission carries them even before a write has
        // (SubmissionKiCadFiles.Collect).
        // ###########################################################################################
        public static List<string> ArrivedKiCadNames(string draftFolder)
        {
            if (string.IsNullOrWhiteSpace(draftFolder))
                return [];

            string folder = Path.GetFullPath(draftFolder);
            string kiCadFolder = Path.Combine(folder, DataTreeUsage.KiCadFolderName);

            if (!Directory.Exists(kiCadFolder))
                return [];

            return Directory.EnumerateFiles(kiCadFolder, "*", SearchOption.AllDirectories)
                .Select(path => Path.GetRelativePath(folder, path).Replace('\\', '/'))
                .Where(DataChecksumManifest.IsSyncable)
                .Order(StringComparer.Ordinal)
                .ToList();
        }

        // ###########################################################################################
        // Where a new file named `name` lies in the draft folder `draftFolder` - or false when the
        // name cannot be one: not shaped as a board's file name (rooted, a ".." segment, a backslash -
        // SubmissionPathRules.IsSafelyShaped), resolving outside the folder, reached through a linked
        // folder, or naming the store, the board file or CRT's bookkeeping.
        //
        // *** A NAME COMES FROM A ROW, AND A ROW CAN SAY ANYTHING (code review, 2026-10-25). *** A file
        // found here is copied into the store and then DELETED from where it lay (RemoveArrived): a
        // File cell holding "C:\Users\me\Pictures\u8.png", or "../<another board>/board.json", had the
        // contributor's own file outside the draft deleted at the next save. Resolved first, then
        // checked for containment - the rule SubmissionPathRules.TryResolve keeps for the server - and
        // no folder on the way may be a link, through which a delete reaches the folder linked to.
        // ###########################################################################################
        public static bool TryArrivedPath(string draftFolder, string name, out string path)
        {
            path = string.Empty;

            if (string.IsNullOrWhiteSpace(draftFolder) || string.IsNullOrWhiteSpace(name) || DraftFiles.IsReserved(name))
                return false;

            if (!SubmissionPathRules.TryResolve(draftFolder, name, out string resolved, out _) ||
                FolderLinks.AnyBetween(draftFolder, resolved))
            {
                return false;
            }

            path = resolved;
            return true;
        }

        // ###########################################################################################
        // Removes from the draft's store every stored file `listed` does not name, and the folders
        // that leaves empty - all but the bytes of `alsoKeep`, what the draft held in memory has seen
        // listed and its history may cite again (ListFor's `remembered`). A failure is untidy, never a
        // failed write.
        // ###########################################################################################
        public static void RemoveUnlisted(
            string boardFile,
            IReadOnlyList<BoardDocument.StoredFileRow> listed,
            IReadOnlyCollection<BoardDocument.StoredFileRow>? alsoKeep = null)
        {
            string store = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(boardFile))!, DataTreeLayout.FilesFolder);

            if (!Directory.Exists(store))
                return;

            HashSet<string> kept = DraftFiles.StoredPaths(Path.GetDirectoryName(store)!, listed.Concat(alsoKeep ?? []));

            try
            {
                foreach (string path in Directory.EnumerateFiles(store, "*", SearchOption.AllDirectories).ToList())
                {
                    if (!kept.Contains(Path.GetFullPath(path)))
                        File.Delete(path);
                }

                foreach (string folder in Directory.EnumerateDirectories(store, "*", SearchOption.AllDirectories).OrderByDescending(path => path.Length).ToList())
                {
                    if (!Directory.EnumerateFileSystemEntries(folder).Any())
                        Directory.Delete(folder);
                }

                if (!Directory.EnumerateFileSystemEntries(store).Any())
                    Directory.Delete(store);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not tidy the draft's stored files in [{store}] - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Where the bytes of `files` lie in the store of the draft folder `draftFolder`, as full paths
        // (ignoring capitals) - the store's own files, for the tidying above and for retirement
        // (DraftRetirement: bytes the board file does not list are no work of the contributor's).
        // ###########################################################################################
        public static HashSet<string> StoredPaths(string draftFolder, IEnumerable<BoardDocument.StoredFileRow> files) =>
            new(
                files.Where(file => DataTreeLayout.IsSha256(file.Sha256))
                    .Select(file => Path.GetFullPath(Path.Combine(draftFolder, DataTreeLayout.StoredFilePath(file.Sha256, file.Name)))),
                StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // A new file lying in the draft folder under its name (`path`): its bytes copied into the
        // draft's store, and its entry - or null when there is none, or the name cannot be one
        // (TryArrivedPath: outside the folder, the store, CRT's own bookkeeping). The original stays
        // until the board file lists it (ListFor).
        // ###########################################################################################
        private static BoardDocument.StoredFileRow? Arrived(string folder, string name, out string path)
        {
            if (!DraftFiles.TryArrivedPath(folder, name, out path) || !File.Exists(path))
                return null;

            string sha = DataTreeLayout.Sha256Of(path);
            long size = new FileInfo(path).Length;
            string stored = Path.Combine(folder, DataTreeLayout.StoredFilePath(sha, name).Replace('/', Path.DirectorySeparatorChar));

            if (!File.Exists(stored))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(stored)!);

                // Copied under a temporary name and moved into place, so a stored file is never
                // half-written under its checksum.
                string temporary = stored + DraftFolderLayout.BookkeepingPrefix + "copying";
                File.Copy(path, temporary, overwrite: true);
                File.Move(temporary, stored, overwrite: true);
            }

            return new BoardDocument.StoredFileRow { Name = name, Sha256 = sha, Size = size };
        }

        // The store, the board file, and CRT's ".crt-" bookkeeping are never a contributor's file.
        private static bool IsReserved(string name)
        {
            string first = name.Split('/')[0];

            return string.Equals(first, DataTreeLayout.FilesFolder, StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, DataTreeLayout.BoardFileName, StringComparison.OrdinalIgnoreCase)
                || first.StartsWith(DraftFolderLayout.BookkeepingPrefix, StringComparison.OrdinalIgnoreCase);
        }

        // The folders a moved file leaves empty, up to (not including) the draft folder.
        private static void RemoveEmptyFolders(string draftFolder, string from)
        {
            string root = Path.GetFullPath(draftFolder);
            string? folder = Path.GetFullPath(from);

            while (folder is not null &&
                   folder.Length > root.Length &&
                   folder.StartsWith(root, StringComparison.OrdinalIgnoreCase) &&
                   Directory.Exists(folder) &&
                   !Directory.EnumerateFileSystemEntries(folder).Any())
            {
                Directory.Delete(folder);
                folder = Path.GetDirectoryName(folder);
            }
        }

        // ---- reading a board file's list, for resolving names ----------------------------------

        private static readonly ConcurrentDictionary<string, (string Stamp, IReadOnlyDictionary<string, BoardDocument.StoredFileRow> Files)> Lists =
            new(StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // The files the board file at `boardFile` lists, by name (ignoring capitals) - remembered
        // while the file is as it was read (its stamp, settled), so resolving a board's dozens of
        // files reads it once. Null when there is no board file there or it cannot be read.
        // ###########################################################################################
        public static IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? ListedIn(string? boardFile)
        {
            if (string.IsNullOrWhiteSpace(boardFile) || !File.Exists(boardFile))
                return null;

            string stamp = BoardDataReader.FileStamp(boardFile);

            if (DraftFiles.Lists.TryGetValue(boardFile, out var known) && known.Stamp == stamp &&
                StampSettling.IsSettled(File.GetLastWriteTimeUtc(boardFile)))
            {
                return known.Files;
            }

            if (BoardFiles.ReadDocument(boardFile) is not BoardDocument document)
                return null;

            var files = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);

            foreach (BoardDocument.StoredFileRow file in document.Files)
                files.TryAdd(file.Name, file);

            if (BoardDataReader.FileStamp(boardFile) == stamp)
                DraftFiles.Lists[boardFile] = (stamp, files);

            return files;
        }
    }

    // ###########################################################################################
    // THE FILES A DRAFT'S BOARD FILE HAS LISTED WHILE THE DRAFT WAS HELD IN MEMORY (code review,
    // 2026-10-26) - DraftSession notes them at every read and write of the file, by name, the latest
    // entry of a name winning. Its history holds rows, not the list, so this is what lets an undo
    // bring back a row citing a file a write in between dropped (DraftFiles.ListFor's `remembered`),
    // with its bytes still in the store (RemoveUnlisted's `alsoKeep`). Kept for the run; forgotten
    // with the history. And the list as it was last read or written (Listed), which an editor
    // compares with the published board's without reading the file again. Any thread.
    // ###########################################################################################
    internal sealed class DraftFileMemory
    {
        private readonly object thisGate = new();
        private readonly Dictionary<string, BoardDocument.StoredFileRow> thisFiles = new(StringComparer.OrdinalIgnoreCase);
        private IReadOnlyDictionary<string, BoardDocument.StoredFileRow> thisListed = new Dictionary<string, BoardDocument.StoredFileRow>();

        public IReadOnlyCollection<BoardDocument.StoredFileRow> Files
        {
            get
            {
                lock (this.thisGate)
                    return this.thisFiles.Values.ToList();
            }
        }

        public IReadOnlyDictionary<string, BoardDocument.StoredFileRow> Listed
        {
            get
            {
                lock (this.thisGate)
                    return this.thisListed;
            }
        }

        // `files`: the board file's whole list, as just read or written.
        public void Note(IReadOnlyCollection<BoardDocument.StoredFileRow> files)
        {
            var listed = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);

            foreach (BoardDocument.StoredFileRow file in files)
                listed.TryAdd(file.Name, file);

            lock (this.thisGate)
            {
                foreach (BoardDocument.StoredFileRow file in files)
                    this.thisFiles[BoardFileNames.Normalise(file.Name)] = file;

                this.thisListed = listed;
            }
        }

        public void Forget()
        {
            lock (this.thisGate)
                this.thisFiles.Clear();
        }
    }
}
