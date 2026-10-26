using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // ONE PARSE OF A WORKBOOK PER VERSION OF IT (code review, 2026-10-13).
    //
    // Coming back to CRT after an Excel edit re-read a changed draft three times on the UI thread -
    // the change count, the problem count and the fingerprint each had their own stamp cache, and
    // each parsed the workbook afresh on its miss. They now share this: a board is parsed once and
    // reused while the workbook AND its sidecar are as they were read (BoardDataReader.StampOf,
    // the reader's own rule), so an edit anywhere - Excel, CRT, a sync - is a new read.
    //
    // *** NOT FOR DELETING, NOT FOR WRITING. *** Retirement compares uncached at the moment it
    // deletes, and every draft write reads the file it replaces (DraftWorkbookStore). This only
    // spares the read-only counts the Drafts tab shows. The BoardData handed out is shared, which
    // is safe because nothing in this codebase mutates a board it was given (BoardData's header).
    //
    // One entry per drafted board and its published copy - a handful; an entry whose file moved is
    // replaced, not added beside, and one whose file is gone is dropped (ForgetVanished).
    // ###########################################################################################
    internal static class BoardReadsByStamp
    {
        private static readonly ConcurrentDictionary<string, (string Stamp, BoardData Board)> Boards =
            new(StringComparer.OrdinalIgnoreCase);

        public static BoardData? Read(string? workbookPath)
        {
            BoardReadsByStamp.ForgetVanished();

            if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(workbookPath))
                return null;

            string stamp = BoardDataReader.StampOf(workbookPath);

            // Trusted only once the files' writes have settled (StampSettling): two saves inside one
            // clock tick can leave one stamp on two contents.
            if (BoardReadsByStamp.Boards.TryGetValue(workbookPath, out var cached) && cached.Stamp == stamp &&
                StampSettling.IsSettled(File.GetLastWriteTimeUtc(workbookPath)) &&
                StampSettling.IsSettled(BoardReadsByStamp.SidecarWriteUtc(workbookPath)))
            {
                return cached.Board;
            }

            BoardData? board = BoardDataReader.ReadWorkbookUncached(workbookPath);

            // Remembered only when nothing moved WHILE it was read.
            if (board is not null && BoardDataReader.StampOf(workbookPath) == stamp)
                BoardReadsByStamp.Boards[workbookPath] = (stamp, board);

            return board;
        }

        // ###########################################################################################
        // *** A WORKBOOK THAT IS GONE IS LET GO (code review, 2026-10-14). *** A discarded or retired
        // draft's parsed board stayed here for the rest of the session - tens of megabytes for a large
        // board, one more for every draft made and discarded. Every read now drops the entries whose
        // file no longer exists: a handful of entries, each one existence check.
        // ###########################################################################################
        private static void ForgetVanished()
        {
            foreach (string path in BoardReadsByStamp.Boards.Keys)
            {
                if (!File.Exists(path))
                    BoardReadsByStamp.Boards.TryRemove(path, out _);
            }
        }

        // Whether a parsed board is kept for `workbookPath` - for tests.
        internal static bool HoldsForTests(string workbookPath) => BoardReadsByStamp.Boards.ContainsKey(workbookPath);

        private static DateTime SidecarWriteUtc(string workbookPath)
        {
            string sidecar = BoardComponentHighlightStorage.GetJsonPath(workbookPath);
            return File.Exists(sidecar) ? File.GetLastWriteTimeUtc(sidecar) : DateTime.MinValue;
        }
    }

    // ###########################################################################################
    // What a draft holds that is NOT a row (code review, 2026-10-13): files whose bytes differ from
    // what the draft was copied from, and KiCad calibrations that differ. The Drafts row used to
    // count rows only, so a draft whose only work was re-calibrating a KiCad overlay, replacing a
    // picture under its own name, or adding KiCad files read "0 rows changed" and its Submit button
    // said "There is nothing to send yet. Add some data to this board first." - while the Wiki's
    // KiCad walkthrough sends exactly such a draft.
    // ###########################################################################################
    public readonly record struct DraftOtherChanges(int Files, int Calibrations)
    {
        public bool Any => this.Files > 0 || this.Calibrations > 0;

        // "2 files changed", "KiCad calibration changed", "1 file and 2 KiCad calibrations changed".
        public string Describe()
        {
            string files = this.Files == 1 ? "1 file" : $"{this.Files} files";

            return (this.Files > 0, this.Calibrations) switch
            {
                (true, 0) => $"{files} changed",
                (true, 1) => $"{files} and a KiCad calibration changed",
                (true, _) => $"{files} and {this.Calibrations} KiCad calibrations changed",
                (false, 1) => "KiCad calibration changed",
                (false, > 1) => $"{this.Calibrations} KiCad calibrations changed",
                _ => string.Empty,
            };
        }
    }

    public static partial class DraftStatusReader
    {
        private static readonly ConcurrentDictionary<string, (string Stamp, DraftOtherChanges Changes)> OtherChangesCache =
            new(StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // The non-row changes of one draft, remembered until the draft folder (every file's length
        // and write time - the base and marker included) or the published workbook moves.
        //
        // Measured against the draft's BASE - what it was copied from - so a picture or calibration
        // the published update replaced since is not counted as the contributor's (DraftBase). A
        // draft with no base (seeded before 2026-10-13) is measured against the published copy as
        // it is now, the rule retirement uses; its published folder is then part of the stamp.
        //
        // A new board counts nothing here: everything in it is new, and its rows say so.
        // ###########################################################################################
        public static DraftOtherChanges CountOtherChangesCached(DraftStatus? status, string draftsRoot)
        {
            if (status is null || status.IsNewBoard ||
                string.IsNullOrWhiteSpace(status.WorkbookPath) || !File.Exists(status.WorkbookPath))
            {
                return default;
            }

            string folder = Path.GetDirectoryName(status.WorkbookPath) ?? string.Empty;
            string? publishedFolder = Path.GetDirectoryName(status.PublishedWorkbookPath);

            bool publishedSettled = true;

            string stamp = string.Join(
                "|",
                DraftRetirement.FolderStamp(folder, out bool draftSettled),
                BoardDataReader.StampOf(status.PublishedWorkbookPath),
                status.HasBase ? "base" : DraftRetirement.FolderStamp(publishedFolder, out publishedSettled));

            // Trusted only once every write in the folders has settled (StampSettling, code review
            // 2026-10-13), and the published workbook's too.
            if (DraftStatusReader.OtherChangesCache.TryGetValue(status.WorkbookPath, out var cached) && cached.Stamp == stamp &&
                draftSettled && publishedSettled &&
                (!File.Exists(status.PublishedWorkbookPath) || StampSettling.IsSettled(File.GetLastWriteTimeUtc(status.PublishedWorkbookPath))))
            {
                return cached.Changes;
            }

            DraftOtherChanges changes = DraftStatusReader.CountOtherChanges(status, draftsRoot, folder, publishedFolder);

            DraftStatusReader.OtherChangesCache[status.WorkbookPath] = (stamp, changes);

            return changes;
        }

        // ###########################################################################################
        // A CRT 4.0 draft's changed files (2026-10-24): its board file's list against its base's -
        // or, with no base, the published board's - each name with other bytes or one the base does
        // not have; and each new file lying in the folder, not yet taken in, unless it is what the
        // list already says. A file the draft dropped shows in its rows.
        // ###########################################################################################
        private static int ChangedListedFiles(DraftStatus status, string folder, DraftBase? based)
        {
            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? draft = DraftFiles.ListedIn(status.WorkbookPath);

            if (draft is null)
                return 0;

            Dictionary<string, string> before = based is not null
                ? based.Files
                    .GroupBy(file => file.Path, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First().Sha256, StringComparer.OrdinalIgnoreCase)
                : (DraftFiles.ListedIn(status.PublishedWorkbookPath) ?? new Dictionary<string, BoardDocument.StoredFileRow>())
                    .ToDictionary(pair => pair.Key, pair => pair.Value.Sha256, StringComparer.OrdinalIgnoreCase);

            int changed = draft.Values.Count(file =>
                !before.TryGetValue(file.Name, out string? sha) || !string.Equals(sha, file.Sha256, StringComparison.OrdinalIgnoreCase));

            foreach (string path in Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories))
            {
                string relative = Path.GetRelativePath(folder, path).Replace('\\', '/');

                if ((DataTreeLayout.IsBoardFile(relative) && !relative.Contains('/')) ||
                    relative.StartsWith(DataTreeLayout.FilesFolder + "/", StringComparison.OrdinalIgnoreCase) ||
                    DraftFingerprint.IsBookkeeping(Path.GetFileName(path)) ||
                    DraftFolderLayout.IsInBookkeepingFolder(relative))
                {
                    continue;
                }

                if (draft.TryGetValue(relative, out BoardDocument.StoredFileRow? listed) &&
                    string.Equals(DraftFingerprint.CachedFileHash(path), listed.Sha256, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                changed++;
            }

            return changed;
        }

        private static DraftOtherChanges CountOtherChanges(DraftStatus status, string draftsRoot, string folder, string? publishedFolder)
        {
            DraftBase? based = status.HasBase ? DraftBaseStore.Load(draftsRoot, status.BoardKey) : null;

            int files = 0;

            try
            {
                files = DraftStatusReader.ChangedFiles(status, folder, publishedFolder, based);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // A file that cannot be read might hold anything: count it as something to send.
                // Submit then stays available - the server's own checks are the gate.
                CrtLog.Warning($"Could not compare the files of the draft [{status.WorkbookPath}] - [{ex.Message}]");
                files = 1;
            }

            return new DraftOtherChanges(files, DraftStatusReader.ChangedCalibrations(status, based));
        }

        private static int ChangedFiles(DraftStatus status, string folder, string? publishedFolder, DraftBase? based)
        {
            if (DataTreeLayout.IsBoardFile(status.WorkbookPath))
                return DraftStatusReader.ChangedListedFiles(status, folder, based);

            string workbook = Path.GetFullPath(status.WorkbookPath);
            string sidecar = Path.GetFullPath(BoardComponentHighlightStorage.GetJsonPath(workbook));

            Dictionary<string, DraftBaseFile>? baseFiles = based?.Files
                .GroupBy(file => file.Path, StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

            int changed = 0;

            foreach (string path in Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories))
            {
                string full = Path.GetFullPath(path);

                if (string.Equals(full, workbook, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(full, sidecar, StringComparison.OrdinalIgnoreCase) ||
                    DraftFingerprint.IsBookkeeping(Path.GetFileName(full)))
                {
                    continue;
                }

                string relative = Path.GetRelativePath(folder, full).Replace('\\', '/');
                var info = new FileInfo(full);

                bool same;

                if (baseFiles is not null)
                {
                    same = baseFiles.TryGetValue(relative, out DraftBaseFile? copied) &&
                           info.Length == copied.Length &&
                           (info.LastWriteTimeUtc.Ticks == copied.DraftWriteUtcTicks ||
                            string.Equals(DraftFingerprint.CachedFileHash(full), copied.Sha256, StringComparison.Ordinal));
                }
                else
                {
                    string published = string.IsNullOrWhiteSpace(publishedFolder)
                        ? string.Empty
                        : Path.Combine(publishedFolder, relative.Replace('/', Path.DirectorySeparatorChar));

                    var publishedInfo = new FileInfo(published.Length > 0 ? published : full);

                    // ###########################################################################################
                    // *** AN UNTOUCHED COPY IS NOT READ (code review, 2026-10-14). *** This runs on the UI
                    // thread for every draft each time CRT comes to the front, and hashing both sides of
                    // every file - about 1,100 a side for a C64 - froze the window for seconds. The seeder
                    // copies a file with its write time (File.Copy keeps it, on every platform CRT runs
                    // on), so a copy with the published file's length AND write time is that file; only a
                    // pair whose stamps differ is read.
                    // ###########################################################################################
                    same = published.Length > 0 &&
                           publishedInfo.Exists &&
                           publishedInfo.Length == info.Length &&
                           (publishedInfo.LastWriteTimeUtc == info.LastWriteTimeUtc ||
                            string.Equals(DraftFingerprint.CachedFileHash(full), DraftFingerprint.CachedFileHash(published), StringComparison.Ordinal));
                }

                if (!same)
                    changed++;
            }

            return changed;
        }

        private static int ChangedCalibrations(DraftStatus status, DraftBase? based)
        {
            BoardData? draft = BoardReadsByStamp.Read(status.WorkbookPath);

            if (draft is null)
                return 0;

            IReadOnlyList<KiCadCalibrationEntry> mine = DraftBoardSource.CollectCalibrations(status.WorkbookPath, draft);
            IReadOnlyList<KiCadCalibrationEntry> theirs;

            if (based is not null)
            {
                theirs = based.Calibrations;
            }
            else
            {
                BoardData? published = BoardReadsByStamp.Read(status.PublishedWorkbookPath);
                theirs = published is null ? [] : DraftBoardSource.CollectCalibrations(status.PublishedWorkbookPath, published);
            }

            var before = new Dictionary<string, KiCadCalibrationEntry>(StringComparer.OrdinalIgnoreCase);

            foreach (KiCadCalibrationEntry entry in theirs)
                before.TryAdd(entry.SchematicName?.Trim() ?? string.Empty, entry);

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int changed = 0;

            foreach (KiCadCalibrationEntry entry in mine)
            {
                string name = entry.SchematicName?.Trim() ?? string.Empty;
                seen.Add(name);

                if (!before.TryGetValue(name, out KiCadCalibrationEntry? old) || !BoardDataDiffer.SameValues(old, entry))
                    changed++;
            }

            // A calibration the draft no longer has - its schematic still there - was removed.
            var schematics = new HashSet<string>(
                draft.Schematics.Select(schematic => schematic.SchematicName?.Trim() ?? string.Empty),
                StringComparer.OrdinalIgnoreCase);

            changed += before.Keys.Count(name => !seen.Contains(name) && schematics.Contains(name));

            return changed;
        }
    }
}
