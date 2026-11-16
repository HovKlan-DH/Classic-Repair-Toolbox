using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** WHAT A DATA CHECK DOWNLOADS FOLLOWS "Visible hardware and boards" (owner request, 2026-11-14: "will it also be
    // possible to limit the synched data, so it honors this list, meaning that it will not download any data for the
    // non-selected entries? If then clicked/checked, and it does not have the data, then it should inform about this").
    // *** A board unticked on the "Configuration" tab - itself, or its hardware - is not downloaded: neither its board file
    // nor the stored files only it names. Everything else is as it always was:
    //
    //   - the main file, which lists every board, and every board file of a board ticked;
    //   - the stored files ("Files/<ab>/<sha256>.<ext>") the ticked boards' board files and the drafts' name - one shared
    //     with an unticked board comes for the ticked one;
    //   - every other file of the tree (the MiniPro tests).
    //
    // What an unticked board has on this computer already STAYS - the orphan cleanup maps the whole manifest - and is
    // simply not brought up to date. A board ticked again whose data is missing or behind is said by Main's notice
    // (Main.DataNeededNotice.cs), whose button runs the data check that then takes it in.
    //
    // With nothing unticked, nothing changes: every board file and every stored file, as before.
    // ###########################################################################################
    internal sealed class DataSyncScope
    {
        private DataSyncScope(IReadOnlySet<string> boardFiles, IReadOnlySet<string> skippedBoardFiles)
        {
            this.BoardFiles = boardFiles;
            this.SkippedBoardFiles = skippedBoardFiles;
        }

        // The board files to download - data-root relative, "/"-separated, as the manifest names them.
        public IReadOnlySet<string> BoardFiles { get; }

        // The published boards' files left out, their boards unticked.
        public IReadOnlySet<string> SkippedBoardFiles { get; }

        public bool SkipsAny => this.SkippedBoardFiles.Count > 0;

        // ###########################################################################################
        // The scope for the boards the main file lists. A board only a draft has is never downloaded - its file is an
        // identity, never on the server (NewBoardIdentity). A board file listed twice, once ticked, is downloaded.
        // `normalize` is the data check's own path rule, so a path compares as the manifest's does.
        // ###########################################################################################
        public static DataSyncScope Of(IEnumerable<HardwareBoardEntry> boards, IReadOnlySet<string> uncheckedKeys, Func<string, string> normalize)
        {
            ArgumentNullException.ThrowIfNull(boards);
            ArgumentNullException.ThrowIfNull(uncheckedKeys);
            ArgumentNullException.ThrowIfNull(normalize);

            var wanted = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var skipped = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (HardwareBoardEntry entry in boards.Where(entry => !entry.IsDraftOnly))
            {
                string file = normalize(entry.ExcelDataFile);

                if (file.Length == 0)
                    continue;

                if (CatalogueVisibility.IsBoardVisible(uncheckedKeys, entry.HardwareName, entry.BoardName))
                {
                    wanted.Add(file);
                }
                else
                {
                    skipped.Add(file);
                }
            }

            skipped.ExceptWith(wanted);
            return new DataSyncScope(wanted, skipped);
        }

        // Whether a data-root relative path is a file of the tree's store - which the boards name by checksum.
        public static bool IsStoredFile(string relativePath) =>
            relativePath.StartsWith(DataTreeLayout.FilesFolder + "/", StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // Whether the background part of a check downloads `relativePath` - any file but the lists (IsListFile, the
        // launch's own), a stored file only when `storedWanted` names it. Null `storedWanted` is every stored file: with
        // nothing skipped, or when what the ticked boards name could not all be read (failing open - a board shown
        // without its pictures is worse than a download too many).
        // ###########################################################################################
        public static bool DownloadsInBackground(string relativePath, Func<string, bool> isListFile, IReadOnlySet<string>? storedWanted) =>
            !isListFile(relativePath) &&
            (storedWanted is null || !DataSyncScope.IsStoredFile(relativePath) || storedWanted.Contains(relativePath));
    }
}
