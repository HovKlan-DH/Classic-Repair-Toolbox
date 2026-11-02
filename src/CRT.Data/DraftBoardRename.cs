using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace Handlers.DataHandling
{
    // What a rename came to - why not, in the contributor's words, when it did not.
    public sealed record DraftRenameResult(bool Renamed, string Reason = "");

    // ###########################################################################################
    // A NEW BOARD'S NAME CHANGED (owner request, 2026-11-01: "When I add a new board, then I get all the
    // input fields and I can add the new board - but how to edit those details again afterwards").
    //
    // A new board's manufacturer, hardware and board ARE its identity - "<Manufacturer>/<Hardware>/
    // <Board>/board.json" (NewBoardIdentity) - and its draft's folder is that path. So a new name moves
    // the folder, whole, its stored files and files not taken in yet with it; then the marker names the
    // new identity and registration, and the board file its new names. Notes alone move nothing.
    //
    //   - ONLY A BOARD THAT EXISTS ONLY AS A DRAFT. A published board's name is how every CRT and the
    //     server know it.
    //   - NEVER OVER ANYTHING: a folder already at the new place refuses the rename.
    //   - ALL OR NOTHING: a step that fails puts everything back, and says why.
    //   - THE DRAFT HELD IN MEMORY IS THE CALLER'S TO SAVE FIRST, AND TO CLOSE ONCE THIS HAS MOVED IT: it
    //     knows its file by the old identity (DraftSessions). Not before - a rename refused or put back
    //     changes nothing, and must not cost the draft its Undo steps. Notes alone leave the board file
    //     as it is, so the draft held in memory goes on as it was.
    //   - A change of capitals only is a rename too, through a name in between: Windows and macOS see
    //     the two as one folder.
    //   - The folders the old path leaves empty go - a manufacturer with no board left is no folder.
    // ###########################################################################################
    public static class DraftBoardRename
    {
        public static DraftRenameResult Run(string draftsRoot, string identity, NewBoardRegistration wanted) =>
            DraftBoardRename.Run(draftsRoot, identity, wanted, Thread.Sleep);

        // `afterMove`: for tests, what happens between the folder's move and its marker's write.
        internal static DraftRenameResult Run(string draftsRoot, string identity, NewBoardRegistration wanted, Action<TimeSpan> pause, Action? afterMove = null)
        {
            ArgumentNullException.ThrowIfNull(wanted);

            string oldFolder = DraftFolderLayout.GetBoardFolder(draftsRoot, identity);
            string newFolder = DraftFolderLayout.GetBoardFolder(draftsRoot, wanted.ExcelDataFile);

            if (oldFolder.Length == 0 || newFolder.Length == 0)
                return new DraftRenameResult(false, "The draft's folder could not be worked out.");

            string oldMarkerPath = DraftFolderLayout.GetMarkerPath(draftsRoot, identity);

            if (DraftMarkerStore.Load(oldMarkerPath) is not { NewBoard: not null } marker)
                return new DraftRenameResult(false, "Only a board that exists only as your draft can be renamed.");

            bool sameFolder = string.Equals(Path.GetFullPath(oldFolder), Path.GetFullPath(newFolder), StringComparison.Ordinal);
            string newMarkerPath = DraftFolderLayout.GetMarkerPath(draftsRoot, wanted.ExcelDataFile);

            // The new path reaching this very folder is a change of capitals on a disk that ignores them.
            bool sameOnDisk = !sameFolder &&
                              string.Equals(Path.GetFullPath(oldFolder), Path.GetFullPath(newFolder), StringComparison.OrdinalIgnoreCase) &&
                              DraftMarkerStore.Load(newMarkerPath)?.BoardKey == marker.BoardKey;

            if (!sameFolder && !sameOnDisk && Directory.Exists(newFolder))
                return new DraftRenameResult(false, $"A folder for [{wanted.HardwareName}] / [{wanted.BoardName}] is there already. Pick a different name.");

            string oldBoardFile = DraftFolderLayout.GetWorkbookPath(draftsRoot, identity);

            if (BoardFiles.ReadDocument(oldBoardFile) is not { } document)
                return new DraftRenameResult(false, "The draft's board file could not be read.");

            string oldHardware = document.HardwareName;
            string oldBoard = document.BoardName;
            bool moved = false;

            try
            {
                if (!sameFolder)
                {
                    DraftBoardRename.Move(oldFolder, newFolder, sameOnDisk, pause);
                    moved = true;
                }

                afterMove?.Invoke();

                DraftMarkerStore.Save(newMarkerPath, marker with { BoardKey = wanted.ExcelDataFile, NewBoard = wanted }, pause);

                if (!string.Equals(oldHardware, wanted.HardwareName, StringComparison.Ordinal) ||
                    !string.Equals(oldBoard, wanted.BoardName, StringComparison.Ordinal))
                {
                    document.HardwareName = wanted.HardwareName;
                    document.BoardName = wanted.BoardName;
                    BoardFiles.Write(DraftFolderLayout.GetWorkbookPath(draftsRoot, wanted.ExcelDataFile), document);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not rename the draft [{identity}] to [{wanted.ExcelDataFile}] - [{ex.Message}]");

                // ###########################################################################################
                // *** "NOTHING WAS CHANGED" ONLY WHEN IT IS SO (code review, 2026-11-02). *** With the move
                // back refused too, the draft lies at its new place with a marker naming the old one - the
                // open draft answers NoDraft at every save, and said to be unchanged, nothing told the
                // contributor. The next start repairs the marker (RepairMovedDrafts).
                // ###########################################################################################
                return DraftBoardRename.PutBack(oldFolder, newFolder, moved, sameOnDisk, oldMarkerPath, newMarkerPath, marker, oldBoardFile, document, oldHardware, oldBoard, pause)
                    ? new DraftRenameResult(false, $"The draft could not be renamed - something may have its folder open. Nothing was changed. ({ex.Message})")
                    : new DraftRenameResult(false, $"The draft could not be renamed, and could not be put back - it may now lie under the new name. Close CRT and start it again to see it there. ({ex.Message})");
            }

            if (moved && !sameOnDisk)
            {
                DraftBoardRename.RemoveEmptyParents(oldFolder, draftsRoot);
            }

            CrtLog.Info($"Renamed the draft [{identity}] to [{wanted.ExcelDataFile}]");

            return new DraftRenameResult(true);
        }

        // ###########################################################################################
        // *** A FILE HELD FOR A MOMENT IS WAITED FOR (found by the Stop hook's run, 2026-11-02). *** Windows
        // will not move a folder with a file open in it, and CRT's own reads off the UI thread or a virus
        // scanner hold one that way for a moment - FileReplacer's rule for a rename, with its tries. A hold
        // that lasts is refused as before.
        // ###########################################################################################
        private static void Move(string from, string to, bool capitalsOnly, Action<TimeSpan> pause)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(to))!);

            if (capitalsOnly)
            {
                string between = from + ".crt-renaming";
                FileReplacer.MoveOver(from, between, Directory.Move, pause);

                try
                {
                    FileReplacer.MoveOver(between, to, Directory.Move, pause);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    FileReplacer.MoveOver(between, from, Directory.Move, pause);
                    throw;
                }

                return;
            }

            FileReplacer.MoveOver(from, to, Directory.Move, pause);
        }

        // Everything as it was, as far as it can be - false when it could not all be, logged. The first
        // failure is what is said, and this says whether "nothing was changed" is true of it.
        private static bool PutBack(
            string oldFolder,
            string newFolder,
            bool moved,
            bool sameOnDisk,
            string oldMarkerPath,
            string newMarkerPath,
            DraftMarker marker,
            string oldBoardFile,
            BoardDocument document,
            string oldHardware,
            string oldBoard,
            Action<TimeSpan> pause)
        {
            try
            {
                if (moved)
                {
                    DraftBoardRename.Move(newFolder, oldFolder, sameOnDisk, pause);
                    DraftBoardRename.RemoveEmptyParents(newFolder, Path.GetDirectoryName(oldFolder) ?? oldFolder);
                }

                DraftMarkerStore.Save(oldMarkerPath, marker, pause);

                document.HardwareName = oldHardware;
                document.BoardName = oldBoard;
                BoardFiles.Write(oldBoardFile, document);

                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not put the draft back at [{oldFolder}] after a rename failed - [{ex.Message}] (marker tried at [{newMarkerPath}])");

                return false;
            }
        }

        // ###########################################################################################
        // *** A MARKER NAMING ANOTHER PLACE IS REPAIRED AT LAUNCH (code review, 2026-11-02) - what a rename
        // stopped between its move and its marker leaves, by a crash or a put-back that failed. *** The
        // folder is at its new place and nothing in it says so: the lists show the old identity, whose
        // folder is gone, and nothing can be saved to it. So a new board's marker naming a folder that is
        // not there is set to the folder it lies in, the board file's names with it. Called before anything
        // is listed (DraftManager.ImportHandPlacedFolders), so it shows under its own name on this launch.
        // What this machine files under the board's old names (BoardKeyMove) stays there - only the page
        // that renamed it could carry that over.
        //
        // NEVER FOR A FOLDER THAT WOULD THEN COLLIDE WITH ANOTHER, nor one that can be anything else:
        //   - the folder its marker names is there: a copy made by hand, which is not that draft and not
        //     this code's to decide about (or this very folder, in other capitals on a disk that ignores them);
        //   - a published board is listed under the folder's names, or another folder has them in other capitals;
        //   - the name in between a change of capitals (Move): no board's name;
        //   - a draft of a published board, never renamed: its marker names the board.
        //
        // The identities repaired, for the log. A failure costs this one folder, logged, and is tried at the
        // next start.
        // ###########################################################################################
        public static IReadOnlyList<string> RepairMovedDrafts(string draftsRoot, IEnumerable<string>? listedBoards)
        {
            var repaired = new List<string>();

            if (string.IsNullOrWhiteSpace(draftsRoot) || !Directory.Exists(draftsRoot))
            {
                return repaired;
            }

            var listed = new HashSet<string>(
                (listedBoards ?? []).Select(BoardDescriptorRules.BoardIdFromExcelDataFile),
                StringComparer.OrdinalIgnoreCase);

            List<string> folders = DraftFolderImport.EnumerateBoardFolders(draftsRoot).ToList();
            List<string> ids = folders.Select(folder => DraftBoardRename.BoardIdOfFolder(draftsRoot, folder)).ToList();

            for (int index = 0; index < folders.Count; index++)
            {
                string id = ids[index];
                string markerPath = Path.Combine(folders[index], DraftFolderLayout.DraftMarkerFileName);

                if (DraftMarkerStore.Load(markerPath) is not { NewBoard: not null } marker ||
                    BoardDescriptorRules.BoardIdFromExcelDataFile(marker.BoardKey) is not { Length: > 0 } named ||
                    string.Equals(named, id, StringComparison.Ordinal))
                {
                    continue;
                }

                string[] parts = id.Split('/');

                if (Directory.Exists(DraftFolderLayout.GetBoardFolder(draftsRoot, marker.BoardKey)) ||
                    listed.Contains(id) ||
                    ids.Where((_, other) => other != index).Contains(id, StringComparer.OrdinalIgnoreCase) ||
                    parts.Any(part => part.Contains(DraftFolderLayout.BookkeepingPrefix, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                string identity = DataTreeLayout.BoardFilePath(id);

                try
                {
                    DraftMarkerStore.Save(markerPath, marker with
                    {
                        BoardKey = identity,
                        NewBoard = new NewBoardRegistration
                        {
                            HardwareName = parts[1],
                            BoardName = parts[2],
                            HardwareNotes = marker.NewBoard.HardwareNotes,
                            ExcelDataFile = identity,
                            CreatedUtc = marker.NewBoard.CreatedUtc,
                        },
                    });

                    string boardFile = Path.Combine(folders[index], DataTreeLayout.BoardFileName);

                    if (BoardFiles.ReadDocument(boardFile) is { } document)
                    {
                        document.HardwareName = parts[1];
                        document.BoardName = parts[2];
                        BoardFiles.Write(boardFile, document);
                    }

                    CrtLog.Info($"The draft in [{folders[index]}] named [{marker.BoardKey}], whose folder is gone - it is now [{identity}]");
                    repaired.Add(identity);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    CrtLog.Warning($"Could not repair the marker of the draft in [{folders[index]}] - [{ex.Message}]");
                }
            }

            return repaired;
        }

        // "<Manufacturer>/<Hardware>/<Board>" of a board folder three levels under the drafts root.
        private static string BoardIdOfFolder(string draftsRoot, string folder) =>
            Path.GetRelativePath(draftsRoot, folder).Replace(Path.DirectorySeparatorChar, '/');

        // The folders above `folder` that hold nothing any more, up to - never including - `root`.
        private static void RemoveEmptyParents(string folder, string root)
        {
            string stop = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string? parent = Path.GetDirectoryName(Path.GetFullPath(folder));

            try
            {
                while (parent is not null &&
                       parent.Length > stop.Length &&
                       parent.StartsWith(stop, StringComparison.OrdinalIgnoreCase) &&
                       Directory.Exists(parent) &&
                       Directory.GetFileSystemEntries(parent).Length == 0)
                {
                    Directory.Delete(parent);
                    parent = Path.GetDirectoryName(parent);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // An empty folder left behind harms nothing - the walk finds no marker in it.
                CrtLog.Warning($"Could not remove the empty folder [{parent}] after a rename - [{ex.Message}]");
            }
        }
    }
}
