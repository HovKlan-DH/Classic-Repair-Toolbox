using System;
using System.IO;

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
        public static DraftRenameResult Run(string draftsRoot, string identity, NewBoardRegistration wanted)
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
                    DraftBoardRename.Move(oldFolder, newFolder, sameOnDisk);
                    moved = true;
                }

                DraftMarkerStore.Save(newMarkerPath, marker with { BoardKey = wanted.ExcelDataFile, NewBoard = wanted });

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

                DraftBoardRename.PutBack(oldFolder, newFolder, moved, sameOnDisk, oldMarkerPath, newMarkerPath, marker, oldBoardFile, document, oldHardware, oldBoard);

                return new DraftRenameResult(false, $"The draft could not be renamed - something may have its folder open. Nothing was changed. ({ex.Message})");
            }

            if (moved && !sameOnDisk)
            {
                DraftBoardRename.RemoveEmptyParents(oldFolder, draftsRoot);
            }

            CrtLog.Info($"Renamed the draft [{identity}] to [{wanted.ExcelDataFile}]");

            return new DraftRenameResult(true);
        }

        private static void Move(string from, string to, bool capitalsOnly)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(to))!);

            if (capitalsOnly)
            {
                string between = from + ".crt-renaming";
                Directory.Move(from, between);

                try
                {
                    Directory.Move(between, to);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    Directory.Move(between, from);
                    throw;
                }

                return;
            }

            Directory.Move(from, to);
        }

        // Everything as it was, as far as it can be - a failure here is only logged, the first is said.
        private static void PutBack(
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
            string oldBoard)
        {
            try
            {
                if (moved)
                {
                    DraftBoardRename.Move(newFolder, oldFolder, sameOnDisk);
                    DraftBoardRename.RemoveEmptyParents(newFolder, Path.GetDirectoryName(oldFolder) ?? oldFolder);
                }

                DraftMarkerStore.Save(oldMarkerPath, marker);

                document.HardwareName = oldHardware;
                document.BoardName = oldBoard;
                BoardFiles.Write(oldBoardFile, document);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not put the draft back at [{oldFolder}] after a rename failed - [{ex.Message}] (marker tried at [{newMarkerPath}])");
            }
        }

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
