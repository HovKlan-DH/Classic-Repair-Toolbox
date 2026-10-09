using System;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Removing one imported KiCad file from a draft's "KiCad data" folder (owner request,
    // 2026-09-24) - the per-file "Remove" in BoardFilesWindow, the KiCad twin of removing a
    // schematic image.
    //
    // *** UNLIKE A SCHEMATIC, THE FILE ITSELF IS DELETED. *** Removing a schematic drops its
    // workbook ROW and keeps the image bytes, because the row is what the board reads. A KiCad file
    // has no row - the app discovers the folder's contents by name (KiCadRawFileScanner), so the
    // file being there IS the data, and the only way to take it out of the board is to delete it.
    // What is lost is the draft's COPY: the import copies, so the contributor's own KiCad project is
    // untouched and re-importing it brings the file back.
    //
    // Purely local. Raw KiCad files are not part of a submission (SubmissionManifestBuilder takes
    // only files the workbook rows name), so there is no other side to keep in step.
    //
    // FOLDERS THE DELETE LEAVES EMPTY ARE REMOVED TOO, up to and including "KiCad data" itself and
    // never beyond it. Removing a multi-sheet project's last page would otherwise leave an empty
    // Pages/ behind, and removing every file an empty "KiCad data" - both of which the component
    // editor's "File location" drop-down would then offer as places to put a file.
    // ###########################################################################################
    internal static class KiCadImportedFiles
    {
        public static bool TryRemove(string kiCadFolder, string filePath)
        {
            if (string.IsNullOrWhiteSpace(kiCadFolder) || string.IsNullOrWhiteSpace(filePath))
            {
                return false;
            }

            try
            {
                string folder = Path.GetFullPath(kiCadFolder);
                string file = Path.GetFullPath(filePath);

                // ###########################################################################################
                // *** CONTAINMENT FIRST, AND BY RELATIVE PATH, NOT BY PREFIX. *** This deletes a file,
                // so it must be impossible to aim it anywhere but inside the KiCad folder. A plain
                // StartsWith would accept a sibling like "KiCad data2/x" - GetRelativePath answers
                // "../KiCad data2/x" for that, which is refused.
                // ###########################################################################################
                if (!KiCadImportedFiles.IsStrictlyInside(folder, file) || !File.Exists(file))
                {
                    return false;
                }

                File.Delete(file);

                KiCadImportedFiles.RemoveEmptyFoldersUpTo(folder, Path.GetDirectoryName(file));

                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
            {
                Logger.Warning($"Could not remove the KiCad file [{filePath}] - [{ex.Message}]");
                return false;
            }
        }

        private static bool IsStrictlyInside(string folder, string path)
        {
            string relative = Path.GetRelativePath(folder, path);

            return relative != "."
                && !Path.IsPathRooted(relative)
                && !relative.Equals("..", StringComparison.Ordinal)
                && !relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal)
                && !relative.StartsWith(".." + Path.AltDirectorySeparatorChar, StringComparison.Ordinal);
        }

        // Walks up from the deleted file's folder, removing each folder that is now empty, and stops
        // at the first one that still holds something or once "KiCad data" itself has been handled.
        private static void RemoveEmptyFoldersUpTo(string kiCadFolder, string? startFolder)
        {
            string? current = startFolder;

            while (!string.IsNullOrWhiteSpace(current))
            {
                bool isKiCadFolder = string.Equals(
                    Path.GetFullPath(current).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    kiCadFolder.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    StringComparison.OrdinalIgnoreCase);

                if (!isKiCadFolder && !KiCadImportedFiles.IsStrictlyInside(kiCadFolder, current))
                {
                    return;
                }

                if (!Directory.Exists(current) || Directory.EnumerateFileSystemEntries(current).Any())
                {
                    return;
                }

                Directory.Delete(current, recursive: false);

                if (isKiCadFolder)
                {
                    return;
                }

                current = Path.GetDirectoryName(current);
            }
        }
    }
}
