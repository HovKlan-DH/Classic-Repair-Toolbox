using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The folders the component editor offers in its "File location" drop-down: every END folder
    // (one with no sub-folders) of the data tree AND of the drafts tree, as "/"-separated paths
    // relative to their own root ("Commodore/C64/250407/Scope baseline"), narrowed by WritableBy to
    // the folders a submission for the system being edited may write.
    //
    // *** THE DRAFTS TREE IS INCLUDED, AND THAT IS THE FIX (owner report, 2026-09-24). ***
    // The list used to be built from the data root alone. A system created with "Add a new system"
    // exists ONLY under the drafts root, so its folders - the "Scope baseline" folder DraftSeeder
    // creates for it in particular - were never offered, and a file could not be filed there.
    //
    // Both trees use the same Manufacturer/Hardware/Board layout (DraftFolderLayout mirrors Data/
    // on purpose), so a path relative to either root means the same place, and a location picked
    // from the drafts half is stored exactly like one from the data half. DraftFileResolver's
    // BuildDraftFileDestination already maps "Manufacturer/Hardware/Board/Scope baseline/x.png"
    // into the draft's own folder, so nothing downstream of this list needed to change.
    //
    // Merged case-INSENSITIVELY: a published board with a seeded draft appears in both trees, and
    // listing each of its folders twice would double the drop-down for nothing. Sorted the same
    // way, since that is the order the window re-sorts into once a row's own location is added.
    //
    // Fails soft per folder: an unreadable directory is skipped rather than losing the whole list,
    // matching what the window did before this was extracted.
    // ###########################################################################################
    internal static class ContributionFileLocations
    {
        public static List<string> FindEndFolders(params string?[] roots)
        {
            var endFolders = new List<string>();

            foreach (string? root in roots)
            {
                if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
                {
                    continue;
                }

                ContributionFileLocations.FindEndFoldersRecursive(root, root, endFolders);
            }

            return endFolders
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(folder => folder, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ###########################################################################################
        // *** ONLY THE FOLDERS A SUBMISSION FOR THIS SYSTEM MAY WRITE (owner request, 2026-09-25). ***
        // The drop-down offered every end folder of every board of every manufacturer - dozens of
        // entries, nearly all of them folders the server refuses a new file in ("[x] belongs to
        // another board"). It now offers the board's own folder and its sub-folders, this
        // manufacturer's "Shared files" folders and the "Generic shared files" folders.
        //
        // The board's own folder is ADDED even when it is not an end folder: its schematic images
        // and other files sit directly in it, and a board with no sub-folders at all would
        // otherwise offer none of its own.
        //
        // An unrecognisable system id filters nothing. The save refuses such a window anyway, and an
        // empty drop-down would only hide why.
        // ###########################################################################################
        public static List<string> WritableBy(string? systemId, IEnumerable<string> folders)
        {
            ArgumentNullException.ThrowIfNull(folders);

            if (!SystemDescriptorRules.IsValidSystemId(systemId))
            {
                return folders.ToList();
            }

            return folders
                .Append(systemId!)
                .Where(folder => ContributionFileLocations.IsWritableFolder(systemId, folder))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(folder => folder, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ###########################################################################################
        // Whether a new file may be filed in this folder by a submission for this system.
        //
        // *** THE SERVER'S OWN RULE DECIDES "WHOSE FOLDER" - SubmissionFileScopes.Classify, which
        // SubmissionFileRules refuses a submission by. *** A second opinion here would drift from it,
        // and the contributor would again find out only when the server refused the submission.
        // Classify judges a FILE path, so a name stands in for the file: a folder is writable exactly
        // when a file directly inside it is.
        //
        // Two more are left out, although the server would take a file in either:
        //   - a folder CRT reads by NAME (DataTreeUsage.FoldersReadByName - the MiniPro IC tests),
        //     which no row cites;
        //   - a draft's COPY of a shared folder ("<system>/Commodore/Shared files/..."). A new file
        //     filed in a shared folder is kept inside the draft under its whole path
        //     (DraftFileResolver.BuildDraftFileDestination), so the drafts-tree scan reports that
        //     copy as a folder of the board itself. It exists nowhere else.
        // ###########################################################################################
        public static bool IsWritableFolder(string? systemId, string? folder)
        {
            string trimmed = folder?.Trim().Replace('\\', '/').Trim('/') ?? string.Empty;

            if (trimmed.Length == 0 || !SystemDescriptorRules.IsValidSystemId(systemId))
            {
                return false;
            }

            string[] system = systemId!.Split('/');

            SubmissionFileScope scope = SubmissionFileScopes.Classify(
                system[0], system[1], system[2], trimmed + "/file");

            if (scope == SubmissionFileScope.Foreign)
            {
                return false;
            }

            bool isReadByName = DataTreeUsage.FoldersReadByName.Any(readByName =>
                string.Equals(trimmed, readByName, StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith(readByName + "/", StringComparison.OrdinalIgnoreCase));

            if (isReadByName)
            {
                return false;
            }

            if (scope == SubmissionFileScope.Own)
            {
                string[] within = trimmed.Length == systemId.Length
                    ? []
                    : trimmed[(systemId.Length + 1)..].Split('/', StringSplitOptions.RemoveEmptyEntries);

                bool isDraftCopyOfSharedFolder =
                    (within.Length > 0 && SubmissionFileScopes.IsSharedFolderName(within[0])) ||
                    (within.Length > 1 && SubmissionFileScopes.IsSharedFolderName(within[1]));

                if (isDraftCopyOfSharedFolder)
                {
                    return false;
                }
            }

            return true;
        }

        // What stops a newly picked file being filed where its row says - see CheckNewFile.
        public enum NewFileProblem
        {
            None,
            NoFolder,
            FolderNotWritable
        }

        // ###########################################################################################
        // *** A NEWLY PICKED FILE MUST BE FILED IN A FOLDER THE SUBMISSION MAY WRITE (owner report,
        // 2026-09-25). *** A file picked from outside the data folder keeps whatever "File location"
        // its row shows, and a new row's starts EMPTY. Saved like that, the row stored the bare file
        // name ("HotCPU.png") - a file at the very top of the data tree, outside every board - and
        // the server refused the whole submission much later with "[HotCPU.png] belongs to another
        // board". "Save to draft" now refuses it instead, while the row is still on screen.
        //
        // `usedInPlace` is a file picked from the data folder and left where it is: it IS the
        // published copy, and citing another board's file unchanged is how real data is shaped (the
        // C128DCR cites the C128's scope baselines), which the server allows too.
        // ###########################################################################################
        public static NewFileProblem CheckNewFile(string? systemId, string? fileLocation, bool usedInPlace)
        {
            if (string.IsNullOrWhiteSpace(fileLocation))
            {
                return NewFileProblem.NoFolder;
            }

            if (usedInPlace || ContributionFileLocations.IsWritableFolder(systemId, fileLocation))
            {
                return NewFileProblem.None;
            }

            return NewFileProblem.FolderNotWritable;
        }

        private static void FindEndFoldersRecursive(string rootPath, string currentPath, List<string> endFolders)
        {
            try
            {
                var subDirs = Directory.GetDirectories(currentPath);
                if (subDirs.Length == 0)
                {
                    string relativePath = Path.GetRelativePath(rootPath, currentPath);
                    if (!string.IsNullOrWhiteSpace(relativePath) && relativePath != ".")
                    {
                        endFolders.Add(relativePath.Replace('\\', '/'));
                    }
                }
                else
                {
                    foreach (var subDir in subDirs)
                    {
                        ContributionFileLocations.FindEndFoldersRecursive(rootPath, subDir, endFolders);
                    }
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Unreadable directories are skipped safely
            }
        }
    }
}
