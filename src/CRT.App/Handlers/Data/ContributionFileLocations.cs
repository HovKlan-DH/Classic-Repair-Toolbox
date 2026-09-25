using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The folders the component editor offers in its "File location" drop-down: every END folder
    // (one with no sub-folders) of the data tree AND of the drafts tree, as "/"-separated paths
    // relative to their own root ("Commodore/C64/250407/Scope baseline").
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
