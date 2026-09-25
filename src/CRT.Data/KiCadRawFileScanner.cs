using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Finding the raw KiCad files inside a folder - the ONE rule, shared by the importer that
    // copies them into a draft and by the app that later reads them back (owner request,
    // 2026-09-24).
    //
    // *** SUB-FOLDERS ARE INCLUDED, AND THAT IS A BUG FIX RATHER THAN A NEW FEATURE. ***
    //
    // A real multi-sheet KiCad project puts its root sheet at the top and every other page in a
    // sub-folder. The shipped Commodore C128 "310378 Open128" board is exactly that shape:
    //
    //     Open128.kicad_pcb           45 MB
    //     Open128.kicad_pro
    //     Open128.kicad_sch           the root sheet, which only REFERENCES the pages
    //     Pages/vic.kicad_sch         22 further sheets, holding the actual circuitry
    //     Pages/cpu-8500.kicad_sch
    //     ...
    //
    // Reading the top level only means reading the root sheet and none of the pages, so nearly
    // every net on that board resolved to nothing. The files were already on every client's disk -
    // DataManager syncs the KiCad folder with SearchOption.AllDirectories - so the sync and the
    // read disagreed, and the read was the half that was wrong.
    //
    // *** THE EXTENSION FILTER IS WHAT KEEPS AN IMPORT SMALL, NOT THE DEPTH LIMIT. *** The concern
    // the top-level rule was meant to answer - "do not drag in footprint libraries, 3D models,
    // gerbers and backups" - is answered by IsSupportedKiCadRawFile, which admits three extensions
    // and nothing else. A .pretty footprint library holds .kicad_mod files, a 3D model folder holds
    // .step/.wrl, gerbers are .gbr/.drl: not one of them is admitted at any depth. So recursing
    // costs nothing in what gets copied while gaining the pages, which is the whole point.
    //
    // The one thing that DOES need care is a backup folder holding copies of the real sheets, which
    // would be admitted on extension - see SkippedFolderNames below.
    // ###########################################################################################
    public static class KiCadRawFileScanner
    {
        // ###########################################################################################
        // Folders never descended into, matched on the folder's own name, case-insensitively.
        //
        // Each one holds files that WOULD pass the extension filter and must not:
        //
        //   - "*-backups" is KiCad's own automatic backup folder (it writes
        //     "<project>-backups/<project>-<timestamp>.zip", and an unzipped one beside it is
        //     common). Importing it duplicates every sheet, and the duplicates parse as real
        //     circuitry - so a net would appear twice over.
        //   - "fp-info-cache" and ".git" are noise rather than a correctness problem, but there is
        //     no reason to walk them.
        //
        // Matched by SUFFIX for the backups case because the folder is named after the project
        // ("Open128-backups"), so the name is not fixed.
        // ###########################################################################################
        public const string BackupFolderSuffix = "-backups";

        private static readonly string[] SkippedFolderNames =
        [
            "fp-info-cache",
        ];

        // ###########################################################################################
        // True when this folder must not be descended into. Takes the folder's own NAME, not a path.
        // ###########################################################################################
        public static bool IsSkippedFolder(string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName))
            {
                return false;
            }

            string name = folderName.Trim();

            // Any dot-folder: ".git", ".vscode", and whatever else a contributor's tooling leaves.
            // The same rule DataChecksumManifest.IsSyncable applies, for the same reason.
            if (name.StartsWith('.'))
            {
                return true;
            }

            if (name.EndsWith(KiCadRawFileScanner.BackupFolderSuffix, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return KiCadRawFileScanner.SkippedFolderNames.Contains(name, StringComparer.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // Every supported raw KiCad file at or below one folder, sorted for a stable result.
        //
        // Returns an empty list rather than throwing for a missing or unreadable folder: both
        // callers run on a user-chosen path, and a board with no KiCad data is the normal case.
        //
        // *** WALKED BY HAND RATHER THAN WITH SearchOption.AllDirectories, *** because that option
        // cannot skip a sub-tree - it would descend into "<project>-backups" and there is no way to
        // tell it not to. It also throws the whole enumeration away on one unreadable folder, where
        // walking lets an unreadable corner be skipped while everything else still imports.
        // ###########################################################################################
        public static List<string> Scan(string folder)
        {
            var found = new List<string>();

            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                return found;
            }

            KiCadRawFileScanner.Walk(folder, found);

            found.Sort(StringComparer.OrdinalIgnoreCase);

            return found;
        }

        private static void Walk(string folder, List<string> found)
        {
            try
            {
                foreach (string path in Directory.EnumerateFiles(folder))
                {
                    if (ComponentListBuilder.IsSupportedKiCadRawFile(path))
                    {
                        found.Add(path);
                    }
                }

                foreach (string subFolder in Directory.EnumerateDirectories(folder))
                {
                    if (!KiCadRawFileScanner.IsSkippedFolder(Path.GetFileName(subFolder)))
                    {
                        KiCadRawFileScanner.Walk(subFolder, found);
                    }
                }
            }
            catch (Exception ex)
            {
                // One unreadable folder must not lose the rest of the project.
                CrtLog.Warning($"Could not read the KiCad folder [{folder}] - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Where one scanned file goes when it is copied into the board's "KiCad data" folder: its
        // path RELATIVE to the folder that was scanned, so "Pages/vic.kicad_sch" stays under Pages.
        //
        // *** THE SUB-FOLDER STRUCTURE IS PRESERVED RATHER THAN FLATTENED, and that is not
        // cosmetic. *** A KiCad root sheet references its pages by relative path, and two pages in
        // different sub-folders may legitimately share a file name. Flattening would break the
        // first and silently drop one of the second.
        //
        // It also makes an import byte-for-byte identical to the hand-made layout the shipped C128
        // board already uses, so imported and hand-placed KiCad data cannot behave differently.
        // ###########################################################################################
        public static string RelativeDestinationFor(string sourceFolder, string sourceFile)
        {
            string relative = Path.GetRelativePath(sourceFolder, sourceFile);

            // A file outside the scanned folder has no sensible destination inside it; fall back to
            // the bare name rather than writing "..\..\something" into the board folder.
            return relative.StartsWith("..", StringComparison.Ordinal) || Path.IsPathRooted(relative)
                ? Path.GetFileName(sourceFile)
                : relative;
        }
    }
}
