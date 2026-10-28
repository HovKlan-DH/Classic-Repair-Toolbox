using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // One walk of a folder an import is made from: the files taken (full paths), every other file
    // (relative to the folder, forward slashes), and which of those an older KiCad wrote.
    // ###########################################################################################
    public sealed record KiCadFolderScan(IReadOnlyList<string> Taken, IReadOnlyList<string> LeftOut, IReadOnlyList<string> OlderKiCad);

    // What a dropped folder is (KiCadRawFileScanner.Probe): whether it holds a KiCad file CRT reads, and
    // if not its .sch and .brd files and which of them an older KiCad wrote.
    public sealed record KiCadFolderProbe(bool HoldsKiCad, IReadOnlyList<string> SchematicsAndBoards, IReadOnlyList<string> OlderKiCad);

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

            KiCadRawFileScanner.Walk(folder, path =>
            {
                if (ComponentListBuilder.IsSupportedKiCadRawFile(path))
                {
                    found.Add(path);
                }

                return true;
            });

            found.Sort(StringComparer.OrdinalIgnoreCase);

            return found;
        }

        // ###########################################################################################
        // Every raw KiCad file a BOARD reads: its published "KiCad data" folder first, then its
        // draft's (Main.KiCadRawPathsFor' rule). Both are read because a contributor
        // can import KiCad data into a draft of a board whose published copy already has some.
        //
        // Shared with the "Schematic images" window (2026-10-10), which offers the view names of
        // exactly the files the Schematics tab will draw that board's traces from.
        // ###########################################################################################
        public static List<string> ScanBoard(string publishedKiCadFolder, string draftKiCadFolder) =>
            KiCadRawFileScanner.Scan(publishedKiCadFolder)
                .Concat(KiCadRawFileScanner.Scan(draftKiCadFolder))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();

        // ###########################################################################################
        // ONE walk of a folder an import is made from (CRT 4.0 step 3, slice 3e; code review,
        // 2026-10-28 - it was walked once for the files taken and again for the rest): the files TAKEN
        // (as Scan finds them), every other file - LEFT OUT - as its path relative to the folder with
        // forward slashes, and which of those are a schematic or board an older KiCad wrote
        // (IsOlderKiCadFile). Walked exactly as Scan walks, so a file in a skipped folder is neither
        // taken nor counted. What a KiCad import's summary says (KiCadImportSummary).
        // ###########################################################################################
        public static KiCadFolderScan ScanSplit(string folder)
        {
            var taken = new List<string>();
            var leftOut = new List<string>();

            if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
            {
                KiCadRawFileScanner.Walk(folder, path =>
                {
                    (ComponentListBuilder.IsSupportedKiCadRawFile(path) ? taken : leftOut).Add(path);
                    return true;
                });
            }

            taken.Sort(StringComparer.OrdinalIgnoreCase);

            List<(string Relative, string Full)> others = leftOut
                .Select(path => (Relative: Path.GetRelativePath(folder, path).Replace(Path.DirectorySeparatorChar, '/'), Full: path))
                .OrderBy(file => file.Relative, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return new KiCadFolderScan(
                taken,
                others.Select(file => file.Relative).ToList(),
                others.Where(file => KiCadRawFileScanner.IsOlderKiCadFile(file.Full)).Select(file => file.Relative).ToList());
        }

        // ###########################################################################################
        // What a folder DROPPED on the Contribute tab is, in ONE walk (second code review, 2026-10-28:
        // a folder holding nothing usable - Documents, a drive's root - was walked in full twice, every
        // path kept): whether it holds a KiCad file CRT reads - the walk stops at the first - and, when
        // it holds none, its .sch and .brd files (relative, forward slashes) and which an older KiCad
        // wrote: all that is needed to say why nothing can be imported (KiCadImportSummary.CannotRead).
        // ###########################################################################################
        public static KiCadFolderProbe Probe(string folder)
        {
            bool holdsKiCad = false;
            var schematicsAndBoards = new List<string>();

            if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
            {
                KiCadRawFileScanner.Walk(folder, path =>
                {
                    if (ComponentListBuilder.IsSupportedKiCadRawFile(path))
                    {
                        holdsKiCad = true;
                        return false;
                    }

                    string extension = Path.GetExtension(path);

                    if (string.Equals(extension, ".sch", StringComparison.OrdinalIgnoreCase) || string.Equals(extension, ".brd", StringComparison.OrdinalIgnoreCase))
                    {
                        schematicsAndBoards.Add(path);
                    }

                    return true;
                });
            }

            if (holdsKiCad)
            {
                return new KiCadFolderProbe(true, [], []);
            }

            List<(string Relative, string Full)> files = schematicsAndBoards
                .Select(path => (Relative: Path.GetRelativePath(folder, path).Replace(Path.DirectorySeparatorChar, '/'), Full: path))
                .OrderBy(file => file.Relative, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return new KiCadFolderProbe(
                false,
                files.Select(file => file.Relative).ToList(),
                files.Where(file => KiCadRawFileScanner.IsOlderKiCadFile(file.Full)).Select(file => file.Relative).ToList());
        }

        // ###########################################################################################
        // Whether a .sch or .brd is one an OLDER KICAD wrote, by its first line: KiCad 5's schematic
        // opens "EESchema Schematic File", KiCad 4's board "PCBNEW-BOARD". EAGLE uses the same two
        // extensions (code review, 2026-10-28: an EAGLE project was told to open itself in KiCad and
        // save), and its files open otherwise. A file that cannot be read is no older KiCad's.
        // ###########################################################################################
        public static bool IsOlderKiCadFile(string path)
        {
            string extension = Path.GetExtension(path ?? string.Empty);
            bool schematic = string.Equals(extension, ".sch", StringComparison.OrdinalIgnoreCase);
            bool board = string.Equals(extension, ".brd", StringComparison.OrdinalIgnoreCase);

            if (!schematic && !board)
            {
                return false;
            }

            try
            {
                using var reader = new StreamReader(path!);
                var start = new char[64];
                int read = reader.ReadBlock(start, 0, start.Length);
                string opening = new string(start, 0, read).TrimStart('﻿', ' ', '\t', '\r', '\n');

                return opening.StartsWith(schematic ? "EESchema Schematic File" : "PCBNEW-BOARD", StringComparison.Ordinal);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return false;
            }
        }

        // ###########################################################################################
        // Every file at or below the folder, past the folders never entered (IsSkippedFolder), handed
        // to `visit` - false from it stops the walk. False when stopped.
        //
        // *** A HIDDEN FILE IS NO FILE OF THE PROJECT (second code review, 2026-10-28), *** as a hidden
        // folder is never entered: macOS leaves "._Open128.kicad_pcb" - binary, its resource fork -
        // beside the real one on a USB stick, and it was taken, parsed and sent, its name refused by
        // the server ("file.hidden_name"). Neither taken nor counted as left out.
        // ###########################################################################################
        private static bool Walk(string folder, Func<string, bool> visit)
        {
            try
            {
                foreach (string path in Directory.EnumerateFiles(folder))
                {
                    if (Path.GetFileName(path).StartsWith('.'))
                    {
                        continue;
                    }

                    if (!visit(path))
                    {
                        return false;
                    }
                }

                foreach (string subFolder in Directory.EnumerateDirectories(folder))
                {
                    if (!KiCadRawFileScanner.IsSkippedFolder(Path.GetFileName(subFolder)) && !KiCadRawFileScanner.Walk(subFolder, visit))
                    {
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                // One unreadable folder must not lose the rest of the project.
                CrtLog.Warning($"Could not read the KiCad folder [{folder}] - [{ex.Message}]");
            }

            return true;
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
