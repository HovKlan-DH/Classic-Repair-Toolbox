using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A KiCad IMPORT FOUND, LEFT OUT AND LEAVES TO DO (CRT 4.0 step 3, slice 3e - the plan:
    // "KiCad: drop the project folder and get a summary of what was found, what was ignored, and what
    // is left to do"). Said at the top of the "KiCad data" window:
    //
    //   - FOUND: the PCB files - the first, the one CRT draws, with its components (the very ones "Add
    //     N from KiCad" counts, KiCadComponentTable) and its nets - and the schematic sheets.
    //   - LEFT OUT: what the folder held that an import does not take (KiCadRawFileScanner.ScanSplit
    //     - every file but .kicad_pcb, .kicad_pro and .kicad_sch), counted by what it is: footprints,
    //     3D models, Gerbers, hidden files never among them... A schematic or board CRT cannot read
    //     (.sch, .brd) is SAID so, with what to do about it - an older KiCad's (told by its first line)
    //     and EAGLE's alike - such a project used to read only "No KiCad files in that folder".
    //   - LEFT TO DO: the getting-ready steps after the KiCad data, by the checklist's own rules
    //     (BoardReadiness) - never a second copy of them.
    //
    // PURE: words only, so they are tested without a display.
    // ###########################################################################################
    internal static class KiCadImportSummary
    {
        // What a file left out is, said in the singular and the plural - in the order they are listed.
        private enum Kind
        {
            Footprint,
            Symbol,
            Model,
            Manufacturing,
            Older,
            Foreign,
            Settings,
            Backup,
            Other,
        }

        private static readonly (Kind Kind, string One, string Many)[] Words =
        [
            (Kind.Footprint, "footprint", "footprints"),
            (Kind.Symbol, "symbol library file", "symbol library files"),
            (Kind.Model, "3D model", "3D models"),
            (Kind.Manufacturing, "Gerber or drill file", "Gerber and drill files"),
            (Kind.Older, "file from an older KiCad", "files from an older KiCad"),
            (Kind.Foreign, "schematic or board file of another program", "schematic and board files of another program"),
            (Kind.Settings, "KiCad settings file", "KiCad settings files"),
            (Kind.Backup, "backup", "backups"),
            (Kind.Other, "other file", "other files"),
        ];

        public const string NothingLeftToDo = "Nothing - the board is in step with its KiCad data.";

        // ###########################################################################################
        // What the KiCad data holds: "Found 1 PCB with 214 components and 380 nets, and 24 schematic
        // sheets." Only the FIRST PCB lights up copper, so of several it is the one counted.
        // ###########################################################################################
        public static string Found(KiCadProjectRoot root)
        {
            ArgumentNullException.ThrowIfNull(root);

            int components = KiCadComponentTable.Of(root).Count;
            int nets = root.Pcb.Count == 0
                ? 0
                : root.Pcb[0].Nets.List.Count(net => !net.IsNoNet && !string.IsNullOrWhiteSpace(net.Name));

            string withWhat = $"{KiCadImportSummary.Count(components, "component", "components")} and {KiCadImportSummary.Count(nets, "net", "nets")}";

            string pcb = root.Pcb.Count switch
            {
                0 => "no PCB file",
                1 => $"1 PCB with {withWhat}",
                _ => $"{KiCadImportSummary.Number(root.Pcb.Count)} PCB files - the first, the one CRT draws, with {withWhat}",
            };

            string sheets = root.Schematics.Count == 0
                ? "no schematic sheet"
                : KiCadImportSummary.Count(root.Schematics.Count, "schematic sheet", "schematic sheets");

            return $"Found {pcb}, and {sheets}.";
        }

        // ###########################################################################################
        // What an import left out, counted by kind - "" when it took everything. A .sch or .brd is an
        // older KiCad's only when `olderKiCad` names it (KiCadRawFileScanner.IsOlderKiCadFile, by its
        // first line) - else another program's, EAGLE's say (code review, 2026-10-28).
        // ###########################################################################################
        public static string LeftOut(IReadOnlyList<string> leftOut, IReadOnlyList<string> olderKiCad)
        {
            ArgumentNullException.ThrowIfNull(leftOut);
            ArgumentNullException.ThrowIfNull(olderKiCad);

            if (leftOut.Count == 0)
            {
                return string.Empty;
            }

            var older = new HashSet<string>(olderKiCad, StringComparer.OrdinalIgnoreCase);

            List<string> parts = leftOut
                .GroupBy(path => KiCadImportSummary.KindOf(path, older))
                .OrderBy(group => group.Key)
                .Select(group =>
                {
                    (_, string one, string many) = KiCadImportSummary.Words.Single(words => words.Kind == group.Key);
                    return KiCadImportSummary.Count(group.Count(), one, many);
                })
                .ToList();

            string listed = parts.Count == 1 ? parts[0] : $"{string.Join(", ", parts.Take(parts.Count - 1))} and {parts[^1]}";

            return $"Left out, as CRT does not read them: {listed}.";
        }

        // ###########################################################################################
        // What of the project CRT cannot read, and what to do about it - none, one or two sentences:
        //
        //   - A schematic (.sch, KiCad 5 or older) or board (.brd, KiCad 4 or older) an OLDER KICAD wrote
        //     (`olderKiCad` - told by its first line): open the project in a newer KiCad and save it.
        //   - A .sch or .brd of ANOTHER PROGRAM - EAGLE uses the same extensions: KiCad imports an EAGLE
        //     project (code review, 2026-10-28: told to open it in KiCad and save, which does nothing).
        //
        // Only what came in that form ALONE: a project saved again in KiCad 6 or later keeps the old
        // files beside the new, and the new are read.
        // ###########################################################################################
        public static List<string> CannotRead(IReadOnlyList<string> leftOut, IReadOnlyList<string> olderKiCad, IReadOnlyList<string> taken)
        {
            ArgumentNullException.ThrowIfNull(leftOut);
            ArgumentNullException.ThrowIfNull(olderKiCad);
            ArgumentNullException.ThrowIfNull(taken);

            bool Has(IEnumerable<string> files, string extension) =>
                files.Any(file => string.Equals(Path.GetExtension(file), extension, StringComparison.OrdinalIgnoreCase));

            var older = new HashSet<string>(olderKiCad, StringComparer.OrdinalIgnoreCase);
            List<string> foreign = leftOut.Where(file => !older.Contains(file)).ToList();

            bool newSheets = Has(taken, ".kicad_sch");
            bool newBoard = Has(taken, ".kicad_pcb");

            bool oldSheets = Has(older, ".sch") && !newSheets;
            bool oldBoard = Has(older, ".brd") && !newBoard;
            bool otherProgram = !newSheets && !newBoard && (Has(foreign, ".sch") || Has(foreign, ".brd"));

            var sentences = new List<string>();

            if (oldSheets || oldBoard)
            {
                string what = oldSheets && oldBoard
                    ? "Its schematic (.sch) and board (.brd) files are"
                    : oldSheets ? "Its schematic files (.sch) are" : "Its board file (.brd) is";

                sentences.Add($"{what} from an older KiCad, which CRT cannot read - open the project in KiCad 6 or later, save it, and import the folder again.");
            }

            if (otherProgram)
            {
                sentences.Add(KiCadImportSummary.OtherProgram);
            }

            return sentences;
        }

        public const string OtherProgram =
            "Its .sch and .brd files are not KiCad's - an EAGLE project, perhaps. KiCad 6 or later can import one (File, Import, Non-KiCad Project): save it there, then import the folder again.";

        // A folder with nothing CRT can read: why, and - for an older KiCad's or another program's
        // project - what to do.
        public static string NothingToImport(IReadOnlyList<string> leftOut, IReadOnlyList<string> olderKiCad) =>
            KiCadImportSummary.CannotRead(leftOut, olderKiCad, []) is { Count: > 0 } why
                ? "No KiCad files CRT can read in that folder. " + string.Join(" ", why)
                : "No KiCad files in that folder. Only modern .kicad_pcb / .kicad_pro / .kicad_sch files are imported - sub-folders are searched too.";

        // ###########################################################################################
        // What is left to do once the KiCad data is in: the getting-ready steps after it that still
        // have something to do, by the checklist's rules (BoardReadiness.Steps) - none for a step that
        // is done, optional, or skipped.
        // ###########################################################################################
        public static List<string> LeftToDo(IReadOnlyList<ReadinessStep> steps)
        {
            ArgumentNullException.ThrowIfNull(steps);

            return steps
                .Where(step => step.State is ReadinessState.Next or ReadinessState.Later)
                .Select(step => step.Kind switch
                {
                    ReadinessStepKind.Match => "Match each picture to the KiCad view it shows.",
                    ReadinessStepKind.Align => "Align the KiCad copper on each picture that shows a view.",
                    ReadinessStepKind.Components when step.FromKiCad > 0 =>
                        $"Add the {KiCadImportSummary.Count(step.FromKiCad, "component", "components")} of the KiCad data not on the board yet.",
                    ReadinessStepKind.Components => "Add the board's components.",
                    ReadinessStepKind.Mark => "Mark the components on the pictures.",
                    ReadinessStepKind.Signals => "Pick the important signals - CRT suggests them from the KiCad net names.",
                    _ => string.Empty,
                })
                .Where(line => line.Length > 0)
                .ToList();
        }

        // ---- the kinds ------------------------------------------------------------------------------

        private static Kind KindOf(string path, ISet<string> olderKiCad)
        {
            string name = (path ?? string.Empty).Replace('\\', '/');
            name = name[(name.LastIndexOf('/') + 1)..].ToLowerInvariant();
            string extension = Path.GetExtension(name);

            if (name.EndsWith("-bak", StringComparison.Ordinal) || name.EndsWith('~') || name.StartsWith("_autosave-", StringComparison.Ordinal) ||
                extension is ".bak" or ".lck" || name.Contains("auto_saved_files", StringComparison.Ordinal))
            {
                return Kind.Backup;
            }

            return extension switch
            {
                ".kicad_mod" => Kind.Footprint,
                ".kicad_sym" or ".lib" or ".dcm" => Kind.Symbol,
                ".step" or ".stp" or ".wrl" or ".igs" or ".iges" or ".x3d" => Kind.Model,
                ".gbr" or ".gbrjob" or ".drl" or ".gtl" or ".gbl" or ".gto" or ".gbo" or ".gts" or ".gbs" or ".gtp" or ".gbp" or ".gm1" or ".gko" or ".gml" or ".xln" => Kind.Manufacturing,
                ".sch" or ".brd" when olderKiCad.Contains(path ?? string.Empty) => Kind.Older,
                ".sch" or ".brd" => Kind.Foreign,
                ".pro" => Kind.Older,
                ".kicad_prl" or ".kicad_dru" or ".kicad_wks" or ".kicad_jobset" or ".net" => Kind.Settings,
                _ when name is "fp-lib-table" or "sym-lib-table" or "fp-info-cache" => Kind.Settings,
                _ => Kind.Other,
            };
        }

        private static string Count(int count, string one, string many) =>
            count == 1 ? $"1 {one}" : $"{KiCadImportSummary.Number(count)} {many}";

        private static string Number(int count) => count.ToString(CultureInfo.InvariantCulture);
    }
}
