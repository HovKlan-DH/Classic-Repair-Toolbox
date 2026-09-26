using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A BOARD'S KiCad DATA TRAVELS IN ITS SUBMISSION (owner decision, 2026-09-26: "when a person
    // submitting anything from his local PC, then I expect that it will send everything the server
    // does not already have").
    //
    // The "KiCad data" folder is the one part of a board no row cites - CRT finds it by NAME
    // (DataTreeUsage.KiCadFolderName) - so the rows-only manifest silently left it behind: a new
    // system was published without its traces, and a re-calibrated project never reached anyone.
    // This class is the ONE rule for which files that folder contributes:
    //
    //   - ONLY the types CRT itself reads (ComponentListBuilder.IsSupportedKiCadRawFile:
    //     .kicad_pcb, .kicad_sch, .kicad_pro) - the same filter the app's own "KiCad data" import
    //     applies (KiCadRawFileScanner). The shipped trees carry two stragglers nothing reads (a
    //     generated KiCad-traces.json report, one legacy .sch); they stay published (the folder is
    //     never emptied by a publish) but are not what a submission carries, and .json stays off
    //     the wire on purpose - see SubmissionFileRules.AllowedExtensions on why.
    //   - ONLY inside the submission's OWN board folder. Another board's KiCad data is that
    //     board's to change.
    //
    // IsSubmittable is that rule for the server (SubmissionFileRules, PublishPlan and
    // AmendSubmissionFlow all ask it); Collect is the client's side of the same rule, gathering
    // the paths a draft's submission should carry.
    // ###########################################################################################
    public static class SubmissionKiCadFiles
    {
        // ###########################################################################################
        // Is this path a KiCad file this submission may carry: inside the submission's own board
        // folder, directly under "KiCad data" (sub-folders included - a multi-sheet project keeps
        // its "Pages/vic.kicad_sch" layout), and of a type CRT reads?
        //
        // The folder name is compared case-insensitively, like every folder comparison
        // (DataTreeUsage); the board prefix is SubmissionFileScopes' own ordinal rule.
        // ###########################################################################################
        public static bool IsSubmittable(SubmissionManifest manifest, string? path)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            if (SubmissionFileScopes.Classify(manifest, path) != SubmissionFileScope.Own)
                return false;

            if (!ComponentListBuilder.IsSupportedKiCadRawFile(path ?? string.Empty))
                return false;

            // Own scope guarantees the three board segments lead; the KiCad folder is the fourth,
            // and the file (or its sub-folders) follow.
            string[] segments = path!.Split('/');

            return segments.Length >= 5 &&
                   string.Equals(segments[3], DataTreeUsage.KiCadFolderName, StringComparison.OrdinalIgnoreCase);
        }

        // The same question for a path already known to be somewhere in a data tree - what the
        // maintainer application asks of a SubmittedFileFact, whose scope the server has already
        // checked. The TYPE check matters here too: a row may cite an ordinary file that happens to
        // sit inside a "KiCad data" folder, and counting it as KiCad data would overstate what the
        // exemption actually admitted.
        public static bool IsKiCadDataPath(string? path) =>
            !string.IsNullOrWhiteSpace(path) &&
            ComponentListBuilder.IsSupportedKiCadRawFile(path) &&
            path.Split('/').Any(segment =>
                string.Equals(segment, DataTreeUsage.KiCadFolderName, StringComparison.OrdinalIgnoreCase));

        // ###########################################################################################
        // Every KiCad file the submission should carry, as data-root-relative paths
        // ("Manu1/Hardware1/Board1/KiCad data/board.kicad_pcb"), sorted for a stable manifest.
        //
        // THE UNION OF BOTH ROOTS, the draft's copy winning - the same two-roots-draft-first rule
        // SubmissionFileLocator applies when the bytes are read, so what is listed here is what
        // will be hashed and sent. The official copy is included because the manifest is the
        // complete intended state: a published board's KiCad files travel as "already held" (the
        // server imports them from its own tree, nothing uploads) rather than reading as absent.
        //
        // A root that does not exist contributes nothing - a draft-only system has no official
        // folder, and most boards have no KiCad data at all.
        // ###########################################################################################
        public static IReadOnlyList<string> Collect(
            string systemId,
            string? draftSystemFolder,
            string? officialSystemFolder)
        {
            if (string.IsNullOrWhiteSpace(systemId))
                return [];

            // ###########################################################################################
            // Relative path inside "KiCad data" -> its data-root-relative path; the DRAFT's copy
            // wins, since a draft's files ARE the board's files.
            //
            // *** MATCHED CASE-INSENSITIVELY, BUT THE OFFICIAL SPELLING IS KEPT (code review,
            // 2026-09-26). *** Windows and macOS fold "Board.kicad_pcb" and "board.kicad_pcb" into
            // one file, so a draft copied from a published board can differ only in case - and
            // sending the draft's spelling would publish a SECOND file beside the official one on
            // the case-sensitive server (a publish never empties the folder), leaving CRT to read
            // two near-identical projects. Keeping the published spelling means the file is
            // replaced, which is what the contributor means. A file the server does not have yet
            // keeps the draft's own spelling, which is then the only one there is.
            //
            // The same rule as SubmissionFileRules' case-variant refusal, applied before the paths
            // are built rather than reported afterwards.
            // ###########################################################################################
            var byRelative = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // The OFFICIAL folder is read first, so its spelling is the one in the dictionary; the
            // draft then contributes only what the published board does not already hold.
            foreach (string? root in new[] { officialSystemFolder, draftSystemFolder })
            {
                if (string.IsNullOrWhiteSpace(root))
                    continue;

                string folder = Path.Combine(root, DataTreeUsage.KiCadFolderName);

                if (!Directory.Exists(folder))
                    continue;

                foreach (string file in Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories))
                {
                    if (!ComponentListBuilder.IsSupportedKiCadRawFile(file))
                        continue;

                    string relative = Path.GetRelativePath(folder, file).Replace(Path.DirectorySeparatorChar, '/');

                    byRelative.TryAdd(relative, $"{systemId}/{DataTreeUsage.KiCadFolderName}/{relative}");
                }
            }

            return byRelative.Values.Order(StringComparer.Ordinal).ToList();
        }
    }
}
