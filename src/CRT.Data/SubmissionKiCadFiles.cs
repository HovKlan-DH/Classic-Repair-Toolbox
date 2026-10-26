using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A BOARD'S KiCad DATA TRAVELS IN ITS SUBMISSION (owner decision, 2026-09-26: "when a person
    // submitting anything from his local PC, then I expect that it will send everything the server
    // does not already have").
    //
    // The "KiCad data" files are the one part of a board no row cites - CRT finds them by NAME
    // (DataTreeUsage.KiCadFolderName) - so a rows-only manifest would leave them behind. In the CRT
    // 4.0 tree (2026-10-25) a board names them in its own file list, "KiCad data/<...>"
    // ("KiCad data/Open128.kicad_pcb", "KiCad data/Pages/vic.kicad_sch" - a project keeps its
    // layout, KiCad's sheets name each other by it), and a submission carries every one the board
    // lists. This class is the ONE rule for which of them a submission may carry:
    //
    //   - a name under "KiCad data/" of a type CRT itself reads (ComponentListBuilder.
    //     IsSupportedKiCadRawFile: .kicad_pcb, .kicad_sch, .kicad_pro) - the same filter the app's
    //     own "KiCad data" import applies (KiCadRawFileScanner);
    //   - OR any other name under "KiCad data/" the published board already lists with the same
    //     bytes. The converted trees carry a few stragglers nothing reads (a generated
    //     KiCad-traces.json, one legacy .sch); a board keeps them for as long as nobody changes
    //     them, and a submission cannot bring a new one - .json stays off the wire on purpose (see
    //     SubmissionFileRules.AllowedExtensions).
    //
    // IsSubmittable is that rule for the server (SubmissionFileRules, PublishPlan and
    // AmendSubmissionFlow all ask it); Collect is the client's side, gathering the names a draft's
    // submission carries.
    // ###########################################################################################
    public static class SubmissionKiCadFiles
    {
        // ###########################################################################################
        // Is this name a KiCad file this submission may carry? `published` is the published board's
        // file of the same name, when there is one - what lets an untouched straggler stay.
        // ###########################################################################################
        public static bool IsSubmittable(SubmissionManifest manifest, string? name, string? publishedSha256 = null, string? sha256 = null)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            if (!SubmissionKiCadFiles.IsInKiCadFolder(name))
                return false;

            if (ComponentListBuilder.IsSupportedKiCadRawFile(name!))
                return true;

            return !string.IsNullOrWhiteSpace(publishedSha256) &&
                   string.Equals(publishedSha256, sha256, StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // A STRAGGLER: a name in the "KiCad data" folder of a type CRT does not read (a generated
        // KiCad-traces.json, a legacy .sch). IsSubmittable admits one ONLY as the bytes the published
        // board already lists under that name, so its content is the published file's - nothing new
        // reaches the tree through it, and no content rule knows its type
        // (SubmissionContentRules.MatchesWholeFile passes it; code review, 2026-10-25: finalise
        // refused it, and with it every submission and maintainer edit of the board).
        // ###########################################################################################
        public static bool IsStraggler(string? name) =>
            SubmissionKiCadFiles.IsInKiCadFolder(name) && !ComponentListBuilder.IsSupportedKiCadRawFile(name!);

        // Whether a board's file name is inside its "KiCad data" folder - "KiCad data/x", with
        // something after the folder. The folder is compared without regard to capitals, like every
        // folder name (DataTreeUsage).
        public static bool IsInKiCadFolder(string? name)
        {
            string normalised = (name ?? string.Empty).Replace('\\', '/').Trim().TrimStart('/');
            string prefix = DataTreeUsage.KiCadFolderName + "/";

            return normalised.Length > prefix.Length &&
                   normalised.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // The same question for a name already known to be a board's file - what the Maintainer tab
        // asks of a SubmittedFileFact. The TYPE check matters here too: counting an ordinary file
        // that happens to sit in the folder as KiCad data would overstate what the rule admitted.
        // ###########################################################################################
        public static bool IsKiCadDataPath(string? path) =>
            !string.IsNullOrWhiteSpace(path) &&
            ComponentListBuilder.IsSupportedKiCadRawFile(path) &&
            path.Split('/').Any(segment =>
                string.Equals(segment, DataTreeUsage.KiCadFolderName, StringComparison.OrdinalIgnoreCase));

        // ###########################################################################################
        // Every KiCad name a draft's submission carries - every file its board file lists under
        // "KiCad data/" (DraftFiles.IsKiCad), sorted for a stable manifest. A draft's board file
        // lists the published board's KiCad files as well as its own, so the complete set goes:
        // the published ones travel as "already held" (the server takes their bytes from its own
        // tree), never reading as removed.
        //
        // *** AND THE ONES NOT LISTED YET (code review, 2026-10-25). *** `arrived` is what lies in
        // "<draft>/KiCad data" for the draft's next write to take in (DraftFiles.ArrivedKiCadNames):
        // a project imported with no write after it was drawn by the Schematics tab and counted by
        // the Drafts row, and then left out of the submission - whose fingerprint covered it, so
        // Submit greyed out as "already sent".
        // ###########################################################################################
        public static IReadOnlyList<string> Collect(BoardDocument? board, IEnumerable<string>? arrived = null) =>
            (board?.Files ?? [])
                .Where(DraftFiles.IsKiCad)
                .Select(file => file.Name)
                .Concat((arrived ?? []).Where(SubmissionKiCadFiles.IsInKiCadFolder).Select(BoardFileNames.Normalise))
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Order(StringComparer.Ordinal)
                .ToList();
    }
}
