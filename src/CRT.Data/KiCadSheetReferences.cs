using System;
using System.Collections.Generic;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The child sheets a KiCad schematic names, and whether CRT may follow each (security review,
    // 2026-10-11).
    //
    // A hierarchical schematic names each child sheet by a path written INSIDE the file - its
    // "Sheetfile" property - and CRT's KiCad loader followed that path wherever it pointed. A
    // contributed file naming "//server/share/x.kicad_sch" therefore made every Windows CRT that
    // showed the board connect to that server and hand it the user's NTLM hash, with no click
    // beyond picking the board. CRT no longer follows such a path (KiCadRawProjectLoader), but the
    // CRTs already installed do - so this is the SERVER's rule, at finalise, at the publish to BETA
    // and at the publish to stable: a submitted .kicad_sch may name only
    //
    //   - a sheet inside its own "KiCad data" folder, and
    //   - a sheet that is itself a .kicad_sch (code review, 2026-10-11). An installed CRT reads ANY
    //     file a sheet names as a schematic and follows the sheets in it, while this rule looks
    //     inside .kicad_sch files only - so a sheet naming "board.kicad_pcb", whose own "Sheetfile"
    //     names a share, went round it. A sheet that is a .kicad_sch is checked by this rule itself.
    //
    // Read with the loader's own tokenizer (KiCadSExpressionTokens), and generously: EVERY token
    // that follows a "Sheetfile" token is checked, wherever it sits in the file, so nothing the
    // loader could follow escapes the check.
    // ###########################################################################################
    public static class KiCadSheetReferences
    {
        // The property a sheet names its file by - what the loader follows, and what is checked here.
        public const string SheetFileProperty = "Sheetfile";

        // The only kind of file a sheet may name.
        public const string SchematicExtension = ".kicad_sch";

        // ###########################################################################################
        // The longest sheet file name accepted, in characters: SubmissionPathRules' longest path in
        // the data tree, which every sheet a schematic names is a file of. Find reads no more than
        // one character past it of any token (code review, 2026-10-11), so a file holding one huge
        // token costs no memory to check - and a name it cut short is refused for its length.
        // ###########################################################################################
        public const int MaximumSheetFileLength = SubmissionPathRules.MaximumPathLength;

        // Does the rule look inside this file at all? Only a schematic names child sheets.
        public static bool AppliesTo(string? path) =>
            string.Equals(Path.GetExtension(path ?? string.Empty), KiCadSheetReferences.SchematicExtension, StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // Every sheet file the schematic names, in file order, read as it is asked for - so a file
        // naming millions of sheets is checked without holding them. A "Sheetfile" followed by a
        // parenthesis names nothing - the loader reads a list there, not a path. A name longer than
        // MaximumSheetFileLength comes back cut to one character more than that.
        // ###########################################################################################
        public static IEnumerable<string> Find(TextReader content)
        {
            ArgumentNullException.ThrowIfNull(content);

            var tokens = new KiCadSExpressionTokens(content, KiCadSheetReferences.MaximumSheetFileLength);
            bool afterSheetFile = false;

            while (tokens.TryRead(out string? token))
            {
                if (afterSheetFile && token is not KiCadSExpressionTokens.Open and not KiCadSExpressionTokens.Close)
                    yield return token;

                afterSheetFile = string.Equals(token, KiCadSheetReferences.SheetFileProperty, StringComparison.OrdinalIgnoreCase);
            }
        }

        // ###########################################################################################
        // Does this sheet file, named by the schematic at `schematicPath` (data-root-relative,
        // "Commodore/C64/250407/KiCad data/Pages/cpu.kicad_sch"), stay inside that board's
        // "KiCad data" folder?
        //
        // Resolved the way KiCad means it - relative to the naming schematic's own folder, "." and
        // ".." taken as folders - and then held to SubmissionPathRules' shape, which refuses what a
        // Windows CRT would read as somewhere else: a leading "/" ("//server/share"), a backslash,
        // a colon ("C:", a stream name), a control character, a trailing dot or space, a device
        // name. ".." may climb as far as "KiCad data" and no further; a schematic outside any
        // "KiCad data" folder may not climb at all. A name longer than MaximumSheetFileLength is
        // refused.
        //
        // A blank name is no reference - the loader skips it - so it passes.
        // ###########################################################################################
        public static bool StaysInKiCadFolder(string schematicPath, string? sheetFile)
        {
            if (string.IsNullOrWhiteSpace(sheetFile))
                return true;

            if (sheetFile.Length > KiCadSheetReferences.MaximumSheetFileLength)
                return false;

            string[] schematicSegments = schematicPath.Split('/');
            var folder = new List<string>(schematicSegments[..^1]);

            // "Manufacturer/Hardware/Board/KiCad data/..." - the folder ".." may not climb out of.
            int floor = schematicSegments.Length >= 5 &&
                        string.Equals(schematicSegments[3], DataTreeUsage.KiCadFolderName, StringComparison.OrdinalIgnoreCase)
                ? 4
                : folder.Count;

            foreach (string segment in sheetFile.Split('/'))
            {
                if (segment.Length == 0)
                    return false;

                if (segment == ".")
                    continue;

                if (segment == "..")
                {
                    if (folder.Count <= floor)
                        return false;

                    folder.RemoveAt(folder.Count - 1);
                    continue;
                }

                folder.Add(segment);
            }

            return SubmissionPathRules.IsSafelyShaped(string.Join('/', folder), out _);
        }

        // ###########################################################################################
        // Is this sheet file a schematic - a name ending ".kicad_sch", in any case? Read off the
        // name as written: a sheet whose last part is "." or ".." names a folder, and has no
        // extension either. A blank name is no reference, so it passes.
        // ###########################################################################################
        public static bool NamesASchematic(string? sheetFile) =>
            string.IsNullOrWhiteSpace(sheetFile) ||
            string.Equals(Path.GetExtension(sheetFile), KiCadSheetReferences.SchematicExtension, StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // The whole rule for one submitted file: false, with the reason a contributor can act on,
        // when the schematic names a sheet CRT may not follow. A file of another type always passes,
        // and `content` is then not read.
        //
        // The reason quotes the name at most 80 characters long, as SubmissionPathRules quotes a
        // path: a whole name of any length made a finding too big to store (code review,
        // 2026-10-11).
        // ###########################################################################################
        public static bool Check(string schematicPath, TextReader content, out string reason)
        {
            reason = string.Empty;

            if (!KiCadSheetReferences.AppliesTo(schematicPath))
                return true;

            foreach (string sheetFile in KiCadSheetReferences.Find(content))
            {
                string shown = SubmissionPathRules.Trim(sheetFile);

                if (sheetFile.Length > KiCadSheetReferences.MaximumSheetFileLength)
                {
                    reason =
                        $"It names a sheet by a path longer than {KiCadSheetReferences.MaximumSheetFileLength} characters: [{shown}]. " +
                        $"Keep every sheet of the project inside \"{DataTreeUsage.KiCadFolderName}\" and save it again from KiCad.";

                    return false;
                }

                if (!KiCadSheetReferences.StaysInKiCadFolder(schematicPath, sheetFile))
                {
                    reason =
                        $"It names a sheet outside the board's \"{DataTreeUsage.KiCadFolderName}\" folder: [{shown}]. " +
                        "KiCad saves a sheet's file relative to the sheet that uses it - keep every sheet of the " +
                        $"project inside \"{DataTreeUsage.KiCadFolderName}\" and save it again from KiCad.";

                    return false;
                }

                if (!KiCadSheetReferences.NamesASchematic(sheetFile))
                {
                    reason =
                        $"It names a sheet that is not a KiCad schematic ({KiCadSheetReferences.SchematicExtension}): [{shown}]. " +
                        "Every sheet of a KiCad project is a schematic file of its own - save the project again from KiCad.";

                    return false;
                }
            }

            return true;
        }
    }
}
