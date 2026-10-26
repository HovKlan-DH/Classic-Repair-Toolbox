using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** AUTOMATIC REMOVAL STAYS WITH THE BOARD'S OWN FILES (owner decision, 2026-09-27). ***
    //
    // "It can delete any files inside its own folder - not outside it - its own board main
    // folder." A publish to BETA, a promotion to production and a push-back each remove the files a
    // board stops using - but only its own: a maintainer can never cause anything of another board
    // to disappear. The administrator clears anything else on purpose with Account > Unused files
    // (UnusedFileFlows), which this rule does not touch.
    //
    // *** IN THE CRT 4.0 TREE A BOARD'S FILES ARE STORED ONCE, UNDER THEIR CHECKSUM (2026-10-25). ***
    // Its folder holds only its board.json; every file's bytes are in "Files/<ab>/<sha256>.<ext>",
    // shared by every board that names the same bytes (DataTreeLayout). Such a stored file is
    // removable too - and the rule still holds, because DataTreeUsage removes one only when NO board
    // of that tree lists it any more: bytes another board names stay, and bytes nobody names can
    // belong to nobody.
    //
    // Paths compare ORDINAL - the server's trees are case-sensitive, so "commodore/c64/250407/x" is
    // not inside "Commodore/C64/250407".
    // ###########################################################################################
    public static class AutomaticRemovalScope
    {
        public static bool IsInsideBoardFolder(string? boardFolder, string? path)
        {
            string folder = AutomaticRemovalScope.Clean(boardFolder);
            string file = AutomaticRemovalScope.Clean(path);

            return folder.Length > 0 && file.StartsWith(folder + "/", StringComparison.Ordinal);
        }

        // A stored file of the 4.0 tree: "Files/<ab>/<sha256>.<ext>".
        public static bool IsStoredFile(string? path) =>
            AutomaticRemovalScope.Clean(path).StartsWith(DataTreeLayout.FilesFolder + "/", StringComparison.Ordinal);

        // Those of `paths` a publish, promotion or push-back of this board may remove.
        public static IReadOnlyList<string> Within(string? boardFolder, IEnumerable<string>? paths) =>
            (paths ?? [])
                .Where(path => AutomaticRemovalScope.IsInsideBoardFolder(boardFolder, path) || AutomaticRemovalScope.IsStoredFile(path))
                .ToList();

        private static string Clean(string? path) =>
            (path ?? string.Empty).Replace('\\', '/').Trim().Trim('/');
    }
}
