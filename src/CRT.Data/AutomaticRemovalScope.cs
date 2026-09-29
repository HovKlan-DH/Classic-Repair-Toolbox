using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** AUTOMATIC REMOVAL STAYS INSIDE THE SYSTEM'S OWN FOLDER (owner decision, 2026-09-27). ***
    //
    // "It can delete any files inside its own folder - not outside it - its own system main
    // folder." A publish to BETA, a promotion to production and a push-back each remove the files a
    // board stops using - but ONLY under that board's own folder ("Commodore/C64/250407/..."). A file
    // in "Generic shared files", "<Manufacturer>/Shared files" or another system's folder is never
    // removed by one of them: it may become unused (an orphan), and the administrator clears those
    // on purpose with Admin > Unused files (UnusedFileFlows), which this rule does not touch.
    //
    // It is what made it safe to drop the administrator's second approval for a submission that
    // merely ADDS a shared file (SubmissionSharedFiles): a board's maintainer can no longer cause
    // anything outside their own board to disappear.
    //
    // Paths compare ORDINAL - the server's trees are case-sensitive, so "commodore/c64/250407/x" is
    // not inside "Commodore/C64/250407".
    // ###########################################################################################
    public static class AutomaticRemovalScope
    {
        public static bool IsInsideSystemFolder(string? systemFolder, string? path)
        {
            string folder = AutomaticRemovalScope.Clean(systemFolder);
            string file = AutomaticRemovalScope.Clean(path);

            return folder.Length > 0 && file.StartsWith(folder + "/", StringComparison.Ordinal);
        }

        // Those of `paths` a publish, promotion or push-back of this system may remove.
        public static IReadOnlyList<string> Within(string? systemFolder, IEnumerable<string>? paths) =>
            (paths ?? []).Where(path => AutomaticRemovalScope.IsInsideSystemFolder(systemFolder, path)).ToList();

        private static string Clean(string? path) =>
            (path ?? string.Empty).Replace('\\', '/').Trim().Trim('/');
    }
}
