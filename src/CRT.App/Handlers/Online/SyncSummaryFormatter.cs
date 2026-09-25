using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.OnlineHandling
{
    // ###########################################################################################
    // Builds the per-file breakdown that follows the one-line sync summary in the log.
    //
    // The summary line (and the banner in the main window) only ever says how MANY files were new,
    // updated, failed or rejected - which is no help when you want to know which board or image
    // actually changed. This turns the same four sets into indented, sorted lists so the log
    // answers that question directly.
    //
    // Pure string work with no I/O, so it lives here rather than inside SyncFilesAsync.
    // Files that were already up to date are deliberately NOT listed: that set is the whole data
    // folder on a normal launch (thousands of entries), and it is the one group the summary does
    // not invite a question about.
    // ###########################################################################################
    internal static class SyncSummaryFormatter
    {
        // Indentation matches the existing "    Sent:[...]" style used elsewhere in the log.
        private const string GroupIndent = "    ";
        private const string FileIndent = "        ";

        // ###########################################################################################
        // Returns the log lines naming every file behind the summary counts, grouped by outcome.
        // Empty groups are skipped entirely, so a clean sync of nothing but new files adds one
        // header and its files - not three empty headers. Returns an empty list when nothing at all
        // was new, updated, failed or rejected.
        // ###########################################################################################
        internal static IReadOnlyList<string> BuildFileBreakdown(
            IReadOnlyList<string>? newFiles,
            IReadOnlyList<string>? updatedFiles,
            IReadOnlyList<string>? failedFiles,
            IReadOnlyList<string>? invalidFiles)
        {
            var lines = new List<string>();

            SyncSummaryFormatter.AppendGroup(lines, "New", newFiles);
            SyncSummaryFormatter.AppendGroup(lines, "Updated", updatedFiles);
            SyncSummaryFormatter.AppendGroup(lines, "Failed", failedFiles);
            SyncSummaryFormatter.AppendGroup(lines, "Rejected", invalidFiles);

            return lines;
        }

        // ###########################################################################################
        // Appends one "<Label> [count]:" header and its file lines, sorted so the same sync always
        // logs the same order regardless of the order the manifest happened to arrive in. Blank
        // entries are dropped - an unnamed file line would be noise, not information.
        // ###########################################################################################
        private static void AppendGroup(List<string> lines, string label, IReadOnlyList<string>? files)
        {
            if (files == null || files.Count == 0)
                return;

            var named = files
                .Where(file => !string.IsNullOrWhiteSpace(file))
                .Select(file => file.Trim())
                .ToList();

            if (named.Count == 0)
                return;

            named.Sort(StringComparer.OrdinalIgnoreCase);

            lines.Add($"{SyncSummaryFormatter.GroupIndent}{label} [{named.Count}]:");

            foreach (var file in named)
                lines.Add($"{SyncSummaryFormatter.FileIndent}[{file}]");
        }
    }
}
