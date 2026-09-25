using System;
using System.Globalization;
using Handlers.DataHandling;

namespace CRT.Maintainer.Handlers
{
    // ###########################################################################################
    // How the administrator's "Unused files" window reads (owner decision, 2026-09-25: "there
    // must be no orphan files").
    //
    // Pure, so the words are tested - the rule ProductionDisplay and MaintainerAssignmentDisplay
    // follow. One rule here is a decision rather than formatting: the Remove button needs a
    // COMPLETE list with something on it AND the administrator's own tick in "I have looked
    // through this list" (CanRemove) - the first clean-up of a tree is looked at before anything
    // goes.
    // ###########################################################################################
    public static class UnusedFilesDisplay
    {
        // "the BETA data" / "production" - the names the rest of the maintainer app uses.
        public static string TreeName(string? tree) =>
            string.Equals(tree, "production", StringComparison.OrdinalIgnoreCase) ? "production" : "the BETA data";

        public static string Summary(UnusedFileListing listing)
        {
            ArgumentNullException.ThrowIfNull(listing);

            string tree = UnusedFilesDisplay.TreeName(listing.Tree);

            if (!listing.IsComplete)
            {
                return $"The files in {tree} could not all be checked, so none can be named as unused and nothing can be removed: " +
                    string.Join(" ", listing.Problems);
            }

            string checkedAgainst =
                $"{UnusedFilesDisplay.Number(listing.FileCount)} files checked against " +
                $"{UnusedFilesDisplay.Count(listing.MasterCount, "master workbook")} and " +
                $"{UnusedFilesDisplay.Count(listing.BoardWorkbookCount, "board workbook")}";

            return listing.Files.Count == 0
                ? $"Nothing in {tree} is unused ({checkedAgainst})."
                : $"{UnusedFilesDisplay.Count(listing.Files.Count, "file")} in {tree} that nothing uses, " +
                  $"{UnusedFilesDisplay.Size(listing.TotalBytes)} ({checkedAgainst}).";
        }

        // "Generic shared files/Component images/7408.jpg  (42.1 KB)"
        public static string Line(UnusedFileEntry entry)
        {
            ArgumentNullException.ThrowIfNull(entry);
            return $"{entry.Path}  ({UnusedFilesDisplay.Size(entry.SizeBytes)})";
        }

        public static string RemoveButton(UnusedFileListing? listing) =>
            listing is null || listing.Files.Count == 0
                ? "Remove unused files"
                : $"Remove {UnusedFilesDisplay.Count(listing.Files.Count, "file")} ({UnusedFilesDisplay.Size(listing.TotalBytes)})";

        public static bool CanRemove(UnusedFileListing? listing, bool lookedThrough) =>
            listing is not null && listing.IsComplete && listing.Files.Count > 0 && lookedThrough;

        // What happened, in one sentence or two.
        public static string Result(UnusedFileRemovalResult result)
        {
            ArgumentNullException.ThrowIfNull(result);

            if (!string.IsNullOrWhiteSpace(result.NotDoneBecause))
                return "Nothing was removed: " + result.NotDoneBecause;

            string removed = result.Removed.Count == 0
                ? "No files were removed."
                : $"{UnusedFilesDisplay.Count(result.Removed.Count, "file")} removed from {UnusedFilesDisplay.TreeName(result.Tree)}.";

            return result.Kept.Count == 0
                ? removed
                : $"{removed} {UnusedFilesDisplay.Count(result.Kept.Count, "file")} kept - something uses it again, or it was already gone.";
        }

        // 812 bytes, 42.1 KB, 8.7 MB - 1024-based, one decimal, the same in every locale.
        public static string Size(long bytes)
        {
            if (bytes < 1024)
                return bytes == 1 ? "1 byte" : $"{bytes.ToString(CultureInfo.InvariantCulture)} bytes";

            double kb = bytes / 1024.0;

            if (kb < 1024)
                return $"{kb.ToString("0.0", CultureInfo.InvariantCulture)} KB";

            return $"{(kb / 1024.0).ToString("0.0", CultureInfo.InvariantCulture)} MB";
        }

        private static string Count(int count, string noun) =>
            count == 1 ? $"1 {noun}" : $"{UnusedFilesDisplay.Number(count)} {noun}s";

        private static string Number(int count) => count.ToString("N0", CultureInfo.InvariantCulture);
    }
}
