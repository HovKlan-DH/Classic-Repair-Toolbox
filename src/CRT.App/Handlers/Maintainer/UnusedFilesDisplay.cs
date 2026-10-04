using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // How the administrator's "Unused files" panel (Account screen; a window until 2026-09-27) reads (owner decision, 2026-09-25: "there
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
        // "the BETA data" / "production" - the names the rest of the Maintainer tab uses.
        public static string TreeName(string? tree) =>
            string.Equals(tree, "production", StringComparison.OrdinalIgnoreCase) ? "the stable data" : "the BETA data";

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

        // 812 bytes, 42.1 KB, 8.7 MB - FileSizeWording's, which every file tree uses too.
        public static string Size(long bytes) => FileSizeWording.Format(bytes);

        // ###########################################################################################
        // The list as the file tree draws it (2026-10-04, owner request: "the exact same tree-view
        // like it does in 'Systems' and 'Files' ... including visualization of images and opening of
        // files"): each file as it is in its tree, opened from there, with its size. Nothing about a
        // change - the list is what is there; removing is the button's.
        // ###########################################################################################
        public static IReadOnlyList<SystemFileEntry> TreeEntries(UnusedFileListing listing)
        {
            ArgumentNullException.ThrowIfNull(listing);

            SystemFileSource source = string.Equals(listing.Tree, "production", StringComparison.OrdinalIgnoreCase)
                ? SystemFileSource.Production
                : SystemFileSource.Beta;

            return listing.Files
                .Select(file => new SystemFileEntry(file.Path, SystemFileChange.Unchanged, source, SizeBytes: file.SizeBytes))
                .ToList();
        }

        private static string Count(int count, string noun) =>
            count == 1 ? $"1 {noun}" : $"{UnusedFilesDisplay.Number(count)} {noun}s";

        private static string Number(int count) => count.ToString("N0", CultureInfo.InvariantCulture);
    }
}
