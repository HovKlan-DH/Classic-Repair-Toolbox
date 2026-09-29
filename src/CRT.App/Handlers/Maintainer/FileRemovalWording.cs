using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // How the files a publish REMOVES are described (owner, 2026-09-25: the list must be
    // "visible BEFORE the maintainer/admin approves it to either BETA or real ... so it is clear
    // what will happen").
    //
    // Used by the submission view (a BETA publish) and the BETA screen (a promotion), so the
    // two say the same thing about the same kind of act. Written only from the server's
    // FileRemovalPreview - the list shown is the list the server then removes, and is sent back
    // with the approval. Pure, so the words are tested.
    // ###########################################################################################
    public static class FileRemovalWording
    {
        // ###########################################################################################
        // The line above the list. `target` names the data: "the BETA data", "production". Null
        // when the server said nothing (an older server, which removed nothing).
        // ###########################################################################################
        public static string? Headline(FileRemovalPreview? preview, string target)
        {
            if (preview is null)
                return null;

            if (preview.IsBlocked)
                return preview.BlockedBecause;

            return preview.Files.Count switch
            {
                0 => $"No files are removed from {target}.",
                1 => $"Publishing REMOVES 1 file from {target} - nothing uses it any more:",
                _ => $"Publishing REMOVES {preview.Files.Count.ToString("N0", CultureInfo.InvariantCulture)} files from {target} - nothing uses them any more:"
            };
        }

        public static string Line(string path) => $"removed  {path}";

        // ###########################################################################################
        // The line beside the Approve / Publish button, so the removal is in front of the maintainer
        // at the moment they press it. Null when nothing is removed.
        // ###########################################################################################
        public static string? ApproveNote(FileRemovalPreview? preview)
        {
            int count = preview?.Files.Count ?? 0;

            return count switch
            {
                0 => null,
                1 => "Approving also REMOVES 1 file that nothing uses any more - it is listed with the files that change.",
                _ => $"Approving also REMOVES {count.ToString("N0", CultureInfo.InvariantCulture)} files that nothing uses any more - they are listed with the files that change."
            };
        }

        // " 2 unused files were removed." - appended to the message after a publish; empty when
        // nothing was.
        public static string Done(IReadOnlyList<string>? removed)
        {
            int count = removed?.Count ?? 0;

            return count switch
            {
                0 => string.Empty,
                1 => " 1 unused file was removed.",
                _ => $" {count.ToString("N0", CultureInfo.InvariantCulture)} unused files were removed."
            };
        }

        // Is `path` one of the files the preview removes? The server's own spelling is compared,
        // ignoring case - a board's citation and the file on disk may differ only in case.
        public static bool IsRemoved(FileRemovalPreview? preview, string path) =>
            preview is not null &&
            preview.Files.Contains(path, StringComparer.OrdinalIgnoreCase);
    }
}
