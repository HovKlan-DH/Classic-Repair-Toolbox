using System;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHERE ONE SUBMITTED FILE'S BYTES ACTUALLY ARE.
    //
    // A submission carries the system as it should READ after the merge, which means its file list
    // is drawn from the MERGED board data - official rows plus drafted ones. Those two kinds of row
    // do not keep their bytes in the same place, and that is not an accident:
    //
    //   - an OFFICIALLY published file lives under Data/, the sync-owned tree;
    //   - a DRAFTED file lives under "<draft system folder>/Files/", because Data/ is freely
    //     overwritten by the next sync and a drafted attachment there would simply vanish.
    //
    // So a draft over a published system - a typo fix that adds one new photo to a board with 240
    // existing ones - references BOTH roots in a single submission. Resolving the whole list
    // against only the draft folder reports 240 published files as "not on disk"; resolving it
    // against only Data/ misses the one file the contributor actually made.
    //
    // *** THIS IS THE SAME TWO-ROOT RULE DraftFileResolver ALREADY APPLIES *** for opening a file
    // to show it on screen, and the order matches deliberately: the official copy first, the
    // drafted copy second. What is submitted must be what the contributor has been looking at, and
    // a different precedence here would mean the bytes uploaded are not the bytes the app
    // displayed.
    //
    // WHAT THIS ADDS over DraftFileResolver is the containment check. DraftFileResolver's input is
    // a path the app itself stored; this one's is about to be read and sent to a public server, so
    // each candidate root is checked through SubmissionPathRules first - a stored path escaping its
    // root ("../../.ssh/id_rsa") must not be readable merely because a File.Exists succeeded.
    //
    // Pure apart from File.Exists, so it is unit tested directly rather than through the HTTP
    // client that calls it.
    // ###########################################################################################
    public static class SubmissionFileLocator
    {
        // ###########################################################################################
        // Resolves one relative path to the absolute file that will be hashed and uploaded.
        //
        // Returns false with a reason fit to show the contributor when the path is unusable or
        // neither root holds it. The reason distinguishes the two cases on purpose: a REFUSED path
        // is a broken row that needs fixing, while a MISSING one is usually a file that was moved
        // or deleted after it was added, and telling someone their file is invalid when it is
        // merely absent sends them looking in the wrong place.
        // ###########################################################################################
        public static bool TryLocate(
            string dataRoot,
            string draftSystemFolder,
            string relativePath,
            out string absolutePath,
            out string reason)
        {
            absolutePath = string.Empty;

            bool anyRootChecked = false;
            string refusalReason = string.Empty;

            // ###########################################################################################
            // *** THE DRAFTED COPY IS TRIED FIRST, AND ITS PATH CHANGED IN PHASE 6 (2026-09-23). ***
            //
            // It used to live under a "Files" subfolder of the draft, with the stored path appended
            // whole. A draft folder is now a BOARD folder, so the file sits where a published board
            // keeps it - which means the system's own "Manufacturer/Hardware/Board/" prefix has to
            // be stripped, because the draft folder already IS those segments.
            //
            // Order flipped to match DraftFileResolver: when a draft exists its files ARE the
            // board's files, including replacements for images that also exist in Data/. Taking the
            // published copy first would upload the OLD picture for an image the contributor has
            // replaced - silently, since both files exist and both are valid.
            //
            // Each root is still resolved through SubmissionPathRules separately. That containment
            // check is the security property here: these bytes are about to be read and sent to a
            // public server, so a stored path escaping its root must not become readable merely
            // because File.Exists succeeded.
            // ###########################################################################################
            string draftedRelative = relativePath;
            string draftRoot = string.Empty;

            if (!string.IsNullOrWhiteSpace(draftSystemFolder))
            {
                string? withinSystem = DraftFolderLayout.RelativeToSystemFolder(
                    SubmissionFileLocator.SystemKeyFromFolder(draftSystemFolder),
                    relativePath);

                if (withinSystem is not null)
                {
                    draftRoot = draftSystemFolder;
                    draftedRelative = withinSystem;
                }
            }

            foreach ((string root, string relative) in new[]
            {
                (draftRoot, draftedRelative),
                (dataRoot, relativePath),
            })
            {
                if (string.IsNullOrWhiteSpace(root))
                    continue;

                anyRootChecked = true;

                if (!SubmissionPathRules.TryResolve(root, relative, out string candidate, out string why))
                {
                    // A path the rules refuse is refused for EVERY root - it is the stored string
                    // that is wrong, not where it was looked for - so the reason is kept and
                    // reported once rather than per root.
                    refusalReason = why;
                    continue;
                }

                if (File.Exists(candidate))
                {
                    absolutePath = candidate;
                    reason = string.Empty;
                    return true;
                }
            }

            if (refusalReason.Length > 0)
            {
                reason = refusalReason;
                return false;
            }

            reason = anyRootChecked
                ? "it is referenced by the data but is not on disk. It may have been moved or deleted."
                : "there is nowhere to look for it - no data folder and no draft folder are known.";

            return false;
        }

        // ###########################################################################################
        // The system identity a draft folder represents - see DraftFileResolver.SystemKeyFromFolder,
        // which does the same job for the same reason. RelativeToSystemFolder ignores everything
        // after the last "/", so only the folder's own last three segments matter.
        // ###########################################################################################
        private static string SystemKeyFromFolder(string draftSystemFolder)
        {
            string[] segments = (draftSystemFolder ?? string.Empty)
                .Split([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar],
                    StringSplitOptions.RemoveEmptyEntries);

            if (segments.Length < 3)
            {
                return string.Empty;
            }

            return string.Join('/', segments[^3..]) + "/placeholder.xlsx";
        }
    }
}
