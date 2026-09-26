using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Handlers.DataHandling;

namespace CRT.Maintainer.Handlers
{
    // ###########################################################################################
    // The two decisions behind the table's file hover card in the maintainer application (owner
    // request, 2026-09-26) - pure, so they are tested. The fetching is ReviewTableFileSource's.
    // ###########################################################################################
    public static class ReviewTableFiles
    {
        // ###########################################################################################
        // The submitted file's hash for a path the table names, or null when the submission carries
        // no file at that path. Ordinal: the server's tree is case-sensitive.
        //
        // The submission's files are listed by the server (the detail's submittedFiles) - every file
        // its rows cite, including those it cites unchanged, which the server imported when the
        // submission was made. A path with no entry is one the table names but the submission does
        // not (typed in the table and not yet saved), and is looked for among the published files.
        // ###########################################################################################
        public static string? SubmittedHashFor(IReadOnlyList<SubmittedFileFact>? files, string? path)
        {
            string wanted = path?.Trim() ?? string.Empty;

            if (wanted.Length == 0 || files is null)
                return null;

            return files.FirstOrDefault(file => string.Equals(file.Path, wanted, StringComparison.Ordinal))?.Sha256;
        }

        // ###########################################################################################
        // The submitted file's hash one side of the hover card reads, or null to read the PUBLISHED
        // file at the path instead.
        //
        // *** A NEW SYSTEM'S "PUBLISHED" SIDE IS THE SUBMISSION TOO. *** Its table is compared with
        // the submission itself as it was opened (MaintainerMain.Table.cs), and the published tree
        // holds nothing of it. Read from the tree, every picture of a new system was shown beside
        // "There is no file at this path" under "Before (published)" - reported, 2026-09-26: "this
        // makes no sense at all". Read from the submission, both sides are the same file, and the
        // card shows it once.
        // ###########################################################################################
        public static string? HashToRead(
            IReadOnlyList<SubmittedFileFact>? files,
            string? path,
            CRT.BoardTableFileSide side,
            bool nothingPublished) =>
            side == CRT.BoardTableFileSide.Current || nothingPublished
                ? ReviewTableFiles.SubmittedHashFor(files, path)
                : null;

        // ###########################################################################################
        // What the two sides are called when the card shows both. For a new system the older side
        // is the submission as it came in and the newer one the maintainer's own change - it is
        // only ever beside another when the maintainer has named another file.
        // ###########################################################################################
        public static (string Published, string Current) SideLabels(bool nothingPublished) =>
            nothingPublished
                ? ("As submitted", "Your change")
                : ("Before (published)", "After (submitted)");

        // ###########################################################################################
        // How long a file opened from the table is left in the temp folder (code review,
        // 2026-09-26). Each open writes the contributor's bytes into a folder of its own and hands
        // it to the operating system, so it cannot be deleted while the viewer may still be
        // reading it - and nothing deleted them at all, so months of reviewing board scans piled up
        // in the maintainer's temp folder for ever.
        //
        // A day is long past any viewer session, and the sweep runs on the NEXT open rather than on
        // a timer, so nothing is deleted while the application is idle with a PDF still on screen.
        // ###########################################################################################
        public static readonly TimeSpan OpenedFileLifetime = TimeSpan.FromDays(1);

        // Which of the temp folders left by earlier opens are old enough to delete: those last
        // written before `now - OpenedFileLifetime`. Pure, so the rule is tested rather than the
        // filesystem.
        public static bool IsStaleOpenedFolder(DateTimeOffset lastWrittenUtc, DateTimeOffset nowUtc) =>
            nowUtc - lastWrittenUtc > ReviewTableFiles.OpenedFileLifetime;

        // ###########################################################################################
        // The name a file is saved under before it is handed to the operating system to open, or
        // false for a type that is not opened at all.
        //
        // *** ONLY THE TYPES A SUBMISSION MAY CARRY (SubmissionFileRules.AllowedExtensions). ***
        // These are a contributor's bytes; nothing else can arrive in a submission, and nothing else
        // is opened from one.
        //
        // *** A WEB PAGE OPENS AS TEXT. *** Opened as .html it would run whatever script the
        // contributor put in it, in the maintainer's browser. Saved as .txt, it is read, which is
        // what a review needs.
        //
        // Only the file's own name is kept - never its folders - and characters no filesystem allows
        // are replaced, so the name cannot place the file anywhere but the folder it is written to.
        // ###########################################################################################
        public static bool TryGetOpenName(string? path, out string fileName)
        {
            fileName = string.Empty;

            string name = Path.GetFileName((path ?? string.Empty).Replace('\\', '/').Split('/').Last()).Trim();
            string extension = Path.GetExtension(name);

            if (name.Length == 0 || !SubmissionFileRules.AllowedExtensions.Contains(extension))
                return false;

            foreach (char invalid in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(invalid, '_');
            }

            bool isWebPage =
                string.Equals(extension, ".html", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(extension, ".htm", StringComparison.OrdinalIgnoreCase);

            fileName = isWebPage ? name + ".txt" : name;
            return true;
        }
    }
}
