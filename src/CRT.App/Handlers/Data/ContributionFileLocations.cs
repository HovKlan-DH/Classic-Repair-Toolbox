using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The folders the component editor offers in its "File location" drop-down - folders of the
    // BOARD'S OWN FILE NAMES (CRT 4.0, 2026-10-25).
    //
    // *** IN 4.0 A ROW NAMES A FILE OF ITS BOARD *** ("Scope baseline/U19_1_PAL.png" -
    // BoardFileNames), and the bytes lie in a store laid out by checksum ("Files/ab/<sha256>.png").
    // So the data tree's folders say nothing about where a file of the board goes, and the list is
    // no longer read off the disk: it was built from the data tree's and the drafts tree's end
    // folders, which in a 4.0 tree are the store's checksum folders and CRT's own bookkeeping.
    //
    // It offers:
    //   - every folder a name the board's rows cite already uses ("Truth tables" on the C64), and
    //   - the four folders every converted board keeps its files in (UsualFolders), so a new board,
    //     or one with no scope baseline yet, offers them too (the trees measured 2026-10-25:
    //     schematics at the top, then exactly these four, and "Truth tables" on two boards).
    //
    // A file is one of the board's whatever folder it is in - the server takes any name inside the
    // board (the Excel tree's "[x] belongs to another board" went with the shared folders) - so
    // there is nothing to narrow by board any more. What is left out is what can never hold a file
    // of a board: see IsWritableFolder.
    //
    // Sorted ignoring capitals, the order the window re-sorts into once a row's own location is
    // added; listed once whatever the capitals, since a name is matched ignoring them.
    // ###########################################################################################
    internal static class ContributionFileLocations
    {
        public static IReadOnlyList<string> UsualFolders { get; } =
        [
            "Board local files",
            "Component images",
            "Component local files",
            DraftFolderLayout.ScopeBaselineFolderName,
        ];

        public static List<string> ForBoard(BoardData? board)
        {
            IEnumerable<string> used = board == null
                ? []
                : DraftFiles.CitedNames(board)
                    .Select(BoardFileNames.Normalise)
                    .Where(name => name.Contains('/'))
                    .Select(name => name[..name.LastIndexOf('/')]);

            return ContributionFileLocations.UsualFolders
                .Concat(used)
                .Where(ContributionFileLocations.IsWritableFolder)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(folder => folder, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ###########################################################################################
        // Whether a new file may be filed in this folder of the board's names. Left out:
        //
        //   - the board's store ("Files/..."): a name there is never a file of the board, so a draft
        //     write would not list it and the file would be lost (DraftFiles takes in only what can
        //     be a contributor's file);
        //   - a folder with a segment starting with a dot - "." and "..", CRT's bookkeeping
        //     (".crt-..."), and every hidden folder: the server refuses any such name
        //     (SubmissionFileRules.TryCheckName, "file.hidden_name"), and a file filed there was
        //     copied in and listed, only to be refused at Submit (code review, 2026-10-25);
        //   - a drive or other ":" a name cannot carry, which would name a file somewhere else.
        //
        // "/"-separated; a backslash and slashes at either end are forgiven, as the writer forgives
        // them (ComponentBoardWriter.JoinFilePath).
        // ###########################################################################################
        public static bool IsWritableFolder(string? folder)
        {
            string trimmed = folder?.Trim().Replace('\\', '/').Trim('/') ?? string.Empty;

            if (trimmed.Length == 0)
            {
                return false;
            }

            string[] segments = trimmed.Split('/');

            if (segments.Any(segment => segment.Trim().Length == 0 || segment.StartsWith('.') || segment.Contains(':')))
            {
                return false;
            }

            return !string.Equals(segments[0], DataTreeLayout.FilesFolder, StringComparison.OrdinalIgnoreCase);
        }

        // The kinds of file row the component editor has - for the folder a picked file goes in when
        // its row names none (DefaultFolderFor).
        public enum FileRowKind
        {
            ComponentPicture,
            ScopeBaseline,
            ComponentFile,
            BoardFile
        }

        // ###########################################################################################
        // *** A PICKED FILE WITH NO FOLDER GOES IN ITS KIND'S USUAL ONE (CRT 4.0 step 1, 2026-10-25).
        // *** "Save to draft" refused a newly picked file with an empty "File location" (owner report,
        // 2026-09-25: saved like that, the file sat at the top of the board); unfinished work is fine
        // now, so the file is filed where every converted board keeps that kind (UsualFolders) - a
        // picture in "Component images", a scope baseline in "Scope baseline", a component's file in
        // "Component local files", a board's in "Board local files" - and never as a bare name.
        // ###########################################################################################
        public static string DefaultFolderFor(FileRowKind kind) => kind switch
        {
            FileRowKind.ComponentPicture => "Component images",
            FileRowKind.ScopeBaseline => DraftFolderLayout.ScopeBaselineFolderName,
            FileRowKind.ComponentFile => "Component local files",
            _ => "Board local files",
        };

        // What stops a newly picked file being filed where its row says - see CheckNewFile.
        public enum NewFileProblem
        {
            None,
            NoFolder,
            FolderNotWritable
        }

        // ###########################################################################################
        // *** A NEWLY PICKED FILE MUST BE FILED IN A FOLDER (owner report, 2026-09-25). *** A new
        // row's "File location" starts EMPTY, and saved like that the file sat at the top of the
        // board beside its schematics. "Save to draft" refuses it, while the row is still on screen.
        // A folder that cannot hold a file of the board (IsWritableFolder) is refused too - the
        // drop-down offers none, but a row keeps its own folder in its list.
        // ###########################################################################################
        public static NewFileProblem CheckNewFile(string? fileLocation)
        {
            if (string.IsNullOrWhiteSpace(fileLocation))
            {
                return NewFileProblem.NoFolder;
            }

            return ContributionFileLocations.IsWritableFolder(fileLocation)
                ? NewFileProblem.None
                : NewFileProblem.FolderNotWritable;
        }
    }
}
