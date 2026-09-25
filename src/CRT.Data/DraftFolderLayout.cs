using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A DRAFT FOLDER CONTAINS, now that a draft IS a board folder
    // (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
    //
    // A draft folder is deliberately INDISTINGUISHABLE from a published one:
    //
    //     <DraftsRoot>/Commodore/C64/250407/
    //         Data C64 250407.xlsx      the board workbook, published schema
    //         Data C64 250407.json      the sidecar: highlights + KiCad calibrations
    //         Sheet1of5.png             images at their real relative paths
    //         KiCad data/
    //         Scope baseline/           created up front for a new system - see ScopeBaselineFolderName
    //         .crt-draft.json           LOCAL ONLY - see DraftMarkerFileName
    //
    // *** THE POINT IS THAT EDITING IN THE APP AND EDITING IN EXCEL ARE INTERCHANGEABLE. *** The
    // previous layout kept row deltas in draft.json and copied attachment bytes into a "Files/"
    // subfolder, which meant the folder on disk looked nothing like the thing it was a draft OF.
    // A contributor could not open the workbook, because there wasn't one.
    //
    // WHAT REPLACED THE DELTAS. draft.json recorded Added/Modified/Deleted per row AT EDIT TIME.
    // That cannot survive someone editing the workbook with the application closed, so the answer
    // is now DERIVED by comparing the draft workbook against the published one - see
    // BoardDataDiffer. This class owns only the file layout; it computes no differences.
    //
    // PURE PATH LOGIC. Every method here is string work on paths, with no filesystem access at
    // all - the copying itself is DraftSeeder's job. That split keeps the naming rules unit
    // testable without a temp folder, and means a caller can ask "where would this go" without
    // creating anything.
    // ###########################################################################################
    public static class DraftFolderLayout
    {
        // ###########################################################################################
        // The local-only marker that makes a folder a DRAFT rather than a copy of a board.
        //
        // *** IT MUST NEVER BE SUBMITTED OR PUBLISHED. *** It records which published revision the
        // draft was taken from and, for a draft-only system, its registration - information that is
        // meaningless to anyone else and that would be noise in the published tree.
        //
        // Two things keep it out, and the belt-and-braces is deliberate because a leak here is
        // silent: SubmissionManifestBuilder.CollectReferencedFiles derives the file list from the
        // ROWS rather than from a directory walk, so a file no row names cannot be picked up; and
        // IsDraftOnlyFile below names it explicitly for anything that does walk the folder.
        //
        // The leading dot sorts it out of the way in a file listing and reads as machine-owned,
        // matching the convention of .gitignore and friends rather than looking like board data.
        // ###########################################################################################
        public const string DraftMarkerFileName = ".crt-draft.json";

        // ###########################################################################################
        // The board's folder of oscilloscope screenshots from a known working board, directly
        // inside the board folder exactly as the published tree has it
        // ("Commodore/C128/310378/Scope baseline/"). Assets/Wiki/Scope-baseline-folder.md names it
        // too, so this string is the contract with contributors - the same standing as
        // AppConfig.KiCadDataFolderName.
        //
        // Only the NAME is a convention: nothing in the app discovers files in it by name. A row
        // references each baseline image by its own path, so a file placed here is reached the same
        // way as any other referenced file (see GetReferencedFilePath).
        // ###########################################################################################
        public const string ScopeBaselineFolderName = "Scope baseline";

        // ###########################################################################################
        // The draft folder for one system, given the drafts root and the system's ExcelDataFile
        // identity ("Commodore/C64/250407/Data C64 250407.xlsx").
        //
        // The manufacturer/hardware/board segments only, dropping the file name - so the draft tree
        // mirrors Data/'s own layout and is easy to find by hand. Splits on '/', matching
        // ExcelDataFile's own convention (the sync manifest's separator) rather than
        // Path.DirectorySeparatorChar, which would not match on Windows.
        //
        // Returns empty rather than a bogus path when either input is unusable; callers must treat
        // that as "no draft folder available" exactly as DraftManager.GetSystemFolder already does.
        // ###########################################################################################
        public static string GetSystemFolder(string draftsRoot, string excelDataFile)
        {
            if (string.IsNullOrWhiteSpace(draftsRoot) || string.IsNullOrWhiteSpace(excelDataFile))
            {
                return string.Empty;
            }

            string[] segments = excelDataFile.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (segments.Length < 2)
            {
                return string.Empty;
            }

            string[] folderSegments = new string[segments.Length - 1];
            Array.Copy(segments, folderSegments, folderSegments.Length);

            return Path.Combine(draftsRoot, Path.Combine(folderSegments));
        }

        // ###########################################################################################
        // The draft's own board workbook - the file a contributor opens in Excel.
        //
        // *** THE SAME FILE NAME THE PUBLISHED SYSTEM USES, not a "draft" variant. *** The whole
        // request was that the folder be indistinguishable from a real board, and a workbook called
        // "Data C64 250407 (draft).xlsx" would announce itself as something else - it would also
        // break BoardDataReader's cache key, which is the ExcelDataFile identity.
        // ###########################################################################################
        public static string GetWorkbookPath(string draftsRoot, string excelDataFile)
        {
            string folder = DraftFolderLayout.GetSystemFolder(draftsRoot, excelDataFile);
            if (folder.Length == 0)
            {
                return string.Empty;
            }

            string fileName = DraftFolderLayout.FileNameOf(excelDataFile);

            return fileName.Length == 0 ? string.Empty : Path.Combine(folder, fileName);
        }

        // ###########################################################################################
        // The draft's JSON sidecar, which carries component highlights AND the KiCad calibrations.
        //
        // *** THIS IS WHY BoardDraft's "eleventh section" CAN RETIRE. *** Calibrations used to live
        // only in draft.json because BoardData has no section for them and nothing else in the
        // draft could hold them. They have always had a published home - the sidecar's "KiCad
        // calibration points" root, which BoardComponentHighlightStorage reads and writes and which
        // the SERVER already publishes into. A draft folder carrying the ordinary sidecar therefore
        // stores them in the exact published format, with no draft-specific mechanism at all.
        //
        // Derived through BoardComponentHighlightStorage.GetJsonPath rather than by swapping the
        // extension here, so there is one definition of "the sidecar beside this workbook".
        // ###########################################################################################
        public static string GetSidecarPath(string draftsRoot, string excelDataFile)
        {
            string workbook = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);

            return workbook.Length == 0
                ? string.Empty
                : BoardComponentHighlightStorage.GetJsonPath(workbook);
        }

        // ###########################################################################################
        // The draft's "Scope baseline" folder, beside the workbook - see ScopeBaselineFolderName.
        // Empty when the system folder cannot be resolved, like every other path here.
        // ###########################################################################################
        public static string GetScopeBaselineFolder(string draftsRoot, string excelDataFile)
        {
            string folder = DraftFolderLayout.GetSystemFolder(draftsRoot, excelDataFile);

            return folder.Length == 0
                ? string.Empty
                : Path.Combine(folder, DraftFolderLayout.ScopeBaselineFolderName);
        }

        // ###########################################################################################
        // The local-only marker file's path - see DraftMarkerFileName.
        // ###########################################################################################
        public static string GetMarkerPath(string draftsRoot, string excelDataFile)
        {
            string folder = DraftFolderLayout.GetSystemFolder(draftsRoot, excelDataFile);

            return folder.Length == 0
                ? string.Empty
                : Path.Combine(folder, DraftFolderLayout.DraftMarkerFileName);
        }

        // ###########################################################################################
        // Where one of the board's referenced files lives inside the draft folder.
        //
        // *** THE PATH IS RELATIVE TO THE SYSTEM FOLDER, NOT TO THE DATA ROOT, and that is the one
        // genuinely subtle thing in this class. *** A BoardData row stores a file as a path from
        // the data root ("Commodore/C64/250407/Sheet1.png"), because that is what the published
        // tree needs. Inside a draft folder the system's own three segments are already the folder
        // itself, so the stored path has to have them stripped or the bytes would land at
        // "<draft>/Commodore/C64/250407/Commodore/C64/250407/Sheet1.png".
        //
        // A file OUTSIDE this system's own folder - a manufacturer "Shared files" image, say, which
        // is referenced as "Commodore/Shared files/7805.jpg" - is deliberately NOT given a draft
        // location, and this returns empty for it: SEEDING never copies a published shared file
        // into a draft. Such a file is shared with other boards and is not the draft's to own; it
        // keeps resolving against Data/ the way it always did. Copying one in would fork it, and a
        // later edit to the draft's copy would silently not reach the boards that actually share it.
        //
        // A NEW shared file the contributor ATTACHES is different - it exists nowhere else yet - and
        // lives in the draft under its whole path (DraftFileResolver.BuildDraftFileDestination).
        // ###########################################################################################
        public static string GetReferencedFilePath(string draftsRoot, string excelDataFile, string? relativeFile)
        {
            string trimmed = relativeFile?.Trim() ?? string.Empty;
            if (trimmed.Length == 0)
            {
                return string.Empty;
            }

            string folder = DraftFolderLayout.GetSystemFolder(draftsRoot, excelDataFile);
            if (folder.Length == 0)
            {
                return string.Empty;
            }

            string? withinSystem = DraftFolderLayout.RelativeToSystemFolder(excelDataFile, trimmed);

            return withinSystem is null
                ? string.Empty
                : Path.Combine(folder, withinSystem.Replace('/', Path.DirectorySeparatorChar));
        }

        // ###########################################################################################
        // Strips the system's own "Manufacturer/Hardware/Board/" prefix off a data-root-relative
        // path, or answers null when the path does not sit under this system at all.
        //
        // Case-INSENSITIVE, matching how the rest of the app compares these paths - a workbook
        // hand-edited to say "commodore/C64/..." names the same folder on Windows, and refusing to
        // recognise it would silently treat every one of that board's own files as shared.
        // ###########################################################################################
        public static string? RelativeToSystemFolder(string excelDataFile, string? relativeFile)
        {
            string trimmed = relativeFile?.Trim() ?? string.Empty;
            if (trimmed.Length == 0 || string.IsNullOrWhiteSpace(excelDataFile))
            {
                return null;
            }

            string[] systemSegments = excelDataFile.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (systemSegments.Length < 2)
            {
                return null;
            }

            string[] fileSegments = trimmed
                .Replace('\\', '/')
                .Split('/', StringSplitOptions.RemoveEmptyEntries);

            // The system's folder is every segment of its ExcelDataFile except the file name.
            int prefixLength = systemSegments.Length - 1;
            if (fileSegments.Length <= prefixLength)
            {
                return null;
            }

            for (int i = 0; i < prefixLength; i++)
            {
                if (!string.Equals(fileSegments[i], systemSegments[i], StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }
            }

            return string.Join('/', fileSegments.Skip(prefixLength));
        }

        // ###########################################################################################
        // Whether a file inside a draft folder is LOCAL BOOKKEEPING that must never be submitted,
        // published, or counted as board content.
        //
        // Only the marker today. It is a method rather than a bare comparison so that anything
        // added later (a lock file, an editor backup convention) is excluded everywhere at once
        // rather than in whichever caller happened to be updated.
        // ###########################################################################################
        public static bool IsDraftOnlyFile(string? fileName)
        {
            string name = Path.GetFileName(fileName?.Trim() ?? string.Empty);

            return string.Equals(
                name,
                DraftFolderLayout.DraftMarkerFileName,
                StringComparison.OrdinalIgnoreCase);
        }

        private static string FileNameOf(string excelDataFile)
        {
            string[] segments = excelDataFile.Split('/', StringSplitOptions.RemoveEmptyEntries);

            return segments.Length == 0 ? string.Empty : segments[^1];
        }
    }
}
