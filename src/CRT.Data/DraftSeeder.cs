using System;
using System.Collections.Generic;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What a seeding attempt did, so the caller can report it rather than guess.
    //
    // FilesCopied and FilesMissing are separate counts on purpose: a published board referencing a
    // file that is not on disk is a real condition (a partial sync, a hand-deleted image), and it
    // must not be reported as a seeding failure. The draft is still usable; that one row just
    // points at nothing, exactly as it did before the draft existed.
    // ###########################################################################################
    public sealed class DraftSeedResult
    {
        public bool Created { get; init; }

        public string SystemFolder { get; init; } = string.Empty;

        public string WorkbookPath { get; init; } = string.Empty;

        public int FilesCopied { get; init; }

        // Referenced files the published tree does not actually hold. Reported, never fatal.
        public IReadOnlyList<string> FilesMissing { get; init; } = [];

        // Files left where they are because they belong to another board - a manufacturer's
        // "Shared files" image, say. See DraftFolderLayout.GetReferencedFilePath.
        public IReadOnlyList<string> FilesShared { get; init; } = [];

        // Why nothing happened, when Created is false. Empty on success.
        public string Reason { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // CREATES A DRAFT FOLDER AS A FULL COPY OF THE PUBLISHED BOARD
    // (NewContributeStrategy.md Phase 6 - maintainer request, 2026-09-23).
    //
    // *** SEEDING IS A COMPLETE COPY, AND THAT IS LOAD-BEARING RATHER THAN GENEROUS. ***
    //
    // The draft folder is what the contributor edits, in the app or in Excel, and what is compared
    // against the published board to work out what they changed (BoardDataDiffer). That comparison
    // reads an ABSENT row as a deletion. If a draft were seeded empty, or partially, then every
    // row the contributor had not yet touched would read as deleted - the report would be noise,
    // and a submission built from it would ask the server to delete most of the board.
    //
    // So: complete copy, deletion means deletion. The maintainer chose this over the cheaper
    // alternative explicitly ("I do believe it is the most transparent thing to do").
    //
    // WHAT IS COPIED: the workbook, its JSON sidecar, and every file the board's own ROWS
    // reference. Derived from the rows via SubmissionManifestBuilder.CollectReferencedFiles, never
    // from a directory walk - the same rule and the same reason as a submission. A walk would
    // sweep up editor backups, thumbnail caches and whatever else happens to sit in a board folder.
    //
    // WHAT IS NOT COPIED: a file belonging to a DIFFERENT board. Manufacturer "Shared files"
    // images are referenced by many boards, and forking one into a draft would mean a later edit
    // silently failed to reach the boards that share it. Those keep resolving against Data/.
    // ###########################################################################################
    public static class DraftSeeder
    {
        // ###########################################################################################
        // Seeds the draft folder for a system that IS published, by copying it.
        //
        // REFUSES TO OVERWRITE AN EXISTING DRAFT. Re-seeding would silently discard whatever the
        // contributor had already changed, which is the single most destructive thing this class
        // could do - so an existing marker means "leave it alone" and the caller is told why.
        //
        // publishedBoard is the board data as loaded from the published workbook; it is passed in
        // rather than read here so this class does not depend on BoardDataReader's async cache and
        // stays testable against a board built in memory.
        // ###########################################################################################
        public static DraftSeedResult SeedFromPublished(
            string draftsRoot,
            string dataRoot,
            string excelDataFile,
            BoardData publishedBoard)
        {
            ArgumentNullException.ThrowIfNull(publishedBoard);

            string folder = DraftFolderLayout.GetSystemFolder(draftsRoot, excelDataFile);
            if (folder.Length == 0)
            {
                return new DraftSeedResult { Reason = "The draft location could not be resolved." };
            }

            string markerPath = DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile);
            if (File.Exists(markerPath))
            {
                return new DraftSeedResult
                {
                    SystemFolder = folder,
                    Reason = "A draft of this system already exists.",
                };
            }

            string workbookPath = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);

            try
            {
                Directory.CreateDirectory(folder);

                // *** THE WORKBOOK IS WRITTEN, NOT FILE-COPIED. *** Copying the published .xlsx
                // byte for byte would carry across whatever else that file holds - stale columns, a
                // macro, a hand-added sheet - and the draft would then not round-trip through this
                // application's own schema. Writing it through BoardWorkbookWriter guarantees the
                // draft is exactly what CRT can read back, which is the promise being made to
                // someone about to edit it in Excel.
                BoardWorkbookWriter.Write(workbookPath, publishedBoard);

                DraftSeeder.CopySidecar(dataRoot, excelDataFile, workbookPath);

                DraftSeeder.CopyReferencedFiles(
                    draftsRoot,
                    dataRoot,
                    excelDataFile,
                    publishedBoard,
                    out int copied,
                    out List<string> missing,
                    out List<string> shared);

                DraftMarkerStore.Save(markerPath, new DraftMarker
                {
                    SystemKey = excelDataFile,
                    BaseRevision = publishedBoard.RevisionDate ?? string.Empty,
                    CreatedUtc = DateTimeOffset.UtcNow.ToString("o"),
                });

                CrtLog.Info(
                    $"Seeded draft for [{excelDataFile}] at [{folder}] - " +
                    $"{copied} file(s) copied, {missing.Count} missing, {shared.Count} shared");

                return new DraftSeedResult
                {
                    Created = true,
                    SystemFolder = folder,
                    WorkbookPath = workbookPath,
                    FilesCopied = copied,
                    FilesMissing = missing,
                    FilesShared = shared,
                };
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not seed draft for [{excelDataFile}] - [{ex.Message}]");

                return new DraftSeedResult
                {
                    SystemFolder = folder,
                    Reason = $"The draft could not be created: {ex.Message}",
                };
            }
        }

        // ###########################################################################################
        // Creates the folder for a system that exists ONLY as a draft - "Add a new system".
        //
        // There is nothing to copy, so this writes an EMPTY board workbook (headers and sheet names
        // only) plus the marker carrying the registration.
        //
        // *** AN EMPTY WORKBOOK IS WRITTEN HERE, AND PHASE 2 EXPLICITLY REJECTED THAT. *** The old
        // reasoning was sound at the time: an .xlsx sitting ALONGSIDE draft.json would have been a
        // second mechanism for one job, with two places to look for the same rows. That is not this.
        // The workbook REPLACES draft.json as the single source, so there is exactly one mechanism -
        // and a new system with no workbook could not be opened in Excel at all, which is the whole
        // point of the change.
        //
        // BaseRevision stays EMPTY: this system has no published counterpart, so there is no
        // revision it is based on and inventing one would make a later drift check compare against
        // something that never existed.
        // ###########################################################################################
        public static DraftSeedResult CreateNewSystem(
            string draftsRoot,
            NewSystemRegistration registration)
        {
            ArgumentNullException.ThrowIfNull(registration);

            string excelDataFile = registration.ExcelDataFile ?? string.Empty;

            string folder = DraftFolderLayout.GetSystemFolder(draftsRoot, excelDataFile);
            if (folder.Length == 0)
            {
                return new DraftSeedResult { Reason = "The draft location could not be resolved." };
            }

            string markerPath = DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile);
            if (File.Exists(markerPath))
            {
                return new DraftSeedResult
                {
                    SystemFolder = folder,
                    Reason = "A draft of this system already exists.",
                };
            }

            string workbookPath = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);

            try
            {
                Directory.CreateDirectory(folder);

                // ###########################################################################################
                // *** AND AN EMPTY "Scope baseline" FOLDER BESIDE THE WORKBOOK (maintainer request,
                // 2026-09-24). ***
                //
                // A new system has no baseline images yet, so there was no such folder - and a
                // contributor referencing their first capture had to create it by hand, with the
                // exact name published boards use, before they could pick it. Now it is simply there
                // to choose. Only for a NEW system: a seeded draft gets the folder from the published
                // board's own files when it has any.
                //
                // An empty folder is inert everywhere else - submissions are built from the rows,
                // not by walking the folder (SubmissionManifestBuilder.CollectReferencedFiles).
                // ###########################################################################################
                Directory.CreateDirectory(DraftFolderLayout.GetScopeBaselineFolder(draftsRoot, excelDataFile));

                // ###########################################################################################
                // Every section empty, but the IDENTITY is carried through (fixed 2026-09-24).
                //
                // BoardWorkbookWriter writes each sheet with its header row and preamble, which is
                // what makes the file hand-editable from the moment it is created - and the
                // preamble's first two lines are "# Hardware:" and "# Board:". Passing a bare
                // `new BoardData()` left both blank, so a brand-new system opened in Excel with no
                // caption at all while every published board had one. Reported with a screenshot.
                //
                // The registration is the only place those names exist for a system that has never
                // been published, which is exactly why it is the thing being written here.
                // ###########################################################################################
                BoardWorkbookWriter.Write(workbookPath, new BoardData
                {
                    HardwareName = registration.HardwareName ?? string.Empty,
                    BoardName = registration.BoardName ?? string.Empty,
                });

                DraftMarkerStore.Save(markerPath, new DraftMarker
                {
                    SystemKey = excelDataFile,
                    BaseRevision = string.Empty,
                    NewSystem = registration,
                    CreatedUtc = DateTimeOffset.UtcNow.ToString("o"),
                });

                CrtLog.Info($"Created draft-only system [{excelDataFile}] at [{folder}]");

                return new DraftSeedResult
                {
                    Created = true,
                    SystemFolder = folder,
                    WorkbookPath = workbookPath,
                };
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not create draft-only system [{excelDataFile}] - [{ex.Message}]");

                return new DraftSeedResult
                {
                    SystemFolder = folder,
                    Reason = $"The draft could not be created: {ex.Message}",
                };
            }
        }

        // ###########################################################################################
        // Copies the published sidecar beside the draft workbook, when there is one.
        //
        // Byte-copied rather than rewritten, unlike the workbook: the sidecar holds roots this
        // application does not model (highlights AND "KiCad calibration points", and potentially
        // others a future version adds). BoardComponentHighlightStorage's own writes are careful to
        // preserve unknown roots for exactly that reason, and rewriting it here from a partial
        // model would discard whatever it did not know about.
        //
        // A missing sidecar is ordinary - plenty of boards have no highlights - so it is silent.
        // ###########################################################################################
        private static void CopySidecar(string dataRoot, string excelDataFile, string draftWorkbookPath)
        {
            if (string.IsNullOrWhiteSpace(dataRoot))
            {
                return;
            }

            string publishedWorkbook = Path.Combine(
                dataRoot,
                excelDataFile.Replace('/', Path.DirectorySeparatorChar));

            string publishedSidecar = BoardComponentHighlightStorage.GetJsonPath(publishedWorkbook);
            if (!File.Exists(publishedSidecar))
            {
                return;
            }

            File.Copy(
                publishedSidecar,
                BoardComponentHighlightStorage.GetJsonPath(draftWorkbookPath),
                overwrite: true);
        }

        // ###########################################################################################
        // Copies every file the board's rows reference into the draft folder.
        //
        // Three outcomes per file, all of them normal:
        //   - copied: it belongs to this system and the published tree holds it;
        //   - shared: it belongs to another board (a manufacturer "Shared files" image), so it is
        //     left alone and keeps resolving against Data/ - see DraftFolderLayout;
        //   - missing: the row names a file the published tree does not have. Reported, not fatal;
        //     the draft is still perfectly usable and that row was already broken.
        // ###########################################################################################
        private static void CopyReferencedFiles(
            string draftsRoot,
            string dataRoot,
            string excelDataFile,
            BoardData board,
            out int copied,
            out List<string> missing,
            out List<string> shared)
        {
            copied = 0;
            missing = [];
            shared = [];

            if (string.IsNullOrWhiteSpace(dataRoot))
            {
                return;
            }

            foreach (string relative in SubmissionManifestBuilder.CollectReferencedFiles(board))
            {
                string destination = DraftFolderLayout.GetReferencedFilePath(
                    draftsRoot,
                    excelDataFile,
                    relative);

                if (destination.Length == 0)
                {
                    shared.Add(relative);
                    continue;
                }

                string source = Path.Combine(
                    dataRoot,
                    relative.Replace('/', Path.DirectorySeparatorChar));

                if (!File.Exists(source))
                {
                    missing.Add(relative);
                    continue;
                }

                string? destinationFolder = Path.GetDirectoryName(destination);
                if (!string.IsNullOrWhiteSpace(destinationFolder))
                {
                    Directory.CreateDirectory(destinationFolder);
                }

                File.Copy(source, destination, overwrite: true);
                copied++;
            }
        }
    }
}
