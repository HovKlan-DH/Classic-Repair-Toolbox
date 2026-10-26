using System;
using System.Collections.Generic;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What a seeding attempt did, so the caller can report it rather than guess. (The Excel-era
    // counts of files copied, missing and shared went with the copying, 2026-10-24: a 4.0 draft
    // names the published board's files and copies none.)
    // ###########################################################################################
    public sealed class DraftSeedResult
    {
        public bool Created { get; init; }

        public string BoardFolder { get; init; } = string.Empty;

        // The draft's board file ("<draft>/board.json").
        public string WorkbookPath { get; init; } = string.Empty;

        // Why nothing happened, when Created is false. Empty on success.
        public string Reason { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // CREATES A DRAFT FOLDER AS A FULL COPY OF THE PUBLISHED BOARD
    // (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
    //
    // *** SEEDING IS A COMPLETE COPY, AND THAT IS LOAD-BEARING RATHER THAN GENEROUS. ***
    //
    // The draft folder is what the contributor edits and what is compared against the published
    // board to work out what they changed (BoardDataDiffer). That comparison reads an ABSENT row as
    // a deletion. If a draft were seeded empty, or partially, then every row the contributor had not
    // yet touched would read as deleted - the report would be noise, and a submission built from it
    // would ask the server to delete most of the board.
    //
    // So: complete copy, deletion means deletion. The project owner chose this over the cheaper
    // alternative explicitly ("I do believe it is the most transparent thing to do").
    //
    // *** IN CRT 4.0 THE COPY IS THE BOARD FILE (2026-10-24). *** The published board.json becomes
    // the draft's board.json - every row, highlight, calibration and rename, and the list of the
    // board's files by name and checksum. No file's bytes are copied: they stay in the tree's store,
    // where the draft's list finds them (DraftFileResolver), until the contributor replaces one. The
    // same board file is kept as the draft's base (DraftBaseStore), what a later update is merged
    // against (DraftRebase).
    // ###########################################################################################
    public static class DraftSeeder
    {
        // ###########################################################################################
        // *** SEEDS FROM THE PUBLISHED BOARD FILE AS IT IS ON DISK RIGHT NOW (owner report,
        // 2026-10-02). *** A draft is a copy of the FILE, never of a cached board read earlier.
        // ###########################################################################################
        public static DraftSeedResult SeedFromPublishedFile(string draftsRoot, string dataRoot, string excelDataFile)
        {
            string publishedPath = DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile);
            BoardDocument? published = BoardFiles.ReadDocument(publishedPath);

            if (published is null)
            {
                return new DraftSeedResult { Reason = "The published board could not be read to copy it into a draft." };
            }

            return DraftSeeder.Seed(draftsRoot, excelDataFile, published, published, publishedPath);
        }

        // ###########################################################################################
        // Seeds the draft folder for a board that IS published, with `publishedBoard` as its rows -
        // the board as read from the published file, or one built in memory by a test. The files it
        // cites are the published board file's (so they resolve), when there is one.
        //
        // REFUSES TO OVERWRITE AN EXISTING DRAFT. Re-seeding would silently discard whatever the
        // contributor had already changed, which is the single most destructive thing this class
        // could do - so an existing marker means "leave it alone" and the caller is told why.
        // ###########################################################################################
        public static DraftSeedResult SeedFromPublished(
            string draftsRoot,
            string dataRoot,
            string excelDataFile,
            BoardData publishedBoard)
        {
            ArgumentNullException.ThrowIfNull(publishedBoard);

            string publishedPath = DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile);
            BoardDocument? published = BoardFiles.ReadDocument(publishedPath);
            BoardContent? publishedContent = published is null ? null : BoardDocumentMapping.ToContent(published);

            BoardDocument document = BoardDocumentMapping.ToDocument(
                new BoardContent(publishedBoard, publishedContent?.Calibrations ?? [], publishedContent?.Renames ?? []),
                name => name,
                published?.Files ?? []);

            // What a later CRT wrote into the published file and this one does not know comes along.
            BoardDocumentUnknowns.CarryOver(published, document);

            return DraftSeeder.Seed(draftsRoot, excelDataFile, document, published ?? document, publishedPath);
        }

        // The draft's board file, its base and its marker - and nothing over an existing draft.
        // `publishedPath` is the published board file, whose content the drift check compares with.
        private static DraftSeedResult Seed(string draftsRoot, string excelDataFile, BoardDocument draft, BoardDocument published, string publishedPath)
        {
            string folder = DraftFolderLayout.GetBoardFolder(draftsRoot, excelDataFile);
            if (folder.Length == 0)
            {
                return new DraftSeedResult { Reason = "The draft location could not be resolved." };
            }

            string markerPath = DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile);
            if (File.Exists(markerPath))
            {
                return new DraftSeedResult
                {
                    BoardFolder = folder,
                    Reason = "A draft of this board already exists.",
                };
            }

            string boardFile = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);

            try
            {
                Directory.CreateDirectory(folder);

                BoardFiles.Write(boardFile, draft);

                // ###########################################################################################
                // *** AND WHAT IT WAS COPIED FROM (code review, 2026-10-13). *** The base - the
                // published board as copied - is what lets this draft be brought up to date when
                // somebody else's work is published later (DraftRebase), instead of sending that work
                // back as removals. A base that cannot be written costs only that: the draft is still a
                // draft, warned about drift the old way.
                // ###########################################################################################
                DraftSeeder.TryWriteBase(draftsRoot, excelDataFile, published);

                DraftMarkerStore.Save(markerPath, new DraftMarker
                {
                    BoardKey = excelDataFile,
                    BaseRevision = published.RevisionDate ?? string.Empty,
                    BaseContentHash = DraftBaseStore.PublishedContentHash(publishedPath),
                    CreatedUtc = DateTimeOffset.UtcNow.ToString("o"),
                });

                CrtLog.Info($"Seeded draft for [{excelDataFile}] at [{folder}] - {draft.Files.Count} file(s) named");

                return new DraftSeedResult
                {
                    Created = true,
                    BoardFolder = folder,
                    WorkbookPath = boardFile,
                };
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not seed draft for [{excelDataFile}] - [{ex.Message}]");

                return new DraftSeedResult
                {
                    BoardFolder = folder,
                    Reason = $"The draft could not be created: {ex.Message}",
                };
            }
        }

        // ###########################################################################################
        // Creates the folder for a board that exists ONLY as a draft - "Add a new board".
        //
        // There is nothing to copy, so this writes an EMPTY board file plus the marker carrying the
        // registration. Every section empty, but the IDENTITY is carried through (fixed 2026-09-24):
        // the board's caption is the registration's names - the only place they exist for a board
        // that has never been published.
        //
        // BaseRevision stays EMPTY: this board has no published counterpart, so there is no revision
        // it is based on and inventing one would make a later drift check compare against something
        // that never existed.
        // ###########################################################################################
        public static DraftSeedResult CreateNewBoard(
            string draftsRoot,
            NewBoardRegistration registration)
        {
            ArgumentNullException.ThrowIfNull(registration);

            string excelDataFile = registration.ExcelDataFile ?? string.Empty;

            string folder = DraftFolderLayout.GetBoardFolder(draftsRoot, excelDataFile);
            if (folder.Length == 0)
            {
                return new DraftSeedResult { Reason = "The draft location could not be resolved." };
            }

            string markerPath = DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile);
            if (File.Exists(markerPath))
            {
                return new DraftSeedResult
                {
                    BoardFolder = folder,
                    Reason = "A draft of this board already exists.",
                };
            }

            string boardFile = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);

            try
            {
                Directory.CreateDirectory(folder);

                // *** AND AN EMPTY "Scope baseline" FOLDER (owner request, 2026-09-24). *** A new
                // board has no baseline images yet; the folder is there to file the first one in.
                // An empty folder is inert everywhere else - a draft's files are what its board file
                // lists (DraftFiles), never what a folder walk finds.
                Directory.CreateDirectory(DraftFolderLayout.GetScopeBaselineFolder(draftsRoot, excelDataFile));

                BoardFiles.Write(boardFile, new BoardDocument
                {
                    HardwareName = registration.HardwareName ?? string.Empty,
                    BoardName = registration.BoardName ?? string.Empty,
                });

                DraftMarkerStore.Save(markerPath, new DraftMarker
                {
                    BoardKey = excelDataFile,
                    BaseRevision = string.Empty,
                    NewBoard = registration,
                    CreatedUtc = DateTimeOffset.UtcNow.ToString("o"),
                });

                CrtLog.Info($"Created draft-only board [{excelDataFile}] at [{folder}]");

                return new DraftSeedResult
                {
                    Created = true,
                    BoardFolder = folder,
                    WorkbookPath = boardFile,
                };
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not create draft-only board [{excelDataFile}] - [{ex.Message}]");

                return new DraftSeedResult
                {
                    BoardFolder = folder,
                    Reason = $"The draft could not be created: {ex.Message}",
                };
            }
        }

        private static void TryWriteBase(string draftsRoot, string excelDataFile, BoardDocument published)
        {
            try
            {
                DraftBaseStore.SaveDocument(draftsRoot, excelDataFile, published);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
            {
                CrtLog.Warning($"Could not keep the base of the draft for [{excelDataFile}] - [{ex.Message}]");
            }
        }
    }
}
