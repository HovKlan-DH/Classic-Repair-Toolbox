using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHICH FOLDER A BOARD'S BOARD DATA IS READ FROM, now that a draft is a real board folder
    // (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
    //
    // *** THIS REPLACES AN OVERLAY WITH A CHOICE, and that is the whole shape of the change. ***
    //
    // Before: the published workbook was read, and a draft's row deltas were merged on top of it
    // (BoardDraftApplier). A drafted board was therefore never a file - it was a computation, and
    // there was nothing a contributor could open in Excel.
    //
    // Now: a draft IS a board folder, so loading a drafted board means reading THAT folder's
    // workbook instead of the published one. No merge, no deltas, nothing to keep in step. The
    // file on disk is the truth - which is what makes editing in the app and editing in Excel
    // interchangeable, the project owner's actual requirement.
    //
    // PURE apart from File.Exists, so the rule is unit tested rather than trusted. It resolves
    // paths and answers "which one"; it opens nothing and parses nothing.
    // ###########################################################################################
    public sealed class BoardSourceSelection
    {
        // The workbook to actually read. Empty only when neither a draft nor a published copy
        // could be resolved at all, which callers already treat as "cannot load this board".
        public string WorkbookPath { get; init; } = string.Empty;

        // True when WorkbookPath is inside the drafts tree rather than the published tree.
        public bool IsDraft { get; init; }

        // The draft's marker, when one was found. Null for a published load. Carries the base
        // revision the drift warning needs and the registration a draft-only board has.
        public DraftMarker? Marker { get; init; }

        // ###########################################################################################
        // True when this board has NO published counterpart - "Add a new board".
        //
        // Read off the MARKER rather than inferred from "the published file is missing", and the
        // difference matters: for a board the main workbook DOES list, a missing file is a real
        // sync failure that must keep failing loudly rather than quietly rendering an empty board.
        // That distinction was already load-bearing in DataManager before this change.
        // ###########################################################################################
        public bool IsNewBoard => this.Marker?.IsNewBoard == true;

        // The published workbook, whether or not it is the one being read. Kept so a caller can
        // compare a draft against what it was drafted from (BoardDataDiffer) without re-deriving
        // the path. EMPTY for a new board's draft, which was drafted from nothing - see
        // DraftBoardSource.ComparisonBaselineOf.
        public string PublishedWorkbookPath { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // Resolves that choice.
    // ###########################################################################################
    public static class DraftBoardSource
    {
        // ###########################################################################################
        // Decides which workbook to read for one board.
        //
        // *** THE MARKER IS WHAT MAKES A FOLDER A DRAFT, not the presence of a workbook. *** A
        // folder under Drafts/ holding an .xlsx but no marker is not a draft HERE. A board folder a
        // contributor put there by hand is given its marker when the board list is loaded
        // (DraftFolderImport, owner request 2026-09-27) - deliberately, and at that one point - so
        // this rule still decides everything and there is no second notion of "draft" to keep in
        // step.
        //
        // preferPublished is "view boards as officially published" (UserSettings). It is honoured
        // HERE, at the one place the choice is made, rather than by each caller - the same reason
        // DataManager applied that toggle in exactly one place when the overlay still existed.
        // The marker is still returned, so a caller can say "you are viewing the published
        // version" without re-resolving anything.
        // ###########################################################################################
        //
        // listedAsPublished: whether the application lists this board as PUBLISHED (the main
        // workbook lists it - HardwareBoardEntry.IsPublished), when the caller knows. It decides
        // what the draft is compared against - see ComparisonBaselineOf.
        public static BoardSourceSelection Resolve(
            string dataRoot,
            string draftsRoot,
            string excelDataFile,
            bool preferPublished = false,
            bool? listedAsPublished = null)
        {
            string publishedPath = DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile);

            string markerPath = DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile);
            DraftMarker? marker = DraftMarkerStore.Load(markerPath);

            if (marker is null)
            {
                return new BoardSourceSelection
                {
                    WorkbookPath = publishedPath,
                    PublishedWorkbookPath = publishedPath,
                    IsDraft = false,
                };
            }

            string draftWorkbook = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);

            // What this draft is compared against - nothing, for a new board.
            string baseline = DraftBoardSource.ComparisonBaselineOf(dataRoot, excelDataFile, marker, listedAsPublished);

            // ###########################################################################################
            // A MARKER WITH NO WORKBOOK BESIDE IT FALLS BACK TO THE PUBLISHED COPY.
            //
            // Reachable by a crash between creating the folder and writing the workbook, or by
            // someone deleting the .xlsx by hand. Falling back shows the published board, which is
            // both true and recoverable; refusing to load would leave the contributor with a board
            // they cannot open and no way to discard the draft from inside the application.
            //
            // The marker is still reported, so the draft is still listed and can be discarded.
            // ###########################################################################################
            if (draftWorkbook.Length == 0 || !File.Exists(draftWorkbook))
            {
                return new BoardSourceSelection
                {
                    WorkbookPath = publishedPath,
                    PublishedWorkbookPath = baseline,
                    IsDraft = false,
                    Marker = marker,
                };
            }

            if (preferPublished)
            {
                return new BoardSourceSelection
                {
                    // *** A DRAFT-ONLY BOARD HAS NO PUBLISHED COPY TO SHOW. *** Pointing at a file
                    // that does not exist is right rather than a bug: officially, this board does
                    // not exist yet, and the load path already knows to render that as a blank
                    // board rather than as a failure (see IsNewBoard's own note).
                    WorkbookPath = publishedPath,
                    PublishedWorkbookPath = baseline,
                    IsDraft = false,
                    Marker = marker,
                };
            }

            return new BoardSourceSelection
            {
                WorkbookPath = draftWorkbook,
                PublishedWorkbookPath = baseline,
                IsDraft = true,
                Marker = marker,
            };
        }

        // ###########################################################################################
        // The draft folder whose FILES belong to the board being SHOWN - or empty when the board
        // on screen is not the draft.
        //
        // *** "VIEW AS OFFICIALLY PUBLISHED" HAS TO SWITCH THE FILES, NOT JUST THE WORKBOOK. ***
        // Resolve above already reads the published workbook when preferPublished is set, but the
        // images, local files and KiCad data a board names were still looked up draft-first
        // through DraftFileResolver, handed the draft folder unconditionally. So with the toggle
        // on, a contributor checking "what everyone else sees" was shown their own replaced
        // schematic image and their own datasheet - unpublished bytes under a label promising the
        // opposite. Every surface that resolves a board's files asks HERE for the folder to
        // search, so the toggle is honoured in the one place that decides it, exactly as it is
        // for the workbook.
        //
        // Empty also for a folder with no marker (not a draft - see Resolve) and for a marker with
        // no workbook beside it (Resolve falls back to the published board, so its files must
        // follow).
        // ###########################################################################################
        public static string ViewedDraftFolder(
            string dataRoot,
            string draftsRoot,
            string excelDataFile,
            bool preferPublished)
        {
            BoardSourceSelection source = DraftBoardSource.Resolve(
                dataRoot, draftsRoot, excelDataFile, preferPublished);

            return source.IsDraft
                ? DraftFolderLayout.GetBoardFolder(draftsRoot, excelDataFile)
                : string.Empty;
        }

        // ###########################################################################################
        // The workbook whose SIDECAR holds the calibrations and highlights being SHOWN - the read
        // twin of ResolveWritablePath below.
        //
        // The two differ only under "view as officially published": an edit still goes to the
        // draft (it is the only place a contributor's change may land), while the screen reads
        // the published sidecar, because that is what the toggle promises. Reading through
        // ResolveWritablePath instead drew the draft's calibration over a board labelled as the
        // published one.
        // ###########################################################################################
        public static string ResolveViewedPath(
            string dataRoot,
            string draftsRoot,
            string excelDataFile,
            bool preferPublished)
        {
            BoardSourceSelection source = DraftBoardSource.Resolve(
                dataRoot, draftsRoot, excelDataFile, preferPublished);

            return source.WorkbookPath.Length > 0
                ? source.WorkbookPath
                : DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile);
        }

        // ###########################################################################################
        // Whether a board has a local draft at all - the question the Drafts tab asks of every
        // board, and the cheapest one to answer.
        //
        // Deliberately does NOT read the workbook or compare anything: listing drafts must not cost
        // a board parse per board. What has CHANGED in each is BoardDataDiffer's job, asked only
        // for the board actually being shown.
        // ###########################################################################################
        public static bool HasDraft(string draftsRoot, string excelDataFile)
        {
            string markerPath = DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile);

            return markerPath.Length > 0 && File.Exists(markerPath);
        }

        // ###########################################################################################
        // The workbook whose SIDECAR a drafted edit should be written to.
        //
        // *** THIS IS WHY BoardDraft's "eleventh section" COULD RETIRE. *** Component highlights
        // and KiCad calibrations live in the JSON beside a board workbook, and a draft folder now
        // carries an ordinary sidecar - so a drafted calibration is written by exactly the same
        // BoardComponentHighlightStorage call a published one is, just pointed at the draft's own
        // workbook path. There is no draft-specific calibration mechanism left at all.
        //
        // Falls back to the published workbook when there is no draft, which is what makes the
        // callers simple: they ask for a path and write to it, without branching on whether a
        // draft exists.
        // ###########################################################################################
        public static string ResolveWritablePath(string dataRoot, string draftsRoot, string excelDataFile)
        {
            string draftWorkbook = DraftFolderLayout.GetWorkbookPath(draftsRoot, excelDataFile);

            if (draftWorkbook.Length > 0
                && DraftBoardSource.HasDraft(draftsRoot, excelDataFile)
                && File.Exists(draftWorkbook))
            {
                return draftWorkbook;
            }

            return DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile);
        }

        // ###########################################################################################
        // Every KiCad calibration a draft carries, for sending with a submission.
        //
        // *** THIS REPLACES KiCadCalibrationDraftWriter.CollectForSubmission. *** That method read
        // the draft's "eleventh section" and had to skip Deleted tombstones - a removal recorded as
        // a row rather than an absence, so reading every row would have sent back the very
        // calibration the contributor removed and publishing would have put it straight back.
        //
        // None of that applies to a sidecar: a calibration that has been removed is simply not in
        // the file. Absence means absence.
        //
        // Calibrations are keyed per SCHEMATIC, so the board's own schematic list is what decides
        // which to look for - a sidecar can carry an entry for a schematic the board no longer has,
        // and sending that would publish a calibration for a page that does not exist.
        //
        // Ordered by schematic name so two submissions of an unchanged draft produce the same
        // manifest; the manifest is hashed and diffed, and an unstable order reports a change
        // nobody made.
        // ###########################################################################################
        public static IReadOnlyList<KiCadCalibrationEntry> CollectCalibrations(
            string workbookPath,
            BoardData? board)
        {
            if (string.IsNullOrWhiteSpace(workbookPath) || board is null)
            {
                return [];
            }

            var entries = new List<KiCadCalibrationEntry>();

            foreach (BoardSchematicEntry schematic in board.Schematics)
            {
                string schematicName = schematic.SchematicName?.Trim() ?? string.Empty;
                if (schematicName.Length == 0)
                {
                    continue;
                }

                if (!BoardComponentHighlightStorage.TryLoadKiCadCalibration(
                        workbookPath,
                        schematicName,
                        out string cadName,
                        out double offsetX,
                        out double offsetY,
                        out double scaleX,
                        out double scaleY,
                        out bool mirrorX,
                        out bool mirrorY))
                {
                    continue;
                }

                entries.Add(new KiCadCalibrationEntry
                {
                    SchematicName = schematicName,
                    CadName = cadName,
                    OffsetX = offsetX,
                    OffsetY = offsetY,
                    ScaleX = scaleX,
                    ScaleY = scaleY,
                    MirrorX = mirrorX,
                    MirrorY = mirrorY,
                });
            }

            return entries
                .OrderBy(entry => entry.SchematicName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ###########################################################################################
        // The published workbook a DRAFT is compared against - or EMPTY for a new board, which by
        // definition has nothing published to compare with (owner report, 2026-09-27).
        //
        // *** NOT "whatever file sits at the published path". *** For a new board that file can
        // only be the contributor's OWN copy: a board registered the old way, in a
        // "_UserContribution" workbook, lives in Data/ at exactly this path, and its copy put into
        // Drafts/ (DraftFolderImport) was compared against itself - every row unchanged, "New
        // board, nothing added yet", and Submit disabled. A new board counts every row as an
        // addition, which is what its "New board, N rows so far" wording already says, and what
        // the table editor already did (it opens a new board with no published board).
        //
        // Retirement does not come through here: it compares against the board the application
        // LISTS (DraftStatusReader.ResolveForBoard), the right question once a board is published.
        //
        // *** THE LISTING DECIDES WHEN THE CALLER KNOWS IT (code review, 2026-09-27). *** "Is there a
        // published board to compare with" is HardwareBoardEntry.IsPublished - the rule retirement
        // uses - and the marker only stands in for it. Two cases the marker gets wrong:
        //   - a NEW board's draft whose board has since been published and listed: compared with
        //     nothing, every row read as an addition, though the published board holds them;
        //   - a legacy "_UserContribution" board drafted through "Save to draft", which seeds an
        //     ORDINARY marker: compared with the contributor's own copy in Data/, so nothing ever
        //     counted and Submit stayed off. The server decides "new board" from its own tree, so
        //     such a board is received as the new board it is.
        // With no listing to ask (null), the marker decides as before.
        // ###########################################################################################
        public static string ComparisonBaselineOf(
            string dataRoot,
            string excelDataFile,
            DraftMarker? marker,
            bool? listedAsPublished = null)
        {
            bool hasPublishedBoard = listedAsPublished ?? marker?.IsNewBoard != true;

            return hasPublishedBoard
                ? DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile)
                : string.Empty;
        }

        // ###########################################################################################
        // The published workbook's absolute path for a board.
        //
        // Splits on '/' because ExcelDataFile uses the sync manifest's separator rather than the
        // platform's - the same conversion every other consumer of this identity makes.
        // ###########################################################################################
        public static string PublishedPathOf(string dataRoot, string excelDataFile)
        {
            if (string.IsNullOrWhiteSpace(dataRoot) || string.IsNullOrWhiteSpace(excelDataFile))
            {
                return string.Empty;
            }

            return Path.Combine(
                dataRoot,
                excelDataFile.Replace('/', Path.DirectorySeparatorChar));
        }
    }
}
