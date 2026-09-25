using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHICH FOLDER A SYSTEM'S BOARD DATA IS READ FROM, now that a draft is a real board folder
    // (NewContributeStrategy.md Phase 6 - maintainer request, 2026-09-23).
    //
    // *** THIS REPLACES AN OVERLAY WITH A CHOICE, and that is the whole shape of the change. ***
    //
    // Before: the published workbook was read, and a draft's row deltas were merged on top of it
    // (BoardDraftApplier). A drafted board was therefore never a file - it was a computation, and
    // there was nothing a contributor could open in Excel.
    //
    // Now: a draft IS a board folder, so loading a drafted system means reading THAT folder's
    // workbook instead of the published one. No merge, no deltas, nothing to keep in step. The
    // file on disk is the truth - which is what makes editing in the app and editing in Excel
    // interchangeable, the maintainer's actual requirement.
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
        // revision the drift warning needs and the registration a draft-only system has.
        public DraftMarker? Marker { get; init; }

        // ###########################################################################################
        // True when this system has NO published counterpart - "Add a new system".
        //
        // Read off the MARKER rather than inferred from "the published file is missing", and the
        // difference matters: for a system the main workbook DOES list, a missing file is a real
        // sync failure that must keep failing loudly rather than quietly rendering an empty board.
        // That distinction was already load-bearing in DataManager before this change.
        // ###########################################################################################
        public bool IsNewSystem => this.Marker?.IsNewSystem == true;

        // The published workbook, whether or not it is the one being read. Kept so a caller can
        // compare a draft against what it was drafted from (BoardDataDiffer) without re-deriving
        // the path, and so "view as officially published" has somewhere to point.
        public string PublishedWorkbookPath { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // Resolves that choice.
    // ###########################################################################################
    public static class DraftBoardSource
    {
        // ###########################################################################################
        // Decides which workbook to read for one system.
        //
        // *** THE MARKER IS WHAT MAKES A FOLDER A DRAFT, not the presence of a workbook. *** A
        // folder under Drafts/ holding an .xlsx but no marker is not a draft - most likely
        // something copied there by hand - and reading it as one would silently substitute
        // unknown data for the published board. Requiring the marker means a draft is only ever
        // something this application created.
        //
        // preferPublished is "view boards as officially published" (UserSettings). It is honoured
        // HERE, at the one place the choice is made, rather than by each caller - the same reason
        // DataManager applied that toggle in exactly one place when the overlay still existed.
        // The marker is still returned, so a caller can say "you are viewing the published
        // version" without re-resolving anything.
        // ###########################################################################################
        public static BoardSourceSelection Resolve(
            string dataRoot,
            string draftsRoot,
            string excelDataFile,
            bool preferPublished = false)
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
                    PublishedWorkbookPath = publishedPath,
                    IsDraft = false,
                    Marker = marker,
                };
            }

            if (preferPublished)
            {
                return new BoardSourceSelection
                {
                    // *** A DRAFT-ONLY SYSTEM HAS NO PUBLISHED COPY TO SHOW. *** Pointing at a file
                    // that does not exist is right rather than a bug: officially, this system does
                    // not exist yet, and the load path already knows to render that as a blank
                    // board rather than as a failure (see IsNewSystem's own note).
                    WorkbookPath = publishedPath,
                    PublishedWorkbookPath = publishedPath,
                    IsDraft = false,
                    Marker = marker,
                };
            }

            return new BoardSourceSelection
            {
                WorkbookPath = draftWorkbook,
                PublishedWorkbookPath = publishedPath,
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
                ? DraftFolderLayout.GetSystemFolder(draftsRoot, excelDataFile)
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
        // Whether a system has a local draft at all - the question the Drafts tab asks of every
        // system, and the cheapest one to answer.
        //
        // Deliberately does NOT read the workbook or compare anything: listing drafts must not cost
        // a board parse per system. What has CHANGED in each is BoardDataDiffer's job, asked only
        // for the system actually being shown.
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
        // The published workbook's absolute path for a system.
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
