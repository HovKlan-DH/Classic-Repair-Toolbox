using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A DRAFT'S BOARD AS EDITABLE SHEETS - the model behind the Drafts tab's "Edit in table
    // format" (owner request, 2026-09-24). Avalonia-free, so the Maintainer tab can put
    // the same table in front of a maintainer later; see BoardTableRow.cs for why it sits here.
    //
    // *** IT SHOWS THE SCHEMA'S SHEETS, NOT THE RAW WORKBOOK. *** One sheet per
    // BoardWorkbookSchema.AllSheets entry, one column per ColumnOrder entry - which is exactly what
    // survives a save. A draft is saved through DraftWorkbookStore.Edit, which regenerates the
    // workbook from BoardData, so an extra column or a note typed into the file by hand is dropped
    // by ANY save, in the app or here. Showing raw cells would promise to keep things the next save
    // throws away.
    //
    // WHAT IS NOT HERE: the component highlights and KiCad calibrations. They live in the JSON
    // beside the workbook rather than in any sheet, are edited by the label editor, and ApplyTo
    // carries them across untouched - EXCEPT the highlights of a component deleted in the table,
    // which go with it (owner request, 2026-09-25; see ApplyTo). The revision date and the
    // preamble names ride along untouched.
    // ###########################################################################################
    public sealed class BoardTableDocument
    {
        private BoardTableDocument(bool hasBaseline)
        {
            this.HasBaseline = hasBaseline;
            this.History = new BoardTableHistory(this);
        }

        // Raised whenever anything a UI summarises may have changed: a sheet's counts after a
        // refresh, rows added or removed, or the unsaved flag flipping.
        public event EventHandler? Changed;

        public IReadOnlyList<BoardTableSheet> Sheets { get; private set; } = [];

        // False for a board with no published copy ("Add a new board"). Nothing is coloured
        // then - see BoardTableSheet's header.
        public bool HasBaseline { get; }

        // What a changed cell's tooltip calls the value it replaced - see Create.
        public const string DefaultBaselineLabel = "Published value";

        // ###########################################################################################
        // *** A TABLE COMPARED WITH A DATA SOURCE NAMES THAT SOURCE (owner request, 2026-10-05). ***
        // "Published value" read as the data everybody has, so a value that BETA already held - newer
        // than the stable source - looked like a mistake until it was worked out where it came from.
        // The Maintainer tab's queue is always compared with BETA (the server reads its BETA tree);
        // the Drafts tab with the data downloaded, from whichever source the Configuration tab
        // picks. The two names are the sources' own (AppConfig.GetOnlineSourceLabel).
        // ###########################################################################################
        public const string BetaSourceBaselineLabel = "BETA source value";

        public const string StableSourceBaselineLabel = "Stable source value";

        public static string SourceBaselineLabel(bool betaSource) =>
            betaSource ? BoardTableDocument.BetaSourceBaselineLabel : BoardTableDocument.StableSourceBaselineLabel;

        public string BaselineLabel { get; private init; } = DefaultBaselineLabel;

        public bool HasUnsavedChanges { get; private set; }

        // Undo and redo across every sheet, back to the last save - see BoardTableHistory.
        public BoardTableHistory History { get; }

        public int TotalChangeCount => this.Sheets.Sum(sheet => sheet.ChangeCount);

        // Every component row the table opened with, and its label then - so a save can tell a
        // component DELETED here (its row object gone) from one RENAMED here (the same row object,
        // a new label), which the rows alone cannot. See ApplyTo.
        private IReadOnlyList<(BoardTableRow Row, string Label)> thisOpenedComponents = [];

        // How many highlights each label had when the table opened, for the message a component
        // delete shows. Case-insensitive, as labels are.
        private Dictionary<string, int> thisHighlightCounts = new(StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // Builds the table for one draft. `published` is null when there is nothing published to
        // compare against.
        //
        // `baselineLabel` is what a changed cell's tooltip CALLS the value it replaced, and it is
        // the caller's to say because only the caller knows what it handed over as `published`. The
        // default names the published board; the Drafts tab and the Maintainer tab's queue name the
        // data source instead (SourceBaselineLabel). The
        // Maintainer tab passes its own for a NEW BOARD: there it compares the submission
        // with ITSELF as it arrived (see TabMaintainer.Table.cs), so `published` is not published
        // at all and "Published value: (empty)" named a board that does not exist - reported by the
        // project owner, 2026-09-26. HasBaseline cannot answer this: it is true in both cases.
        // ###########################################################################################
        //
        // `files` says whether a file a row names is there (BoardDataChecks) - the draft's own
        // folder and the downloaded data for the Drafts tab (DiskFileLookup), nothing for the
        // Maintainer tab, whose files the server checked when the submission arrived.
        public static BoardTableDocument Create(
            BoardData? published,
            BoardData draft,
            string? baselineLabel = null,
            IBoardFileLookup? files = null)
        {
            ArgumentNullException.ThrowIfNull(draft);

            var document = new BoardTableDocument(published is not null)
            {
                BaselineLabel = string.IsNullOrWhiteSpace(baselineLabel) ? DefaultBaselineLabel : baselineLabel,
                thisFiles = files,
                thisHighlights = draft.ComponentHighlights.ToList(),
                thisCreating = true,
            };

            var sheets = new List<BoardTableSheet>();
            foreach (BoardWorkbookSchema.SheetDefinition definition in BoardWorkbookSchema.AllSheets)
            {
                var sheet = new BoardTableSheet(document, definition, published, draft);
                sheet.Changed += (_, _) => document.Changed?.Invoke(document, EventArgs.Empty);
                sheets.Add(sheet);
            }

            document.Sheets = sheets;

            foreach (BoardTableSheet sheet in sheets)
            {
                sheet.Refresh();
            }

            if (document.FindSheet(BoardWorkbookSchema.SheetComponents) is BoardTableSheet components)
            {
                document.thisOpenedComponents = components.Rows
                    .Where(row => !row.IsDeleted)
                    .Select(row => (row, components.CellText(row, BoardWorkbookSchema.ColBoardLabel)))
                    .Where(opened => opened.Item2.Length > 0)
                    .ToList();
            }

            document.thisHighlightCounts = draft.ComponentHighlights
                .Where(highlight => !string.IsNullOrWhiteSpace(highlight.BoardLabel))
                .GroupBy(highlight => highlight.BoardLabel.Trim(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

            // Once, with every sheet built - not after each of the refreshes above.
            document.thisCreating = false;
            document.RefreshProblems();

            return document;
        }

        public BoardTableSheet? FindSheet(string sheetName) =>
            this.Sheets.FirstOrDefault(sheet => string.Equals(sheet.Name, sheetName, StringComparison.Ordinal));

        // ###########################################################################################
        // WHICH SHEET TABS SHOW (owner requests, 2026-09-26: with "Show changes only" ticked, "it
        // should also hide the sheets/tabs having no changes" - and 2026-10-02, the colour key's
        // counts as the filter, BoardTableRowFilter). Every sheet with nothing picked; otherwise the
        // sheets with a row of a picked kind (BoardTableSheet.HasRowsShownBy) - and always
        // `current`, so undoing the last change of the sheet being worked on does not pull its tab
        // away. A table with nothing to show anywhere keeps every tab: hiding them all but one would
        // look broken, and there is nothing to narrow down to.
        //
        // *** AND THE SEARCH BOX'S (2026-10-02, BoardTableSearch). *** A sheet's tab shows when it
        // has a row both the pills and the search show; with neither in use, every tab.
        //
        // *** A SEARCH FINDING NOTHING IN ANY SHEET KEEPS ONLY `current` (owner request, 2026-10-03:
        // "it should work in all sheets, only showing sheets with filtered data"). *** Every tab
        // stayed until then, each sheet empty, which read as a search of the one sheet on screen;
        // the editor says why the sheet is empty (BoardTableSearch.NothingFoundLine). With no
        // `current`, no sheet at all. Picked pills alone keep the rule above - every tab.
        // ###########################################################################################
        public IReadOnlyList<BoardTableSheet> SheetsShown(BoardTableRowKinds kinds, BoardTableSheet? current) =>
            this.SheetsShown(kinds, BoardTableSearch.None, current);

        public IReadOnlyList<BoardTableSheet> SheetsShown(BoardTableRowKinds kinds, BoardTableSearch search, BoardTableSheet? current)
        {
            ArgumentNullException.ThrowIfNull(search);

            if (!this.HasRowsShownBy(kinds, search))
            {
                if (!search.IsActive)
                    return this.Sheets;

                return current is not null && this.Sheets.Contains(current) ? [current] : [];
            }

            return this.Sheets
                .Where(sheet => sheet.HasRowsShownBy(kinds, search) || ReferenceEquals(sheet, current))
                .ToList();
        }

        // Whether any sheet has a row of a picked kind - false with nothing picked.
        public bool HasRowsShownBy(BoardTableRowKinds kinds) =>
            this.HasRowsShownBy(kinds, BoardTableSearch.None);

        // The same, for the pills and the search together - false with neither in use.
        public bool HasRowsShownBy(BoardTableRowKinds kinds, BoardTableSearch search) =>
            this.Sheets.Any(sheet => sheet.HasRowsShownBy(kinds, search));

        // ###########################################################################################
        // The sheet to show: `wanted` (by name - the sheet on screen before, or the one last looked
        // at in another table) while its tab shows, else the first sheet whose tab shows. When a
        // search finds nothing anywhere no tab shows of itself, and the table stays on `wanted`.
        // ###########################################################################################
        public BoardTableSheet SheetToShow(string? wanted, BoardTableRowKinds kinds) =>
            this.SheetToShow(wanted, kinds, BoardTableSearch.None);

        public BoardTableSheet SheetToShow(string? wanted, BoardTableRowKinds kinds, BoardTableSearch search)
        {
            IReadOnlyList<BoardTableSheet> shown = this.SheetsShown(kinds, search, current: null);

            return shown.FirstOrDefault(sheet => string.Equals(sheet.Name, wanted, StringComparison.Ordinal))
                ?? shown.FirstOrDefault()
                ?? (wanted is null ? null : this.FindSheet(wanted))
                ?? this.Sheets[0];
        }

        // ###########################################################################################
        // *** THE COLOUR KEY COUNTS THE WHOLE DRAFT (owner decision, 2026-10-02: "Added, Modified
        // and Deleted should work per system like Error and Warning"). *** Every sheet added
        // together, for all five pills, on whichever sheet is on screen. They counted that sheet
        // alone, which read "0 Errors" on a sheet without any while the draft's row said "8 errors";
        // a sheet's own numbers stay on the sheets, for its tab's change count.
        // ###########################################################################################
        public int AddedCount => this.Sheets.Sum(sheet => sheet.AddedCount);

        public int ModifiedCount => this.Sheets.Sum(sheet => sheet.ModifiedCount);

        public int DeletedCount => this.Sheets.Sum(sheet => sheet.DeletedCount);

        public int ErrorRowCount => this.Sheets.Sum(sheet => sheet.ErrorRowCount);

        public int WarningRowCount => this.Sheets.Sum(sheet => sheet.WarningRowCount);

        // The colour key's count for ONE kind - the matching count above.
        public int CountOf(BoardTableRowKinds kind) => kind switch
        {
            BoardTableRowKinds.Added => this.AddedCount,
            BoardTableRowKinds.Modified => this.ModifiedCount,
            BoardTableRowKinds.Deleted => this.DeletedCount,
            BoardTableRowKinds.Errors => this.ErrorRowCount,
            BoardTableRowKinds.Warnings => this.WarningRowCount,
            _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, "One kind at a time.")
        };

        // ###########################################################################################
        // Whether clicking ONE kind's pill does anything while `picked` is the filter (owner
        // request, 2026-10-03): a pill counting nothing cannot be picked - it could only show an
        // empty table - but a picked one can always be put back, also once its count has dropped
        // to 0 (the last error fixed), or it could never be turned off.
        // ###########################################################################################
        public bool CanToggle(BoardTableRowKinds picked, BoardTableRowKinds kind) =>
            picked.HasFlag(kind) || this.CountOf(kind) > 0;

        // ###########################################################################################
        // *** THE CHECKS (owner request, 2026-10-02: "integrate the existing validation check into
        // the "Draft" system table view"). *** BoardDataChecks over the board as a save would write
        // it, every problem placed on its cell - the row the entry was built from, in the problem's
        // column (the row's first cell when the problem names none). A problem in no sheet (a
        // highlight) is kept apart, in ProblemsOutsideSheets.
        //
        // Run after every refresh of any sheet (BoardTableSheet.Refresh calls OnSheetRefreshed),
        // since the rules look across sheets. Cheap after the first run: the file answers are kept
        // by the lookup, and a cell is only told when its own problems changed.
        // ###########################################################################################
        private IBoardFileLookup? thisFiles;
        private IReadOnlyList<ComponentHighlightEntry> thisHighlights = [];
        private bool thisCreating;

        public IReadOnlyList<BoardDataProblem> Problems { get; private set; } = [];

        public IReadOnlyList<BoardDataProblem> ProblemsOutsideSheets { get; private set; } = [];

        public int ErrorCount => this.Problems.Count(problem => problem.Level == BoardProblemLevel.Error);

        public int WarningCount => this.Problems.Count(problem => problem.Level == BoardProblemLevel.Warning);

        internal void OnSheetRefreshed()
        {
            if (this.thisCreating)
            {
                return;
            }

            if (this.thisProblemsDeferred > 0)
            {
                this.thisProblemsPending = true;
                return;
            }

            this.RefreshProblems();
        }

        // ###########################################################################################
        // *** ONE CHECK FOR A CHANGE THAT REFRESHES SEVERAL SHEETS (code review, 2026-10-04). ***
        // Deleting 8 components refreshed Components and then up to three more sheets per component
        // - up to 25 whole-board checks on the UI thread, each re-mapping every row of all nine
        // sheets. Inside DeferProblems the sheets refresh as before and the checks run ONCE, when the
        // outermost scope ends, followed by Changed so whatever repaints reads the new problems.
        // Create already did this for its own refreshes (thisCreating).
        // ###########################################################################################
        private int thisProblemsDeferred;
        private bool thisProblemsPending;

        internal IDisposable DeferProblems()
        {
            this.thisProblemsDeferred++;
            return new ProblemsDeferral(this);
        }

        private void EndProblemsDeferral()
        {
            this.thisProblemsDeferred--;

            if (this.thisProblemsDeferred > 0 || !this.thisProblemsPending)
            {
                return;
            }

            this.thisProblemsPending = false;
            this.RefreshProblems();
            this.Changed?.Invoke(this, EventArgs.Empty);
        }

        private sealed class ProblemsDeferral(BoardTableDocument document) : IDisposable
        {
            private bool thisEnded;

            public void Dispose()
            {
                if (this.thisEnded)
                {
                    return;
                }

                this.thisEnded = true;
                document.EndProblemsDeferral();
            }
        }

        // How many times the whole board has been checked - for tests (DeferProblems).
        internal int ProblemChecksRun { get; private set; }

        // How many rows have been mapped by the schema afresh rather than remembered - for tests
        // (BoardTableRow.MappedForChecks).
        internal int RowsMapped { get; set; }

        public void RefreshProblems()
        {
            this.ProblemChecksRun++;

            // Each sheet's rows as a save writes them - live, not blank, and mapped (an incomplete
            // important signal is dropped on save, so it is not checked either) - with the row each
            // entry came from, so a problem's index finds its row again. Each row's mapping is
            // remembered until one of its cells changes (code review, 2026-10-04: every edit mapped
            // every row of all nine sheets again), so an edit maps just its own row.
            var rowsBySheet = new Dictionary<string, List<BoardTableRow>>(StringComparer.Ordinal);
            var entriesBySheet = new Dictionary<string, List<object>>(StringComparer.Ordinal);
            var leftOut = new List<(BoardTableCell Cell, BoardDataProblem Problem)>();

            foreach (BoardTableSheet sheet in this.Sheets)
            {
                var rows = new List<BoardTableRow>();
                var entries = new List<object>();

                foreach (BoardTableRow row in sheet.Rows)
                {
                    if (row.IsDeleted || row.IsBlank)
                    {
                        continue;
                    }

                    IReadOnlyList<object> mapped = row.MappedForChecks();
                    if (mapped.Count == 1)
                    {
                        rows.Add(row);
                        entries.Add(mapped[0]);
                    }
                    else
                    {
                        leftOut.Add(BoardTableDocument.LeftOutWhenSaving(row));
                    }
                }

                rowsBySheet[sheet.Name] = rows;
                entriesBySheet[sheet.Name] = entries;
            }

            List<ComponentEntry> components = BoardTableDocument.EntriesOf<ComponentEntry>(entriesBySheet, BoardWorkbookSchema.SheetComponents);

            var checkRows = new BoardCheckRows(
                BoardTableDocument.EntriesOf<BoardSchematicEntry>(entriesBySheet, BoardWorkbookSchema.SheetBoardSchematics),
                components,
                BoardTableDocument.EntriesOf<ComponentImageEntry>(entriesBySheet, BoardWorkbookSchema.SheetComponentImages),
                this.HighlightsAsSaved(components),
                BoardTableDocument.EntriesOf<ComponentLocalFileEntry>(entriesBySheet, BoardWorkbookSchema.SheetComponentLocalFiles),
                BoardTableDocument.EntriesOf<BoardLocalFileEntry>(entriesBySheet, BoardWorkbookSchema.SheetBoardLocalFiles),
                BoardTableDocument.EntriesOf<ComponentLinkEntry>(entriesBySheet, BoardWorkbookSchema.SheetComponentLinks),
                BoardTableDocument.EntriesOf<BoardLinkEntry>(entriesBySheet, BoardWorkbookSchema.SheetBoardLinks),
                BoardTableDocument.EntriesOf<CreditEntry>(entriesBySheet, BoardWorkbookSchema.SheetCredits),
                BoardTableDocument.EntriesOf<KiCadImportantSignalEntry>(entriesBySheet, BoardWorkbookSchema.SheetKiCadImportantSignals),
                BoardTableDocument.EntriesOf<ComponentPinEntry>(entriesBySheet, BoardWorkbookSchema.SheetComponentPins),
                BoardTableDocument.EntriesOf<ComponentReplacementEntry>(entriesBySheet, BoardWorkbookSchema.SheetComponentReplacements));

            IReadOnlyList<BoardDataProblem> problems = BoardDataChecks.Check(checkRows, this.thisFiles, BoardCheckScope.Everything);

            var byCell = new Dictionary<BoardTableCell, List<BoardDataProblem>>();
            var outside = new List<BoardDataProblem>();

            foreach ((BoardTableCell cell, BoardDataProblem problem) in leftOut)
            {
                byCell[cell] = [problem];
            }

            foreach (BoardDataProblem problem in problems)
            {
                if (problem.Sheet is null ||
                    !rowsBySheet.TryGetValue(problem.Sheet, out List<BoardTableRow>? rows) ||
                    problem.Index < 0 ||
                    problem.Index >= rows.Count)
                {
                    outside.Add(problem);
                    continue;
                }

                BoardTableRow row = rows[problem.Index];
                int column = problem.Column is null ? -1 : BoardTableDocument.ColumnIndex(row.Sheet, problem.Column);
                BoardTableCell cell = row.Cells[Math.Max(column, 0)];

                if (!byCell.TryGetValue(cell, out List<BoardDataProblem>? onCell))
                {
                    onCell = [];
                    byCell[cell] = onCell;
                }

                onCell.Add(problem);
            }

            foreach (BoardTableSheet sheet in this.Sheets)
            {
                foreach (BoardTableRow row in sheet.Rows)
                {
                    foreach (BoardTableCell cell in row.Cells)
                    {
                        cell.SetProblems(byCell.TryGetValue(cell, out List<BoardDataProblem>? onCell) ? onCell : []);
                    }
                }

                sheet.CountProblemRows();
            }

            this.Problems = [.. problems, .. leftOut.Select(item => item.Problem)];
            this.ProblemsOutsideSheets = outside;
        }

        // ###########################################################################################
        // *** A ROW THE SAVE LEAVES OUT IS A WARNING (owner request, 2026-10-03: "Should flagged now
        // be treated as warnings?"; cases agreed with the project owner). *** It was a violet
        // "Flagged" row. The checks never see it - they read the rows a save writes, and this is
        // the one kind it does not - so the table tells it itself, on the first empty value it
        // needs (BoardWorkbookSchema.RequiredColumns: only an important signal can be one today).
        // ###########################################################################################
        private static (BoardTableCell Cell, BoardDataProblem Problem) LeftOutWhenSaving(BoardTableRow row)
        {
            List<string> missing = BoardWorkbookSchema.RequiredColumns(row.Sheet.Name)
                .Where(column =>
                {
                    int index = BoardTableDocument.ColumnIndex(row.Sheet, column);
                    return index >= 0 && row.Cells[index].Text.Length == 0;
                })
                .ToList();

            string message = missing.Count == 0
                ? "This row is left out when saving, because a value it needs is missing."
                : $"This row is left out when saving, because {BoardDataChecks.ListOf(missing)} {(missing.Count == 1 ? "is" : "are")} empty.";

            int column = missing.Count == 0 ? 0 : BoardTableDocument.ColumnIndex(row.Sheet, missing[0]);

            var problem = new BoardDataProblem(
                BoardProblemLevel.Warning,
                "row.incomplete",
                row.Cells[0].Text,
                message,
                row.Sheet.Name,
                -1,
                missing.Count == 0 ? null : missing[0]);

            return (row.Cells[Math.Max(column, 0)], problem);
        }

        private static List<T> EntriesOf<T>(Dictionary<string, List<object>> entriesBySheet, string sheet) =>
            entriesBySheet.TryGetValue(sheet, out List<object>? entries) ? entries.Cast<T>().ToList() : [];

        private static int ColumnIndex(BoardTableSheet sheet, string column)
        {
            for (int i = 0; i < sheet.Columns.Count; i++)
            {
                if (string.Equals(sheet.Columns[i], column, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        // ###########################################################################################
        // The highlights a save would keep: the draft's, less those of components deleted in this
        // table (ApplyTo drops them) - so deleting U8 does not also warn about U8's highlights.
        // ###########################################################################################
        private IReadOnlyList<ComponentHighlightEntry> HighlightsAsSaved(IReadOnlyList<ComponentEntry> components)
        {
            BoardTableSheet? sheet = this.FindSheet(BoardWorkbookSchema.SheetComponents);
            var live = new HashSet<BoardTableRow>(sheet?.Rows.Where(row => !row.IsDeleted) ?? []);

            var remaining = new HashSet<string>(
                components.Select(component => (component.BoardLabel ?? string.Empty).Trim()),
                StringComparer.OrdinalIgnoreCase);

            var deleted = this.thisOpenedComponents
                .Where(opened => !live.Contains(opened.Row) && !remaining.Contains(opened.Label))
                .Select(opened => opened.Label)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            return deleted.Count == 0
                ? this.thisHighlights
                : this.thisHighlights
                    .Where(highlight => !deleted.Contains((highlight.BoardLabel ?? string.Empty).Trim()))
                    .ToList();
        }

        // ###########################################################################################
        // The board to save: `current` - the board as it is on disk right now, handed over by
        // DraftWorkbookStore.Edit - with every sheet replaced by this table's rows.
        //
        // *** ALL NINE SHEETS ARE REPLACED, EDITED OR NOT. *** That is only safe because `current`
        // is what the table was built from: the draft's session, which every other change goes
        // through and the table is built again after (BoardTableEditor.DraftSession.cs) - otherwise a
        // sheet the contributor never touched here would silently overwrite an edit made elsewhere.
        //
        // *** AND THE HIGHLIGHTS OF A COMPONENT DELETED HERE GO WITH IT (2026-09-25). *** A component
        // whose row was deleted in the table, and whose label no row still carries, is gone from the
        // board, and so are the rectangles naming it. A component RENAMED here keeps its highlights
        // (its row object is still live) - telling the two apart is why the opened rows are kept.
        // ###########################################################################################
        public BoardData ApplyTo(BoardData current) => this.TakeSaveSnapshot()(current);

        // ###########################################################################################
        // *** THE TABLE'S CHANGE AS AN EDIT OF THE DRAFT HELD IN MEMORY (CRT 4.0 step 1, 2026-10-25;
        // DraftSession.Apply). *** The session's content with every sheet replaced by this table's
        // rows, as a save writes them (TakeSaveSnapshot) - the highlights, calibrations and renames
        // carried along - or null when that changes no row: an empty row inserted, a value typed back
        // to what it was. Null is "nothing to do" to the session, so no step that changes nothing
        // lands in the draft's history. Every sheet needing it is brought up to date first.
        //
        // *** ONLY THE SHEETS CHANGED SINCE THE LAST HAND-OVER (code review, 2026-10-26). *** Every
        // settled cell edit built, put in and compared all nine sheets - their rows three times over -
        // on the UI thread, a lag on the largest boards. A sheet no change touched since MarkSaved is
        // the session's already (the table is built again from it after any other change), so it is
        // left as the session has it, its very rows, and only the changed ones are worked on.
        // ###########################################################################################
        public Func<BoardContent, BoardContent?> TakeSessionEdit()
        {
            this.RefreshPendingSheets();

            List<BoardTableSheet> changed = this.Sheets.Where(this.thisChangedSheets.Contains).ToList();
            Func<BoardData, BoardData> apply = this.TakeSaveSnapshot(changed);
            List<BoardWorkbookSchema.SheetDefinition> compared = changed.Select(sheet => sheet.Definition).ToList();

            return content =>
            {
                ArgumentNullException.ThrowIfNull(content);

                BoardData board = apply(content.Board);

                return BoardTableDocument.HasSameRows(content.Board, board, compared) ? null : content with { Board = board };
            };
        }

        // Every sheet whose rows changed since its last refresh, refreshed - the board checked once.
        public void RefreshPendingSheets()
        {
            using (this.DeferProblems())
            {
                foreach (BoardTableSheet sheet in this.Sheets.Where(sheet => sheet.NeedsRefresh))
                {
                    sheet.Refresh();
                }
            }
        }

        // Whether two boards hold the same rows on `sheets` and the same highlights - a table's
        // change can only drop highlights (a deleted component's), so their number says it.
        private static bool HasSameRows(BoardData before, BoardData after, IEnumerable<BoardWorkbookSchema.SheetDefinition> sheets) =>
            before.ComponentHighlights.Count == after.ComponentHighlights.Count &&
            sheets.All(sheet =>
            {
                IReadOnlyList<IReadOnlyDictionary<string, string>> was = BoardWorkbookSchema.BuildRows(sheet, before);
                IReadOnlyList<IReadOnlyDictionary<string, string>> now = BoardWorkbookSchema.BuildRows(sheet, after);

                return was.Count == now.Count &&
                       was.Zip(now).All(pair =>
                           pair.First.Count == pair.Second.Count &&
                           pair.First.All(cell => pair.Second.TryGetValue(cell.Key, out string? value) && string.Equals(cell.Value, value, StringComparison.Ordinal)));
            });

        // ###########################################################################################
        // *** WHAT A SAVE WRITES, TAKEN NOW - AND APPLIED LATER, ON ANY THREAD (2026-09-30). ***
        // The table's rows are read HERE, on the thread that edits them; the function returned
        // touches nothing of this document, so it can be applied off the UI thread (owner
        // question, 2026-09-30: a save of the largest board takes 0.6-2 seconds). An edit made
        // after this is taken is not in it.
        //
        // ApplyTo is exactly this, applied at once: same rows, same deleted components. `only`: just
        // those sheets, the rest of `current` left as it is (TakeSessionEdit) - every sheet otherwise.
        // ###########################################################################################
        public Func<BoardData, BoardData> TakeSaveSnapshot(IReadOnlyCollection<BoardTableSheet>? only = null)
        {
            List<(BoardWorkbookSchema.SheetDefinition Definition, IReadOnlyList<IReadOnlyDictionary<string, string>> Rows)> sheets = (only ?? this.Sheets)
                .Select(sheet => (sheet.Definition, sheet.BuildRows()))
                .ToList();

            // The components opened with a row that is no longer live - deleted, unless some row the
            // save writes still carries the label (LabelsOfDeletedComponents).
            BoardTableSheet? components = this.FindSheet(BoardWorkbookSchema.SheetComponents);
            var live = new HashSet<BoardTableRow>(components?.Rows.Where(row => !row.IsDeleted) ?? []);

            List<string> gone = this.thisOpenedComponents
                .Where(opened => !live.Contains(opened.Row))
                .Select(opened => opened.Label)
                .ToList();

            return current =>
            {
                ArgumentNullException.ThrowIfNull(current);

                BoardData result = current;
                foreach ((BoardWorkbookSchema.SheetDefinition definition, IReadOnlyList<IReadOnlyDictionary<string, string>> rows) in sheets)
                {
                    result = BoardWorkbookSchema.WithRows(result, definition, rows);
                }

                IReadOnlySet<string> deleted = BoardTableDocument.LabelsOfDeletedComponents(result, gone);

                if (deleted.Count == 0)
                {
                    return result;
                }

                return result.WithComponentHighlights(result.ComponentHighlights
                    .Where(highlight => !deleted.Contains((highlight.BoardLabel ?? string.Empty).Trim()))
                    .ToList());
            };
        }

        // The labels of components deleted in this table: opened with a row that is no longer
        // live (`gone`), and carried by no row the save writes.
        private static IReadOnlySet<string> LabelsOfDeletedComponents(BoardData saved, IReadOnlyList<string> gone)
        {
            var remaining = new HashSet<string>(
                saved.Components.Select(component => (component.BoardLabel ?? string.Empty).Trim()),
                StringComparer.OrdinalIgnoreCase);

            return gone
                .Where(label => !remaining.Contains(label))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // The rows a deleted component leaves on the other component sheets - deleted with it (see
        // BoardTableSheet.DeleteRow). `label` and `region` are the deleted Components row's.
        //
        //   - No row still has the label: every image, local file, link, pin and replacement row for
        //     it goes.
        //   - Another row has the label but none has this REGION (a regional variant deleted, its
        //     twin kept): only the image, pin and replacement rows for that label AND region go.
        //     Files and links are not per region, and still belong to the twin; so do rows with a
        //     blank region.
        //   - Otherwise (a duplicate row deleted): nothing else goes.
        //
        // Labels and regions compare ignoring case, as everywhere else. Null when nothing went.
        // ###########################################################################################
        internal BoardTableDeletedWith? DeleteRowsOfComponent(string label, string region)
        {
            if (string.IsNullOrWhiteSpace(label) || this.FindSheet(BoardWorkbookSchema.SheetComponents) is not BoardTableSheet components)
            {
                return null;
            }

            var others = components.Rows
                .Where(row => !row.IsDeleted)
                .Select(row => (Label: components.CellText(row, BoardWorkbookSchema.ColBoardLabel), Region: components.CellText(row, BoardWorkbookSchema.ColRegion)))
                .Where(other => string.Equals(other.Label, label, StringComparison.OrdinalIgnoreCase))
                .ToList();

            bool labelRemains = others.Count > 0;
            bool regionGone = region.Length > 0 && !others.Any(other => string.Equals(other.Region, region, StringComparison.OrdinalIgnoreCase));

            if (labelRemains && !regionGone)
            {
                return null;
            }

            var counts = new List<BoardTableSheetCount>();

            // Pins and replacement parts (CRT 4.0, 2026-10-25) are filed by region as images are,
            // so a regional variant deleted takes its own along.
            bool PerRegion(string sheetName) =>
                sheetName is BoardWorkbookSchema.SheetComponentImages
                    or BoardWorkbookSchema.SheetComponentPins
                    or BoardWorkbookSchema.SheetComponentReplacements;

            foreach (string sheetName in new[]
            {
                BoardWorkbookSchema.SheetComponentImages,
                BoardWorkbookSchema.SheetComponentLocalFiles,
                BoardWorkbookSchema.SheetComponentLinks,
                BoardWorkbookSchema.SheetComponentPins,
                BoardWorkbookSchema.SheetComponentReplacements
            })
            {
                if (labelRemains && !PerRegion(sheetName))
                {
                    continue;
                }

                if (this.FindSheet(sheetName) is not BoardTableSheet sheet)
                {
                    continue;
                }

                List<BoardTableRow> rows = sheet.Rows
                    .Where(row => !row.IsDeleted)
                    .Where(row => string.Equals(sheet.CellText(row, BoardWorkbookSchema.ColBoardLabel), label, StringComparison.OrdinalIgnoreCase))
                    .Where(row => !labelRemains || string.Equals(sheet.CellText(row, BoardWorkbookSchema.ColRegion), region, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                sheet.RemoveLiveRows(rows);

                if (rows.Count > 0)
                {
                    counts.Add(new BoardTableSheetCount(sheetName, rows.Count));
                }
            }

            int highlights = labelRemains ? 0 : this.thisHighlightCounts.GetValueOrDefault(label.Trim());

            return counts.Count == 0 && highlights == 0
                ? null
                : new BoardTableDeletedWith(label.Trim(), counts, highlights);
        }

        // ###########################################################################################
        // *** FIND & REPLACE (CRT 4.0, 2026-10-25 - the plan's step 0). *** Every occurrence of
        // `find` in the cells of `sheets` - ignoring capitals, as the search box does - becomes
        // `replacement`, as typed. Only rows `shows` lets through, which is the table's own view (a
        // picked pill, a search): what changes is what can be seen. Red rows are never touched.
        //
        // ONE change across every sheet it touches - one undo step - and the board checked once.
        // A replacement that leaves a cell as it was (the same text in other capitals) is not
        // counted.
        // ###########################################################################################
        public BoardTableReplaceResult ReplaceAll(
            string? find,
            string? replacement,
            IReadOnlyList<BoardTableSheet> sheets,
            Func<BoardTableRow, bool> shows)
        {
            ArgumentNullException.ThrowIfNull(sheets);
            ArgumentNullException.ThrowIfNull(shows);

            if (string.IsNullOrEmpty(find))
            {
                return new BoardTableReplaceResult(0, 0, 0);
            }

            int occurrences = 0;
            int cells = 0;
            var touched = new List<BoardTableSheet>();

            this.History.BeginGroup();

            try
            {
                using (this.DeferProblems())
                {
                    foreach (BoardTableSheet sheet in sheets.Where(sheet => this.Sheets.Contains(sheet)).Distinct())
                    {
                        int cellsBefore = cells;

                        foreach (BoardTableRow row in sheet.Rows.Where(row => !row.IsDeleted && shows(row)).ToList())
                        {
                            foreach (BoardTableCell cell in row.Cells)
                            {
                                int count = BoardTableDocument.CountOccurrences(cell.Text, find);

                                if (count > 0 && BoardTableSheet.Write(cell, cell.Text.Replace(find, replacement ?? string.Empty, StringComparison.OrdinalIgnoreCase)))
                                {
                                    occurrences += count;
                                    cells++;
                                }
                            }
                        }

                        if (cells > cellsBefore)
                        {
                            touched.Add(sheet);
                            sheet.Refresh();
                        }
                    }
                }
            }
            finally
            {
                // The step belongs to the first sheet written - where Ctrl+Z puts the cursor. Ended
                // whatever happened, or every later edit would join this one step.
                this.History.EndGroup(touched.Count > 0 ? touched[0] : this.Sheets[0]);
            }

            return new BoardTableReplaceResult(occurrences, cells, touched.Count);
        }

        // How often `find` occurs in `text`, ignoring capitals, not overlapping - as Replace counts.
        private static int CountOccurrences(string text, string find)
        {
            int count = 0;

            for (int at = text.IndexOf(find, StringComparison.OrdinalIgnoreCase); at >= 0; at = text.IndexOf(find, at + find.Length, StringComparison.OrdinalIgnoreCase))
            {
                count++;
            }

            return count;
        }

        public void MarkSaved()
        {
            this.History.MarkSaved();
            this.thisChangedSheets.Clear();
            this.SetUnsavedChanges(false);
        }

        // A sheet changed - every change a sheet makes says so, which is what makes the table hand
        // anything over at all - and which (TakeSessionEdit works on those alone).
        internal void MarkDirty(BoardTableSheet sheet)
        {
            this.NoteSheetChanged(sheet);
            this.SetUnsavedChanges(true);
        }

        // The sheets changed since the last MarkSaved - see TakeSessionEdit.
        internal void NoteSheetChanged(BoardTableSheet sheet) => this.thisChangedSheets.Add(sheet);

        private readonly HashSet<BoardTableSheet> thisChangedSheets = new(ReferenceEqualityComparer.Instance);

        // Also how an undo or redo that lands back on the saved state clears the flag again.
        internal void SetUnsavedChanges(bool value)
        {
            if (this.HasUnsavedChanges == value)
            {
                return;
            }

            this.HasUnsavedChanges = value;
            this.Changed?.Invoke(this, EventArgs.Empty);
        }
    }
}
