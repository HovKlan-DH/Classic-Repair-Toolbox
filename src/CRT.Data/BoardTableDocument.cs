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

        // False for a system with no published copy ("Add a new system"). Nothing is coloured
        // then - see BoardTableSheet's header.
        public bool HasBaseline { get; }

        // What a changed cell's tooltip calls the value it replaced - see Create.
        public const string DefaultBaselineLabel = "Published value";

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
        // default names the published board, which is what every draft compares against. The
        // Maintainer tab passes its own for a NEW SYSTEM: there it compares the submission
        // with ITSELF as it arrived (see TabMaintainer.Table.cs), so `published` is not published
        // at all and "Published value: (empty)" named a board that does not exist - reported by the
        // project owner, 2026-09-26. HasBaseline cannot answer this: it is true in both cases.
        // ###########################################################################################
        public static BoardTableDocument Create(BoardData? published, BoardData draft, string? baselineLabel = null)
        {
            ArgumentNullException.ThrowIfNull(draft);

            var document = new BoardTableDocument(published is not null)
            {
                BaselineLabel = string.IsNullOrWhiteSpace(baselineLabel) ? DefaultBaselineLabel : baselineLabel,
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

            return document;
        }

        public BoardTableSheet? FindSheet(string sheetName) =>
            this.Sheets.FirstOrDefault(sheet => string.Equals(sheet.Name, sheetName, StringComparison.Ordinal));

        // ###########################################################################################
        // WHICH SHEET TABS SHOW (owner request, 2026-09-26: with "Show changes only" ticked, "it
        // should also hide the sheets/tabs having no changes"). Every sheet without the filter; with
        // it, the sheets with a row the filter shows (BoardTableSheet.HasChangeRows) - and always
        // `current`, so undoing the last change of the sheet being worked on does not pull its tab
        // away. A table with nothing to show anywhere keeps every tab: hiding them all but one would
        // look broken, and there is nothing to narrow down to.
        // ###########################################################################################
        public IReadOnlyList<BoardTableSheet> SheetsShown(bool onlyChanges, BoardTableSheet? current)
        {
            if (!onlyChanges || !this.Sheets.Any(sheet => sheet.HasChangeRows))
                return this.Sheets;

            return this.Sheets
                .Where(sheet => sheet.HasChangeRows || ReferenceEquals(sheet, current))
                .ToList();
        }

        // ###########################################################################################
        // The sheet to show: `wanted` (by name - the sheet on screen before, or the one last looked
        // at in another table) while its tab shows, else the first sheet whose tab shows.
        // ###########################################################################################
        public BoardTableSheet SheetToShow(string? wanted, bool onlyChanges)
        {
            IReadOnlyList<BoardTableSheet> shown = this.SheetsShown(onlyChanges, current: null);

            return shown.FirstOrDefault(sheet => string.Equals(sheet.Name, wanted, StringComparison.Ordinal))
                ?? shown.FirstOrDefault()
                ?? this.Sheets[0];
        }

        // ###########################################################################################
        // The board to save: `current` - the board as it is on disk right now, handed over by
        // DraftWorkbookStore.Edit - with every sheet replaced by this table's rows.
        //
        // *** ALL NINE SHEETS ARE REPLACED, EDITED OR NOT. *** That is only safe because the save
        // path refuses when the workbook has changed since this table was read (see
        // DraftTableSession) - otherwise a sheet the contributor never touched here would silently
        // overwrite an edit made to it in Excel meanwhile.
        //
        // *** AND THE HIGHLIGHTS OF A COMPONENT DELETED HERE GO WITH IT (2026-09-25). *** A component
        // whose row was deleted in the table, and whose label no row still carries, is gone from the
        // board, and so are the rectangles naming it. A component RENAMED here keeps its highlights
        // (its row object is still live) - telling the two apart is why the opened rows are kept.
        // ###########################################################################################
        public BoardData ApplyTo(BoardData current)
        {
            ArgumentNullException.ThrowIfNull(current);

            BoardData result = current;
            foreach (BoardTableSheet sheet in this.Sheets)
            {
                result = BoardWorkbookSchema.WithRows(result, sheet.Definition, sheet.BuildRows());
            }

            IReadOnlySet<string> deleted = this.LabelsOfDeletedComponents(result);

            if (deleted.Count == 0)
            {
                return result;
            }

            return result.WithComponentHighlights(result.ComponentHighlights
                .Where(highlight => !deleted.Contains((highlight.BoardLabel ?? string.Empty).Trim()))
                .ToList());
        }

        // The labels of components deleted in this table: opened with a row that is no longer
        // live, and carried by no row the save writes.
        private IReadOnlySet<string> LabelsOfDeletedComponents(BoardData saved)
        {
            var remaining = new HashSet<string>(
                saved.Components.Select(component => (component.BoardLabel ?? string.Empty).Trim()),
                StringComparer.OrdinalIgnoreCase);

            BoardTableSheet? components = this.FindSheet(BoardWorkbookSchema.SheetComponents);
            var live = new HashSet<BoardTableRow>(components?.Rows.Where(row => !row.IsDeleted) ?? []);

            return this.thisOpenedComponents
                .Where(opened => !live.Contains(opened.Row) && !remaining.Contains(opened.Label))
                .Select(opened => opened.Label)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // The rows a deleted component leaves on the other component sheets - deleted with it (see
        // BoardTableSheet.DeleteRow). `label` and `region` are the deleted Components row's.
        //
        //   - No row still has the label: every image, local file and link row for it goes.
        //   - Another row has the label but none has this REGION (a regional variant deleted, its
        //     twin kept): only the image rows for that label AND region go. Files and links are not
        //     per region, and still belong to the twin; so do images with a blank region.
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

            foreach (string sheetName in new[]
            {
                BoardWorkbookSchema.SheetComponentImages,
                BoardWorkbookSchema.SheetComponentLocalFiles,
                BoardWorkbookSchema.SheetComponentLinks
            })
            {
                if (labelRemains && sheetName != BoardWorkbookSchema.SheetComponentImages)
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

        public void MarkSaved()
        {
            this.History.MarkSaved();
            this.SetUnsavedChanges(false);
        }

        internal void MarkDirty() => this.SetUnsavedChanges(true);

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
