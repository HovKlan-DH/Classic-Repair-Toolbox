using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A DRAFT'S BOARD AS EDITABLE SHEETS - the model behind the Drafts tab's "Edit in table
    // format" (maintainer request, 2026-09-24). Avalonia-free, so the review application can put
    // the same table in front of a reviewer later; see BoardTableRow.cs for why it sits here.
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
    // carries them across untouched. The revision date and the preamble names ride along the same
    // way.
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

        public bool HasUnsavedChanges { get; private set; }

        // Undo and redo across every sheet, back to the last save - see BoardTableHistory.
        public BoardTableHistory History { get; }

        public int TotalChangeCount => this.Sheets.Sum(sheet => sheet.ChangeCount);

        // ###########################################################################################
        // Builds the table for one draft. `published` is null when there is nothing published to
        // compare against.
        // ###########################################################################################
        public static BoardTableDocument Create(BoardData? published, BoardData draft)
        {
            ArgumentNullException.ThrowIfNull(draft);

            var document = new BoardTableDocument(published is not null);

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

            return document;
        }

        public BoardTableSheet? FindSheet(string sheetName) =>
            this.Sheets.FirstOrDefault(sheet => string.Equals(sheet.Name, sheetName, StringComparison.Ordinal));

        // ###########################################################################################
        // The board to save: `current` - the board as it is on disk right now, handed over by
        // DraftWorkbookStore.Edit - with every sheet replaced by this table's rows.
        //
        // *** ALL NINE SHEETS ARE REPLACED, EDITED OR NOT. *** That is only safe because the save
        // path refuses when the workbook has changed since this table was read (see
        // DraftTableSession) - otherwise a sheet the contributor never touched here would silently
        // overwrite an edit made to it in Excel meanwhile.
        // ###########################################################################################
        public BoardData ApplyTo(BoardData current)
        {
            ArgumentNullException.ThrowIfNull(current);

            BoardData result = current;
            foreach (BoardTableSheet sheet in this.Sheets)
            {
                result = BoardWorkbookSchema.WithRows(result, sheet.Definition, sheet.BuildRows());
            }

            return result;
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
