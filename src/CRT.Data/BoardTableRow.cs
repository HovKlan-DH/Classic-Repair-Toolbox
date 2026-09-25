using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE CELLS AND ROWS OF THE BOARD TABLE EDITOR (maintainer request, 2026-09-24) - the Drafts
    // tab's "Edit in table format", which shows a draft's workbook as editable sheets with every
    // difference from the published data coloured in.
    //
    // *** WHY THIS LIVES IN CRT.Data AND NOT IN THE APP. *** The maintainer wants the same table
    // in the review application later, so a reviewer can make the same edits. Everything that
    // decides WHAT a cell is - its text, whether it differs, what the published value was - is
    // here, Avalonia-free; the app's BoardTableEditor only paints it. INotifyPropertyChanged is
    // plain .NET, so the grid can bind straight to these without a view-model layer in between
    // (this codebase has none - see CLAUDE.md, "No MVVM").
    //
    // The pairing and the colouring rules are BoardTableSheet's; these types only carry the
    // result. See that file's header before changing what a state means.
    // ###########################################################################################

    // What one cell shows, compared against the published data.
    public enum BoardTableCellState
    {
        Unchanged,

        // The whole row is new - the cell is green.
        Added,

        // This cell differs from the published row it pairs with - orange, on the cell alone.
        Modified,

        // The row exists officially and not in the draft - red, and read-only.
        Deleted,

        // The row is marked "!" - a duplicate of a row above, or incomplete - violet across the
        // whole row (maintainer request, 2026-09-24: the "!" alone was too easy to miss). Not a
        // change: it counts towards none of the other three, and it is coloured with or without a
        // published board, since it is about the data rather than about publishing.
        Flagged
    }

    // What one row is, which drives the one-character marker at its start. Deliberately richer
    // than the cell states: a duplicate and a row that will not survive a save are both simply
    // Flagged (violet), and the marker's tooltip is where the difference is told.
    public enum BoardTableRowState
    {
        Unchanged,
        Added,
        Modified,

        // A ghost: published, absent from the draft. Shown where it used to be.
        Deleted,

        // Nothing typed in it yet. Dropped on save, exactly as the reader skips an empty row.
        Blank,

        // The same natural key as a row ABOVE it. BoardDataDiffer counts only the first row per
        // key, so this one is not a counted change - but it is still saved, so it is flagged.
        Duplicate,

        // A row the schema's own mapper drops on save (an important signal missing one of its two
        // halves). Flagged rather than silently lost.
        Incomplete
    }

    // ###########################################################################################
    // One cell.
    //
    // *** TEXT IS TRIMMED ON THE WAY IN. *** BoardDataReader trims every cell it reads, so a value
    // saved with a trailing space would come back without it - trimming here means the table
    // never shows something a reload would change. It also absorbs the line break Excel puts
    // after a copied cell, which would otherwise ride into the cell on a paste.
    // ###########################################################################################
    public sealed class BoardTableCell : INotifyPropertyChanged
    {
        private string thisText;
        private BoardTableCellState thisState;
        private string thisPublishedText = string.Empty;

        internal BoardTableCell(BoardTableRow row, int columnIndex, string? text)
        {
            this.Row = row;
            this.ColumnIndex = columnIndex;
            this.thisText = BoardTableCell.NormaliseText(text);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public BoardTableRow Row { get; }

        public int ColumnIndex { get; }

        // ###########################################################################################
        // The cell's value. Setting it on a DELETED row is ignored - a ghost shows what was
        // published, and the only way to change it is to restore it first.
        //
        // Setting it does NOT re-pair the sheet. That is BoardTableSheet.Refresh's job, and it is
        // deliberately left to the caller: a grid commits a cell edit from inside its own edit
        // handling, and a refresh can insert or remove ghost rows - changing the grid's rows under
        // it mid-commit is asking for trouble. The sheet raises CellEdited so the UI can schedule
        // the refresh for just after.
        // ###########################################################################################
        public string Text
        {
            get => this.thisText;
            set
            {
                if (this.Row.IsDeleted)
                {
                    return;
                }

                string normalised = BoardTableCell.NormaliseText(value);
                if (string.Equals(normalised, this.thisText, StringComparison.Ordinal))
                {
                    return;
                }

                // The step is recorded BEFORE the change, so undoing it returns to the text as it was.
                this.Row.Sheet.RecordCellEdit(this);

                this.thisText = normalised;
                this.Raise(nameof(this.Text));

                this.Row.Sheet.OnCellEdited(this);
            }
        }

        // Puts text back during an undo or redo: no step recorded, no edit raised - the history
        // refreshes the sheet itself once every row is back.
        internal void RestoreText(string text)
        {
            string normalised = BoardTableCell.NormaliseText(text);
            if (string.Equals(normalised, this.thisText, StringComparison.Ordinal))
            {
                return;
            }

            this.thisText = normalised;
            this.Raise(nameof(this.Text));
        }

        public BoardTableCellState State
        {
            get => this.thisState;
            internal set
            {
                if (this.thisState == value)
                {
                    return;
                }

                this.thisState = value;
                this.Raise(nameof(this.State));
                this.Raise(nameof(this.ToolTip));
            }
        }

        // What the paired published row holds in this column. Empty when the row has no published
        // counterpart. Kept for every paired cell, not just modified ones, so "Revert" always has
        // something to revert to.
        public string PublishedText
        {
            get => this.thisPublishedText;
            internal set
            {
                string normalised = value ?? string.Empty;
                if (string.Equals(normalised, this.thisPublishedText, StringComparison.Ordinal))
                {
                    return;
                }

                this.thisPublishedText = normalised;
                this.Raise(nameof(this.PublishedText));
                this.Raise(nameof(this.ToolTip));
            }
        }

        // The hover text: on a changed cell the published value it replaced, and on a flagged
        // row's cells why the row is flagged - the same words as its "!" marker, so the reason is
        // under the pointer wherever it lands on the row. Null otherwise, so an unchanged cell
        // shows no tooltip at all rather than an empty box.
        public string? ToolTip => this.thisState switch
        {
            BoardTableCellState.Modified =>
                $"Published value: {(this.thisPublishedText.Length == 0 ? "(empty)" : this.thisPublishedText)}",
            BoardTableCellState.Flagged => this.Row.MarkerToolTip,
            _ => null,
        };

        // A flagged row can change WHY it is flagged (duplicate to incomplete) with this cell's own
        // state staying Flagged, so the row tells its cells when their tooltip may have changed.
        internal void RaiseToolTipChanged() => this.Raise(nameof(this.ToolTip));

        public static string NormaliseText(string? text) => (text ?? string.Empty).Trim();

        private void Raise(string propertyName) =>
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    // ###########################################################################################
    // One row: a live row of the draft, or - when IsDeleted - a read-only ghost of a published row
    // the draft no longer has.
    //
    // IsDeleted is fixed for the object's life. Restoring a ghost REPLACES it with a new live row
    // rather than flipping a flag, so a grid never has to notice a row changing kind under it.
    // ###########################################################################################
    public sealed class BoardTableRow : INotifyPropertyChanged
    {
        private BoardTableRowState thisState;

        internal BoardTableRow(BoardTableSheet sheet, IReadOnlyList<string> values, bool isDeleted)
        {
            this.Sheet = sheet;
            this.IsDeleted = isDeleted;

            var cells = new List<BoardTableCell>(sheet.Columns.Count);
            for (int i = 0; i < sheet.Columns.Count; i++)
            {
                cells.Add(new BoardTableCell(this, i, i < values.Count ? values[i] : string.Empty));
            }

            this.Cells = cells;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public BoardTableSheet Sheet { get; }

        public IReadOnlyList<BoardTableCell> Cells { get; }

        public bool IsDeleted { get; }

        // For a live row, the published row it pairs with; for a ghost, the published row it
        // shows. -1 when there is none.
        internal int PublishedIndex { get; set; } = -1;

        // Added with "Insert row" in THIS sitting. On the Components sheet such a row is placed in
        // its category, in label order, when saved - the same rule every other way of adding a
        // component follows (ComponentPlacement) - unless the contributor has moved it by hand.
        internal bool IsInsertedThisSession { get; set; }

        // Moved by the contributor (Move up/down, or dragged). Their placement wins over the
        // automatic one.
        internal bool IsPlacedByUser { get; set; }

        public BoardTableRowState State
        {
            get => this.thisState;
            internal set
            {
                if (this.thisState == value)
                {
                    return;
                }

                this.thisState = value;
                this.Raise(nameof(this.State));
                this.Raise(nameof(this.Marker));
                this.Raise(nameof(this.MarkerToolTip));

                foreach (BoardTableCell cell in this.Cells)
                {
                    cell.RaiseToolTipChanged();
                }
            }
        }

        public bool IsBlank => this.Cells.All(cell => cell.Text.Length == 0);

        // ###########################################################################################
        // The one character at the start of the row. It says the same thing as the colour, in a
        // form that does not depend on telling red from green - roughly one man in twelve cannot
        // reliably do that, and a deleted row must never be mistaken for an added one.
        // ###########################################################################################
        public string Marker => this.thisState switch
        {
            BoardTableRowState.Added or BoardTableRowState.Blank => "+",
            BoardTableRowState.Modified => "~",
            BoardTableRowState.Deleted => "-",
            BoardTableRowState.Duplicate or BoardTableRowState.Incomplete => "!",
            _ => string.Empty,
        };

        public string? MarkerToolTip => this.thisState switch
        {
            BoardTableRowState.Added => "Added - this row is not in the published data",
            BoardTableRowState.Modified => "Changed - hover an orange cell to see its published value",
            BoardTableRowState.Deleted => "Deleted - this row is in the published data but not in your draft. Ctrl+Z brings it back if you deleted it since your last save.",
            BoardTableRowState.Blank => "New, empty row - it is not saved until something is typed in it",
            BoardTableRowState.Duplicate => "Another row above identifies the same thing, so this one is not counted as a change. Check it is not a duplicate.",
            BoardTableRowState.Incomplete => "Incomplete - this row is left out when saving, because a value it needs is missing",
            _ => null,
        };

        // The row as the schema's mappers read it: column name to cell text.
        internal IReadOnlyDictionary<string, string> ToDictionary()
        {
            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < this.Cells.Count; i++)
            {
                row[this.Sheet.Columns[i]] = this.Cells[i].Text;
            }

            return row;
        }

        internal IReadOnlyList<string> Values() => this.Cells.Select(cell => cell.Text).ToList();

        private void Raise(string propertyName) =>
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
