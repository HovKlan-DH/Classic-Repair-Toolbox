using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE CELLS AND ROWS OF THE BOARD TABLE EDITOR (owner request, 2026-09-24) - the Drafts
    // tab's "Edit in table format", which shows a draft's workbook as editable sheets with every
    // difference from the published data coloured in.
    //
    // *** WHY THIS LIVES IN CRT.Data AND NOT IN THE APP. *** The project owner wants the same table
    // in the Maintainer tab later, so a maintainer can make the same edits. Everything that
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
        Deleted

        // There was a fifth, Flagged - violet, for a duplicate row and a row the save leaves out.
        // Both are WARNINGS since 2026-10-03 (owner request: "Should flagged now be treated as
        // warnings?"), so the colours mean only what changed.
    }

    // What one row is, which drives the one-character marker at its start. Deliberately richer
    // than the cell states: a duplicate and a row that will not survive a save are not changes,
    // so neither is coloured - the warning on a cell says what is wrong with them.
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
        // key, so this one is not a counted change - but it is still saved, so it is warned about
        // (BoardDataChecks, "row.duplicate").
        Duplicate,

        // A row the schema's own mapper drops on save (an important signal missing one of its two
        // halves). Warned about rather than silently lost (BoardTableDocument, "row.incomplete").
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
                this.Row.ForgetMapped();
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
            this.Row.ForgetMapped();
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

        // ###########################################################################################
        // The hover text: what is WRONG with the cell first (its errors and warnings - see
        // Problems), then on a changed cell the value it replaced. Null when there is nothing to
        // say, so an unchanged cell shows no tooltip at all rather than an empty box.
        // ###########################################################################################
        public string? ToolTip
        {
            get
            {
                string? state = this.thisState switch
                {
                    BoardTableCellState.Modified =>
                        $"{this.Row.Sheet.ReplacedValueLabel}: {(this.thisPublishedText.Length == 0 ? "(empty)" : this.thisPublishedText)}",
                    _ => null,
                };

                return BoardTableCell.JoinLines(this.ProblemToolTip, state);
            }
        }

        // ###########################################################################################
        // *** THE CHECKS' PROBLEMS ON THIS CELL (owner request, 2026-10-02) *** - BoardDataChecks,
        // placed here by BoardTableDocument.RefreshProblems. Empty for a deleted row or a new,
        // still empty one: neither is part of what a save writes, so neither is checked.
        // ###########################################################################################
        public IReadOnlyList<BoardDataProblem> Problems => this.thisProblems;

        // The worst of them - what the cell's corner mark shows.
        public BoardProblemLevel ProblemLevel => BoardDataChecks.Worst(this.thisProblems);

        // Only the problems, one per line, "Error: ..." / "Warning: ..." - for a file cell, whose
        // other hover text gives way to the file card (BoardTableEditor.FilePreview.cs) but whose
        // problem must still be readable. Null with none.
        public string? ProblemToolTip => BoardTableCell.DescribeProblems(this.thisProblems);

        private IReadOnlyList<BoardDataProblem> thisProblems = [];

        // True when the problems changed.
        internal bool SetProblems(IReadOnlyList<BoardDataProblem> problems)
        {
            if (this.thisProblems.SequenceEqual(problems))
            {
                return false;
            }

            this.thisProblems = problems;
            this.Raise(nameof(this.Problems));
            this.Raise(nameof(this.ProblemLevel));
            this.Raise(nameof(this.ProblemToolTip));
            this.Raise(nameof(this.ToolTip));

            return true;
        }

        public static string? DescribeProblems(IReadOnlyList<BoardDataProblem> problems) =>
            problems.Count == 0
                ? null
                : string.Join(
                    Environment.NewLine,
                    problems
                        .OrderByDescending(problem => problem.Level)
                        .Select(problem => $"{(problem.Level == BoardProblemLevel.Error ? "Error" : "Warning")}: {problem.Message}"));

        private static string? JoinLines(string? first, string? second) =>
            (first, second) switch
            {
                (null, _) => second,
                (_, null) => first,
                _ => first + Environment.NewLine + second
            };

        // A row's state can change with this cell's own state staying the same, so the row tells
        // its cells when their tooltip may have changed.
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

        // Whether any cell carries an error, or a warning - BoardDataChecks' problems, which the
        // colour key counts and filters on (BoardTableRowFilter).
        public bool HasErrors => this.Cells.Any(cell => cell.Problems.Any(problem => problem.Level == BoardProblemLevel.Error));

        public bool HasWarnings => this.Cells.Any(cell => cell.Problems.Any(problem => problem.Level == BoardProblemLevel.Warning));

        // ###########################################################################################
        // The one character at the start of the row. It says the same thing as the colour, in a
        // form that does not depend on telling red from green - roughly one man in twelve cannot
        // reliably do that, and a deleted row must never be mistaken for an added one.
        // ###########################################################################################
        //
        // A duplicate row has no sign, and a row the save leaves out has a new row's (2026-10-03):
        // neither is a change, and the warning on its cell says what is wrong. Both were "!" until
        // then.
        public string Marker => this.thisState switch
        {
            BoardTableRowState.Added or BoardTableRowState.Blank or BoardTableRowState.Incomplete => "+",
            BoardTableRowState.Modified => "~",
            BoardTableRowState.Deleted => "-",
            _ => string.Empty,
        };

        public string? MarkerToolTip => this.thisState switch
        {
            BoardTableRowState.Added => "Added - this row is not in the published data",
            BoardTableRowState.Modified => "Changed - hover an orange cell to see its published value",
            BoardTableRowState.Deleted => "Deleted - this row is in the published data but not in your draft. Ctrl+Z brings it back if you deleted it since your last save.",
            BoardTableRowState.Blank => "New, empty row - it is not saved until something is typed in it",
            BoardTableRowState.Incomplete => "New row - it is not saved until the values it needs are filled in",
            _ => null,
        };

        // ###########################################################################################
        // *** THE ROW AS A SAVE WOULD WRITE IT, REMEMBERED (code review, 2026-10-04). *** The checks
        // run after every cell edit (BoardTableDocument.RefreshProblems) and mapped every live row of
        // all nine sheets each time - thousands of rows re-read for one typed cell. The mapping is
        // kept until a cell of THIS row changes its text (ForgetMapped, from both of the cell's text
        // writers), so an edit maps one row again. Empty when the schema's mapper drops the row.
        // ###########################################################################################
        private IReadOnlyList<object>? thisMapped;

        internal IReadOnlyList<object> MappedForChecks()
        {
            if (this.thisMapped is null)
            {
                this.thisMapped = BoardWorkbookSchema.MapRows(this.Sheet.Definition, [this.ToDictionary()]);
                this.Sheet.CountRowMapped();
            }

            return this.thisMapped;
        }

        internal void ForgetMapped() => this.thisMapped = null;

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
