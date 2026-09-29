using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // ONE SHEET OF THE BOARD TABLE EDITOR - its rows, and how each compares with the published
    // data (owner request, 2026-09-24).
    //
    // *** THE PAIRING RULE IS BoardDataDiffer'S, AND IT MUST STAY THAT WAY. *** The Drafts tab
    // row already says "N rows changed", counted by BoardDataDiffer. This table shows the same
    // changes cell by cell, so the two must agree on which rows are "the same row", or the sheet
    // tabs would promise a different number from the line directly above them. So:
    //
    //   - a row's identity is BoardDraftNaturalKeys.ForRow, on the entry the schema's own mapper
    //     builds from the cells - never a second opinion about which columns form a key;
    //   - keys compare case-INSENSITIVELY, cell values trimmed and case-SENSITIVELY;
    //   - the FIRST row per key pairs, on both sides; a later one is a Duplicate and not counted.
    //
    // BoardTableDocumentTests asserts ChangeCount against BoardDataDiffer on the saved board, so a
    // drift between the two fails a test rather than a contributor's trust.
    //
    // *** A KEY EDIT IS A DELETE PLUS AN ADD, deliberately. *** Changing a Board label from U8 to
    // U9 shows U9 as added and the old U8 as a deleted ghost. That is how BoardDataDiffer, the
    // submission and the maintainer all see it, and the table showing anything friendlier would be
    // describing a change nobody downstream agrees happened.
    //
    // *** DELETED ROWS ARE SHOWN WHERE THEY WERE. *** Each ghost is placed straight after the
    // nearest EARLIER published row that still exists in the draft, or at the top when none does
    // - so a removed row appears between the neighbours it used to have. When rows have been
    // reordered in Excel the placement is approximate, which is fine for what it is: a visual cue
    // about where the row came from, not a claim about order (row order is not a change at all).
    //
    // With NO published board (a system created with "Add a new system") nothing is coloured
    // green, orange or red: every row would be green, which says nothing. Duplicates and incomplete
    // rows are still FLAGGED - "!" and violet - since those are about the data rather than about
    // publishing.
    // ###########################################################################################
    public sealed class BoardTableSheet
    {
        private readonly BoardTableDocument thisDocument;

        // The published side, fixed for the sheet's life: each row's cells (for the per-cell
        // comparison and for ghosts), and the first index per natural key (for pairing).
        private readonly IReadOnlyList<IReadOnlyList<string>> thisPublishedRows;
        private readonly Dictionary<string, int> thisPublishedFirstIndexByKey =
            new(StringComparer.OrdinalIgnoreCase);

        // Ghost rows by the published index they show, reused across refreshes so a grid does not
        // see a deleted row vanish and reappear every time a cell elsewhere is edited.
        private Dictionary<int, BoardTableRow> thisGhostsByPublishedIndex = new();

        internal BoardTableSheet(
            BoardTableDocument document,
            BoardWorkbookSchema.SheetDefinition definition,
            BoardData? published,
            BoardData draft)
        {
            this.thisDocument = document;
            this.Definition = definition;

            var publishedRows = new List<IReadOnlyList<string>>();

            if (published is not null)
            {
                IReadOnlyList<IReadOnlyDictionary<string, string>> rows = BoardWorkbookSchema.BuildRows(definition, published);
                IReadOnlyList<object> entries = BoardWorkbookSchema.EntriesOf(definition, published);

                for (int i = 0; i < rows.Count; i++)
                {
                    publishedRows.Add(this.Columns.Select(column => BoardTableSheet.CellOf(rows[i], column)).ToList());

                    // First wins, as in BoardDataDiffer.IndexByKey.
                    string key = BoardDraftNaturalKeys.ForRow(entries[i]);
                    this.thisPublishedFirstIndexByKey.TryAdd(key, i);
                }
            }

            this.thisPublishedRows = publishedRows;

            foreach (IReadOnlyDictionary<string, string> row in BoardWorkbookSchema.BuildRows(definition, draft))
            {
                this.Rows.Add(new BoardTableRow(
                    this,
                    this.Columns.Select(column => BoardTableSheet.CellOf(row, column)).ToList(),
                    isDeleted: false));
            }
        }

        // Raised after Refresh has recomputed states and counts, and after any row is added,
        // removed or restored - whatever a UI showing counts or a dirty flag needs to repaint for.
        public event EventHandler? Changed;

        // Raised when a cell's text changes, BEFORE the sheet has been re-paired. The UI answers it
        // by scheduling Refresh - see BoardTableCell.Text for why that is not done synchronously.
        public event EventHandler<BoardTableCell>? CellEdited;

        public BoardWorkbookSchema.SheetDefinition Definition { get; }

        public string Name => this.Definition.SheetName;

        public IReadOnlyList<string> Columns => this.Definition.ColumnOrder;

        // Every row the grid shows, live rows and deleted ghosts interleaved.
        public ObservableCollection<BoardTableRow> Rows { get; } = new();

        public bool HasBaseline => this.thisDocument.HasBaseline;

        // ###########################################################################################
        // *** WHAT A CHANGED CELL'S TOOLTIP CALLS THE VALUE IT REPLACED (owner request, 2026-09-26).
        // *** It always said "Published value: ...", which on a NEW SYSTEM names something that does
        // not exist, so every such tooltip read "Published value: (empty)" and told the maintainer
        // nothing: "yes, it will always be empty, but it should state [the contributor's value] if
        // it was changed from empty to something".
        //
        // The answer is the whole table's, and it belongs to whoever BUILT it - HasBaseline cannot
        // give it, because the maintainer's new-system table compares the submission with itself
        // and so has a baseline that is not published. See BoardTableDocument.Create.
        // ###########################################################################################
        public string ReplacedValueLabel => this.thisDocument.BaselineLabel;

        // Rows added, modified or deleted - the number on the sheet's tab. Blank, duplicate and
        // incomplete rows are not counted, matching BoardDataDiffer on the board a save produces.
        public int ChangeCount { get; private set; }

        // The same total split three ways, for the colour key's badges (2026-09-24). They always add
        // up to ChangeCount: rows added, rows with at least one changed cell, and deleted ghosts.
        public int AddedCount { get; private set; }

        public int ModifiedCount { get; private set; }

        public int DeletedCount { get; private set; }

        // Rows marked "!" - duplicates and incomplete rows - for the colour key's fourth badge. NOT
        // part of ChangeCount: BoardDataDiffer counts neither (a duplicate is ignored, an
        // incomplete row is dropped on save), and the tab's number must keep agreeing with it.
        public int FlaggedCount { get; private set; }

        // Whether a row is one the "Show changes only" view shows: anything not plainly unchanged - an
        // added, changed or deleted row, and also a row that needs a second look (blank, duplicate
        // or incomplete), since those are the contributor's own unfinished work.
        public static bool IsChangeRow(BoardTableRow row) =>
            row is not null && row.State != BoardTableRowState.Unchanged;

        // Whether "Show changes only" shows anything of this sheet - and so whether its tab shows
        // (BoardTableDocument.SheetsShown). Read live: an edit can add or remove the last such row.
        public bool HasChangeRows => this.Rows.Any(BoardTableSheet.IsChangeRow);

        // True between a cell edit and the Refresh that re-pairs the sheet.
        public bool NeedsRefresh { get; private set; }

        // ###########################################################################################
        // Adds an empty row straight after `after` ("Insert row below"), or at the end when that is
        // null (or not in this sheet). An empty row is not a change yet and is not saved until
        // something is typed in.
        // ###########################################################################################
        public BoardTableRow InsertRow(BoardTableRow? after)
        {
            int index = after is null ? -1 : this.Rows.IndexOf(after);
            int at = index < 0 ? this.Rows.Count : index + 1;

            // Undone, the cursor goes back to the row it was on when the button was pressed.
            return this.InsertRowAt(at, focusIndex: at - 1);
        }

        // ###########################################################################################
        // "Insert row above" (owner request, 2026-09-24): the empty row goes straight BEFORE
        // `before`, or at the very top when that is null or not in this sheet.
        //
        // Next to a red deleted row the new row can land on the far side of it: ghosts are placed
        // after the published row they followed, whatever sits around them (see PlaceGhosts).
        // ###########################################################################################
        public BoardTableRow InsertRowAbove(BoardTableRow? before)
        {
            int index = before is null ? -1 : this.Rows.IndexOf(before);
            int at = index < 0 ? 0 : index;

            // Undone, `before` is back at `at` - the row the cursor was on.
            return this.InsertRowAt(at, focusIndex: at);
        }

        private BoardTableRow InsertRowAt(int at, int focusIndex)
        {
            var row = new BoardTableRow(this, [], isDeleted: false) { IsInsertedThisSession = true };

            this.thisDocument.History.Record(this, row, focusIndex, -1);
            this.Rows.Insert(at, row);

            this.thisDocument.MarkDirty();
            this.Refresh();

            return row;
        }

        // ###########################################################################################
        // Starts dragging a live row - see BoardTableRowDrag. Null for a red ghost or a row not in
        // this sheet: a deleted row cannot be moved, since where it shows is worked out from the
        // published order, not chosen.
        // ###########################################################################################
        public BoardTableRowDrag? BeginRowDrag(BoardTableRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            if (row.IsDeleted || !this.Rows.Contains(row))
            {
                return null;
            }

            this.thisDocument.History.BeginGroup();
            return new BoardTableRowDrag(this, row);
        }

        internal bool EndRowDrag() => this.thisDocument.History.EndGroup(this);

        // ###########################################################################################
        // Moves a live row so it lands where `insertIndex` points in Rows as it is NOW - the usual
        // meaning of a drop position: 0 is the top, Rows.Count the bottom, and dropping a row on
        // its own position is a no-op. False for a ghost, a row not in this sheet, or no movement.
        //
        // *** ORDER IS NOT A CHANGE the comparison counts. *** BoardDataDiffer pairs rows by key and
        // ignores their order, so a moved row stays uncoloured. It IS saved - the order is what the
        // application's component list shows - and the table is marked unsaved.
        //
        // Deleted ghosts are re-placed after every move (Refresh), so they keep following the
        // published row before them.
        // ###########################################################################################
        public bool MoveRow(BoardTableRow row, int insertIndex)
        {
            ArgumentNullException.ThrowIfNull(row);

            int from = this.Rows.IndexOf(row);
            if (row.IsDeleted || from < 0)
            {
                return false;
            }

            int to = Math.Clamp(insertIndex > from ? insertIndex - 1 : insertIndex, 0, this.Rows.Count - 1);
            if (to == from)
            {
                return false;
            }

            this.thisDocument.History.Record(this, row, from, -1);
            this.Rows.Move(from, to);
            row.IsPlacedByUser = true;

            this.thisDocument.MarkDirty();
            this.Refresh();

            return true;
        }

        // ###########################################################################################
        // One step up or down, past the neighbouring LIVE row. Ghosts are skipped: moving past a
        // deleted row changes nothing that is saved, so a press that only did that would look
        // broken.
        // ###########################################################################################
        public bool MoveRowUp(BoardTableRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            for (int i = this.Rows.IndexOf(row) - 1; i >= 0; i--)
            {
                if (!this.Rows[i].IsDeleted)
                {
                    return this.MoveRow(row, i);
                }
            }

            return false;
        }

        // Whether MoveRowUp / MoveRowDown would do anything - for enabling the UI's buttons.
        public bool CanMoveRowUp(BoardTableRow? row) =>
            row is { IsDeleted: false } && this.Rows.Take(Math.Max(0, this.Rows.IndexOf(row))).Any(other => !other.IsDeleted);

        public bool CanMoveRowDown(BoardTableRow? row) =>
            row is { IsDeleted: false } && this.Rows.IndexOf(row) >= 0 &&
            this.Rows.Skip(this.Rows.IndexOf(row) + 1).Any(other => !other.IsDeleted);

        public bool MoveRowDown(BoardTableRow row)
        {
            ArgumentNullException.ThrowIfNull(row);

            int from = this.Rows.IndexOf(row);
            if (from < 0)
            {
                return false;
            }

            for (int i = from + 1; i < this.Rows.Count; i++)
            {
                if (!this.Rows[i].IsDeleted)
                {
                    return this.MoveRow(row, i + 1);
                }
            }

            return false;
        }

        // ###########################################################################################
        // Removes a live row. A row that was published comes straight back as a red ghost in its
        // place; one that was added simply goes. Returns false for a ghost, which is already gone.
        //
        // *** A DELETED COMPONENT TAKES EVERYTHING OF ITS OWN WITH IT (owner request,
        // 2026-09-25): "if a component really is deleted, then it should remove EVERYTHING related
        // to this component." *** On the Components sheet, the component's rows on the image, local
        // file and link sheets are deleted too - shown red on their own sheets, and all of it ONE
        // undo step - and its highlights go when the table is saved (BoardTableDocument.ApplyTo).
        // `deletedWith` says what else went, for the editor to tell the user; null when nothing did.
        // ###########################################################################################
        public bool DeleteRow(BoardTableRow row) => this.DeleteRow(row, out _);

        public bool DeleteRow(BoardTableRow row, out BoardTableDeletedWith? deletedWith)
        {
            ArgumentNullException.ThrowIfNull(row);
            deletedWith = null;

            int index = this.Rows.IndexOf(row);
            if (row.IsDeleted || index < 0)
            {
                return false;
            }

            if (!string.Equals(this.Name, BoardWorkbookSchema.SheetComponents, StringComparison.Ordinal))
            {
                this.RemoveLiveRows([row]);
                return true;
            }

            string label = this.CellText(row, BoardWorkbookSchema.ColBoardLabel);
            string region = this.CellText(row, BoardWorkbookSchema.ColRegion);

            BoardTableHistory history = this.thisDocument.History;
            history.BeginGroup();

            try
            {
                this.RemoveLiveRows([row]);
                deletedWith = this.thisDocument.DeleteRowsOfComponent(label, region);
            }
            finally
            {
                history.EndGroup(this);
            }

            return true;
        }

        // ###########################################################################################
        // Removes these live rows as ONE change - recorded once, refreshed once. The rows must be
        // live rows of this sheet; the first one is where an undo puts the cursor.
        // ###########################################################################################
        internal void RemoveLiveRows(IReadOnlyList<BoardTableRow> rows)
        {
            if (rows.Count == 0)
            {
                return;
            }

            this.thisDocument.History.Record(this, rows[0], this.Rows.IndexOf(rows[0]), -1);

            foreach (BoardTableRow row in rows)
            {
                this.Rows.Remove(row);
            }

            this.thisDocument.MarkDirty();
            this.Refresh();
        }

        // A cell's text by COLUMN NAME, trimmed - "" when the sheet has no such column.
        internal string CellText(BoardTableRow row, string column)
        {
            int index = -1;

            for (int i = 0; i < this.Columns.Count; i++)
            {
                if (string.Equals(this.Columns[i], column, StringComparison.OrdinalIgnoreCase))
                {
                    index = i;
                    break;
                }
            }

            return index < 0 || index >= row.Cells.Count
                ? string.Empty
                : BoardTableCell.NormaliseText(row.Cells[index].Text);
        }

        // ###########################################################################################
        // Brings a deleted row back with its published values, in the ghost's own position.
        // Returns the new live row, or null when `ghost` is not a ghost of this sheet.
        // ###########################################################################################
        public BoardTableRow? RestoreRow(BoardTableRow ghost)
        {
            ArgumentNullException.ThrowIfNull(ghost);

            int index = this.Rows.IndexOf(ghost);
            if (!ghost.IsDeleted || index < 0)
            {
                return null;
            }

            var restored = new BoardTableRow(this, ghost.Values(), isDeleted: false);

            this.thisDocument.History.Record(this, ghost, index, -1);
            this.Rows[index] = restored;
            this.thisGhostsByPublishedIndex.Remove(ghost.PublishedIndex);

            this.thisDocument.MarkDirty();
            this.Refresh();

            return restored;
        }

        // ###########################################################################################
        // Puts one changed cell back to its published value. False when there is nothing to revert
        // - an unchanged cell, a cell on an added row (no published value exists), or a ghost.
        // ###########################################################################################
        public bool RevertCell(BoardTableCell cell)
        {
            ArgumentNullException.ThrowIfNull(cell);

            if (cell.Row.IsDeleted || cell.State != BoardTableCellState.Modified)
            {
                return false;
            }

            cell.Text = cell.PublishedText;
            this.Refresh();

            return true;
        }

        // ###########################################################################################
        // Re-pairs every live row with the published data, recolours every cell, and places the
        // deleted ghosts. Cheap enough to run after every edit - a sheet is a few thousand rows at
        // most, and this is one pass over them plus one over the published rows.
        // ###########################################################################################
        public void Refresh()
        {
            bool hasBaseline = this.HasBaseline;
            var liveRows = this.Rows.Where(row => !row.IsDeleted).ToList();

            var seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var liveByPublishedIndex = new Dictionary<int, BoardTableRow>();
            int changes = 0;

            foreach (BoardTableRow row in liveRows)
            {
                row.PublishedIndex = -1;

                if (row.IsBlank)
                {
                    BoardTableSheet.Colour(row, BoardTableRowState.Blank, hasBaseline ? BoardTableCellState.Added : BoardTableCellState.Unchanged);
                    continue;
                }

                // The entry the schema builds from these cells is what is saved, and what the key
                // is built from. None at all means the mapper drops the row on save.
                IReadOnlyList<object> mapped = BoardWorkbookSchema.MapRows(this.Definition, [row.ToDictionary()]);
                if (mapped.Count == 0)
                {
                    BoardTableSheet.Colour(row, BoardTableRowState.Incomplete, BoardTableCellState.Flagged);
                    continue;
                }

                string key = BoardDraftNaturalKeys.ForRow(mapped[0]);

                if (!seenKeys.Add(key))
                {
                    BoardTableSheet.Colour(row, BoardTableRowState.Duplicate, BoardTableCellState.Flagged);
                    continue;
                }

                if (!hasBaseline)
                {
                    BoardTableSheet.Colour(row, BoardTableRowState.Unchanged, BoardTableCellState.Unchanged);
                    continue;
                }

                if (!this.thisPublishedFirstIndexByKey.TryGetValue(key, out int publishedIndex))
                {
                    BoardTableSheet.Colour(row, BoardTableRowState.Added, BoardTableCellState.Added);
                    changes++;
                    continue;
                }

                row.PublishedIndex = publishedIndex;
                liveByPublishedIndex[publishedIndex] = row;

                if (this.ComparePaired(row, this.thisPublishedRows[publishedIndex]))
                {
                    changes++;
                }
            }

            int deleted = this.PlaceGhosts(liveRows, liveByPublishedIndex, hasBaseline);
            changes += deleted;

            this.ChangeCount = changes;
            this.AddedCount = liveRows.Count(row => row.State == BoardTableRowState.Added);
            this.ModifiedCount = liveRows.Count(row => row.State == BoardTableRowState.Modified);
            this.DeletedCount = deleted;
            this.FlaggedCount = liveRows.Count(row => row.State is BoardTableRowState.Duplicate or BoardTableRowState.Incomplete);
            this.NeedsRefresh = false;

            this.Changed?.Invoke(this, EventArgs.Empty);
        }

        // ###########################################################################################
        // The live rows as the schema's mappers want them, for saving: ghosts and blank rows left
        // out, exactly as the reader skips an empty row.
        //
        // *** ON THE COMPONENTS SHEET, NEW ROWS ARE PLACED (owner request, 2026-09-24). *** A
        // row added with "Insert row" in this sitting, and not moved by hand since, goes into its
        // category in label order - the same rule the Contribute window and the label editor
        // follow (ComponentPlacement), so a component lands in the same place however it was
        // added. Every other row keeps the order on screen. Placing on SAVE rather than as the
        // label is typed keeps a row from jumping away from the cursor mid-edit.
        // ###########################################################################################
        public IReadOnlyList<IReadOnlyDictionary<string, string>> BuildRows() =>
            this.OrderedForSave()
                .Select(row => row.ToDictionary())
                .ToList();

        // True when saving would move at least one new component into place - so the UI can say
        // why rows moved.
        public bool HasRowsToPlaceOnSave =>
            this.Definition.SheetName == BoardWorkbookSchema.SheetComponents &&
            this.Rows.Any(BoardTableSheet.IsToPlace);

        private static bool IsToPlace(BoardTableRow row) =>
            !row.IsDeleted && !row.IsBlank && row.IsInsertedThisSession && !row.IsPlacedByUser;

        private List<BoardTableRow> OrderedForSave()
        {
            List<BoardTableRow> live = this.Rows.Where(row => !row.IsDeleted && !row.IsBlank).ToList();

            if (this.Definition.SheetName != BoardWorkbookSchema.SheetComponents)
            {
                return live;
            }

            List<BoardTableRow> ordered = live.Where(row => !BoardTableSheet.IsToPlace(row)).ToList();

            foreach (BoardTableRow row in live.Where(BoardTableSheet.IsToPlace))
            {
                List<ComponentEntry> placed = BoardWorkbookSchema.MapComponents(ordered.Select(existing => existing.ToDictionary()));
                ComponentEntry entry = BoardWorkbookSchema.MapComponents([row.ToDictionary()])[0];

                ordered.Insert(ComponentPlacement.InsertionIndex(placed, entry.Category, entry.BoardLabel), row);
            }

            return ordered;
        }

        internal void RecordCellEdit(BoardTableCell cell) =>
            this.thisDocument.History.Record(this, cell.Row, this.Rows.IndexOf(cell.Row), cell.ColumnIndex);

        // The live rows as they are now, for the history - see BoardTableHistory.
        internal BoardTableSheetSnapshot TakeSnapshot() =>
            BoardTableSheetSnapshot.Of(this.Rows.Where(row => !row.IsDeleted));

        // ###########################################################################################
        // Puts the live rows back exactly as a snapshot had them - the same row objects, in the same
        // order, with the same values and placement flags - then refreshes, which re-colours them
        // and works the deleted ghosts out again. An undo or a redo, never an edit: nothing here
        // records a step.
        // ###########################################################################################
        internal void RestoreSnapshot(BoardTableSheetSnapshot snapshot)
        {
            foreach (BoardTableSheetSnapshot.Entry entry in snapshot.Rows)
            {
                for (int i = 0; i < entry.Row.Cells.Count; i++)
                {
                    entry.Row.Cells[i].RestoreText(i < entry.Values.Length ? entry.Values[i] : string.Empty);
                }

                entry.Row.IsInsertedThisSession = entry.IsInsertedThisSession;
                entry.Row.IsPlacedByUser = entry.IsPlacedByUser;
            }

            // Live rows only: Refresh places the ghosts again around them.
            this.SyncRows(snapshot.Rows.Select(entry => entry.Row).ToList());
            this.Refresh();
        }

        internal void OnCellEdited(BoardTableCell cell)
        {
            this.NeedsRefresh = true;
            this.thisDocument.MarkDirty();

            this.CellEdited?.Invoke(this, cell);
        }

        // Compares a paired row cell by cell. True when anything differs.
        private bool ComparePaired(BoardTableRow row, IReadOnlyList<string> published)
        {
            bool anyModified = false;

            for (int i = 0; i < row.Cells.Count; i++)
            {
                BoardTableCell cell = row.Cells[i];
                string publishedText = i < published.Count ? published[i] : string.Empty;

                cell.PublishedText = publishedText;

                // Trimmed and ordinal - BoardDataDiffer.DifferingFields' rule exactly.
                bool modified = !string.Equals(cell.Text.Trim(), publishedText.Trim(), StringComparison.Ordinal);
                cell.State = modified ? BoardTableCellState.Modified : BoardTableCellState.Unchanged;

                anyModified |= modified;
            }

            row.State = anyModified ? BoardTableRowState.Modified : BoardTableRowState.Unchanged;

            return anyModified;
        }

        // ###########################################################################################
        // Works out which published rows are deleted, where each ghost goes, and brings Rows into
        // that order. Returns how many ghosts there are.
        // ###########################################################################################
        private int PlaceGhosts(
            List<BoardTableRow> liveRows,
            Dictionary<int, BoardTableRow> liveByPublishedIndex,
            bool hasBaseline)
        {
            var ghosts = new Dictionary<int, BoardTableRow>();

            if (hasBaseline)
            {
                // Only a key's FIRST published row can be a ghost - a later duplicate is ignored,
                // as BoardDataDiffer ignores it.
                foreach (int publishedIndex in this.thisPublishedFirstIndexByKey.Values)
                {
                    if (liveByPublishedIndex.ContainsKey(publishedIndex))
                    {
                        continue;
                    }

                    if (!this.thisGhostsByPublishedIndex.TryGetValue(publishedIndex, out BoardTableRow? ghost))
                    {
                        ghost = new BoardTableRow(this, this.thisPublishedRows[publishedIndex], isDeleted: true)
                        {
                            PublishedIndex = publishedIndex,
                        };
                    }

                    BoardTableSheet.Colour(ghost, BoardTableRowState.Deleted, BoardTableCellState.Deleted);
                    foreach (BoardTableCell cell in ghost.Cells)
                    {
                        cell.PublishedText = cell.Text;
                    }

                    ghosts[publishedIndex] = ghost;
                }
            }

            this.thisGhostsByPublishedIndex = ghosts;

            // Walk the published order, remembering the latest row that survives; each ghost goes
            // after it (or at the top), in published order.
            var topGhosts = new List<BoardTableRow>();
            var ghostsAfter = new Dictionary<BoardTableRow, List<BoardTableRow>>(ReferenceEqualityComparer.Instance);
            BoardTableRow? anchor = null;

            for (int publishedIndex = 0; publishedIndex < this.thisPublishedRows.Count; publishedIndex++)
            {
                if (liveByPublishedIndex.TryGetValue(publishedIndex, out BoardTableRow? surviving))
                {
                    anchor = surviving;
                }
                else if (ghosts.TryGetValue(publishedIndex, out BoardTableRow? ghost))
                {
                    if (anchor is null)
                    {
                        topGhosts.Add(ghost);
                    }
                    else
                    {
                        if (!ghostsAfter.TryGetValue(anchor, out List<BoardTableRow>? after))
                        {
                            after = new List<BoardTableRow>();
                            ghostsAfter[anchor] = after;
                        }

                        after.Add(ghost);
                    }
                }
            }

            var desired = new List<BoardTableRow>(liveRows.Count + ghosts.Count);
            desired.AddRange(topGhosts);

            foreach (BoardTableRow live in liveRows)
            {
                desired.Add(live);

                if (ghostsAfter.TryGetValue(live, out List<BoardTableRow>? after))
                {
                    desired.AddRange(after);
                }
            }

            this.SyncRows(desired);

            return ghosts.Count;
        }

        // ###########################################################################################
        // Brings Rows into the desired order with as FEW collection changes as possible - a grid
        // redraws on every one, and a rebuild would lose its scroll position and current cell.
        //
        // Stale rows go first, so the live rows (whose relative order never changes here) are
        // already in sequence and the second pass only has to slot ghosts in.
        // ###########################################################################################
        private void SyncRows(List<BoardTableRow> desired)
        {
            var wanted = new HashSet<BoardTableRow>(desired, ReferenceEqualityComparer.Instance);

            for (int i = this.Rows.Count - 1; i >= 0; i--)
            {
                if (!wanted.Contains(this.Rows[i]))
                {
                    this.Rows.RemoveAt(i);
                }
            }

            for (int i = 0; i < desired.Count; i++)
            {
                if (i < this.Rows.Count && ReferenceEquals(this.Rows[i], desired[i]))
                {
                    continue;
                }

                int existing = BoardTableSheet.IndexOf(this.Rows, desired[i], i + 1);
                if (existing >= 0)
                {
                    this.Rows.Move(existing, i);
                }
                else
                {
                    this.Rows.Insert(i, desired[i]);
                }
            }
        }

        private static int IndexOf(ObservableCollection<BoardTableRow> rows, BoardTableRow row, int from)
        {
            for (int i = from; i < rows.Count; i++)
            {
                if (ReferenceEquals(rows[i], row))
                {
                    return i;
                }
            }

            return -1;
        }

        private static void Colour(BoardTableRow row, BoardTableRowState rowState, BoardTableCellState cellState)
        {
            row.State = rowState;

            foreach (BoardTableCell cell in row.Cells)
            {
                cell.State = cellState;

                if (rowState != BoardTableRowState.Deleted)
                {
                    cell.PublishedText = string.Empty;
                }
            }
        }

        private static string CellOf(IReadOnlyDictionary<string, string> row, string column) =>
            row.TryGetValue(column, out string? value) ? value ?? string.Empty : string.Empty;
    }
}
