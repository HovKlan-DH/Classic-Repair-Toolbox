using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // UNDO AND REDO FOR THE BOARD TABLE EDITOR (owner request, 2026-09-24) - Ctrl+Z and
    // Ctrl+Y in the Drafts tab's "Edit in table format".
    //
    // *** EACH STEP IS A SNAPSHOT OF ONE SHEET, NOT A RECIPE FOR REVERSING ONE OPERATION. *** A
    // step holds the sheet's live rows as they were just before the change - the row OBJECTS, their
    // values and their two placement flags. Undoing puts exactly that back and lets Refresh redo
    // the colours and the deleted ghosts. So every kind of change - a cell, an inserted, deleted,
    // restored or moved row, a revert, a paste - is undone by the same few lines, and a change
    // added later is undoable the moment it records a step. Per-operation inverses would each be
    // a chance to get one case subtly wrong (a restored row is a NEW object replacing a ghost; a
    // move is re-anchored against ghosts), and a wrong inverse corrupts the table silently.
    //
    // Keeping the row OBJECTS, not copies, is what lets the UI put the cursor back on the row that
    // changed, and lets a redo bring back the very row an undo removed.
    //
    // A sheet is a few thousand rows at most and a snapshot shares the strings rather than copying
    // them, so a step costs little; MaxSteps bounds the total all the same.
    //
    // *** IT REACHES BACK TO THE LAST SAVE, NOT FURTHER. *** A save reloads the table from the file
    // (DraftTableSession, BoardTableEditor.Save), which makes a new document with a new history.
    //
    // *** A STEP CAN SPAN SHEETS (2026-09-25). *** Deleting a component also deletes its rows on the
    // other component sheets (BoardTableDocument.DeleteRowsOfComponent), and one Ctrl+Z must bring
    // all of it back. Within a group, the first change on EACH sheet adds that sheet's snapshot to
    // the group's one step; undo and redo restore every sheet in it. The step's FIRST sheet is the
    // one the gesture was made on, and is where the cursor goes.
    //
    // Avalonia-free, like the rest of the model, so the Maintainer tab gets it too.
    // ###########################################################################################
    public sealed class BoardTableHistory
    {
        // Enough for a long sitting; the oldest step is dropped beyond it.
        public const int MaxSteps = 200;

        private readonly BoardTableDocument thisDocument;
        private readonly List<Step> thisUndo = new();
        private readonly List<Step> thisRedo = new();

        // How many undo steps stand between the opening state and the last save: the document is
        // clean exactly when the undo stack is that deep. -1 once that state can no longer be
        // reached - its step was dropped past MaxSteps, or it sat on the redo stack when a new
        // change threw that away.
        private int thisSavedDepth;

        // True while a step is being put back, so the changes that makes are not recorded as new
        // steps of their own.
        private bool thisApplying;

        // A GROUP: one gesture made of many small changes - a row dragged past several others moves
        // once per row it passes - kept as ONE step. Only the first change in a group records a
        // snapshot; the rest ride along. See BeginGroup.
        private bool thisGrouping;
        private bool thisGroupRecorded;

        internal BoardTableHistory(BoardTableDocument document)
        {
            this.thisDocument = document;
        }

        public bool CanUndo => this.thisUndo.Count > 0;

        public bool CanRedo => this.thisRedo.Count > 0;

        public int UndoCount => this.thisUndo.Count;

        public int RedoCount => this.thisRedo.Count;

        // Whether the table is exactly as last saved (or as opened), by the history's account.
        internal bool IsAtSavedState => this.thisSavedDepth == this.thisUndo.Count;

        // ###########################################################################################
        // Called by the sheet JUST BEFORE it changes. `focusRow` is the row the change is about and
        // `focusIndex` where to put the cursor when that row does not exist after an undo or redo
        // (an undone insert). `focusColumn` is -1 when the change is to a whole row, meaning "keep
        // whichever column the cursor is in".
        // ###########################################################################################
        internal void Record(BoardTableSheet sheet, BoardTableRow? focusRow, int focusIndex, int focusColumn)
        {
            if (this.thisApplying)
            {
                return;
            }

            // A group already has its step: this sheet's state from before the group joins it, the
            // first time the group touches this sheet.
            if (this.thisGrouping && this.thisGroupRecorded)
            {
                Step current = this.thisUndo[^1];

                if (!current.Sheets.Any(state => ReferenceEquals(state.Sheet, sheet)))
                {
                    this.thisUndo[^1] = current with { Sheets = [.. current.Sheets, new SheetState(sheet, sheet.TakeSnapshot())] };
                }

                return;
            }

            this.thisGroupRecorded = this.thisGrouping;

            // The saved state was on the redo stack: this change throws it away for good.
            if (this.thisSavedDepth > this.thisUndo.Count)
            {
                this.thisSavedDepth = -1;
            }

            this.thisRedo.Clear();
            this.thisUndo.Add(new Step([new SheetState(sheet, sheet.TakeSnapshot())], focusRow, focusIndex, focusColumn));

            if (this.thisUndo.Count > BoardTableHistory.MaxSteps)
            {
                this.thisUndo.RemoveAt(0);
                this.thisSavedDepth = this.thisSavedDepth > 0 ? this.thisSavedDepth - 1 : -1;
            }
        }

        // ###########################################################################################
        // Takes back the latest change and says where it was, so the UI can show that sheet and put
        // the cursor on that row. Null when there is nothing to undo.
        // ###########################################################################################
        public BoardTableHistoryResult? Undo() => this.Apply(this.thisUndo, this.thisRedo);

        public BoardTableHistoryResult? Redo() => this.Apply(this.thisRedo, this.thisUndo);

        internal void MarkSaved() => this.thisSavedDepth = this.thisUndo.Count;

        // ###########################################################################################
        // Starts a group: every change until EndGroup is ONE undo step (a whole row drag).
        // ###########################################################################################
        internal void BeginGroup()
        {
            this.thisGrouping = true;
            this.thisGroupRecorded = false;
        }

        // ###########################################################################################
        // Ends the group. Returns whether it changed anything.
        //
        // A group that came to NOTHING - a row dragged away and back to where it started - leaves no
        // step behind: an undo that visibly did nothing would look broken. Its snapshot is put back
        // too, which also undoes the "moved by hand" mark the moves set on the row, so a new row
        // that merely took a trip still gets placed automatically on save.
        // ###########################################################################################
        internal bool EndGroup(BoardTableSheet sheet)
        {
            bool recorded = this.thisGrouping && this.thisGroupRecorded;
            this.thisGrouping = false;
            this.thisGroupRecorded = false;

            if (!recorded || this.thisUndo.Count == 0 || !ReferenceEquals(this.thisUndo[^1].Sheet, sheet))
            {
                return false;
            }

            Step step = this.thisUndo[^1];
            if (step.Sheets.Any(state => !state.Snapshot.HasSameRowsAs(state.Sheet.TakeSnapshot())))
            {
                return true;
            }

            this.thisUndo.RemoveAt(this.thisUndo.Count - 1);

            this.thisApplying = true;
            try
            {
                foreach (SheetState state in step.Sheets)
                {
                    state.Sheet.RestoreSnapshot(state.Snapshot);
                }
            }
            finally
            {
                this.thisApplying = false;
            }

            this.thisDocument.SetUnsavedChanges(!this.IsAtSavedState);
            return false;
        }

        private BoardTableHistoryResult? Apply(List<Step> from, List<Step> to)
        {
            if (from.Count == 0)
            {
                return null;
            }

            Step step = from[^1];
            from.RemoveAt(from.Count - 1);

            // The state being left goes on the other stack, so the opposite command returns to it.
            to.Add(step with
            {
                Sheets = step.Sheets.Select(state => new SheetState(state.Sheet, state.Sheet.TakeSnapshot())).ToList()
            });

            this.thisApplying = true;
            try
            {
                foreach (SheetState state in step.Sheets)
                {
                    state.Sheet.RestoreSnapshot(state.Snapshot);
                }
            }
            finally
            {
                this.thisApplying = false;
            }

            this.thisDocument.SetUnsavedChanges(!this.IsAtSavedState);

            return new BoardTableHistoryResult(step.Sheet, BoardTableHistory.FocusRowAfter(step), step.FocusColumn);
        }

        // The row the step was about when it is still in the sheet, otherwise the row now where it
        // was - the nearest thing to "the same place".
        private static BoardTableRow? FocusRowAfter(Step step)
        {
            IList<BoardTableRow> rows = step.Sheet.Rows;

            if (step.FocusRow is not null && rows.Contains(step.FocusRow))
            {
                return step.FocusRow;
            }

            return rows.Count == 0 ? null : rows[Math.Clamp(step.FocusIndex, 0, rows.Count - 1)];
        }

        // Every sheet the step changed, each with its state from before; the first is the sheet
        // the change was made on.
        private sealed record Step(
            IReadOnlyList<SheetState> Sheets,
            BoardTableRow? FocusRow,
            int FocusIndex,
            int FocusColumn)
        {
            public BoardTableSheet Sheet => this.Sheets[0].Sheet;
        }

        private sealed record SheetState(BoardTableSheet Sheet, BoardTableSheetSnapshot Snapshot);
    }

    // Where an undo or redo happened: the sheet, the row to put the cursor on, and the column (-1
    // when the change was to a whole row - keep the cursor's own column).
    public sealed record BoardTableHistoryResult(BoardTableSheet Sheet, BoardTableRow? Row, int Column);

    // One sheet's live rows at a moment: each row object with its values and placement flags. Ghosts
    // are not kept - Refresh works them out again from the live rows.
    internal sealed class BoardTableSheetSnapshot
    {
        internal BoardTableSheetSnapshot(IReadOnlyList<Entry> rows)
        {
            this.Rows = rows;
        }

        internal IReadOnlyList<Entry> Rows { get; }

        // The same row objects, in the same order, holding the same values - the placement flags
        // aside, which are bookkeeping about HOW a row got where it is, not what the table shows.
        internal bool HasSameRowsAs(BoardTableSheetSnapshot other) =>
            this.Rows.Count == other.Rows.Count &&
            this.Rows.Zip(other.Rows).All(pair =>
                ReferenceEquals(pair.First.Row, pair.Second.Row) &&
                pair.First.Values.SequenceEqual(pair.Second.Values, StringComparer.Ordinal));

        internal static BoardTableSheetSnapshot Of(IEnumerable<BoardTableRow> liveRows) =>
            new(liveRows
                .Select(row => new Entry(row, row.Values().ToArray(), row.IsInsertedThisSession, row.IsPlacedByUser))
                .ToList());

        internal sealed record Entry(BoardTableRow Row, string[] Values, bool IsInsertedThisSession, bool IsPlacedByUser);
    }
}
