using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // ONE ROW BEING DRAGGED IN THE BOARD TABLE EDITOR (owner request, 2026-09-24): "insert a
    // placeholder when I drag a row, like when moving an image in the worklog".
    //
    // The row moves LIVE while it is dragged - into the place of the row under the pointer - so
    // the rows around it shift into the order a drop will produce, and the UI draws the dragged
    // row as an empty dashed slot: what is seen is the space the row will take, exactly as the
    // worklog's photo and file lists do it.
    //
    // THE RULES, here rather than in the UI so the maintainer application's table gets them too:
    //
    //   - A red deleted row is never a place to drop onto. Where a ghost shows is worked out from
    //     the published order after every move (BoardTableSheet.PlaceGhosts), so "dropping onto" one
    //     would be undone by the very refresh that followed - and the dragged row would flicker
    //     between the two places for as long as the pointer stayed there.
    //   - The whole drag is ONE undo step, however many rows it passed (BoardTableHistory's group).
    //   - A drag that ends where it started changes nothing: no undo step, not unsaved, and the row
    //     keeps its automatic placement on save (see BoardTableHistory.EndGroup).
    //
    // Started by BoardTableSheet.BeginRowDrag and ended by Finish, always - an unfinished drag
    // would leave the history grouping every later change into one step.
    // ###########################################################################################
    public sealed class BoardTableRowDrag
    {
        private readonly BoardTableSheet thisSheet;

        internal BoardTableRowDrag(BoardTableSheet sheet, BoardTableRow row)
        {
            this.thisSheet = sheet;
            this.Row = row;
        }

        public BoardTableRow Row { get; }

        public bool IsFinished { get; private set; }

        // ###########################################################################################
        // Moves the dragged row into `target`'s place - the row under the pointer - so the target
        // shifts one place towards where the dragged row came from. False when nothing moved: the
        // target is the dragged row itself, a red ghost, or not in the sheet.
        //
        // With every row the same height that is stable: after the move the dragged row IS under
        // the pointer, so the next pointer event over it does nothing.
        // ###########################################################################################
        public bool MoveOnto(BoardTableRow target)
        {
            ArgumentNullException.ThrowIfNull(target);

            if (this.IsFinished || target.IsDeleted || ReferenceEquals(target, this.Row))
            {
                return false;
            }

            int from = this.thisSheet.Rows.IndexOf(this.Row);
            int to = this.thisSheet.Rows.IndexOf(target);
            if (from < 0 || to < 0)
            {
                return false;
            }

            // MoveRow takes an insert position in the rows as they are NOW, so going down it is the
            // slot after the target.
            return this.thisSheet.MoveRow(this.Row, to > from ? to + 1 : to);
        }

        // One live row further up or down - for a pointer dragged past the top or bottom of the
        // rows on screen, where there is no row under it to move onto.
        public bool Step(bool up)
        {
            if (this.IsFinished)
            {
                return false;
            }

            return up ? this.thisSheet.MoveRowUp(this.Row) : this.thisSheet.MoveRowDown(this.Row);
        }

        // Ends the drag. True when the row really ended up somewhere else.
        public bool Finish()
        {
            if (this.IsFinished)
            {
                return false;
            }

            this.IsFinished = true;
            return this.thisSheet.EndRowDrag();
        }
    }
}
