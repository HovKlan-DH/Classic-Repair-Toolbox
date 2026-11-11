using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.VisualTree;
using System;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE FIRST ROW AS TALL AS THE CELLS IT SHOWS, AS EVERY OTHER ROW IS (owner report, 2026-11-11: "I do not understand why
    // the first row in table is higher than the other rows? I think it should have same height as the other rows").
    //
    // A row of the table grows to its tallest WRAPPED cell among the columns on screen: ProDataGrid hides the cells of columns
    // scrolled out of view, so they count for nothing - scrolled to a long note, every row grows to show it; scrolled back, they
    // shrink again. Every row but the first. ProDataGrid shows every cell of the row at index 0 whatever the scroll
    // (DataGridCellsPresenter.MeasureOverride: "ShouldDisplayCell(...) || OwningRow.Index == 0" - that row is what it sizes
    // columns by), so that row alone grew to a three-line "Purpose or note" far off to the right, and stood twice as tall as the
    // same rows under it.
    //
    // So after each layout the first row's cells outside the columns on screen are held to a plain row's height (MaxHeight);
    // a cell scrolled into view is let go again, and the row measured as the others are. Only the row at index 0 - a container
    // the grid reuses for another row is let go of as it loads.
    //
    // Part of the BoardTableEditor partial class - see BoardTableEditor.axaml.cs for the file map.
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        // The container showing the first row, while there is one - and its cells' panel, and the grid's rows' area.
        private DataGridRow? thisFirstRow;
        private DataGridCellsPresenter? thisFirstRowCells;
        private DataGridRowsPresenter? thisRowsPresenter;

        private void WireFirstRowHeight()
        {
            this.TableGrid.LoadingRow += (_, e) => this.FollowLoadedRow(e.Row);
            this.TableGrid.UnloadingRow += (_, e) =>
            {
                if (ReferenceEquals(e.Row, this.thisFirstRow))
                {
                    BoardTableEditor.LetGoOfCells(e.Row);
                    this.thisFirstRow = null;
                }
            };
            this.TableGrid.LayoutUpdated += (_, _) => this.HoldFirstRowToCellsShown();
        }

        // ###########################################################################################
        // Only the container held as the first row ever has cells held (HoldFirstRowToCellsShown), so only it is let go of -
        // never a walk through every row's cells as the grid loads and unloads them while scrolling. A NEW container loaded at
        // index 0 while another is held - a row put in above the first, the grid making a container for it and moving the old
        // one down without loading it again - lets go of the old one first: replaced unseen, it kept its cells held, and a long
        // note scrolled into view on it stood cut to one line.
        // ###########################################################################################
        private void FollowLoadedRow(DataGridRow row)
        {
            if (this.thisFirstRow is { } held && (ReferenceEquals(held, row) || row.Index == 0))
            {
                BoardTableEditor.LetGoOfCells(held);
                this.thisFirstRow = null;
            }

            if (row.Index == 0)
            {
                this.thisFirstRow = row;
            }
        }

        // A container's cells with no height of their own held - the first row's, as it stops being the first.
        private static void LetGoOfCells(DataGridRow row)
        {
            foreach (DataGridCell cell in row.GetVisualDescendants().OfType<DataGridCell>())
            {
                if (!double.IsPositiveInfinity(cell.MaxHeight))
                {
                    cell.MaxHeight = double.PositiveInfinity;
                }
            }
        }

        // ###########################################################################################
        // The first row's cells outside the columns on screen held to a plain row's height, and those on screen let go - set only
        // where it changes, as each setting lays the row out again. "On screen" is the grid's own rule, read off where the
        // cells stand: a frozen column (the marker) always; another while some of it lies between the frozen columns' edge and
        // the right edge of the rows' area.
        // ###########################################################################################
        private void HoldFirstRowToCellsShown()
        {
            // LayoutUpdated comes after every layout pass anywhere in the window: a table not shown (the Contribute tab's "UI"
            // view, a Maintainer view switched away) has nothing to hold, and is looked at again as it is shown and laid out.
            if (!this.TableGrid.IsEffectivelyVisible)
            {
                return;
            }

            // A container the grid moved off index 0 without loading it again (a row put in above it, or the first row
            // deleted): its cells let go - held, a long note scrolled into view on it stood cut to one line - and the
            // container now at index 0 found among the rows shown.
            if (this.thisFirstRow is { } moved && moved.Index != 0)
            {
                BoardTableEditor.LetGoOfCells(moved);
                this.thisFirstRow = null;
            }

            // Looked up once each, not at every layout: this runs after every one.
            if (this.thisRowsPresenter is null || this.thisRowsPresenter.FindAncestorOfType<DataGrid>() is null)
            {
                this.thisRowsPresenter = this.TableGrid.GetVisualDescendants().OfType<DataGridRowsPresenter>().FirstOrDefault();
            }

            this.thisFirstRow ??= this.thisRowsPresenter?.Children.OfType<DataGridRow>().FirstOrDefault(shown => shown.IsVisible && shown.Index == 0);

            if (this.thisFirstRow is not { } row)
            {
                return;
            }

            if (this.thisFirstRowCells is null || !ReferenceEquals(this.thisFirstRowCells.GetVisualParent()?.FindAncestorOfType<DataGridRow>(includeSelf: true), row))
            {
                this.thisFirstRowCells = row.GetVisualDescendants().OfType<DataGridCellsPresenter>().FirstOrDefault();
            }

            if (this.thisFirstRowCells is not DataGridCellsPresenter cells ||
                this.thisRowsPresenter is not DataGridRowsPresenter rows ||
                cells.TranslatePoint(default, rows) is not Point at)
            {
                return;
            }

            double frozen = this.TableGrid.Columns.Where(column => column.IsVisible && column.IsFrozen).Sum(column => column.ActualWidth);
            double right = rows.Bounds.Width - at.X;

            foreach (DataGridCell cell in cells.Children.OfType<DataGridCell>())
            {
                if (cell.OwningColumn is not { } column)
                {
                    continue;
                }

                bool shown = column.IsFrozen || (cell.Bounds.Right > frozen && cell.Bounds.Left < right);
                double wanted = shown ? double.PositiveInfinity : BoardTableEditor.MinRowHeight;

                if (!cell.MaxHeight.Equals(wanted))
                {
                    cell.MaxHeight = wanted;
                }
            }
        }

        internal DataGridRow? FirstRowForTests => this.thisFirstRow;
    }
}
