using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.TextFormatting;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // BoardTableEditor - SEEING ALL OF A LONG TEXT.
    //
    //   - EVERY CELL WRAPS ITS TEXT within its column's width, and each row grows to its tallest
    //     cell (owner request, 2026-10-04: "remove that checkbox again and then always apply wrap
    //     text on cells, if the column size gets reduced"). The grid's CellsWrap class turns the
    //     cells' wrapping on (the editor's styles), and the grid has no fixed row height. A row with
    //     nothing long stays MinRowHeight tall. It was a "Wrap text" check box, remembered by the
    //     host, for the same day before that.
    //
    //   - A COLUMN CAN BE DRAGGED AS WIDE AS WANTED. Why it could not be: every column is sized to its
    //     text up to MaxColumnWidth, so one long note does not push every other column off screen -
    //     and that cap held for a drag too. Pressing the header row frees it: each column keeps the
    //     width it has, with no upper limit, until the next sheet or table. A header is pressed for
    //     nothing else here - sorting is off - so a column then stops growing to fit rows scrolled
    //     into view only once somebody has set out to size them.
    //
    //   - DOUBLE-CLICKING A HEADING'S EDGE FITS THE COLUMN to the widest text in EVERY row shown, as
    //     Excel does (owner request, same day) - never wider than the table can show, since the cells
    //     wrap, and then scrolled wholly into view. The decision is
    //     Handlers/Geometry/ColumnAutoFitGeometry; this measures. The grid's
    //     own double-click fit is OFF (CanUserResizeColumnsOnDoubleClick in the markup): it measures
    //     only the rows on screen.
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        // The widest a column sizes ITSELF to - a drag or a double-click goes past it (see the header).
        internal const double MaxColumnWidth = 420;

        // The height of a row with nothing long in it - dense, close to Excel's (owner request,
        // 2026-09-24). The editor's styles give it to the cells (a MinHeight on the row changed
        // nothing), so keep the two in step.
        internal const double MinRowHeight = 26;

        // Whether the columns of the sheet on screen have been freed of MaxColumnWidth.
        private bool thisColumnWidthsFreed;

        private void WireColumnWidths()
        {
            // Tunnel and handled too: the header starts its resize on this press itself, and a
            // double-click on an edge must be answered before it starts another.
            this.TableGrid.AddHandler(PointerPressedEvent, this.OnGridPointerPressedForColumnWidths, RoutingStrategies.Tunnel, handledEventsToo: true);
        }

        private void OnGridPointerPressedForColumnWidths(object? sender, PointerPressedEventArgs e)
        {
            if (e.Source is not Visual source || source.FindAncestorOfType<DataGridColumnHeader>(includeSelf: true) is not DataGridColumnHeader header)
            {
                return;
            }

            this.FreeColumnWidths();

            if (e.ClickCount != 2 || !e.GetCurrentPoint(header).Properties.IsLeftButtonPressed)
            {
                return;
            }

            if (this.ColumnAtEdge(header, e.GetPosition(header).X) is DataGridColumn column)
            {
                this.FitColumnToText(column);

                // Or the header starts a resize drag on this second press.
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Every data column at the width it has now, with no upper limit - so a drag can take it as
        // wide as the text needs. Once per sheet shown (RebuildColumns builds them capped again).
        // ###########################################################################################
        internal void FreeColumnWidths()
        {
            if (this.thisColumnWidthsFreed)
            {
                return;
            }

            foreach (DataGridColumn column in this.TableGrid.Columns)
            {
                if (column.Tag is not int index || index < 0)
                {
                    continue;
                }

                double width = column.ActualWidth;

                column.MaxWidth = double.PositiveInfinity;

                if (width > 0)
                {
                    column.Width = new DataGridLength(width);
                }
            }

            this.thisColumnWidthsFreed = true;
        }

        // The DATA column whose edge a press at x on this heading is on - the marker column is never
        // resized, so its edge answers nothing.
        private DataGridColumn? ColumnAtEdge(DataGridColumnHeader header, double x)
        {
            DataGridColumn? owning = header.OwningColumn;

            if (owning is null)
            {
                return null;
            }

            DataGridColumn? column = ColumnAutoFitGeometry.EdgeAt(x, header.Bounds.Width) switch
            {
                ColumnEdge.ThisColumn => owning,
                ColumnEdge.PreviousColumn => this.TableGrid.Columns
                    .Where(candidate => candidate.IsVisible && candidate.DisplayIndex < owning.DisplayIndex)
                    .OrderBy(candidate => candidate.DisplayIndex)
                    .LastOrDefault(),
                _ => null,
            };

            return column?.Tag is int index && index >= 0 ? column : null;
        }

        // ###########################################################################################
        // The column as wide as its heading and the widest text in every row the grid shows (a
        // picked pill's or a search's rows - hidden rows are not looked at), within the room the
        // table has. Each text is measured in the cells' own font; what a cell adds around its text
        // is read off a cell the grid has built.
        // ###########################################################################################
        internal void FitColumnToText(DataGridColumn column)
        {
            if (column.Tag is not int index || index < 0)
            {
                return;
            }

            this.FreeColumnWidths();

            double headerWidth = 0;

            if (this.TableGrid.GetVisualDescendants().OfType<DataGridColumnHeader>().FirstOrDefault(header => ReferenceEquals(header.OwningColumn, column)) is DataGridColumnHeader shownHeader)
            {
                shownHeader.Measure(Size.Infinity);
                headerWidth = shownHeader.DesiredSize.Width;
            }

            // A cell the grid has built, for the font and the margins around the text. With none -
            // nothing shown - only the heading counts.
            DataGridCell? sample = this.TableGrid.GetVisualDescendants()
                .OfType<DataGridCell>()
                .FirstOrDefault(cell => ReferenceEquals(cell.OwningColumn, column) && cell.IsEffectivelyVisible && cell.DataContext is BoardTableRow);

            TextBlock? sampleText = sample?.GetVisualDescendants().OfType<TextBlock>().FirstOrDefault();

            var textWidths = new List<double>();
            double chrome = 0;

            if (sample is not null && sampleText is not null)
            {
                var typeface = new Typeface(sampleText.FontFamily, sampleText.FontStyle, sampleText.FontWeight, sampleText.FontStretch);
                double fontSize = sampleText.FontSize;

                sample.Measure(Size.Infinity);
                chrome = sample.DesiredSize.Width - BoardTableEditor.TextWidth(BoardTableEditor.CellText((BoardTableRow)sample.DataContext!, index), typeface, fontSize);

                foreach (BoardTableRow row in this.TableGrid.ItemsSource?.OfType<BoardTableRow>() ?? [])
                {
                    textWidths.Add(BoardTableEditor.TextWidth(BoardTableEditor.CellText(row, index), typeface, fontSize));
                }
            }

            double width = ColumnAutoFitGeometry.FitWidth(headerWidth, textWidths, chrome, column.MinWidth, this.RoomForOneColumn());

            column.Width = new DataGridLength(width);

            // *** AND THE WHOLE COLUMN IN VIEW. *** A column grows to the right, and the grid keeps its
            // scroll position - so a long note fitted to the table's width had its right half past
            // the edge, and the text the fit was for still could not be read (seen in a render).
            // Sideways only (no row), and after the new width is laid out, or the grid scrolls by
            // the old one.
            this.TableGrid.UpdateLayout();
            this.TableGrid.ScrollIntoView(null!, column);
        }

        private static string CellText(BoardTableRow row, int index) =>
            index < row.Cells.Count ? row.Cells[index].Text : string.Empty;

        // The width a TextBlock gives this text on one line - each line of a text with line breaks
        // measured on its own, the widest counting.
        private static double TextWidth(string text, Typeface typeface, double fontSize)
        {
            if (text.Length == 0)
            {
                return 0;
            }

            using var layout = new TextLayout(text, typeface, fontSize, foreground: null);

            return layout.WidthIncludingTrailingWhitespace;
        }

        // ###########################################################################################
        // The widest a column can be and still be seen whole: the scrolling part of the table (the
        // header row's width, past the row grips and short of the scroll bar) less the frozen marker
        // column. Unknown - the grid not laid out - is no limit.
        // ###########################################################################################
        private double RoomForOneColumn()
        {
            if (this.TableGrid.GetVisualDescendants().OfType<DataGridColumnHeadersPresenter>().FirstOrDefault() is not DataGridColumnHeadersPresenter headers ||
                headers.Bounds.Width <= 0)
            {
                return double.PositiveInfinity;
            }

            double frozen = this.TableGrid.Columns.Where(column => column.IsVisible && column.IsFrozen).Sum(column => column.ActualWidth);
            double room = headers.Bounds.Width - frozen;

            return room > 0 ? room : double.PositiveInfinity;
        }
    }
}
