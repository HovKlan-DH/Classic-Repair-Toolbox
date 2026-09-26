using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using System;

namespace CRT
{
    // ###########################################################################################
    // HOVERING A FILE CELL SHOWS THE FILE (owner request, 2026-09-26) - the picture, the published
    // and the new picture side by side when it changed, or a link that opens a PDF. The same in the
    // Drafts tab and in the maintainer application, which is why it is here and not in either.
    //
    // *** AT ONCE, BOTH WAYS (owner request, 2026-09-26: "show image/tooltip instantly, instead of
    // the small delay ... and likewise instantly NOT show the helper when moving mouse outside"). ***
    // The card opens the moment the pointer is on a file cell, moves to the next file cell as the
    // pointer does, and closes the moment the pointer is on neither its cell nor the card. It used
    // to wait 450 ms before opening, and - as a flyout - to close only once the pointer was about
    // 100 px away from it.
    //
    // *** THE POINTER MAY CROSS ONTO THE CARD, so a PDF's link can be clicked. *** That is why the
    // card is a popup in the window's overlay layer (FilePreviewPopup in the markup): there "the
    // pointer is on the card" is the card's own IsPointerOver, known exactly. Leaving the grid or
    // the card is judged once the move has been handled (posted), since the card is only marked as
    // under the pointer after the grid has been told the pointer left it.
    //
    // Never during a row drag or with a button held, and the mouse wheel closes it (the rows move
    // under a still pointer). A file column's text tooltip ("Published value: ...") gives way to the
    // card, which shows both paths anyway.
    //
    // Nothing happens until the host sets FileSource - which is how each application supplies the
    // bytes (IBoardTableFileSource).
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        private IBoardTableFileSource? thisFileSource;

        // The card on screen, and the cell it is for.
        private BoardTableFilePreview? thisFilePreview;
        private (BoardTableRow Row, int Column, Control Anchor)? thisFilePreviewShownFor;

        // Set while the card is closed only to be opened again at another cell, so its Closed
        // handler does not let go of the NEW card's pictures.
        private bool thisMovingFilePreview;

        // ###########################################################################################
        // Where the preview reads files from - null (the default) shows none. Setting it rebuilds
        // the columns, since a file column then carries the card instead of the text tooltip.
        // ###########################################################################################
        public IBoardTableFileSource? FileSource
        {
            get => this.thisFileSource;
            set
            {
                if (ReferenceEquals(this.thisFileSource, value))
                {
                    return;
                }

                this.thisFileSource = value;
                this.HideFilePreview();
                this.RebuildColumns();
            }
        }

        // The card open on screen, or null - for tests.
        internal BoardTableFilePreview? ShownFilePreviewForTests => this.FilePreviewPopup.IsOpen ? this.thisFilePreview : null;

        // The card's frame, where the pointer is when it is "on the card" - for tests.
        internal Control FilePreviewFrameForTests => this.FilePreviewFrame;

        private void WireFilePreview()
        {
            this.TableGrid.AddHandler(PointerMovedEvent, this.OnFilePreviewPointerMoved, RoutingStrategies.Tunnel, handledEventsToo: true);
            this.TableGrid.AddHandler(PointerWheelChangedEvent, (_, _) => this.HideFilePreview(), RoutingStrategies.Tunnel, handledEventsToo: true);

            // Leaving the grid - onto the card, or anywhere else - and leaving the card.
            this.TableGrid.PointerExited += (_, _) => this.HideFilePreviewUnlessPointerIsOnIt();
            this.FilePreviewFrame.PointerExited += (_, _) => this.HideFilePreviewUnlessPointerIsOnIt();

            // Also closed when the editor leaves the window; its pictures go with it.
            this.FilePreviewPopup.Closed += (_, _) =>
            {
                if (!this.thisMovingFilePreview && !this.FilePreviewPopup.IsOpen)
                {
                    this.ReleaseFilePreview();
                }
            };
        }

        private void OnFilePreviewPointerMoved(object? sender, PointerEventArgs e)
        {
            if (this.thisFileSource is null ||
                this.thisRowDrag is not null ||
                e.GetCurrentPoint(this.TableGrid).Properties.IsLeftButtonPressed)
            {
                this.HideFilePreview();
                return;
            }

            DataGridCell? cell = (e.Source as Visual)?.FindAncestorOfType<DataGridCell>(includeSelf: true);
            (BoardTableRow Row, int Column)? target = cell is null ? null : BoardTableEditor.FileCellOf(cell);

            if (target is null)
            {
                this.HideFilePreview();
                return;
            }

            // Still on the cell whose card is open.
            if (this.thisFilePreviewShownFor is { } shown &&
                ReferenceEquals(shown.Row, target.Value.Row) &&
                shown.Column == target.Value.Column &&
                ReferenceEquals(shown.Anchor, cell))
            {
                return;
            }

            this.ShowFilePreview(target.Value.Row, target.Value.Column, cell!);
        }

        // ###########################################################################################
        // After the pointer left the grid or the card: closes the card unless the pointer is now on
        // the card or back on its cell. Posted, because the card is marked as under the pointer only
        // after the grid is told the pointer left it. A move onto another cell of the grid has been
        // handled by then - OnFilePreviewPointerMoved moved or closed the card already.
        // ###########################################################################################
        private void HideFilePreviewUnlessPointerIsOnIt()
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (this.thisFilePreviewShownFor is not { } shown ||
                    this.FilePreviewFrame.IsPointerOver ||
                    shown.Anchor.IsPointerOver)
                {
                    return;
                }

                this.HideFilePreview();
            });
        }

        // The row and data-column index of a grid cell - only when it is a file cell that names a
        // file. The marker column (Tag -1) and every other column answer null.
        private static (BoardTableRow Row, int Column)? FileCellOf(DataGridCell cell)
        {
            if (DataGridColumn.GetColumnContainingElement(cell)?.Tag is not int column ||
                cell.DataContext is not BoardTableRow row ||
                column < 0 ||
                column >= row.Cells.Count ||
                BoardTableFileCells.Of(row.Cells[column]) is null)
            {
                return null;
            }

            return (row, column);
        }

        // ###########################################################################################
        // Opens the card for one cell, beside `anchor` - replacing the card of another cell. Returns
        // the preview, or null (and no card) when the cell names no file or there is no source.
        //
        // A card already open is closed and opened again: that is what places it at the new cell.
        // ###########################################################################################
        internal BoardTableFilePreview? ShowFilePreview(BoardTableRow row, int columnIndex, Control anchor)
        {
            BoardTableFilePreview? preview = this.BuildFilePreview(row, columnIndex);

            if (preview is null)
            {
                this.HideFilePreview();
                return null;
            }

            BoardTableFilePreview? previous = this.thisFilePreview;

            this.thisFilePreview = preview;
            this.thisFilePreviewShownFor = (row, columnIndex, anchor);
            this.FilePreviewFrame.Child = preview;

            previous?.ReleaseImages();

            this.thisMovingFilePreview = true;

            try
            {
                this.FilePreviewPopup.IsOpen = false;
                this.FilePreviewPopup.PlacementTarget = anchor;
                this.FilePreviewPopup.IsOpen = true;
            }
            finally
            {
                this.thisMovingFilePreview = false;
            }

            return preview;
        }

        // ###########################################################################################
        // What the card holds for one cell - the seam the tests use for the card's contents.
        //
        // A flagged row's reason rides along: its cells' tooltip gives way to the card on a file
        // column, and the reason must not disappear with it.
        // ###########################################################################################
        internal BoardTableFilePreview? BuildFilePreview(BoardTableRow row, int columnIndex)
        {
            if (this.thisFileSource is null || columnIndex < 0 || columnIndex >= row.Cells.Count)
            {
                return null;
            }

            BoardTableCell cell = row.Cells[columnIndex];
            BoardTableFileCell? file = BoardTableFileCells.Of(cell);

            if (file is null)
            {
                return null;
            }

            string? note = cell.State == BoardTableCellState.Flagged ? cell.ToolTip : null;

            return new BoardTableFilePreview(file, this.thisFileSource, note);
        }

        // Whether this column's cells show the card rather than their text tooltip.
        private bool PreviewsFilesIn(int columnIndex) =>
            this.thisFileSource is not null &&
            this.thisCurrentSheet is not null &&
            columnIndex >= 0 &&
            columnIndex < this.thisCurrentSheet.Columns.Count &&
            BoardTableFileCells.IsFileColumn(this.thisCurrentSheet.Name, this.thisCurrentSheet.Columns[columnIndex]);

        private void HideFilePreview()
        {
            if (this.FilePreviewPopup.IsOpen)
            {
                this.FilePreviewPopup.IsOpen = false;
            }

            this.ReleaseFilePreview();
        }

        // Takes the card out of the frame and lets go of its pictures - the Image first, then the
        // bitmap (BoardTableFilePreview.ReleaseImages).
        private void ReleaseFilePreview()
        {
            BoardTableFilePreview? preview = this.thisFilePreview;

            this.thisFilePreview = null;
            this.thisFilePreviewShownFor = null;
            this.FilePreviewFrame.Child = null;

            preview?.ReleaseImages();
        }
    }
}
