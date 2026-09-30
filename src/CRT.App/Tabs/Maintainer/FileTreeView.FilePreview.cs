using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // POINTING AT A FILE SHOWS IT (owner request, 2026-09-28: "the same hover functionality per
    // file, as in the table format, so I can view an image or open a PDF etc. It should not show any
    // text, if it has changed or whatever ... only the relative path and then the other
    // functionality from the table format").
    //
    // So this is the table's card - BoardTableFilePreview, with one side and no headline
    // (FileTreePreviewSource) - and the table's behaviour, kept the same on purpose (see
    // BoardTableEditor.FilePreview.cs for why each part is as it is):
    //
    //   - at once, both ways: open the moment the pointer is on a file's row, moved to the next
    //     file's as the pointer goes, closed the moment it is on neither the row nor the card;
    //   - *** ONLY ON WHAT THE ROW SAYS (owner request, 2026-09-28: "It gets confusing when I am in
    //     the middle of an empty screen, that it then reacts of the file on that row"). *** The icon,
    //     the name and the marker - the row's RowContent panel - not the empty width of the list
    //     beside them, and not the indent before them. The row's hover shading follows the same
    //     part (the markup's styles), so what lights up is what reacts;
    //   - the pointer may cross onto the card, so its "Open PDF file" can be clicked - a popup in
    //     the window's overlay layer, whose IsPointerOver is known exactly;
    //   - never with a button held, and the wheel closes it (the rows move under a still pointer).
    //
    // A file with nothing to open yet has no card - its tooltip says why.
    // ###########################################################################################
    public partial class FileTreeView
    {
        // The card on screen, and the row it is for.
        private CRT.BoardTableFilePreview? thisFilePreview;
        private (FileTreeRowView Row, Control Anchor)? thisFilePreviewShownFor;

        // Set while the card is closed only to be opened again at another row, so its Closed
        // handler does not let go of the NEW card's pictures.
        private bool thisMovingFilePreview;

        // Found by name: this control loads its markup itself, so the generated fields stay empty.
        private Popup PreviewPopup => this.FindControl<Popup>("FilePreviewPopup")!;

        private Border PreviewFrame => this.FindControl<Border>("FilePreviewFrame")!;

        // The card open on screen, or null - for tests.
        internal CRT.BoardTableFilePreview? ShownFilePreviewForTests => this.PreviewPopup.IsOpen ? this.thisFilePreview : null;

        private void WireFilePreview()
        {
            if (this.FindControl<ListBox>("RowsList") is not ListBox list)
                return;

            list.AddHandler(PointerMovedEvent, this.OnFilePreviewPointerMoved, RoutingStrategies.Tunnel, handledEventsToo: true);
            list.AddHandler(PointerWheelChangedEvent, (_, _) => this.HideFilePreview(), RoutingStrategies.Tunnel, handledEventsToo: true);

            // Leaving the list - onto the card, or anywhere else - and leaving the card.
            list.PointerExited += (_, _) => this.HideFilePreviewUnlessPointerIsOnIt();
            this.PreviewFrame.PointerExited += (_, _) => this.HideFilePreviewUnlessPointerIsOnIt();

            // Also closed when the tree leaves the window; its pictures go with it.
            this.PreviewPopup.Closed += (_, _) =>
            {
                if (!this.thisMovingFilePreview && !this.PreviewPopup.IsOpen)
                    this.ReleaseFilePreview();
            };

            this.DetachedFromVisualTree += (_, _) => this.HideFilePreview();
        }

        private void OnFilePreviewPointerMoved(object? sender, PointerEventArgs e)
        {
            if (this.Files is null || e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            {
                this.HideFilePreview();
                return;
            }

            StackPanel? content = FileTreeView.RowContentAt(e.Source);

            if (content?.DataContext is not FileTreeRowView { Node.IsFolder: false } row ||
                row.Node.File!.OpenFrom == SystemFileSource.NotWrittenYet)
            {
                this.HideFilePreview();
                return;
            }

            // Still on the row whose card is open.
            if (this.thisFilePreviewShownFor is { } shown && ReferenceEquals(shown.Row, row))
                return;

            // Beside what the row says, rather than beside the whole row, which is as wide as the
            // list and would put the card past the window's edge.
            this.ShowFilePreview(row, content);
        }

        // What the row SAYS - its RowContent panel - under the pointer, or null over the empty
        // width of the list, the indent, or anything that is not a row.
        private static StackPanel? RowContentAt(object? source) =>
            (source as Visual)?.GetSelfAndVisualAncestors()
                .OfType<StackPanel>()
                .FirstOrDefault(panel => panel.Classes.Contains("RowContent"));

        // ###########################################################################################
        // After the pointer left the list or the card: closes the card unless the pointer is now on
        // the card or back on its row. Posted, because the card is marked as under the pointer only
        // after the list is told the pointer left it.
        // ###########################################################################################
        private void HideFilePreviewUnlessPointerIsOnIt()
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (this.thisFilePreviewShownFor is not { } shown ||
                    this.PreviewFrame.IsPointerOver ||
                    shown.Anchor.IsPointerOver)
                {
                    return;
                }

                this.HideFilePreview();
            });
        }

        // ###########################################################################################
        // Opens the card for one file's row, beside `anchor` - replacing the card of another row.
        // A card already open is closed and opened again: that is what places it at the new row.
        // ###########################################################################################
        internal CRT.BoardTableFilePreview? ShowFilePreview(FileTreeRowView row, Control anchor)
        {
            if (this.Files is null || row.Node.File is not SystemFileEntry file || file.OpenFrom == SystemFileSource.NotWrittenYet)
            {
                this.HideFilePreview();
                return null;
            }

            var source = new FileTreePreviewSource(this.Files, file);
            var preview = new CRT.BoardTableFilePreview(source.Cell, source, FileTreeWording.Note(file), showHeadline: false);

            CRT.BoardTableFilePreview? previous = this.thisFilePreview;

            this.thisFilePreview = preview;
            this.thisFilePreviewShownFor = (row, anchor);
            this.PreviewFrame.Child = preview;

            previous?.ReleaseImages();

            this.thisMovingFilePreview = true;

            try
            {
                this.PreviewPopup.IsOpen = false;
                this.PreviewPopup.PlacementTarget = anchor;
                this.PreviewPopup.IsOpen = true;
            }
            finally
            {
                this.thisMovingFilePreview = false;
            }

            return preview;
        }

        private void HideFilePreview()
        {
            if (this.PreviewPopup.IsOpen)
                this.PreviewPopup.IsOpen = false;

            this.ReleaseFilePreview();
        }

        // Takes the card out of the frame and lets go of its pictures - the Image first, then the
        // bitmap (BoardTableFilePreview.ReleaseImages).
        private void ReleaseFilePreview()
        {
            CRT.BoardTableFilePreview? preview = this.thisFilePreview;

            this.thisFilePreview = null;
            this.thisFilePreviewShownFor = null;
            this.PreviewFrame.Child = null;

            preview?.ReleaseImages();
        }
    }
}
