using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // Draws a schematic with a moved highlight's BEFORE and AFTER rectangles on it
    // (NewContributeStrategy.md Phase 5, task 4).
    //
    // *** BOTH RECTANGLES GO ON ONE COPY OF THE BOARD, not two side by side. *** A move is small
    // relative to a board scan, and two images a few hundred pixels apart force the maintainer to
    // hold one in their head while looking at the other. Overlaying them makes the distance itself
    // the thing on screen, which is the question being asked.
    //
    // *** THE MATHS IS NOT HERE. *** ReviewHighlightGeometry converts the stored pixel coordinates
    // into fractions of the image and is unit tested; this control multiplies those fractions by
    // the size it has actually been given. That split is the whole reason the placement can be
    // verified by the suite rather than by holding the screen next to a board - and it is the same
    // split ExportOverlayGeometry uses for the PDF, after the version that confused fractions with
    // absolute lengths shipped and drew a tenth of a board as most of it.
    //
    // A custom control rather than a Canvas with children: the image is letterboxed inside
    // whatever space the panel gives it (Stretch.Uniform), so the rectangles have to be placed
    // against the DRAWN image rather than against the control - and that rect is only known during
    // rendering. Positioning children would need a layout pass that re-derives the same thing.
    // ###########################################################################################
    public sealed class ReviewHighlightCanvas : Control
    {
        // ###########################################################################################
        // The colours. RED is where the highlight WAS, GREEN is where it is being PUT.
        //
        // Deliberately the same vocabulary the change summary already uses - removals red,
        // additions green - so a maintainer does not have to learn a second colour language for the
        // one screen that draws geometry.
        //
        // Both are drawn as an outline over a light wash rather than a solid fill: a filled
        // rectangle hides the very circuitry the maintainer is trying to judge the placement against.
        // ###########################################################################################
        private static readonly IBrush BeforeStroke = new SolidColorBrush(Color.FromRgb(205, 92, 92));
        private static readonly IBrush AfterStroke = new SolidColorBrush(Color.FromRgb(46, 139, 87));

        private static readonly IBrush BeforeFill = new SolidColorBrush(Color.FromArgb(48, 205, 92, 92));
        private static readonly IBrush AfterFill = new SolidColorBrush(Color.FromArgb(48, 46, 139, 87));

        private Bitmap? thisBoard;
        private ReviewHighlightMove thisMove;

        // ###########################################################################################
        // The schematic to draw on. Set once the bytes have arrived and been decoded.
        //
        // The control OWNS this bitmap and disposes it on detach - unlike CRT's Workbooks tab,
        // where the bitmaps are shared across previews and a dispose under a still-open editor
        // was a fatal render-thread crash. Here each canvas fetched its own copy, so nothing else
        // can be holding it.
        // ###########################################################################################
        public Bitmap? Board
        {
            get => this.thisBoard;
            set
            {
                this.thisBoard = value;
                this.InvalidateMeasure();
                this.InvalidateVisual();
            }
        }

        public ReviewHighlightMove Move
        {
            get => this.thisMove;
            set
            {
                this.thisMove = value;
                this.InvalidateVisual();
            }
        }

        // ###########################################################################################
        // Asks for the board's own aspect ratio inside whatever width is available.
        //
        // Without this the control has no natural size and collapses to nothing in a StackPanel,
        // which would show the caption and no picture at all.
        // ###########################################################################################
        protected override Size MeasureOverride(Size availableSize)
        {
            if (this.thisBoard is null)
                return default;

            double imageWidth = this.thisBoard.PixelSize.Width;
            double imageHeight = this.thisBoard.PixelSize.Height;

            if (imageWidth <= 0 || imageHeight <= 0)
                return default;

            // A StackPanel measures with infinite height, so only the width constrains. MaxHeight
            // set by the caller still applies on top of whatever is returned here.
            double width = double.IsInfinity(availableSize.Width) ? imageWidth : availableSize.Width;

            return new Size(width, width * imageHeight / imageWidth);
        }

        public override void Render(DrawingContext context)
        {
            base.Render(context);

            if (this.thisBoard is null)
                return;

            double imageWidth = this.thisBoard.PixelSize.Width;
            double imageHeight = this.thisBoard.PixelSize.Height;

            if (imageWidth <= 0 || imageHeight <= 0)
                return;

            // Where the image is actually DRAWN inside this control. Uniform scaling letterboxes
            // it, and the rectangles must be placed against the picture rather than against the
            // control - otherwise every highlight on a letterboxed board is offset by the margin.
            Rect drawn = ReviewHighlightCanvas.Letterbox(this.Bounds.Size, imageWidth, imageHeight);

            if (drawn.Width <= 0 || drawn.Height <= 0)
                return;

            context.DrawImage(this.thisBoard, drawn);

            // BEFORE first, so the AFTER rectangle is on top where the two overlap - which they
            // usually do, since most moves are small. The proposal is what the maintainer is being
            // asked about, so it must not be hidden behind the thing it replaces.
            ReviewHighlightCanvas.DrawBox(
                context,
                drawn,
                imageWidth,
                imageHeight,
                this.thisMove.BeforeX, this.thisMove.BeforeY,
                this.thisMove.BeforeWidth, this.thisMove.BeforeHeight,
                ReviewHighlightCanvas.BeforeStroke,
                ReviewHighlightCanvas.BeforeFill);

            ReviewHighlightCanvas.DrawBox(
                context,
                drawn,
                imageWidth,
                imageHeight,
                this.thisMove.AfterX, this.thisMove.AfterY,
                this.thisMove.AfterWidth, this.thisMove.AfterHeight,
                ReviewHighlightCanvas.AfterStroke,
                ReviewHighlightCanvas.AfterFill);
        }

        // ###########################################################################################
        // One rectangle, placed proportionally.
        //
        // The fractions come from ReviewHighlightGeometry, which refuses anything unusable - a
        // malformed number, a zero size, a rectangle entirely off the board. So a side that cannot
        // be drawn simply is not, rather than being drawn wrongly.
        // ###########################################################################################
        private static void DrawBox(
            DrawingContext context,
            Rect drawn,
            double imageWidth,
            double imageHeight,
            string x,
            string y,
            string width,
            string height,
            IBrush stroke,
            IBrush fill)
        {
            if (!ReviewHighlightGeometry.TryBuild(
                    x, y, width, height, imageWidth, imageHeight, out ReviewHighlightBox box))
            {
                return;
            }

            var rect = new Rect(
                drawn.X + (box.Left * drawn.Width),
                drawn.Y + (box.Top * drawn.Height),
                box.Width * drawn.Width,
                box.Height * drawn.Height);

            context.FillRectangle(fill, rect);
            context.DrawRectangle(new Pen(stroke, 2), rect);
        }

        // ###########################################################################################
        // Where a uniformly-scaled image sits inside a control - the letterboxed content rect.
        //
        // Pure enough to reason about here rather than pulled into Handlers/: it is four lines with
        // no board-data vocabulary in it, and the thing worth testing (turning stored coordinates
        // into fractions) is already there.
        // ###########################################################################################
        private static Rect Letterbox(Size available, double imageWidth, double imageHeight)
        {
            if (available.Width <= 0 || available.Height <= 0)
                return default;

            double scale = System.Math.Min(
                available.Width / imageWidth,
                available.Height / imageHeight);

            double width = imageWidth * scale;
            double height = imageHeight * scale;

            return new Rect(
                (available.Width - width) / 2,
                (available.Height - height) / 2,
                width,
                height);
        }

        // ###########################################################################################
        // Releases the board bitmap.
        //
        // A full-resolution schematic is tens of megabytes decoded (a 4220x2941 scan is ~47 MB of
        // BGRA), and a maintainer working through a queue opens one submission after another. The
        // control owns its own copy, so nothing else can be rendering it when this runs - the
        // condition that makes CRT's Workbooks tab dispose its shared cache so carefully.
        // ###########################################################################################
        protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnDetachedFromVisualTree(e);

            this.thisBoard?.Dispose();
            this.thisBoard = null;
        }
    }
}
