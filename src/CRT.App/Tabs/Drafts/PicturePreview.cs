using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.Geometry;
using System;
using System.Collections.Generic;

namespace CRT
{
    // ###########################################################################################
    // ONE OF THE BOARD'S SCHEMATIC IMAGES WITH ITS MARKED COMPONENTS, DRAWN IN THE COLOURS PICKED
    // (CRT 4.0 step 2, slice 4b - the plan: "with a live preview on the picture"). The component pane's
    // Schematic images part hands it the image made small, the image's own pixel size, its highlights
    // in those pixels, and the colour and opacity - and it draws again as they change. Every position
    // is PicturePreviewGeometry's, so this only paints - and, since 2026-10-29 (owner request: "maybe
    // to zoom in/out in that small thumbnail so I can view relevant parts of it"):
    //
    //   - IT ZOOMS (IsZoomable): Ctrl+scroll at the pointer - a plain scroll goes on scrolling the
    //     cards, as on a web page's map - drag to move, double-click to see it whole. Zoomed in, it
    //     asks for the image decoded sharper (SharperWanted, with the pixels across it needs).
    //   - IT DRAWS EITHER SIDE'S TRACES when handed them (in the image's pixels) - this side's in the
    //     highlight's trace colour over the other side's in the trace colour - the card's "Traces on the
    //     preview".
    //   - THE HIGHLIGHTS ARE DRAWN AS THE SCHEMATICS TAB DRAWS THEM: the fill at the opacity, the edge
    //     a little stronger (SchematicHighlightsOverlay's rule), so the preview is the real look - the
    //     thumbnail-sized one too.
    // ###########################################################################################
    public sealed class PicturePreview : Control
    {
        public static readonly StyledProperty<Bitmap?> SourceProperty =
            AvaloniaProperty.Register<PicturePreview, Bitmap?>(nameof(Source));

        public static readonly StyledProperty<Size> ImageSizeProperty =
            AvaloniaProperty.Register<PicturePreview, Size>(nameof(ImageSize));

        public static readonly StyledProperty<IReadOnlyList<Rect>?> HighlightsProperty =
            AvaloniaProperty.Register<PicturePreview, IReadOnlyList<Rect>?>(nameof(Highlights));

        public static readonly StyledProperty<Color> HighlightColorProperty =
            AvaloniaProperty.Register<PicturePreview, Color>(nameof(HighlightColor), Colors.IndianRed);

        public static readonly StyledProperty<double> HighlightOpacityProperty =
            AvaloniaProperty.Register<PicturePreview, double>(nameof(HighlightOpacity), 0.2);

        public static readonly StyledProperty<IReadOnlyList<PreviewTrace>?> ThisSideTracesProperty =
            AvaloniaProperty.Register<PicturePreview, IReadOnlyList<PreviewTrace>?>(nameof(ThisSideTraces));

        public static readonly StyledProperty<Color> ThisSideTraceColorProperty =
            AvaloniaProperty.Register<PicturePreview, Color>(nameof(ThisSideTraceColor), Colors.DeepSkyBlue);

        public static readonly StyledProperty<IReadOnlyList<PreviewTrace>?> OtherSideTracesProperty =
            AvaloniaProperty.Register<PicturePreview, IReadOnlyList<PreviewTrace>?>(nameof(OtherSideTraces));

        public static readonly StyledProperty<Color> OtherSideTraceColorProperty =
            AvaloniaProperty.Register<PicturePreview, Color>(nameof(OtherSideTraceColor), Colors.DodgerBlue);

        public static readonly StyledProperty<double> ThisSideTraceOpacityProperty =
            AvaloniaProperty.Register<PicturePreview, double>(nameof(ThisSideTraceOpacity), 0.45);

        public static readonly StyledProperty<double> OtherSideTraceOpacityProperty =
            AvaloniaProperty.Register<PicturePreview, double>(nameof(OtherSideTraceOpacity), 0.45);

        public static readonly StyledProperty<bool> IsZoomableProperty =
            AvaloniaProperty.Register<PicturePreview, bool>(nameof(IsZoomable));

        public static readonly StyledProperty<Action<int>?> SharperWantedProperty =
            AvaloniaProperty.Register<PicturePreview, Action<int>?>(nameof(SharperWanted));

        // How much one notch of the wheel zooms.
        private const double ZoomStep = 1.25;

        private double thisZoom = 1.0;
        private Point thisCentre = PicturePreviewGeometry.Middle;
        private Point? thisDragFrom;
        private int thisLastWanted;

        static PicturePreview()
        {
            Control.AffectsRender<PicturePreview>(
                SourceProperty, ImageSizeProperty, HighlightsProperty, HighlightColorProperty, HighlightOpacityProperty,
                ThisSideTracesProperty, ThisSideTraceColorProperty, OtherSideTracesProperty, OtherSideTraceColorProperty,
                ThisSideTraceOpacityProperty, OtherSideTraceOpacityProperty);

            // Another image is shown whole again.
            ImageSizeProperty.Changed.AddClassHandler<PicturePreview>((preview, _) => preview.ShowWhole());

            // A sharper image taken away again - another card took the one sharp image - is asked for
            // again (code review, 2026-10-30: only zooming past the widest asked before did, so the
            // preview stayed soft at the zoom it was left at). Asked at the next zoom, or when the
            // pointer comes back and rests on it - never at once, or two zoomed cards would take it from
            // each other.
            SourceProperty.Changed.AddClassHandler<PicturePreview>((preview, change) => preview.OnSourceChanged(change));
        }

        public PicturePreview()
        {
            this.ClipToBounds = true;
        }

        public Bitmap? Source { get => this.GetValue(SourceProperty); set => this.SetValue(SourceProperty, value); }

        // The image's own pixel size - the highlights are in those pixels, the Source is smaller.
        public Size ImageSize { get => this.GetValue(ImageSizeProperty); set => this.SetValue(ImageSizeProperty, value); }

        public IReadOnlyList<Rect>? Highlights { get => this.GetValue(HighlightsProperty); set => this.SetValue(HighlightsProperty, value); }

        public Color HighlightColor { get => this.GetValue(HighlightColorProperty); set => this.SetValue(HighlightColorProperty, value); }

        // 0..1, as the Schematics tab draws it.
        public double HighlightOpacity { get => this.GetValue(HighlightOpacityProperty); set => this.SetValue(HighlightOpacityProperty, value); }

        // The traces on the image's own side and on its other side, in its pixels - null draws none.
        public IReadOnlyList<PreviewTrace>? ThisSideTraces { get => this.GetValue(ThisSideTracesProperty); set => this.SetValue(ThisSideTracesProperty, value); }

        public Color ThisSideTraceColor { get => this.GetValue(ThisSideTraceColorProperty); set => this.SetValue(ThisSideTraceColorProperty, value); }

        public IReadOnlyList<PreviewTrace>? OtherSideTraces { get => this.GetValue(OtherSideTracesProperty); set => this.SetValue(OtherSideTracesProperty, value); }

        public Color OtherSideTraceColor { get => this.GetValue(OtherSideTraceColorProperty); set => this.SetValue(OtherSideTraceColorProperty, value); }

        // 0..1 - each side's, as on the Schematics tab (TraceColours - the same unless the picture says
        // otherwise, 2026-10-30).
        public double ThisSideTraceOpacity { get => this.GetValue(ThisSideTraceOpacityProperty); set => this.SetValue(ThisSideTraceOpacityProperty, value); }

        public double OtherSideTraceOpacity { get => this.GetValue(OtherSideTraceOpacityProperty); set => this.SetValue(OtherSideTraceOpacityProperty, value); }

        public bool IsZoomable { get => this.GetValue(IsZoomableProperty); set => this.SetValue(IsZoomableProperty, value); }

        // Told the pixels across the image it needs to be drawn sharp, when that is more than Source has.
        public Action<int>? SharperWanted { get => this.GetValue(SharperWantedProperty); set => this.SetValue(SharperWantedProperty, value); }

        internal double ZoomForTests => this.thisZoom;

        internal Point CentreForTests => this.thisCentre;

        // ---- zooming and moving -----------------------------------------------------------------

        // The image whole, as it first shows.
        internal void ShowWhole()
        {
            this.thisZoom = 1.0;
            this.thisCentre = PicturePreviewGeometry.Middle;
            this.thisDragFrom = null;
            this.thisLastWanted = 0;
            this.Cursor = null;
            this.InvalidateVisual();
        }

        // Zoomed by `factor` at `pointer` - for the wheel, and for a test.
        internal void ZoomAt(Point pointer, double factor)
        {
            if (!this.IsZoomable || this.DrawnSize() is not { } image)
            {
                return;
            }

            (this.thisZoom, this.thisCentre) = PicturePreviewGeometry.ZoomAt(image, this.Bounds.Size, this.thisZoom, this.thisCentre, pointer, factor);
            this.Cursor = this.thisZoom > 1.0 ? new Cursor(StandardCursorType.Hand) : null;
            this.InvalidateVisual();
            this.AskSharperIfNeeded();
        }

        protected override void OnPointerWheelChanged(PointerWheelEventArgs e)
        {
            base.OnPointerWheelChanged(e);

            // Ctrl (Cmd on a Mac) and the wheel zoom; the wheel alone is the cards' scroll.
            if (!this.IsZoomable || (e.KeyModifiers & (KeyModifiers.Control | KeyModifiers.Meta)) == 0 || e.Delta.Y == 0)
            {
                return;
            }

            this.ZoomAt(e.GetPosition(this), Math.Pow(ZoomStep, e.Delta.Y));
            e.Handled = true;
        }

        protected override void OnPointerPressed(PointerPressedEventArgs e)
        {
            base.OnPointerPressed(e);

            if (!this.IsZoomable || !e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            {
                return;
            }

            if (e.ClickCount >= 2)
            {
                this.ShowWhole();
                e.Handled = true;
                return;
            }

            if (this.thisZoom > 1.0)
            {
                this.thisDragFrom = e.GetPosition(this);
                e.Pointer.Capture(this);
                this.Cursor = new Cursor(StandardCursorType.SizeAll);
                e.Handled = true;
            }
        }

        protected override void OnPointerMoved(PointerEventArgs e)
        {
            base.OnPointerMoved(e);

            if (this.thisDragFrom is not { } from || this.DrawnSize() is not { } image)
            {
                return;
            }

            Point now = e.GetPosition(this);
            this.thisCentre = PicturePreviewGeometry.Pan(image, this.Bounds.Size, this.thisZoom, this.thisCentre, now - from);
            this.thisDragFrom = now;
            this.InvalidateVisual();
        }

        protected override void OnPointerReleased(PointerReleasedEventArgs e)
        {
            base.OnPointerReleased(e);

            if (this.thisDragFrom is null)
            {
                return;
            }

            this.thisDragFrom = null;
            e.Pointer.Capture(null);
            this.Cursor = this.thisZoom > 1.0 ? new Cursor(StandardCursorType.Hand) : null;
        }

        protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e)
        {
            base.OnPointerCaptureLost(e);
            this.thisDragFrom = null;
        }

        protected override void OnSizeChanged(SizeChangedEventArgs e)
        {
            base.OnSizeChanged(e);
            this.AskSharperIfNeeded();
        }

        // ###########################################################################################
        // Back over a zoomed preview whose sharp image another card took: asked for again - once the
        // pointer has RESTED on it (PointerRest), never as it passes (code review, 2026-10-30: with two
        // cards zoomed, every pass of the pointer took the one sharp image from the other card, each
        // a decode of up to 25 MB).
        // ###########################################################################################
        internal static readonly TimeSpan PointerRest = TimeSpan.FromMilliseconds(500);

        private DispatcherTimer? thisRestTimer;

        internal bool IsWaitingForPointerRestForTests => this.thisRestTimer?.IsEnabled == true;

        protected override void OnPointerEntered(PointerEventArgs e)
        {
            base.OnPointerEntered(e);

            if (!this.IsZoomable || this.thisZoom <= 1.0)
            {
                return;
            }

            if (this.thisRestTimer is null)
            {
                this.thisRestTimer = new DispatcherTimer { Interval = PicturePreview.PointerRest };
                this.thisRestTimer.Tick += (_, _) => this.PointerRested();
            }

            this.thisRestTimer.Stop();
            this.thisRestTimer.Start();
        }

        protected override void OnPointerExited(PointerEventArgs e)
        {
            base.OnPointerExited(e);
            this.thisRestTimer?.Stop();
        }

        protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnDetachedFromVisualTree(e);
            this.thisRestTimer?.Stop();
        }

        // The pointer rested on the preview - what it needs is asked for.
        internal void PointerRested()
        {
            this.thisRestTimer?.Stop();
            this.AskSharperIfNeeded();
        }

        // A narrower image than before - the sharp one let go of - is no answer to what was asked.
        private void OnSourceChanged(AvaloniaPropertyChangedEventArgs change)
        {
            if (change.NewValue is not Bitmap now || (change.OldValue is Bitmap before && now.PixelSize.Width < before.PixelSize.Width))
            {
                this.thisLastWanted = 0;
            }
        }

        // ###########################################################################################
        // The image ZOOMED larger than Source can show it sharp - its pixels across, asked for once per
        // amount (the pane decodes it in steps and keeps one sharp image at a time). Never for a preview
        // shown whole: widening the previews resizes every card, and each would have been decoded again.
        // ###########################################################################################
        private void AskSharperIfNeeded()
        {
            if (!this.IsZoomable || this.thisZoom <= 1.0 ||
                this.SharperWanted is not { } ask || this.Source is not { } bitmap || this.DrawnSize() is not { } image)
            {
                return;
            }

            double scaling = TopLevel.GetTopLevel(this)?.RenderScaling ?? 1.0;
            int wanted = PicturePreviewGeometry.PixelsWanted(image, this.Bounds.Size, this.thisZoom, scaling);

            if (wanted > bitmap.PixelSize.Width * 1.15 && wanted > this.thisLastWanted)
            {
                this.thisLastWanted = wanted;
                ask(wanted);
            }
        }

        // ---- drawing ----------------------------------------------------------------------------

        // The size the image is drawn by - its own when known, else the decoded one (no highlights then).
        private Size? DrawnSize()
        {
            if (this.Source is not { } bitmap)
            {
                return null;
            }

            return this.ImageSize.Width > 0 && this.ImageSize.Height > 0 ? this.ImageSize : bitmap.Size;
        }

        public override void Render(DrawingContext context)
        {
            if (this.Source is not { } bitmap || this.DrawnSize() is not { } image)
            {
                return;
            }

            // The highlights are in the image's OWN pixels: with its size not known, none is drawn -
            // placed against the size it was decoded to, they came out far too big (code review,
            // 2026-10-26). ImagePixelSize reads every type CRT shows, so this is a header it could not read.
            bool sizeKnown = this.ImageSize.Width > 0 && this.ImageSize.Height > 0;
            Rect target = PicturePreviewGeometry.View(image, this.Bounds.Size, this.thisZoom, this.thisCentre);

            if (target.Width <= 0)
            {
                return;
            }

            context.DrawImage(bitmap, new Rect(bitmap.Size), target);

            if (!sizeKnown)
            {
                return;
            }

            // The other side first, this side over it - the Schematics tab's order.
            this.DrawTraces(context, this.OtherSideTraces, this.OtherSideTraceColor, this.OtherSideTraceOpacity, image, target);
            this.DrawTraces(context, this.ThisSideTraces, this.ThisSideTraceColor, this.ThisSideTraceOpacity, image, target);

            double opacity = Math.Clamp(this.HighlightOpacity, 0.0, 1.0);
            var fill = new SolidColorBrush(this.HighlightColor, opacity);
            var outline = new Pen(new SolidColorBrush(this.HighlightColor, Math.Min(1.0, opacity * 1.4)), 1);

            foreach (Rect highlight in this.Highlights ?? [])
            {
                context.DrawRectangle(fill, outline, PicturePreviewGeometry.Place(highlight, image, target));
            }
        }

        // One side's traces, in `colour` at that side's `opacity` - at least a pixel wide, the Schematics tab's least too.
        private void DrawTraces(DrawingContext context, IReadOnlyList<PreviewTrace>? traces, Color colour, double opacity, Size image, Rect target)
        {
            if (traces is not { Count: > 0 })
            {
                return;
            }

            double scale = target.Width / image.Width;
            var brush = new SolidColorBrush(colour, Math.Clamp(opacity, 0.0, 1.0));
            var pens = new Dictionary<double, Pen>();

            foreach (PreviewTrace trace in traces)
            {
                double thickness = Math.Round(Math.Max(1.0, trace.Width * scale), 1);

                if (!pens.TryGetValue(thickness, out Pen? pen))
                {
                    pen = new Pen(brush, thickness, lineCap: PenLineCap.Round);
                    pens[thickness] = pen;
                }

                context.DrawLine(pen, PicturePreviewGeometry.Place(trace.From, image, target), PicturePreviewGeometry.Place(trace.To, image, target));
            }
        }
    }
}
