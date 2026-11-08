using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Handlers.Geometry;
using System;

namespace CRT
{
    // ###########################################################################################
    // A PHOTO SHOWN WHOLE, ZOOMED WITH THE WHEEL (owner request, 2026-11-06: "when I click a photo in here
    // and in the worklog, that I then can zoom in/out with scrollwheel") - the photo viewer's
    // (WorklogPhotoViewerWindow), opened from a worklog's photos and a device's.
    //
    // The wheel zooms at the pointer, keeping the point under it where it is; a drag with the right or the
    // middle button moves the photo while zoomed in, as on the Schematics tab (PanGesture, CRT's one rule,
    // owner request 2026-11-06); a double-click shows it whole again. Every position is the Contribute tab's
    // picture preview's own (PicturePreviewGeometry: fitted whole and centred, zoom 1 to its MaxZoom,
    // the photo always covering the area along a side it is larger than) - never a second copy of it.
    //
    // DrawImage's SOURCE rectangle is in the photo's PIXELS, never its DIPs (CLAUDE.md: on a display at
    // 125% the two differ, and DIPs drew part of the picture over the whole area).
    // ###########################################################################################
    public sealed class PhotoZoomView : Control
    {
        public static readonly StyledProperty<Bitmap?> SourceProperty =
            AvaloniaProperty.Register<PhotoZoomView, Bitmap?>(nameof(Source));

        // One notch of the wheel: a quarter larger or smaller - the picture preview's step.
        private const double ZoomStep = 1.25;

        private double thisZoom = 1.0;
        private Point thisCentre = PicturePreviewGeometry.Middle;

        // Where a drag last was - null while none is under way.
        private Point? thisDragFrom;

        static PhotoZoomView()
        {
            AffectsRender<PhotoZoomView>(SourceProperty);

            // Zoomed in, the photo is larger than the area - what lies outside it is cut off.
            ClipToBoundsProperty.OverrideDefaultValue<PhotoZoomView>(true);
        }

        public Bitmap? Source
        {
            get => this.GetValue(SourceProperty);
            set => this.SetValue(SourceProperty, value);
        }

        // How far in it is zoomed - 1 for whole - and the point of the photo in the middle. For a test, as PicturePreview's.
        internal double ZoomForTests => this.thisZoom;

        internal Point CentreForTests => this.thisCentre;

        protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
        {
            base.OnPropertyChanged(change);

            // Another photo starts whole.
            if (change.Property == SourceProperty)
            {
                this.ShowWhole();
            }
        }

        // ###########################################################################################
        // Zoomed by `factor` at `pointer` (in this control) - the wheel's step, or a test's.
        // ###########################################################################################
        internal void ZoomAt(Point pointer, double factor)
        {
            if (this.Source is not { } source)
            {
                return;
            }

            (this.thisZoom, this.thisCentre) = PicturePreviewGeometry.ZoomAt(
                PixelSizeOf(source), this.Bounds.Size, this.thisZoom, this.thisCentre, pointer, factor);
            this.Cursor = this.thisZoom > 1.0 ? new Cursor(StandardCursorType.Hand) : null;
            this.InvalidateVisual();
        }

        // Whole again, centred.
        internal void ShowWhole()
        {
            this.thisZoom = 1.0;
            this.thisCentre = PicturePreviewGeometry.Middle;
            this.Cursor = null;
            this.InvalidateVisual();
        }

        protected override void OnPointerWheelChanged(PointerWheelEventArgs e)
        {
            base.OnPointerWheelChanged(e);

            if (e.Delta.Y == 0 || this.Source is null)
            {
                return;
            }

            this.ZoomAt(e.GetPosition(this), Math.Pow(ZoomStep, e.Delta.Y));
            e.Handled = true;
        }

        // ###########################################################################################
        // A drag with the RIGHT or the MIDDLE button moves the zoomed photo - the Schematics tab's way, CRT's one way
        // (owner request, 2026-11-06: "the same pan functionality as present in the 'Schematics' tab"; PanGesture) - and
        // a left double-click shows it whole.
        // ###########################################################################################
        protected override void OnPointerPressed(PointerPressedEventArgs e)
        {
            base.OnPointerPressed(e);

            PointerPointProperties properties = e.GetCurrentPoint(this).Properties;

            if (properties.IsLeftButtonPressed && e.ClickCount >= 2)
            {
                this.ShowWhole();
                e.Handled = true;
                return;
            }

            if (PanGesture.IsPanPress(properties) && this.thisZoom > 1.0)
            {
                this.thisDragFrom = e.GetPosition(this);
                this.Cursor = new Cursor(StandardCursorType.SizeAll);
                e.Pointer.Capture(this);
                e.Handled = true;
            }
        }

        protected override void OnPointerMoved(PointerEventArgs e)
        {
            base.OnPointerMoved(e);

            if (this.thisDragFrom is not { } from || this.Source is not { } source)
            {
                return;
            }

            Point now = e.GetPosition(this);
            this.thisCentre = PicturePreviewGeometry.Pan(PixelSizeOf(source), this.Bounds.Size, this.thisZoom, this.thisCentre, now - from);
            this.thisDragFrom = now;
            this.InvalidateVisual();
        }

        protected override void OnPointerReleased(PointerReleasedEventArgs e)
        {
            base.OnPointerReleased(e);
            this.EndDrag(e.Pointer);
        }

        // A drag ended by losing the pointer - the window leaving the front, another control taking it - ends as a
        // release does, the drag's cursor going with it (code review, 2026-11-08: it stayed the four arrows).
        protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e)
        {
            base.OnPointerCaptureLost(e);
            this.EndDrag(pointer: null);
        }

        private void EndDrag(IPointer? pointer)
        {
            if (this.thisDragFrom is not null)
            {
                this.thisDragFrom = null;
                this.Cursor = this.thisZoom > 1.0 ? new Cursor(StandardCursorType.Hand) : null;
                pointer?.Capture(null);
            }
        }

        public override void Render(DrawingContext context)
        {
            base.Render(context);

            // The whole area hit-testable, the wheel and a press taken anywhere on it.
            context.FillRectangle(Brushes.Transparent, new Rect(this.Bounds.Size));

            if (this.Source is not { } source || source.PixelSize.Width <= 0 || source.PixelSize.Height <= 0)
            {
                return;
            }

            Rect target = PicturePreviewGeometry.View(PixelSizeOf(source), this.Bounds.Size, this.thisZoom, this.thisCentre);

            if (target.Width <= 0 || target.Height <= 0)
            {
                return;
            }

            using (context.PushRenderOptions(new RenderOptions { BitmapInterpolationMode = BitmapInterpolationMode.HighQuality }))
            {
                context.DrawImage(source, new Rect(0, 0, source.PixelSize.Width, source.PixelSize.Height), target);
            }
        }

        // The photo's own size, in its pixels - only its proportions matter to where it is drawn.
        private static Size PixelSizeOf(Bitmap source) => new(source.PixelSize.Width, source.PixelSize.Height);
    }
}
