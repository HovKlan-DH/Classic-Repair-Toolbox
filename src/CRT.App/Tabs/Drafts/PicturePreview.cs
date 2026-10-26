using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Handlers.Geometry;
using System.Collections.Generic;

namespace CRT
{
    // ###########################################################################################
    // ONE OF THE BOARD'S PICTURES WITH ITS MARKED COMPONENTS, DRAWN IN THE COLOURS PICKED (CRT 4.0
    // step 2, slice 4b - the plan: "with a live preview on the picture"). The component pane's
    // Pictures part hands it the picture made small, the picture's own pixel size, its highlights in
    // those pixels, and the colour and opacity - and it draws again as they change. Every position is
    // PicturePreviewGeometry's, so this only paints.
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

        static PicturePreview()
        {
            Control.AffectsRender<PicturePreview>(SourceProperty, ImageSizeProperty, HighlightsProperty, HighlightColorProperty, HighlightOpacityProperty);
        }

        public Bitmap? Source { get => this.GetValue(SourceProperty); set => this.SetValue(SourceProperty, value); }

        // The picture's own pixel size - the highlights are in those pixels, the Source is smaller.
        public Size ImageSize { get => this.GetValue(ImageSizeProperty); set => this.SetValue(ImageSizeProperty, value); }

        public IReadOnlyList<Rect>? Highlights { get => this.GetValue(HighlightsProperty); set => this.SetValue(HighlightsProperty, value); }

        public Color HighlightColor { get => this.GetValue(HighlightColorProperty); set => this.SetValue(HighlightColorProperty, value); }

        // 0..1, as the Schematics tab draws it.
        public double HighlightOpacity { get => this.GetValue(HighlightOpacityProperty); set => this.SetValue(HighlightOpacityProperty, value); }

        public override void Render(DrawingContext context)
        {
            if (this.Source is not { } bitmap)
            {
                return;
            }

            // The highlights are in the picture's OWN pixels: with its size not known, none is drawn -
            // placed against the size it was decoded to, they came out far too big (code review,
            // 2026-10-26). ImagePixelSize reads every type CRT shows, so this is a header it could not read.
            bool sizeKnown = this.ImageSize.Width > 0 && this.ImageSize.Height > 0;
            Size image = sizeKnown ? this.ImageSize : bitmap.Size;
            Rect target = PicturePreviewGeometry.Fit(image, this.Bounds.Size);

            if (target.Width <= 0)
            {
                return;
            }

            context.DrawImage(bitmap, new Rect(bitmap.Size), target);

            if (!sizeKnown)
            {
                return;
            }

            var fill = new SolidColorBrush(this.HighlightColor, this.HighlightOpacity);
            var outline = new Pen(new SolidColorBrush(this.HighlightColor), 1);

            foreach (Rect highlight in this.Highlights ?? [])
            {
                context.DrawRectangle(fill, outline, PicturePreviewGeometry.Place(highlight, image, target));
            }
        }
    }
}
