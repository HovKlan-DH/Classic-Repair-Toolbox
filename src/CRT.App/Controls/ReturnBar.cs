using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

namespace CRT
{
    // ###########################################################################################
    // THE BAR THAT TAKES A PERSON BACK WHERE THEY CAME FROM (owner request, 2026-11-11: "This panel should be identical
    // everywhere, and please let us try and make this lightgrey background with 2px black dashed border, to give it some
    // attention" - and a white background in place of the lightgray, the same day). The Contribute tab sends a person to the Schematics tab - "Show on board", lighting a net, "Mark components
    // manually", "Align KiCad data" - and there a bar says what is shown and offers the way back: "Back to Contribute", or a
    // picture tool's Done and Cancel. Every such bar is this control, so they look alike.
    //
    // A Decorator drawing its own ground and frame: Avalonia's Border has no dashed edge, and only shapes carry a dash
    // array. Colors are App.axaml's ReturnBar_*, the same in both themes - the bar is meant to stand out from either.
    // ###########################################################################################
    public sealed class ReturnBar : Decorator
    {
        // The frame's thickness and its dash, in device-independent pixels.
        public const double FrameThickness = 2;

        public const double Rounding = 6;

        private static readonly double[] Dash = [3, 2];

        public static readonly StyledProperty<IBrush?> BackgroundProperty =
            AvaloniaProperty.Register<ReturnBar, IBrush?>(nameof(Background));

        public static readonly StyledProperty<IBrush?> FrameBrushProperty =
            AvaloniaProperty.Register<ReturnBar, IBrush?>(nameof(FrameBrush));

        static ReturnBar()
        {
            AffectsRender<ReturnBar>(BackgroundProperty, FrameBrushProperty);
        }

        public ReturnBar()
        {
            this.Padding = new Thickness(12, 8);
            this.Bind(BackgroundProperty, this.GetResourceObservable("ReturnBar_Bg"));
            this.Bind(FrameBrushProperty, this.GetResourceObservable("ReturnBar_Border"));
        }

        public IBrush? Background
        {
            get => this.GetValue(BackgroundProperty);
            set => this.SetValue(BackgroundProperty, value);
        }

        public IBrush? FrameBrush
        {
            get => this.GetValue(FrameBrushProperty);
            set => this.SetValue(FrameBrushProperty, value);
        }

        public override void Render(DrawingContext context)
        {
            // The frame drawn on the inside of the bounds, so none of it is clipped.
            double half = ReturnBar.FrameThickness / 2;
            Rect bounds = new Rect(this.Bounds.Size).Deflate(half);

            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                return;
            }

            var frame = new RoundedRect(bounds, ReturnBar.Rounding);
            Pen? pen = this.FrameBrush is { } brush ? new Pen(brush, ReturnBar.FrameThickness, new DashStyle(ReturnBar.Dash, 0)) : null;

            context.DrawRectangle(this.Background, pen, frame);
        }
    }
}
