using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Data;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;

namespace CRT
{
    // ###########################################################################################
    // THE COLOR DROPPER (owner request, 2026-11-03, "a general thing": "I really miss the 'color picker'
    // option, as it then would be easy to select the same color as something existing"). Avalonia's picker
    // has none, so one is put in its list (ColorPickerDropper). Pressed, the pointer is a cross over CRT's
    // whole window, a lens beside it says the color under it, and a click takes that color - any right or
    // middle click, Esc, the window losing the front or changing its size lets go. Bind Color two-way or
    // listen to Picked.
    //
    // *** ONLY CRT'S OWN WINDOW. *** It reads a picture of the window (RenderTargetBitmap) taken as it is
    // pressed, so it works alike on Windows, Linux and macOS - which offer no common way to read the rest
    // of the screen. An open flyout is a window of its own, so the picker's is not in the picture.
    //
    // In Controls/: it touches nothing of CRT's orchestration (SharedControlsIndependenceTests).
    // ###########################################################################################
    public sealed class ColorDropperButton : Button
    {
        public const string Tip = "Pick a color from anything in CRT's window - the next click takes the color under the pointer, Esc cancels";

        public const string LensHint = "Click to use it - Esc cancels";

        // Font Awesome's eye-dropper.
        public const string Icon = "\uF1FB";

        public static readonly StyledProperty<Color> ColorProperty =
            AvaloniaProperty.Register<ColorDropperButton, Color>(nameof(Color), defaultBindingMode: BindingMode.TwoWay);

        private Canvas? thisCover;
        private Border? thisLens;
        private Border? thisSwatch;
        private TextBlock? thisHex;
        private TopLevel? thisTop;
        private ScreenPixels? thisPixels;

        public ColorDropperButton()
        {
            this.Classes.Add("ColorDropper");
            this.Content = new TextBlock { Text = ColorDropperButton.Icon, Classes = { "ColorDropperIcon" } };
            ToolTip.SetTip(this, ColorDropperButton.Tip);
            this.Click += (_, _) => this.Begin();
        }

        // What it picks for, when it is not in that window itself: the color picker whose list it is in -
        // a flyout is a window of its own, with nothing in it to pick from.
        public Visual? Anchor { get; set; }

        // The flyout it is in - closed before the picking starts, as it would stand over the window.
        public Flyout? FlyoutToClose { get; set; }

        // A subclass, drawn as the Button it is.
        protected override Type StyleKeyOverride => typeof(Button);

        public Color Color
        {
            get => this.GetValue(ColorProperty);
            set => this.SetValue(ColorProperty, value);
        }

        // Told the color taken - after Color has it.
        public event EventHandler<Color>? Picked;

        // ###########################################################################################
        // A dropper is on: the keys are its own. Those handlers of CRT's that take a key whatever is
        // handled ask this (one dropper at a time - it covers the window), or the Esc that ends it would
        // end "Draw trace" under it too.
        // ###########################################################################################
        public static bool IsAnyPicking { get; private set; }

        // Stands in for the picture of the window - a test's display draws nothing to read.
        internal Func<ScreenPixels?>? PixelsOverrideForTests { get; set; }

        internal bool IsPicking => this.thisCover is not null;

        internal Control? CoverForTests => this.thisCover;

        internal string? LensTextForTests => this.thisHex?.Text;

        protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
        {
            this.End();
            base.OnDetachedFromVisualTree(e);
        }

        private void Begin()
        {
            if (this.FlyoutToClose is not { } flyout)
            {
                this.BeginPicking();
                return;
            }

            // Once it has gone: this button leaves the visual tree with it, and that ends a picking begun.
            flyout.Hide();
            Dispatcher.UIThread.Post(this.BeginPicking, DispatcherPriority.Background);
        }

        private void BeginPicking()
        {
            Visual anchor = this.Anchor ?? this;

            if (this.IsPicking || ColorDropperButton.IsAnyPicking || TopLevel.GetTopLevel(anchor) is not { } top || OverlayLayer.GetOverlayLayer(anchor) is not { } layer)
            {
                return;
            }

            ScreenPixels? pixels = this.PixelsOverrideForTests is { } standIn ? standIn() : ColorDropperButton.PictureOf(top);

            if (pixels is null)
            {
                return;
            }

            this.thisTop = top;
            this.thisPixels = pixels;

            this.thisSwatch = new Border { Width = 26, Height = 26, Classes = { "ColorDropperSwatch" } };
            this.thisHex = new TextBlock { FontWeight = FontWeight.SemiBold };

            this.thisLens = new Border
            {
                Classes = { "ColorDropperLens" },
                IsHitTestVisible = false,
                IsVisible = false,
                Child = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 8,
                    Children =
                    {
                        this.thisSwatch,
                        new StackPanel { Children = { this.thisHex, new TextBlock { Text = ColorDropperButton.LensHint, FontSize = 11 } } },
                    },
                },
            };

            // The window's size, not the layer's: a layer nothing has used yet is made as it is asked for,
            // and has no size until it is laid out - the cover had none, and no pointer ever found it.
            this.thisCover = new Canvas
            {
                Background = Brushes.Transparent,
                Cursor = new Cursor(StandardCursorType.Cross),
                Width = top.ClientSize.Width,
                Height = top.ClientSize.Height,
                Children = { this.thisLens },
            };

            this.thisCover.PointerMoved += this.OnCoverPointerMoved;
            this.thisCover.PointerPressed += this.OnCoverPointerPressed;

            layer.Children.Add(this.thisCover);

            anchor.DetachedFromVisualTree += this.OnAnchorDetached;

            ColorDropperButton.IsAnyPicking = true;
            top.AddHandler(KeyDownEvent, this.OnKeyWhilePicking, RoutingStrategies.Tunnel, handledEventsToo: true);
            top.SizeChanged += this.OnWindowResized;

            if (top is WindowBase window)
            {
                window.Deactivated += this.OnWindowDeactivated;
            }
        }

        private void End()
        {
            if (this.thisCover is not { } cover)
            {
                return;
            }

            cover.PointerMoved -= this.OnCoverPointerMoved;
            cover.PointerPressed -= this.OnCoverPointerPressed;
            (cover.Parent as Panel)?.Children.Remove(cover);

            (this.Anchor ?? this).DetachedFromVisualTree -= this.OnAnchorDetached;

            if (this.thisTop is { } top)
            {
                top.RemoveHandler(KeyDownEvent, this.OnKeyWhilePicking);
                top.SizeChanged -= this.OnWindowResized;

                if (top is WindowBase window)
                {
                    window.Deactivated -= this.OnWindowDeactivated;
                }
            }

            this.thisCover = null;
            this.thisLens = null;
            this.thisSwatch = null;
            this.thisHex = null;
            this.thisTop = null;
            this.thisPixels = null;

            ColorDropperButton.IsAnyPicking = false;
        }

        private void OnWindowDeactivated(object? sender, EventArgs e) => this.End();

        // The cover and the picture are of the window as it was pressed (code review, 2026-11-04): maximized
        // from its title bar, which does not deactivate it, what the cover no longer reached took clicks
        // while every key was still the dropper's, and what it did reach read colors from the old layout.
        private void OnWindowResized(object? sender, SizeChangedEventArgs e) => this.End();

        // What it picks for has gone - a picker's card made again, another board shown.
        private void OnAnchorDetached(object? sender, VisualTreeAttachmentEventArgs e) => this.End();

        // Every key is the dropper's while it is on - what it was pressed on must not take them.
        private void OnKeyWhilePicking(object? sender, KeyEventArgs e)
        {
            e.Handled = true;

            // Once this key's route has run: the handlers below the window, which take a key whatever is
            // handled, still have to see the dropper on (IsAnyPicking).
            if (e.Key == Key.Escape)
            {
                Dispatcher.UIThread.Post(this.End);
            }
        }

        private void OnCoverPointerMoved(object? sender, PointerEventArgs e)
        {
            if (this.thisCover is not { } cover || this.thisLens is not { } lens || this.thisPixels is not { } pixels)
            {
                return;
            }

            Point point = e.GetPosition(cover);

            if (pixels.ColorAt(point) is not { } color)
            {
                lens.IsVisible = false;
                return;
            }

            this.thisSwatch!.Background = new SolidColorBrush(color);
            this.thisHex!.Text = ScreenPixels.HexOf(color);
            lens.IsVisible = true;

            // Its own size once measured - a first guess of the size it has with its words, until then.
            Size size = lens.Bounds.Width > 0 ? lens.Bounds.Size : new Size(220, 46);
            Point place = ScreenPixels.LensPlace(point, size, cover.Bounds.Size);

            Canvas.SetLeft(lens, place.X);
            Canvas.SetTop(lens, place.Y);
        }

        private void OnCoverPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            e.Handled = true;

            Color? color = this.thisCover is { } cover && e.GetCurrentPoint(cover).Properties.IsLeftButtonPressed
                ? this.thisPixels?.ColorAt(e.GetPosition(cover))
                : null;

            this.End();

            if (color is { } picked)
            {
                this.SetCurrentValue(ColorProperty, picked);
                this.Picked?.Invoke(this, picked);
            }
        }

        // ###########################################################################################
        // CRT's window as a picture - on the UI thread, as a press: measured at 5 ms for 1920x1080.
        // ###########################################################################################
        private static ScreenPixels? PictureOf(TopLevel top)
        {
            try
            {
                double scale = top.RenderScaling;
                PixelSize size = PixelSize.FromSize(top.ClientSize, scale);

                if (size.Width <= 0 || size.Height <= 0)
                {
                    return null;
                }

                using var target = new RenderTargetBitmap(size, new Vector(96 * scale, 96 * scale));
                target.Render(top);

                return ScreenPixels.Of(target, scale);
            }
            catch (Exception ex)
            {
                // A picture that cannot be taken is a button that does nothing - never a CRT that falls over.
                CrtLog.Warning($"The color dropper could not read the window - [{ex.Message}]");
                return null;
            }
        }
    }
}
