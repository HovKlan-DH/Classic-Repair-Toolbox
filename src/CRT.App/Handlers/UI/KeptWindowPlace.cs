using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Platform;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace CRT
{
    // ###########################################################################################
    // A WINDOW THAT OPENS WHERE IT WAS LEFT - "Capture from oscilloscope pin-by-pin" (2026-11-20) and a collection's window
    // (owner request, 2026-11-28: "do save its size and position, so it will open same size and place the next time the user opens
    // it"). Its Normal size and position followed as it is moved and sized (never while maximized, so a window maximized and
    // closed opens maximized over the size it had); a place kept used as it opens, FITTED to the screen it opens on - a smaller
    // screen than where it was left kept its buttons off the bottom (code review, 2026-11-21) - and a place no screen shows any
    // more given up, the window opening over CRT's as the first time. Avalonia-bound, so no Handlers/Geometry class; the
    // arithmetic is DialogPlacement.KeptPlaceOn's, tested there.
    // ###########################################################################################
    internal sealed class KeptWindowPlace
    {
        private readonly Window thisWindow;

        private double thisNormalWidth;

        private double thisNormalHeight;

        private PixelPoint thisNormalPosition;

        public KeptWindowPlace(Window window)
        {
            ArgumentNullException.ThrowIfNull(window);

            this.thisWindow = window;
            this.thisNormalWidth = window.Width;
            this.thisNormalHeight = window.Height;

            window.SizeChanged += (_, _) =>
            {
                if (window.WindowState == WindowState.Normal && window.ClientSize is { Width: > 0, Height: > 0 } size)
                {
                    this.thisNormalWidth = size.Width;
                    this.thisNormalHeight = size.Height;
                }
            };

            window.PositionChanged += (_, _) =>
            {
                if (window.WindowState == WindowState.Normal)
                    this.thisNormalPosition = window.Position;
            };
        }

        // ###########################################################################################
        // No place kept: its own size, fitted to the screen of CRT's window it opens over.
        // ###########################################################################################
        public void FitTo((PixelRect WorkingArea, double Scaling) screen)
        {
            Window window = this.thisWindow;

            (Size size, _) = DialogPlacement.KeptPlaceOn(
                new Size(window.Width, window.Height), new Size(window.MinWidth, window.MinHeight), default, screen.WorkingArea, screen.Scaling);

            window.Width = size.Width;
            window.Height = size.Height;
            this.thisNormalWidth = window.Width;
            this.thisNormalHeight = window.Height;
        }

        // ###########################################################################################
        // The place kept, before the window is shown: its size fitted to the screen the place is on (else `ownerScreen`, the screen
        // of CRT's window), there - unless no screen shows that place - and maximized when it was left so.
        // ###########################################################################################
        public void Use(WindowPlace place, (PixelRect WorkingArea, double Scaling)? ownerScreen)
        {
            ArgumentNullException.ThrowIfNull(place);

            Window window = this.thisWindow;
            var saved = new PixelPoint(place.X, place.Y);
            Screen? onScreen = window.Screens.ScreenFromPoint(saved);

            (Size size, PixelPoint position) = DialogPlacement.KeptPlaceOn(
                new Size(place.Width, place.Height),
                new Size(window.MinWidth, window.MinHeight),
                saved,
                onScreen?.WorkingArea ?? ownerScreen?.WorkingArea ?? default,
                onScreen?.Scaling ?? ownerScreen?.Scaling ?? 1.0);

            window.Width = size.Width;
            window.Height = size.Height;
            this.thisNormalWidth = window.Width;
            this.thisNormalHeight = window.Height;

            if (onScreen is not null)
            {
                window.WindowStartupLocation = WindowStartupLocation.Manual;
                window.Position = position;
                this.thisNormalPosition = position;
            }

            if (string.Equals(place.State, "Maximized", StringComparison.Ordinal))
                window.WindowState = WindowState.Maximized;
        }

        // ###########################################################################################
        // Where it is left - its Normal size and position as they are now, or, maximized, as they were before.
        // ###########################################################################################
        public WindowPlace Current()
        {
            Window window = this.thisWindow;

            if (window.WindowState == WindowState.Normal && window.ClientSize is { Width: > 0, Height: > 0 } size)
            {
                this.thisNormalWidth = size.Width;
                this.thisNormalHeight = size.Height;
                this.thisNormalPosition = window.Position;
            }

            return new WindowPlace(
                window.WindowState == WindowState.Maximized ? "Maximized" : "Normal",
                this.thisNormalWidth,
                this.thisNormalHeight,
                this.thisNormalPosition.X,
                this.thisNormalPosition.Y);
        }
    }
}
