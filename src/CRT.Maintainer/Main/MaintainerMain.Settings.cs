using System;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using CRT.Maintainer.Handlers;

namespace CRT.Maintainer
{
    // ###########################################################################################
    // WHAT THE WINDOW REMEMBERS BETWEEN RUNS (owner requests, 2026-09-26): its place - "if it is
    // maximized, then start again in maximized, and if it is windowed, then start same position and
    // size next time" - and "Show changes only". MaintainerSettings holds them; this applies them
    // before the window is shown, and saves them when it closes.
    //
    // *** THE NORMAL POSITION AND SIZE ARE TRACKED WHILE THE WINDOW IS NORMAL. *** A maximized
    // window reports the maximized ones, so they are recorded only in the Normal state - which is
    // what un-maximizing next time returns to. A maximized window is placed on its saved screen
    // first and then maximized, or the system maximizes it on whichever monitor it picks - CRT's
    // own main window does the same, for the same reason.
    //
    // Only the running application calls UseSettings (MaintainerApp), so a test that builds this
    // window never reads or writes the user's real settings file.
    // ###########################################################################################
    public partial class MaintainerMain
    {
        private MaintainerSettings? thisSettings;
        private Action<MaintainerSettings>? thisSaveSettings;

        private PixelPoint thisNormalPosition;
        private Size thisNormalSize;
        private bool thisHasNormalBounds;

        // Set while UseSettings is placing the window, so the placement's own PositionChanged and
        // ClientSize changes are not mistaken for the user moving or resizing it.
        private bool thisRestoring;

        // ###########################################################################################
        // Applies the remembered settings - before the window is shown, so it opens where it was
        // rather than jumping there - and keeps `save` for when it closes.
        // ###########################################################################################
        internal void UseSettings(MaintainerSettings settings, Action<MaintainerSettings> save)
        {
            ArgumentNullException.ThrowIfNull(settings);
            ArgumentNullException.ThrowIfNull(save);

            this.thisSettings = settings;
            this.thisSaveSettings = save;

            this.TableEditor.OnlyChanges = settings.ShowChangesOnly;

            // ###########################################################################################
            // *** THE RESTORE'S OWN MOVES ARE NOT THE USER'S (code review, 2026-09-26). *** A
            // window restored maximized is first placed on its saved screen and maximized after -
            // and that placement raised PositionChanged while the state was still Normal, so the
            // synthetic "screen corner + 100" point overwrote the real remembered position and was
            // saved back on close. A few maximized-only runs later, un-maximizing landed in the
            // corner instead of where the window last was.
            //
            // thisRestoring is set across the placement below and cleared once the window is up
            // (SettleRestoredPlacement), after which every move is the user's own.
            // ###########################################################################################
            this.PositionChanged += (_, e) =>
            {
                if (this.WindowState == WindowState.Normal && !this.thisRestoring)
                    this.thisNormalPosition = e.Point;
            };

            if (!settings.HasWindowPlacement)
                return;

            this.WindowStartupLocation = WindowStartupLocation.Manual;
            this.Width = Math.Max(this.MinWidth, settings.WindowWidth);
            this.Height = Math.Max(this.MinHeight, settings.WindowHeight);

            this.thisNormalPosition = new PixelPoint(settings.WindowX, settings.WindowY);
            this.thisNormalSize = new Size(this.Width, this.Height);
            this.thisHasNormalBounds = true;

            this.thisRestoring = true;

            if (settings.WindowMaximized)
            {
                // Anywhere on the saved screen, so it maximizes THERE. The remembered NORMAL
                // position is kept as it is - see the PositionChanged note above.
                this.Position = new PixelPoint(settings.ScreenX + 100, settings.ScreenY + 100);
                this.WindowState = WindowState.Maximized;
            }
            else
            {
                this.Position = this.thisNormalPosition;
            }
        }

        // ###########################################################################################
        // Once shown: a maximized window is maximized again (some Linux window managers ignore it
        // before the window is on screen), and a normal one whose remembered place is on no screen
        // any more - a monitor unplugged - is centred on the main screen instead.
        // ###########################################################################################
        private void SettleRestoredPlacement()
        {
            if (this.thisSettings is not { HasWindowPlacement: true } settings)
            {
                this.thisRestoring = false;
                return;
            }

            if (settings.WindowMaximized)
            {
                this.WindowState = WindowState.Maximized;

                // The window is up and maximized; every move from here is the user's own.
                this.thisRestoring = false;
                return;
            }

            this.thisRestoring = false;

            bool visible = WindowPlacementRules.IsCentreOnAScreen(
                this.thisNormalPosition.X,
                this.thisNormalPosition.Y,
                this.thisNormalSize.Width,
                this.thisNormalSize.Height,
                this.RenderScaling,
                this.Screens.All.Select(screen => (screen.Bounds.X, screen.Bounds.Y, screen.Bounds.Width, screen.Bounds.Height)));

            if (!visible && this.Screens.Primary is { } primary)
            {
                PixelRect area = primary.WorkingArea;
                double scaling = primary.Scaling > 0 ? primary.Scaling : 1;

                this.Position = new PixelPoint(
                    area.X + Math.Max(0, (area.Width - (int)(this.Width * scaling)) / 2),
                    area.Y + Math.Max(0, (area.Height - (int)(this.Height * scaling)) / 2));
            }
        }

        protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
        {
            base.OnPropertyChanged(change);

            if (change.Property == Window.ClientSizeProperty && this.WindowState == WindowState.Normal && this.IsVisible)
            {
                this.thisNormalSize = this.ClientSize;
                this.thisHasNormalBounds = true;
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            this.SaveSettings();
        }

        // What is saved: the normal place (the current one, when the window closes normal), whether
        // it was maximized and on which screen, and the "Show changes only" choice.
        private void SaveSettings()
        {
            if (this.thisSettings is null || this.thisSaveSettings is null)
                return;

            this.thisSaveSettings(this.CurrentSettings());
        }

        internal MaintainerSettings CurrentSettings()
        {
            MaintainerSettings settings = this.thisSettings ?? new MaintainerSettings();

            if (this.WindowState == WindowState.Normal && this.IsVisible)
            {
                this.thisNormalPosition = this.Position;
                this.thisNormalSize = this.ClientSize;
                this.thisHasNormalBounds = true;
            }

            settings.HasWindowPlacement = this.thisHasNormalBounds;
            settings.WindowMaximized = this.WindowState is WindowState.Maximized or WindowState.FullScreen;
            settings.WindowX = this.thisNormalPosition.X;
            settings.WindowY = this.thisNormalPosition.Y;
            settings.WindowWidth = this.thisNormalSize.Width;
            settings.WindowHeight = this.thisNormalSize.Height;

            if (this.IsVisible && this.Screens.ScreenFromWindow(this) is { } screen)
            {
                settings.ScreenX = screen.Bounds.X;
                settings.ScreenY = screen.Bounds.Y;
            }

            settings.ShowChangesOnly = this.TableEditor.OnlyChangesWanted;

            return settings;
        }
    }
}
