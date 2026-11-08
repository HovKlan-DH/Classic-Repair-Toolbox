using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;

namespace CRT
{
    // ###########################################################################################
    // THE PINOUT IN A WINDOW OF ITS OWN (owner request, 2026-11-07: 'a new checkbox option for "Detach pinout image in
    // component info popup". If this is enabled, then do show the pinout and its information in another popup window and
    // do make sure to save size and location etc. (one global, and not per component/board)').
    //
    // The popup's own pinout section - the case drawn and every pin beside it, its clicks going to the pin's images as
    // they do in the popup - is MOVED into the window while "Detach pinout image in component info popup" is ticked and
    // the component has pins, and put back where it stood otherwise: one section, never a copy kept in step. The window
    // goes with its popup, shows the popup's title, and takes no focus as it opens. Closed by the person, the pinout comes
    // back into that popup for as long as it is open. Its size and place are one for every popup (UserSettings
    // .SavePinoutWindowLayout); a place no screen shows any more is given up for one beside the popup.
    //
    // A popup that closes as it loses the front (CloseOnDeactivate) stays while its pinout window has the front.
    //
    // Part of the ComponentInfoWindow partial class - see ComponentInfoWindow.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentInfoWindow
    {
        private ComponentPinoutWindow? thisPinoutWindow;

        // Where the section stands in the popup, to go back to.
        private Panel? thisPinoutHome;

        private int thisPinoutHomeIndex;

        // Closed by its person: the pinout stays in this popup while it is open.
        private bool thisPinoutWindowDismissed;

        internal ComponentPinoutWindow? PinoutWindowForTests => this.thisPinoutWindow;

        // Following the setting from the popup's opening to its closing - a popup never shown holds on to nothing.
        private void WirePinoutWindow() => this.Opened += (_, _) =>
        {
            UserSettings.DetachPinoutChanged -= this.OnDetachPinoutChanged;
            UserSettings.DetachPinoutChanged += this.OnDetachPinoutChanged;
            this.FollowPinoutDetachment();
        };

        // The setting changed while the popup is open - asked for, so a window closed before comes back.
        private void OnDetachPinoutChanged() => Dispatcher.UIThread.Post(() =>
        {
            this.thisPinoutWindowDismissed = false;
            this.FollowPinoutDetachment();
        });

        private void LetGoOfPinoutWindow()
        {
            UserSettings.DetachPinoutChanged -= this.OnDetachPinoutChanged;
            this.TakePinoutBack(closeWindow: true);
        }

        // ###########################################################################################
        // The pinout where it belongs now - in its own window while the setting is on, the component has pins and the
        // popup is open, in the popup otherwise. Its window follows the popup's title.
        // ###########################################################################################
        private void FollowPinoutDetachment()
        {
            bool detached = UserSettings.DetachPinoutInComponentPopup &&
                            !this.thisPinoutWindowDismissed &&
                            this._pinRows.Count > 0 &&
                            this.IsVisible;

            if (!detached)
            {
                this.TakePinoutBack(closeWindow: true);
                return;
            }

            if (this.thisPinoutWindow is null)
            {
                // Its place in the popup, found once: on its way between the windows it stands in neither.
                if (this.thisPinoutHome is null)
                {
                    if (this.PinoutSection.Parent is not Panel home)
                    {
                        return;
                    }

                    this.thisPinoutHome = home;
                    this.thisPinoutHomeIndex = home.Children.IndexOf(this.PinoutSection);
                }

                TopLevel? from = TopLevel.GetTopLevel(this.PinoutSection);
                this.thisPinoutHome.Children.Remove(this.PinoutSection);

                var window = new ComponentPinoutWindow(this);
                window.Closed += (_, _) =>
                {
                    // Closed by its person (CRT closing it lets go of it first).
                    if (ReferenceEquals(this.thisPinoutWindow, window))
                    {
                        this.thisPinoutWindowDismissed = true;
                        this.TakePinoutBack(closeWindow: false);
                    }
                };
                window.Deactivated += (_, _) => this.CloseUnlessPinoutWindowTakesOver();

                this.thisPinoutWindow = window;
                window.Show(this);
                this.PlaceSectionAfterLayoutOf(from);
            }

            this.thisPinoutWindow.Title = ComponentPinoutWindow.TitleFor(this.TitleText.Text);
        }

        // The section back in the popup, where it stood; its window closed when CRT closes it.
        private void TakePinoutBack(bool closeWindow)
        {
            if (this.thisPinoutWindow is not { } window)
            {
                return;
            }

            TopLevel? from = TopLevel.GetTopLevel(this.PinoutSection);
            this.thisPinoutWindow = null;
            window.LetGoOfSection();

            if (closeWindow)
            {
                window.Close();
            }

            this.PlaceSectionAfterLayoutOf(from);
        }

        // ###########################################################################################
        // *** TAKEN OUT AT ONCE, PUT IN ONLY ONCE THE WINDOW IT LEFT HAS HAD ITS LAYOUT (owner report, 2026-11-07: a
        // resistor clicked with the pinout detached, and CRT ended on "Attempt to call InvalidateArrange on wrong
        // LayoutManager"). *** Avalonia keeps a control queued for the layout of the window it stood in after it is taken
        // out (LayoutManager's _toMeasure/_toArrange), and arranging it there once it stands in ANOTHER window throws - a
        // move made in the middle of that window's own layout (a size changed, the pin list going beside or under) or before
        // it had run. Taken out, a control drops out of the queue at that window's next pass (Arrange: NotVisible), so the
        // window it left runs its layout first - never inside a pass, as this is a job of its own - and the section is put
        // where it belongs THEN, its window or the popup; a later move decides alone, after the layout of every window the
        // section left since it last stood somewhere.
        // ###########################################################################################
        private int thisSectionMoves;

        private readonly List<TopLevel> thisWindowsSectionLeft = [];

        private void PlaceSectionAfterLayoutOf(TopLevel? from)
        {
            if (from is not null && !this.thisWindowsSectionLeft.Contains(from))
            {
                this.thisWindowsSectionLeft.Add(from);
            }

            int move = ++this.thisSectionMoves;

            Dispatcher.UIThread.Post(() =>
            {
                if (move != this.thisSectionMoves || this.PinoutSection.Parent is not null)
                {
                    return;
                }

                foreach (TopLevel left in this.thisWindowsSectionLeft)
                {
                    left.UpdateLayout();
                }

                this.thisWindowsSectionLeft.Clear();

                if (this.thisPinoutWindow is { } window)
                {
                    window.TakeSection(this.PinoutSection);
                }
                else if (this.thisPinoutHome is { } home)
                {
                    home.Children.Insert(Math.Clamp(this.thisPinoutHomeIndex, 0, home.Children.Count), this.PinoutSection);
                }
            }, DispatcherPriority.Background);
        }

        // ###########################################################################################
        // The popup or its pinout window lost the front: with a pinout window open the popup closes - as it does by itself
        // when CloseOnDeactivate - only once neither has it, decided a moment later, when the other has taken it. True when
        // that is what happens, so the caller closes nothing at once.
        // ###########################################################################################
        private bool CloseUnlessPinoutWindowTakesOver()
        {
            if (this.thisPinoutWindow is null)
            {
                return false;
            }

            Dispatcher.UIThread.Post(() =>
            {
                if (!this.CloseOnDeactivate || !this.IsVisible || this.IsActive || this.thisPinoutWindow?.IsActive == true)
                {
                    return;
                }

                if (this.Owner is Main mainOwner && mainOwner.isHoveringComponent)
                {
                    return;
                }

                this.Close();
            }, DispatcherPriority.Background);

            return true;
        }
    }

    // ###########################################################################################
    // The window the popup's pinout stands in while detached (ComponentInfoWindow.Pinout.cs) - its size and place one for
    // all, saved as it closes.
    // ###########################################################################################
    internal sealed class ComponentPinoutWindow : Window
    {
        private readonly Border thisHost;

        private double thisNormalWidth;
        private double thisNormalHeight;
        private PixelPoint thisNormalPosition;

        // Its size the first time (owner request, 2026-11-07: "have the default size be aprox 870x960px"), no larger than
        // the screen it opens on.
        internal const double DefaultWidth = 870;
        internal const double DefaultHeight = 960;

        // Made empty: the section comes in once the popup has had its layout (TakeSection).
        public ComponentPinoutWindow(Window popup)
        {
            this.thisHost = new Border { Padding = new Thickness(12) };
            this.Content = new ScrollViewer { HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled, Content = this.thisHost };

            this.ShowInTaskbar = false;
            this.ShowActivated = false;
            this.MinWidth = 260;
            this.MinHeight = 200;
            this.WindowStartupLocation = WindowStartupLocation.Manual;
            this.Bind(BackgroundProperty, this.GetResourceObservable("Bg"));
            this.Bind(ForegroundProperty, this.GetResourceObservable("Fg"));

            this.Width = UserSettings.HasPinoutWindowLayout ? UserSettings.PinoutWindowWidth : ComponentPinoutWindow.DefaultWidth;
            this.Height = UserSettings.HasPinoutWindowLayout ? UserSettings.PinoutWindowHeight : ComponentPinoutWindow.DefaultHeight;
            this.Position = UserSettings.HasPinoutWindowLayout
                ? new PixelPoint(UserSettings.PinoutWindowX, UserSettings.PinoutWindowY)
                : ComponentPinoutWindow.BesidePopup(popup);

            this.thisNormalWidth = this.Width;
            this.thisNormalHeight = this.Height;
            this.thisNormalPosition = this.Position;

            // A place no screen shows any more - a screen taken away - is given up for one beside the popup.
            this.Opened += (_, _) =>
            {
                if (this.Screens.ScreenFromPoint(this.Position) is null)
                {
                    this.Position = ComponentPinoutWindow.BesidePopup(popup);
                }

                // 960 high is more than a laptop's screen has: kept within the screen it opened on.
                if ((this.Screens.ScreenFromPoint(this.Position) ?? this.Screens.Primary) is { } screen && screen.Scaling > 0)
                {
                    double screenWidth = screen.WorkingArea.Width / screen.Scaling;
                    double screenHeight = screen.WorkingArea.Height / screen.Scaling;

                    this.Width = Math.Min(this.Width, screenWidth);
                    this.Height = Math.Min(this.Height, screenHeight);
                }
            };

            this.SizeChanged += (_, _) =>
            {
                if (this.WindowState == WindowState.Normal)
                {
                    this.thisNormalWidth = this.Width;
                    this.thisNormalHeight = this.Height;
                }
            };

            this.PositionChanged += (_, _) =>
            {
                if (this.WindowState == WindowState.Normal)
                {
                    this.thisNormalPosition = this.Position;
                }
            };

            this.Closing += (_, _) => UserSettings.SavePinoutWindowLayout(this.thisNormalWidth, this.thisNormalHeight, this.thisNormalPosition.X, this.thisNormalPosition.Y);
        }

        // "Pinout - U8 | CPU | 6510": the popup's title after the word.
        internal static string TitleFor(string? popupTitle) =>
            string.IsNullOrWhiteSpace(popupTitle) ? "Pinout" : "Pinout - " + popupTitle.Trim();

        internal void TakeSection(Control section) => this.thisHost.Child = section;

        // The section handed back - the window keeps nothing of the popup's.
        internal void LetGoOfSection() => this.thisHost.Child = null;

        // Right of the popup, its top level with the popup's.
        private static PixelPoint BesidePopup(Window popup) =>
            new(popup.Position.X + (int)Math.Round(popup.Bounds.Width * popup.RenderScaling) + 8, popup.Position.Y);
    }
}
