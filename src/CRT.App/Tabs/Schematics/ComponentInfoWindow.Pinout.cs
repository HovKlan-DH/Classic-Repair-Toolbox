using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.Layout;
using Avalonia.VisualTree;
using System.Linq;
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
    // they do in the popup - is MOVED into the window while "Detach pinout image in component info popup" is ticked, the
    // component has pins and the popup has no room for them (below, PinoutFitsInPopup), and put back where it stood
    // otherwise: one section, never a copy kept in step. The window
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
        // *** DECIDED ONCE MAIN HAS DONE WITH THE POPUP (code review, 2026-11-09). *** After SetComponent, Main shows or hides
        // "Edit component" in the title row (OfferEdit) and the oscilloscope rows in the footer
        // (UpdateOscilloscopeSessionTitleState) - both change the info column's room, so a decision made inside SetComponent
        // could keep a pinout in a popup that then had to scroll for it. Posted at Normal, it runs once all of that is done
        // and before anything is laid out or drawn (Normal is above Render); several asks in a row are one decision.
        // ###########################################################################################
        private bool thisPinoutDecisionPosted;

        private void FollowPinoutDetachmentSoon()
        {
            if (this.thisPinoutDecisionPosted)
            {
                return;
            }

            this.thisPinoutDecisionPosted = true;

            Dispatcher.UIThread.Post(() =>
            {
                this.thisPinoutDecisionPosted = false;
                this.FollowPinoutDetachment();
            }, DispatcherPriority.Normal);
        }

        // ###########################################################################################
        // The pinout where it belongs now - in its own window while the setting is on, the component has pins, the popup
        // is open and has no room for them, in the popup otherwise. Its window follows the popup's title.
        // ###########################################################################################
        private void FollowPinoutDetachment()
        {
            bool wanted = UserSettings.DetachPinoutInComponentPopup &&
                          !this.thisPinoutWindowDismissed &&
                          this._pinRows.Count > 0 &&
                          this.IsVisible;

            // On its way between the windows it can be measured in neither: decided once it stands somewhere - the move's
            // own job, posted before this one at the same priority, puts it there first.
            if (wanted && this.PinoutSection.Parent is null)
            {
                Dispatcher.UIThread.Post(this.FollowPinoutDetachment, DispatcherPriority.Background);
                return;
            }

            bool detached = wanted && !this.PinoutFitsInPopup();

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

        // ###########################################################################################
        // *** ONLY A PINOUT THE POPUP HAS NO ROOM FOR GETS A WINDOW OF ITS OWN (owner request, 2026-11-09: "most of the pinout
        // windows does not have that much information inside, and that it could easily fit into the normal component info
        // popup ... if no vertical scrolling will be implied with this info inside of component info popup, then you will not
        // open a new window ... so the user can always see the pins when navigating in the oscilloscope baseline"). ***
        //
        // Measured before a window opens: the popup's other sections as they are now, the pinout as it would stand in the
        // popup at its width - the pin list beside the drawing or under it - and the TALLEST image note and baseline section
        // among the images the component shows, so the pins stay in sight whichever image is navigated to. Fitting the room
        // the popup's info column has, the pinout stays in the popup. Decided when a component or a region is shown, the
        // popup opens or the setting is ticked - so it holds while the images are gone through, and never as the popup is
        // resized, which would open and close a window under the pointer. Not laid out yet, it cannot be told: a window, as
        // before.
        // ###########################################################################################
        private bool PinoutFitsInPopup()
        {
            this.UpdateLayout();

            double room = this.InfoScrollViewer.Viewport.Height;
            double width = this.InfoSectionsPanel.Bounds.Width;

            if (room <= 0 || width <= 0)
            {
                return false;
            }

            // Each section shown adds its height and the panel's spacing; the first one has none before it.
            double spacing = this.InfoSectionsPanel.Spacing;
            double height = -spacing;

            foreach (Control section in this.InfoSectionsPanel.Children)
            {
                if (section.IsVisible &&
                    !ReferenceEquals(section, this.PinoutSection) &&
                    !ReferenceEquals(section, this.ImageNoteSection) &&
                    !ReferenceEquals(section, this.BaselineConditionsSection))
                {
                    height += section.DesiredSize.Height + spacing;
                }
            }

            height += this.PinoutHeightAt(width) + spacing;
            height += this.TallestImageSectionsAt(width, spacing);

            return height <= room + 0.5;
        }

        // The pinout's height standing in the popup at this width. Laid out there, as it is; in its own window, measured as it
        // would stand here - its pin list placed by this width - and then left to that window as it was.
        private double PinoutHeightAt(double width)
        {
            if (ReferenceEquals(this.PinoutSection.Parent, this.InfoSectionsPanel))
            {
                return this.PinoutSection.DesiredSize.Height;
            }

            this.PlacePinList(width);

            // Put back whatever happens (code review, 2026-11-09): left placed for the popup's width, the window's pin list
            // stood under the drawing until the grid next changed size.
            try
            {
                return ComponentInfoWindow.MeasuredAfresh(this.PinoutSection, width);
            }
            finally
            {
                this.PlacePinList(this.PinoutGrid.Bounds.Width);
                this.PinoutSection.InvalidateMeasure();
            }
        }

        // A section's height at this width, measured afresh: Avalonia keeps a control's own measure while only something
        // inside it changed - measured again at the same width, a note section with a new note still had its old height -
        // so every part of it is measured again.
        private static double MeasuredAfresh(Control section, double width)
        {
            foreach (Layoutable part in section.GetVisualDescendants().OfType<Layoutable>())
            {
                part.InvalidateMeasure();
            }

            section.InvalidateMeasure();
            section.Measure(new Size(width, double.PositiveInfinity));
            return section.DesiredSize.Height;
        }

        // ###########################################################################################
        // The most the image note and the baseline section take, with their spacing, on any image the component shows - what
        // the popup shows put back after. *** NOT EVERY IMAGE MEASURED *** (code review, 2026-11-09: each image's note and
        // machine state measured together - a 40-pin chip with forty baselines, each its own note and machine state, took
        // 135 ms more a click, measured): the note and the baseline section apart, each over its distinct wordings - a
        // component's images mostly share them - and of those only the MeasuredAtMost that look tallest, the most lines and
        // then the most characters, a shorter text in the same box seldom being taller. The two tallest added: more than one
        // image needs only when the longest note and the tallest baseline section are on different images, which only leaves
        // a window where the popup might just have held the pins.
        // ###########################################################################################
        private const int MeasuredAtMost = 3;

        private double TallestImageSectionsAt(double width, double spacing)
        {
            double tallestNote = 0;
            double tallestBaseline = 0;
            var shown = this.ThumbnailList.SelectedItem as ComponentImageItem;
            List<ComponentImageEntry> images = this.ImageEntriesShown();

            static int LinesOf(string? text) => text is null ? 0 : text.Count(character => character == '\n') + 1;

            List<string> notes = images
                .Select(image => image.Note ?? string.Empty)
                .Distinct(StringComparer.Ordinal)
                .OrderByDescending(LinesOf)
                .ThenByDescending(note => note.Length)
                .Take(MeasuredAtMost)
                .ToList();

            List<ComponentImageEntry> baselines = images
                .GroupBy(image => (ComponentPopupContent.BaselineConditions(image), ComponentImageKinds.IsScopeBaseline(image)))
                .Select(group => group.First())
                .OrderByDescending(image => ComponentImageKinds.IsScopeBaseline(image))
                .ThenByDescending(image => ComponentPopupContent.BaselineConditions(image)?.Length ?? 0)
                .Take(MeasuredAtMost)
                .ToList();

            double HeightOf(Control section) => section.IsVisible ? ComponentInfoWindow.MeasuredAfresh(section, width) + spacing : 0;

            try
            {
                // The note alone: no image behind it, so no baseline section.
                foreach (string note in notes)
                {
                    this.UpdateImageNote(new ComponentImageItem { Note = note });
                    tallestNote = Math.Max(tallestNote, HeightOf(this.ImageNoteSection));
                }

                // The baseline section alone: no note.
                foreach (ComponentImageEntry baseline in baselines)
                {
                    this.UpdateImageNote(new ComponentImageItem { SourceEntry = baseline });
                    tallestBaseline = Math.Max(tallestBaseline, HeightOf(this.BaselineConditionsSection));
                }
            }
            finally
            {
                this.UpdateImageNote(shown);
            }

            return tallestNote + tallestBaseline;
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
