using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // The component filter/search list, the component info popup(s) it opens, the "blink
    // selected" highlight timer, the region (PAL/NTSC) toggle, and what the Contribute tab's
    // component pane asks of them: "Show on board" and "Preview popup" (CRT 4.0 step 2, slice 3c -
    // the component editor's window they replace is gone). See Main.axaml.cs for the file map of the
    // whole partial class.
    // ###########################################################################################
    public partial class Main
    {
        // ###########################################################################################
        // Clears the component search text, removes all selected items in the Component filter list,
        // and resets schematic highlights back to an empty selection state.
        // ###########################################################################################
        private void OnClearComponentsClick(object? sender, RoutedEventArgs e)
        {
            if (this._contributeListsShown)
            {
                this.ClearContributeLists();
                return;
            }

            this.ShowComponentSearchText(string.Empty);

            if (this.TabSchematicsControl.IsLabelEditorActive)
            {
                this.TabSchematicsControl.ApplyLabelEditorSearchFilter(string.Empty);
            }
            else
            {
                this.ApplyNormalComponentSearchFilter(string.Empty);
            }

            this.TabSchematicsControl.ClearSchematicsOnlySelectedComponents();
            this.ComponentFilterListBox.SelectedItems?.Clear();
            this.TabSchematicsControl.UpdateHighlightsForComponents(new List<string>());
        }

        // ###########################################################################################
        // Selects all available items currently populated within the Component filter list box - none
        // while a board is being read, when the rows on screen are about to go (code review, 2026-11-05).
        // ###########################################################################################
        private void OnMarkAllComponentsClick(object? sender, RoutedEventArgs e)
        {
            if (this._currentBoardData is null)
            {
                return;
            }

            this._markingAll = true;

            try
            {
                this.ComponentFilterListBox.SelectAll();
            }
            catch (OutOfMemoryException ex)
            {
                Logger.Debug(ex, "Failed to mark all components. The selection was too large to process");
            }
            finally
            {
                this._markingAll = false;
            }

            this._markAllPressed = true;
        }

        // ###########################################################################################
        // Whether the selection is "Mark all"'s - the button pressed, nothing chosen since (OnComponentFilterSelectionChanged,
        // ApplyNormalComponentSearchFilter). What the Contribute tab shows and never opens (IsEveryComponentRowSelected; code
        // review, 2026-11-09: told by every row being selected, the filter box's own selecting all it finds, or Ctrl on both of
        // two rows, closed the tab's pane and spoke of a "Mark all" nobody pressed).
        // ###########################################################################################
        private bool _markAllPressed;

        private bool _markingAll;

        // ###########################################################################################
        // Switches the local region to PAL and reloads images.
        // ###########################################################################################
        private void OnPalRegionClick(object? sender, RoutedEventArgs e)
        {
            if (this._suppressRegionToggle)
                return;

            this.PickRegion(BoardRegions.Pal);
        }

        // ###########################################################################################
        // Switches the local region to NTSC and reloads images.
        // ###########################################################################################
        private void OnNtscRegionClick(object? sender, RoutedEventArgs e)
        {
            if (this._suppressRegionToggle)
                return;

            this.PickRegion(BoardRegions.Ntsc);
        }

        // ###########################################################################################
        // The region CRT shows, from the buttons under the component list - the Contribute tab's too.
        // ###########################################################################################
        internal void PickRegion(string region)
        {
            this.RememberRegionPicked(region);
            this.ShowRegion(region);
        }

        // ###########################################################################################
        // *** THE REGION SHOWN IS ONE OF THE BOARD'S, WHATEVER CHANGES ITS LIST (code review, 2026-11-09). *** "Board: Regions"
        // deleting the region CRT showed left every button unlit and the list built for a region the board no longer had -
        // a component told apart per region gone from it - until a button was pressed or the board read again. Asked as the
        // Contribute tab's lists are read again (OnContributeListsShown), which every change of its draft does, and as the tab
        // is entered or left (third code review, 2026-11-09: the draft the tab shows and the board on screen can come in other
        // regions - "view boards as officially published"): a region no longer the board's - the one the buttons read
        // (UpdateRegionButtonsState) - gives way to the one a board load would show (RegionChoice.ToShow), nobody having
        // picked it, so nothing is remembered. True when it did - the lists were read again for it.
        // ###########################################################################################
        private bool FollowRegionsOfBoardShown()
        {
            BoardData? board = this._contributeListsShown ? this.TabDrafts.BoardForList() ?? this._currentBoardData : this._currentBoardData;

            if (board is null)
            {
                return false;
            }

            IReadOnlyList<string> regions = BoardRegions.Of(board);

            if (regions.Count == 0 || regions.Contains(this._localRegion, StringComparer.OrdinalIgnoreCase))
            {
                return false;
            }

            this.ShowRegion(RegionChoice.ToShow(regions, UserSettings.Region, UserSettings.RegionOfBoard(this.GetCurrentBoardKey())));
            return true;
        }

        // The region CRT shows, everywhere - the buttons, the pictures, the Contribute tab, the "Overview" tab (each component
        // as that region has it - OverviewHtmlBuilder.BuildRows) and the "Resources" tab.
        private void ShowRegion(string region)
        {
            this._localRegion = region;

            this.UpdateRegionButtonsState();
            this.RefreshImages();
            this.TabSchematicsControl.UpdateOverlayLabels();
            this.TabDrafts.UseRegion(region);
            this.TabOverview.ShowRegion(region);

            if (this._currentBoardData is { } board && this.GetCurrentBoardEntry() is { } entry)
            {
                this.ShowResources(board, entry, DataManager.DataRoot, DraftManager.GetViewedBoardFolder(entry.ExcelDataFile));
            }
        }

        // ###########################################################################################
        // Updates the region toggle and button states to match the current local region.
        // Hides the entire region toggle area when the current board has no regions (BoardRegions).
        // ###########################################################################################
        private void UpdateRegionButtonsState()
        {
            this._suppressRegionToggle = true;

            // On the Contribute tab, by what it reads: the board as the draft now has it - its colors too.
            BoardData? board = this.BoardOfRegionButtons();
            IReadOnlyList<string> regions = BoardRegions.Of(board);

            this.RegionButtonsGrid.IsVisible = regions.Count > 0;

            this.LayOutRegionButtons(regions, board);

            foreach (Button button in this._ownRegionButtons.Prepend(this.NtscRegionButton).Prepend(this.PalRegionButton))
            {
                button.Classes.Set("active", string.Equals(button.Tag as string, this._localRegion, StringComparison.OrdinalIgnoreCase));
            }

            this._suppressRegionToggle = false;
        }

        // ###########################################################################################
        // What was picked is remembered: PAL or NTSC as the one region picked anywhere (UserSettings.Region, which CRT 3.x
        // reads too), a region of the board's own on that board alone (2026-11-09).
        // ###########################################################################################
        private void RememberRegionPicked(string region)
        {
            string boardKey = this.GetCurrentBoardKey();

            if (BoardRegions.IsPalOrNtsc(region))
            {
                UserSettings.Region = region.Trim().ToUpperInvariant();

                if (boardKey.Length > 0)
                {
                    UserSettings.SetRegionOfBoard(boardKey, null);
                }
            }
            else if (boardKey.Length > 0)
            {
                UserSettings.SetRegionOfBoard(boardKey, region);
            }
        }

        // The buttons of a board's own regions, made as a board needs them and kept for the next.
        private readonly List<Button> _ownRegionButtons = new();

        // ###########################################################################################
        // THE REGION BUTTONS, ONE PER REGION OF THE BOARD, IN ITS ORDER (2026-11-09: a board may come in regions of its own):
        // "PAL" and "NTSC" are the buttons they always were - a board in PAL and NTSC looks as it always did, two buttons
        // side by side - and a region of the board's own a button made for it, in one of RegionChoice's colors
        // (RegionButtonGrid.ButtonsFor, the popup's rule too). At most three to a row, 6px apart, as the panel is narrow.
        // ###########################################################################################
        private void LayOutRegionButtons(IReadOnlyList<string> regions, BoardData? board)
        {
            List<Button> buttons = RegionButtonGrid.ButtonsFor(
                regions,
                this.PalRegionButton,
                this.NtscRegionButton,
                this._ownRegionButtons,
                this.MakeOwnRegionButton,
                this.RegionButtonsGrid,
                region => RegionColors.Of(board, region));

            // An own region's button says its name - only its text changes (code review, 2026-11-09: a new block at every
            // read of the lists).
            foreach (Button button in this._ownRegionButtons.Where(buttons.Contains))
            {
                RegionButtonGrid.ShowText(button, button.Tag as string ?? string.Empty);
            }

            RegionButtonGrid.Place(buttons);
        }

        // A button for a region of the board's own: as wide as its share of the panel, a long name wrapping.
        private Button MakeOwnRegionButton()
        {
            var made = new Button
            {
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Stretch,
                HorizontalContentAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                FontSize = 12,
                Padding = new Avalonia.Thickness(4),
            };

            made.Content = RegionButtonGrid.NameBlock(made);
            made.Click += this.OnOwnRegionClick;

            return made;
        }

        private void OnOwnRegionClick(object? sender, RoutedEventArgs e)
        {
            if (!this._suppressRegionToggle && sender is Button { Tag: string region })
            {
                this.PickRegion(region);
            }
        }

        // The regions of the board on screen, as its buttons show them (BoardRegions.Of).
        internal IReadOnlyList<string> CurrentBoardRegions() => BoardRegions.Of(this._currentBoardData);

        // The board the region buttons are of: on the Contribute tab, the board as the draft now has it.
        private BoardData? BoardOfRegionButtons() =>
            this._contributeListsShown ? this.TabDrafts.BoardForList() ?? this._currentBoardData : this._currentBoardData;

        // `region`'s own colors on the board the buttons are of (RegionColors) - null for the colors CRT gives it.
        internal (string Background, string Text)? RegionColorsShown(string? region) => RegionColors.Of(this.BoardOfRegionButtons(), region);

        // ###########################################################################################
        // Positions a new popup on the same screen as the main window with a slight staggered offset.
        // ###########################################################################################
        private void PositionPopupOnSameScreen(Window popup)
        {
            popup.WindowStartupLocation = WindowStartupLocation.Manual;

            double scaling = this.RenderScaling > 0 ? this.RenderScaling : 1.0;
            double w = this.Bounds.Width > 0 ? this.Bounds.Width : this._restoreWidth;
            double h = this.Bounds.Height > 0 ? this.Bounds.Height : this._restoreHeight;

            int centerX = this.Position.X + (int)((w * scaling) / 2);
            int centerY = this.Position.Y + (int)((h * scaling) / 2);

            var screen = this.Screens.All.FirstOrDefault(s =>
                centerX >= s.Bounds.X &&
                centerY >= s.Bounds.Y &&
                centerX < s.Bounds.X + s.Bounds.Width &&
                centerY < s.Bounds.Y + s.Bounds.Height)
                ?? this.Screens.Primary;

            if (screen != null)
            {
                int cascadeStep = (int)(32 * scaling);
                int maxCascade = (int)(256 * scaling);

                int offsetX = this._popupCascadeOffset * cascadeStep;
                int offsetY = this._popupCascadeOffset * cascadeStep;

                if (offsetX > maxCascade)
                {
                    this._popupCascadeOffset = 0;
                    offsetX = 0;
                    offsetY = 0;
                }

                // Base it off the owner window's position slightly indented
                int px = this.Position.X + (int)(40 * scaling) + offsetX;
                int py = this.Position.Y + (int)(40 * scaling) + offsetY;

                // Adjust slightly if it forces itself off the edges of this target screen
                if (px + (popup.Width * scaling) > screen.Bounds.Right)
                    px = screen.Bounds.X + offsetX;
                if (py + (popup.Height * scaling) > screen.Bounds.Bottom)
                    py = screen.Bounds.Y + offsetY;

                popup.Position = new PixelPoint(px, py);
                this._popupCascadeOffset++;
            }
        }

        // ###########################################################################################
        // Opens a component info popup according to user settings.
        // ###########################################################################################
        internal void OpenComponentInfoPopup(string boardLabel, string displayText)
        {
            // Component images and local files are resolved draft-first through the draft folder -
            // see DraftFileResolver. The VIEWED folder, so it is empty under "view boards as
            // officially published" and the popup then shows only published bytes.
            string draftBoardFolder = this.GetCurrentBoardEntry() is { } currentEntry
                ? DraftManager.GetViewedBoardFolder(currentEntry.ExcelDataFile)
                : string.Empty;

            // The board whose published file lists the names a row cites (DraftFileResolver).
            string boardIdentity = this.GetCurrentBoardEntry()?.ExcelDataFile ?? string.Empty;

            this.ShowComponentInfoPopup(this._currentBoardData, boardLabel, displayText, draftBoardFolder, boardIdentity, offersEdit: true);
        }

        // ###########################################################################################
        // "Edit component" on a popup. The popup has closed itself; CRT's window comes to the front.
        // ###########################################################################################
        private void EditComponentFromPopup(string boardLabel)
        {
            this.Activate();
            _ = this.EditComponentAsync(boardLabel);
        }

        // Another board chosen: a popup still open is of the board before, and offers no editing of this one.
        private void FollowBoardInPopups()
        {
            string boardShown = this.GetCurrentBoardEntry()?.ExcelDataFile ?? string.Empty;

            this._singleComponentInfoWindow?.FollowBoardShown(boardShown);

            foreach (ComponentInfoWindow popup in this._componentInfoWindowsByKey.Values)
            {
                popup.FollowBoardShown(boardShown);
            }
        }

        internal ComponentInfoWindow? SingleComponentInfoWindowForTests => this._singleComponentInfoWindow;

        // ###########################################################################################
        // The Contribute tab's "Preview popup" (CRT 4.0 step 2, slice 3c - the plan: "Preview popup
        // shows exactly what will be published"): the popup on `board`, the draft as it is in memory,
        // its pictures from the draft's folder `draftBoardFolder`. The board on screen may not have
        // caught up with the draft yet (TabDrafts.CatchUpBoardOnScreen), so it is not read from there.
        // ###########################################################################################
        internal void OpenComponentPreview(BoardData board, string boardLabel, string draftBoardFolder, string boardIdentity)
        {
            ArgumentNullException.ThrowIfNull(board);

            ComponentEntry? row = ComponentPaneEdits.RowShown(board, boardLabel, this._localRegion);
            string displayText = string.Join(
                " | ",
                new[] { row?.BoardLabel ?? boardLabel, row?.FriendlyName, row?.TechnicalNameOrValue }
                    .Select(part => part?.Trim() ?? string.Empty)
                    .Where(part => part.Length > 0));

            // What will be published, seen from the pane the component is open in: nothing to go and edit.
            this.ShowComponentInfoPopup(board, boardLabel, displayText, draftBoardFolder, boardIdentity, offersEdit: false);
        }

        // ###########################################################################################
        // The popup of `boardLabel` on `boardData`, according to the user's settings - one popup
        // reused, or one per component.
        // ###########################################################################################
        private void ShowComponentInfoPopup(BoardData? boardData, string boardLabel, string displayText, string draftBoardFolder, string boardIdentity, bool offersEdit)
        {
            Action<string>? edit = offersEdit ? this.EditComponentFromPopup : null;
            string boardShown = this.GetCurrentBoardEntry()?.ExcelDataFile ?? string.Empty;

            string componentKey = $"{boardLabel}\u001F{displayText}";
            IReadOnlyList<string> regions = BoardRegions.Of(boardData);
            var images = boardData?.ComponentImages ?? new List<ComponentImageEntry>();
            var localFiles = boardData?.ComponentLocalFiles ?? new List<ComponentLocalFileEntry>();
            var links = boardData?.ComponentLinks ?? new List<ComponentLinkEntry>();
            var pins = boardData?.ComponentPins ?? new List<ComponentPinEntry>();
            var replacements = boardData?.ComponentReplacements ?? new List<ComponentReplacementEntry>();
            var boardFiles = boardData?.BoardLocalFiles ?? new List<BoardLocalFileEntry>();

            // Each pin's net on the board shown, from the KiCad data the Schematics tab has loaded for
            // it - the board every popup and the Contribute tab's preview are of.
            var pinNets = this.TabSchematicsControl.KiCadPinNetsOf(boardLabel);

            var componentEntries = boardData?.Components
                .Where(c => string.Equals(c.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase))
                .ToList() ?? new List<ComponentEntry>();

            // The board's own region colors (RegionColors, 2026-11-11), handed over before the buttons are laid out.
            IReadOnlyList<RegionColorEntry> regionColors = boardData?.RegionColors ?? [];

            if (UserSettings.MultipleInstancesForComponentPopup)
            {
                if (!this._componentInfoWindowsByKey.TryGetValue(componentKey, out var popup) || !popup.IsVisible)
                {
                    popup = new ComponentInfoWindow();
                    this._componentInfoWindowsByKey[componentKey] = popup;

                    popup.Closed += (_, _) =>
                    {
                        if (this._componentInfoWindowsByKey.TryGetValue(componentKey, out var existing) && ReferenceEquals(existing, popup))
                            this._componentInfoWindowsByKey.Remove(componentKey);
                    };
                }

                popup.RegionColorRows = regionColors;
                popup.SetComponent(
                    boardLabel,
                    displayText,
                    componentEntries,
                    images,
                    localFiles,
                    links,
                    this._localRegion,
                    DataManager.DataRoot,
                    regions,
                    draftBoardFolder,
                    boardIdentity,
                    pins,
                    replacements,
                    boardFiles,
                    pinNets);

                popup.OfferEdit(edit, boardShown);

                popup.UpdateOscilloscopeSessionTitleState(
                    this.TabOscilloscopeControl.HasSeenEstablishedOscilloscopeSessionForTitleState(),
                    this.TabOscilloscopeControl.HasActiveEstablishedOscilloscopeSessionForTitleState());

                if (!popup.IsVisible)
                {
                    this.PositionPopupOnSameScreen(popup);
                    popup.Show(this);
                    popup.Focus();
                }
                else
                {
                    popup.Activate();
                    popup.Focus();
                }

                return;
            }

            if (this._singleComponentInfoWindow == null)
            {
                this._singleComponentInfoWindow = new ComponentInfoWindow();
                this._singleComponentInfoWindow.Closed += (_, _) => this._singleComponentInfoWindow = null;
            }

            this._singleComponentInfoWindow.CloseOnDeactivate = false;
            this._singleComponentInfoWindow.RegionColorRows = regionColors;
            this._singleComponentInfoWindow.SetComponent(
                boardLabel,
                displayText,
                componentEntries,
                images,
                localFiles,
                links,
                this._localRegion,
                DataManager.DataRoot,
                regions,
                draftBoardFolder,
                boardIdentity,
                pins,
                replacements,
                boardFiles,
                pinNets);

            this._singleComponentInfoWindow.OfferEdit(edit, boardShown);

            this._singleComponentInfoWindow.UpdateOscilloscopeSessionTitleState(
                this.TabOscilloscopeControl.HasSeenEstablishedOscilloscopeSessionForTitleState(),
                this.TabOscilloscopeControl.HasActiveEstablishedOscilloscopeSessionForTitleState());

            if (!this._singleComponentInfoWindow.IsVisible)
            {
                // Single-instance popups reuse the last on-screen position (when not maximized)
                // instead of re-cascading from the main window every time they are reopened.
                if (UserSettings.HasComponentInfoWindowLayout &&
                    this._singleComponentInfoWindow.WindowState != Avalonia.Controls.WindowState.Maximized)
                {
                    this._singleComponentInfoWindow.Position =
                        new PixelPoint(UserSettings.ComponentInfoWindowX, UserSettings.ComponentInfoWindowY);
                }
                else
                {
                    this.PositionPopupOnSameScreen(this._singleComponentInfoWindow);
                }

                this._singleComponentInfoWindow.Show(this);
                this._singleComponentInfoWindow.Focus();
            }
            else
            {
                this._singleComponentInfoWindow.Activate();
                this._singleComponentInfoWindow.Focus();
            }
        }

        // ###########################################################################################
        // Handles "Blink selected" checkbox changes and refreshes highlight visuals immediately.
        // ###########################################################################################
        private void OnBlinkSelectedChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.BlinkSelected = BlinkSelectedCheckBox.IsChecked ?? false;

            this._blinkSelectedEnabled = this.BlinkSelectedCheckBox.IsChecked == true;

            bool hasBlinkEligibleSelection = this.TabSchematicsControl.HasBlinkEligibleSelection();
            bool hasComponentSelection = this.TabSchematicsControl.highlightIndexBySchematic.Count > 0;

            if (this._blinkSelectedEnabled && hasBlinkEligibleSelection)
            {
                this._blinkSelectedPhaseVisible = false;
                this.TabSchematicsControl.ApplyHighlightVisuals(
                    hasComponentSelection,
                    this.GetCurrentBlinkFactor(true));
                this.UpdateBlinkTimer(true);
                return;
            }

            this._blinkSelectedPhaseVisible = true;
            this.UpdateBlinkTimer(hasBlinkEligibleSelection);
            this.TabSchematicsControl.ApplyHighlightVisuals(
                hasComponentSelection,
                this.GetCurrentBlinkFactor(hasBlinkEligibleSelection));
        }

        // ###########################################################################################
        // Starts or stops the blink timer depending on current checkbox state and selection state.
        // ###########################################################################################
        internal void UpdateBlinkTimer(bool hasSelection)
        {
            bool shouldBlink = this._blinkSelectedEnabled && hasSelection;

            if (!shouldBlink)
            {
                this._blinkSelectedTimer?.Stop();
                this._blinkSelectedPhaseVisible = true;
                return;
            }

            if (this._blinkSelectedTimer == null)
            {
                this._blinkSelectedTimer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(450)
                };
                this._blinkSelectedTimer.Tick += this.OnBlinkSelectedTimerTick;
            }

            if (!this._blinkSelectedTimer.IsEnabled)
                this._blinkSelectedTimer.Start();
        }

        // ###########################################################################################
        // Advances blink phase and re-applies highlight visuals while selection exists.
        // ###########################################################################################
        private void OnBlinkSelectedTimerTick(object? sender, EventArgs e)
        {
            bool hasBlinkEligibleSelection = this.TabSchematicsControl.HasBlinkEligibleSelection();
            bool hasComponentSelection = this.TabSchematicsControl.highlightIndexBySchematic.Count > 0;

            if (!hasBlinkEligibleSelection)
            {
                this.UpdateBlinkTimer(false);
                this.TabSchematicsControl.ApplyHighlightVisuals(hasComponentSelection, 1.0);
                return;
            }

            // The phase alone - an opacity on what is already drawn, nothing rendered or built again
            // (code review, 2026-10-13; TabSchematics.ApplyBlinkPhase). It rendered every highlighted
            // thumbnail and rebuilt the KiCad overlay twice a second.
            this._blinkSelectedPhaseVisible = !this._blinkSelectedPhaseVisible;
            this.TabSchematicsControl.ApplyBlinkPhase(this.GetCurrentBlinkFactor(true));
        }

        // ###########################################################################################
        // Computes effective blink multiplier for current frame.
        // ###########################################################################################
        internal double GetCurrentBlinkFactor(bool hasSelection)
        {
            if (!hasSelection || !this._blinkSelectedEnabled)
                return 1.0;

            return this._blinkSelectedPhaseVisible ? 1.0 : 0.0;
        }

        // ###########################################################################################
        // Closes single popup when clicking the main window outside a component hit target.
        // ###########################################################################################
        private void OnMainPointerPressedCloseSinglePopup(object? sender, PointerPressedEventArgs e)
        {
            if (UserSettings.MultipleInstancesForComponentPopup)
                return;

            var popup = this._singleComponentInfoWindow;
            if (popup == null || !popup.IsVisible)
                return;

            if (this.isHoveringComponent)
                return;

            popup.Close();
        }

        // ###########################################################################################
        // Hands the arrow keys and Escape to "Align KiCad" while it is on and the keyboard is outside
        // the Schematics tab (KeyIsForKiCadAlignment, below - 2026-11-06), Home, End, Page Up and Page
        // Down to the "About" page while it is shown (KeyIsForAboutPage),
        // cancels "Add worklog" area-marking mode on Escape, closes the single component popup on
        // Escape otherwise, and toggles the schematics fullscreen window on F11 (ToggleSchematicsFullscreenWindow -
        // opens it if closed, closes it if already open, so repeatedly pressing F11 here behaves the
        // same as pressing Escape while it is open).
        //
        // WHY THE WORKLOG BRANCH IS HERE AND NOT ONLY IN TabSchematics: the mode is started by the
        // "Add worklog" button in this bar, which KEEPS keyboard focus afterwards, so the tab's own
        // OnSchematicsKeyDown - which also handles Escape for this mode - never receives the press
        // unless the user first clicks the schematic. Escape did nothing in exactly the moment it
        // is most wanted: right after clicking the button, before committing to a drag. This
        // handler is registered on the WINDOW at the Tunnel phase, so it sees the key wherever
        // focus happens to be.
        //
        // The tab's own handler is deliberately KEPT rather than replaced: it is the one that runs
        // once the user has clicked into the schematic (which takes focus), and it also covers the
        // fullscreen schematics window, which is a different window and never routes through here.
        // CancelWorklogEntryMode is safe to call twice, so the overlap costs nothing.
        //
        // The DETACHED THUMBNAILS window is a THIRD window and also never routes through here - see
        // SchematicsThumbnailsWindow's own F11 handler, which calls this same
        // ToggleSchematicsFullscreenWindow via TabSchematics.MainWindow.
        //
        // Ordered ABOVE the popup close, and gated on the mode actually being active so Escape
        // still reaches the popup at every other time: while an area is being marked the mode is
        // what Escape means, and cancelling both at once on one press would be two undos for one
        // keystroke.
        // ###########################################################################################
        private void OnMainKeyDownCloseSinglePopup(object? sender, KeyEventArgs e)
        {
            // A color dropper on: the key is its own (2026-11-03). And so is every key while an introduction tour shows
            // (Main.Tours.cs) - F11 would open the schematics' fullscreen under it.
            if (ColorDropperButton.IsAnyPicking || this.TourOverlay.IsShowing)
                return;

            // The arrow keys move the box "Align KiCad" aligns wherever the keyboard is (owner report,
            // 2026-11-06: "it seems not always to react when I use the keys for fine-tuning") - see
            // KeyIsForKiCadAlignment. Each press, a held key's every repeat included.
            if (e.Key is Key.Left or Key.Right or Key.Up or Key.Down && this.KeyIsForKiCadAlignment())
            {
                if (this.TabSchematicsControl.TakeKiCadCalibrationKey(e.Key, e.KeyModifiers))
                {
                    e.Handled = true;
                }

                return;
            }

            // Home, End, Page Up and Page Down scroll the "About" tab from wherever the keyboard is (owner request,
            // 2026-11-06) - see KeyIsForAboutPage.
            if (e.Key is Key.Home or Key.End or Key.PageUp or Key.PageDown && e.KeyModifiers == KeyModifiers.None && this.KeyIsForAboutPage())
            {
                if (this.TabAbout.TakeScrollKey(e.Key))
                {
                    e.Handled = true;
                }

                return;
            }

            // Not over the Maintainer tab (2026-09-29): F11 opens the SCHEMATICS fullscreen, which has
            // nothing to do with reviewing, over a window whose sidebar that tab has collapsed.
            if (e.Key == Key.F11 && !ReferenceEquals(this.MainTabControl?.SelectedItem, this.MaintainerTabItem))
            {
                this.ToggleSchematicsFullscreenWindow();
                e.Handled = true;
                return;
            }

            if (e.Key != Key.Escape)
                return;

            if (this.TabSchematicsControl.IsWorklogEntryModeActive)
            {
                // Resets this bar's own buttons too, via the ResetWorklogEntryModeButtons callback
                // the tab makes on the way out - so cancelling by keyboard leaves the bar in
                // exactly the state the Cancel button would.
                this.TabSchematicsControl.CancelWorklogEntryMode();
                e.Handled = true;
                return;
            }

            // "Draw trace" stops on Escape wherever the keyboard is (TabSchematics.ViewTools.cs): the
            // tab's own handler sees no key while the component filter holds the focus. Only while
            // the Schematics tab is shown (code review, 2026-10-29): on another tab, Escape belongs to
            // the box it was pressed in - a picture's name put back, a table cell.
            if (this.TabSchematicsControl.IsDrawingTraces &&
                ReferenceEquals(this.MainTabControl?.SelectedItem, this.SchematicsTabItem))
            {
                this.TabSchematicsControl.SetDrawingTraces(false);
                e.Handled = true;
                return;
            }

            var popup = this._singleComponentInfoWindow;
            if (!UserSettings.MultipleInstancesForComponentPopup && popup != null && popup.IsVisible)
            {
                popup.Close();
                e.Handled = true;
                return;
            }

            // "Align KiCad" ends on Escape wherever the keyboard is, as its arrow keys work (above) -
            // after the popup, which Escape has always closed first.
            if (this.KeyIsForKiCadAlignment() && this.TabSchematicsControl.TakeKiCadCalibrationKey(e.Key, e.KeyModifiers))
            {
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // *** A KEY FOR "ALIGN KICAD" FROM WHEREVER THE KEYBOARD IS IN THIS WINDOW (owner report,
        // 2026-11-06). *** The alignment's keys reached the box only through the Schematics tab's own
        // handler, which sees a key only while the keyboard is in that tab: after a press on the
        // component list, a drop-down, a tab header or a button of the left panel - or with nothing
        // focused at all, once the Contribute tab that asked for the tool let go of its button - the
        // arrow keys moved a list's selection or nothing, and the box stayed where it was.
        //
        // So while the tool is on and the Schematics tab is shown, Main hands them over - unless:
        //   - the keyboard is in the Schematics tab itself: its own handler takes the key next on the
        //     route, and handing it over here too would move the box twice;
        //   - it is in a box being typed in (the component filter, as the Delete key's rule has it,
        //     MainPictureToolsTests): there an arrow moves the caret;
        //   - it is in an open drop-down list (code review, 2026-11-08): there the arrows move through it and
        //     Escape closes it - handed over, Escape ended the alignment, the box's place thrown away, with the
        //     list still open. A closed drop-down is not spared, as KeyIsForAboutPage has it.
        // ###########################################################################################
        private bool KeyIsForKiCadAlignment() =>
            this.TabSchematicsControl.IsAligningKiCad &&
            ReferenceEquals(this.MainTabControl?.SelectedItem, this.SchematicsTabItem) &&
            !this.TabSchematicsControl.IsKeyboardFocusWithin &&
            TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement() is not (TextBox or ComboBox { IsDropDownOpen: true } or ComboBoxItem);

        // ###########################################################################################
        // *** HOME, END, PAGE UP AND PAGE DOWN ARE THE "ABOUT" PAGE'S WHILE IT IS SHOWN (owner request,
        // 2026-11-06: "When I am placed anywhere in the 'About' tab and I press 'Home', then I want it to navigate
        // to the top of the page"). *** The tab has nothing that takes the keyboard, and every click on it sends
        // the focus back to the component filter (ShouldReturnFocusToComponentSearch) - so the keys moved the
        // filter's caret, or a drop-down's selection, and the page stayed where it was. Main hands them over
        // unless the keyboard is where those keys mean something else:
        //   - a box being typed in other than the component filter (none on this tab, but a tab switch with
        //     the focus left in one);
        //   - an open drop-down list, which they move through;
        //   - a list on the left (components, categories), which they move through too.
        // A closed drop-down on the left is not spared: there Home or End would pick another hardware or board.
        // ###########################################################################################
        private bool KeyIsForAboutPage()
        {
            if (!ReferenceEquals(this.MainTabControl?.SelectedItem, this.AboutTabItem))
            {
                return false;
            }

            object? focused = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement();

            return focused switch
            {
                TextBox box => ReferenceEquals(box, this.ComponentSearchTextBox),
                ComboBox combo => !combo.IsDropDownOpen,
                ListBoxItem or ListBox => false,
                _ => true,
            };
        }

        // ###########################################################################################
        // Updates the UI with info specific to the current board's revision date and credits.
        // ###########################################################################################
        private void PopulateBoardInfoSection(string? revisionDate, List<CreditEntry>? credits)
        {
            this.TabAbout.SetBoardInfo(revisionDate, credits);
        }

        // ###########################################################################################
        // Refreshes the component list and highlight data for the current local region.
        // ###########################################################################################
        private void RefreshImages()
        {
            _ = this.ApplyRegionFilterAsync();
        }

        // ###########################################################################################
        // Refresh the component list according to the active region while recovering any matching
        // existing selection, similar to category filter switching.
        // ###########################################################################################
        private async Task ApplyRegionFilterAsync()
        {
            if (this._currentBoardData is not { } board)
                return;

            string region = this._localRegion;

            var highlightRects = await Task.Run(() => HighlightRectBuilder.BuildHighlightRects(board, region));

            // A load began or ended meanwhile, or another region was picked: the rectangles are that one's to set
            // (code review, 2026-11-05: the board's field read on the pool, and these put in before asking).
            if (!ReferenceEquals(board, this._currentBoardData) || !string.Equals(region, this._localRegion, StringComparison.OrdinalIgnoreCase))
                return;

            this.SetComponentHighlightRects(highlightRects);

            // What is held rather than what the list has selected now, as a load has it: a row of one region has a
            // key of its own, so switching there and back on the Contribute tab would leave nothing to select.
            var previouslySelectedKeys = this.IsSelectionHeld
                ? this._selectionToKeep!.Value.ComponentKeys
                : new HashSet<string>(
                    this.ComponentFilterListBox.SelectedItems?.Cast<ComponentListItem>()
                        .Select(i => i.SelectionKey) ?? Enumerable.Empty<string>(),
                    StringComparer.OrdinalIgnoreCase);

            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;
            var componentItems = ComponentListBuilder.BuildComponentItems(
                this._currentBoardData, this._localRegion, activeCategories, searchTerm, this._currentBoardDraftSummary.DraftedComponentBoardLabels);

            // A word selects all it finds - but for what is held, as a load has it (ShowListsOfBoard): the region
            // can be switched on the Contribute tab, and what the other tabs had selected is waiting there.
            bool selectsAllItFinds = !string.IsNullOrWhiteSpace(searchTerm) && !this.IsSelectionHeld;

            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            if (selectsAllItFinds)
            {
                // ItemsSource was just reassigned above; SelectAll can throw if Avalonia's
                // selection model has not yet caught up with the new item count.
                try { this.ComponentFilterListBox.SelectAll(); } catch { }
            }
            else
            {
                // A component told apart per region is selected again by its row of this region (owner report, 2026-11-13:
                // it was let go of, its key holding the region switched from - ComponentListBuilder.RowsToSelect).
                this.SelectInComponentList(componentItems, previouslySelectedKeys);
            }
            this._suppressComponentHighlightUpdate = false;

            // What the list now has selected - a row found again by its label and region too (SelectInComponentList).
            List<string> survivingLabels = this.SelectedComponentLabels();

            if (selectsAllItFinds)
            {
                survivingLabels = componentItems
                    .Select(item => item.BoardLabel)
                    .Where(l => !string.IsNullOrEmpty(l))
                    .ToList();
            }

            this.TabSchematicsControl.UpdateHighlightsForComponents(survivingLabels);
            this.TabOverview.ApplyFilter(searchTerm);
        }

        // How long the component filter waits after the last keystroke before it filters - the
        // Workbooks tab's search waits the same.
        internal static readonly TimeSpan ComponentSearchDebounceInterval = TimeSpan.FromMilliseconds(200);

        private DispatcherTimer? _componentSearchDebounceTimer;

        internal bool IsComponentSearchPendingForTests =>
            this._componentSearchDebounceTimer?.IsEnabled == true || this._contributeSearchTimer?.IsEnabled == true;

        // Applies a waiting filter of the box shown at once, as its timer would.
        internal void FlushComponentSearchForTests()
        {
            if (this._contributeListsShown)
            {
                this.ApplyContributeSearchNow();
            }
            else
            {
                this.ApplyComponentSearchNow();
            }
        }

        // The words the lists show now. Null: not known, so whatever the box says next is applied.
        private string? _componentSearchInEffect = string.Empty;

        // ###########################################################################################
        // *** THE BOX'S WORDS SET BY CRT, THE LISTS ALREADY SHOWING THEM (code review, 2026-11-04; measured).
        // *** Avalonia POSTS TextChanged, so a flag set around the assignment had long been lifted when it
        // came: 200 ms later the filter ran on lists nobody had touched, replaced them, let go of the selection
        // and took its highlights. So what the lists show is remembered, and a TextChanged saying the same is
        // nothing to do - every place setting the box goes through here.
        // ###########################################################################################
        private void ShowComponentSearchText(string text)
        {
            this._componentSearchInEffect = text;
            this.ComponentSearchTextBox.Text = text;
        }

        // ###########################################################################################
        // The search box's TextChanged: the filter runs 200 ms after the LAST keystroke (code review,
        // 2026-10-13). Each keystroke rebuilt the component list, every highlight, the thumbnails'
        // highlight layers and the KiCad overlay; typing "U12" did all of it three times. A direct
        // call with no sender (the label editor's) is applied at once.
        // ###########################################################################################
        public void OnComponentSearchTextChanged(object? sender, global::Avalonia.Controls.TextChangedEventArgs e)
        {
            if (sender == null)
            {
                this.ApplyComponentSearchNow();
                return;
            }

            // What the lists show already - set by CRT, or typed and taken back before the filter ran.
            if (string.Equals(this.ComponentSearchTextBox.Text ?? string.Empty, this._componentSearchInEffect, StringComparison.Ordinal))
            {
                this._componentSearchDebounceTimer?.Stop();
                return;
            }

            if (this._componentSearchDebounceTimer == null)
            {
                this._componentSearchDebounceTimer = new DispatcherTimer { Interval = Main.ComponentSearchDebounceInterval };
                this._componentSearchDebounceTimer.Tick += (_, _) => this.ApplyComponentSearchNow();
            }

            this._componentSearchDebounceTimer.Stop();
            this._componentSearchDebounceTimer.Start();
        }

        // ###########################################################################################
        // Refreshes the component filter list based on search text, or switches to label-editor
        // search behavior while the component label editor is active.
        // ###########################################################################################
        internal void ApplyComponentSearchNow()
        {
            this._componentSearchDebounceTimer?.Stop();

            if (this._currentBoardData == null || this._suppressCategoryFilterSave)
            {
                this._componentSearchInEffect = null;
                return;
            }

            string searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;

            this._componentSearchInEffect = searchTerm;

            if (this.TabSchematicsControl.IsLabelEditorActive)
            {
                this.TabSchematicsControl.ApplyLabelEditorSearchFilter(searchTerm);
                return;
            }

            this.ApplyNormalComponentSearchFilter(searchTerm);
        }

        // ###########################################################################################
        // Applies the normal component search behavior used outside the label editor.
        // ###########################################################################################
        private void ApplyNormalComponentSearchFilter(string searchTerm)
        {
            // Somebody's choice, as a selection made is: nothing held comes back in its place at the next load - and no longer
            // "Mark all"'s.
            this._selectionToKeep = null;
            this._markAllPressed = false;

            // *** NO BOARD READ, NOTHING TO FILTER (code review, 2026-11-05). *** "Clear" comes here without
            // ApplyComponentSearchNow's check, while a board loads or when one could not be read, and an
            // exception in a click handler ends CRT. A load in progress reads the box itself (ShowListsOfBoard).
            if (this._currentBoardData is not { } board)
            {
                return;
            }

            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var componentItems = ComponentListBuilder.BuildComponentItems(
                board, this._localRegion, activeCategories, searchTerm, this._currentBoardDraftSummary.DraftedComponentBoardLabels);

            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            var highlightLabels = new List<string>();

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                // ItemsSource was just reassigned above; SelectAll can throw if Avalonia's
                // selection model has not yet caught up with the new item count.
                try { this.ComponentFilterListBox.SelectAll(); } catch { }

                highlightLabels = componentItems
                    .Select(item => item.BoardLabel)
                    .Where(label => !string.IsNullOrEmpty(label))
                    .ToList();
            }

            this._suppressComponentHighlightUpdate = false;

            this.TabSchematicsControl.UpdateHighlightsForComponents(highlightLabels);

            // Forward the search term to filter the Overview tab's list
            this.TabOverview.ApplyFilter(searchTerm);
        }

        // ###########################################################################################
        // Updates the component search box text hint so its current behavior is obvious.
        // ###########################################################################################
        internal void UpdateComponentSearchTextBoxMode()
        {
            if (this.ComponentSearchTextBox == null)
            {
                return;
            }

            this.ComponentSearchTextBox.PlaceholderText = this.TabSchematicsControl.IsLabelEditorActive
                ? "Find component label or category"
                : "Filter components";
        }

        // ###########################################################################################
        // Returns true when the current board has regions (BoardRegions).
        // ###########################################################################################
        internal bool CurrentBoardHasExplicitRegionComponents()
        {
            return BoardRegions.Has(this._currentBoardData);
        }

        // ###########################################################################################
        // The Contribute tab's "Show on board" (CRT 4.0 step 2, slice 3c): the Schematics tab with only
        // `boardLabel` selected. `readBoardFirst` when the board on screen has not caught up with the
        // draft yet - it is read again, and the component selected once that load has its highlights
        // (DoWhatWasAskedAfterLoad): a reload drops the selection, and a component the draft added or
        // marked is not on the board as it was.
        // ###########################################################################################
        internal void ShowComponentOnBoard(string boardLabel, bool readBoardFirst)
        {
            this.SwitchToSchematicsTab();
            this.OnBoardShownUpToDate(readBoardFirst, this.TabSchematicsControl.GetCurrentSchematicName(), () => this.TabSchematicsControl.ShowOnlyComponent(boardLabel));
            this.OfferBackToContribute(PictureToolWording.ComponentShown(boardLabel));
        }

        // ###########################################################################################
        // The Contribute tab's "Show on board" on an important signal (CRT 4.0 step 2 - "click-to-light-
        // up"): the Schematics tab with only the net `netName` lit up, on the board read again first
        // when it has not caught up with the draft - as for a component.
        // ###########################################################################################
        internal void ShowNetOnBoard(string netName, bool readBoardFirst)
        {
            this.SwitchToSchematicsTab();
            this.OnBoardShownUpToDate(readBoardFirst, this.TabSchematicsControl.GetCurrentSchematicName(), () => this.TabSchematicsControl.ShowOnlyNet(netName));
            this.OfferBackToContribute(PictureToolWording.NetShown(netName));
        }

        // ###########################################################################################
        // `then` on the board on screen as its draft now is: at once, or - when the Contribute tab's
        // saves left it behind - once the board is read again, `schematic` selected. Used by "Show on
        // board" and the picture tools (Main.PictureTools.cs).
        // ###########################################################################################
        private void OnBoardShownUpToDate(bool readBoardFirst, string schematic, Action then)
        {
            if (readBoardFirst)
            {
                this._afterBoardLoad = (this.GetCurrentBoardKey(), then, 0);
                this.ReloadCurrentBoardFromDisk(schematic);
            }
            else
            {
                this._afterBoardLoad = null;
                then();
            }
        }

        // What waits for the board on screen to be read again - for the load of that board only, and
        // only the load that TOOK it (`Load`, 0 while none has): one that never reached its end - the
        // board file being replaced as it was read, a failure - leaves nothing for a later reload of the
        // same board to run out of the blue (code review, 2026-10-26).
        private (string BoardKey, Action Then, int Load)? _afterBoardLoad;

        // At the start of every board load: taken by it, or - taken by an earlier one - dropped.
        private void TakeWhatWaitsForLoad(int loadVersion)
        {
            if (this._afterBoardLoad is not { } asked)
            {
                return;
            }

            this._afterBoardLoad = asked.Load == 0 ? asked with { Load = loadVersion } : null;
        }

        // The load that took it has the board in: run, when that load is the latest and of its board.
        internal void DoWhatWasAskedAfterLoad(string boardKey)
        {
            if (this._afterBoardLoad is not { } asked)
            {
                return;
            }

            this._afterBoardLoad = null;

            if (asked.Load == this._boardSelectionLoadVersion && string.Equals(asked.BoardKey, boardKey, StringComparison.OrdinalIgnoreCase))
            {
                asked.Then();
            }
        }

        // ###########################################################################################
        // Pushes the current oscilloscope session title state into any open component info popup
        // windows so their title suffix stays aligned with the oscilloscope tab.
        // ###########################################################################################
        internal void UpdateComponentInfoWindowsOscilloscopeSessionState(
            bool hasSeenOscilloscopeSession,
            bool hasActiveOscilloscopeSession)
        {
            this._singleComponentInfoWindow?.UpdateOscilloscopeSessionTitleState(
                hasSeenOscilloscopeSession,
                hasActiveOscilloscopeSession);

            foreach (var popup in this._componentInfoWindowsByKey.Values)
            {
                popup.UpdateOscilloscopeSessionTitleState(
                    hasSeenOscilloscopeSession,
                    hasActiveOscilloscopeSession);
            }

            // The Contribute tab's component pane offers "Capture from my oscilloscope" while it can.
            this.TabDrafts.ShowScopeConnected(hasActiveOscilloscopeSession);
        }
    }
}
