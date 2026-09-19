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
    // selected" highlight timer, the region (PAL/NTSC) toggle, and the contribution editor
    // window opened from a popup. See Main.axaml.cs for the file map of the whole partial class.
    // ###########################################################################################
    public partial class Main
    {
        // ###########################################################################################
        // Clears the component search text, removes all selected items in the Component filter list,
        // and resets schematic highlights back to an empty selection state.
        // ###########################################################################################
        private void OnClearComponentsClick(object? sender, RoutedEventArgs e)
        {
            this._suppressComponentSearchRefresh = true;
            this.ComponentSearchTextBox.Text = string.Empty;
            this._suppressComponentSearchRefresh = false;

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
        // Selects all available items currently populated within the Component filter list box.
        // ###########################################################################################
        private void OnMarkAllComponentsClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                this.ComponentFilterListBox.SelectAll();
            }
            catch (OutOfMemoryException ex)
            {
                Logger.Debug(ex, "Failed to mark all components. The selection was too large to process");
            }
        }

        // ###########################################################################################
        // Switches the local region to PAL and reloads images.
        // ###########################################################################################
        private void OnPalRegionClick(object? sender, RoutedEventArgs e)
        {
            if (this._suppressRegionToggle)
                return;

            this._localRegion = "PAL";
            UserSettings.Region = "PAL";

            this.UpdateRegionButtonsState();
            this.RefreshImages();
            this.TabSchematicsControl.UpdateOverlayLabels();
        }

        // ###########################################################################################
        // Switches the local region to NTSC and reloads images.
        // ###########################################################################################
        private void OnNtscRegionClick(object? sender, RoutedEventArgs e)
        {
            if (this._suppressRegionToggle)
                return;

            this._localRegion = "NTSC";
            UserSettings.Region = "NTSC";

            this.UpdateRegionButtonsState();
            this.RefreshImages();
            this.TabSchematicsControl.UpdateOverlayLabels();
        }

        // ###########################################################################################
        // Updates the region toggle and button states to match the current local region.
        // Hides the entire region toggle area when the current board has no explicit PAL/NTSC components.
        // ###########################################################################################
        private void UpdateRegionButtonsState()
        {
            this._suppressRegionToggle = true;
            bool isNtsc = string.Equals(this._localRegion, "NTSC", StringComparison.OrdinalIgnoreCase);
            bool hasExplicitRegionComponents = ComponentListBuilder.HasExplicitRegionComponents(this._currentBoardData);

            this.RegionButtonsGrid.IsVisible = hasExplicitRegionComponents;

            this.NtscRegionButton.Classes.Set("active", isNtsc);
            this.PalRegionButton.Classes.Set("active", !isNtsc);

            this._suppressRegionToggle = false;
        }

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
            string componentKey = $"{boardLabel}\u001F{displayText}";
            var boardData = this._currentBoardData;
            bool hasExplicitRegionComponents = ComponentListBuilder.HasExplicitRegionComponents(boardData);
            var images = boardData?.ComponentImages ?? new List<ComponentImageEntry>();
            var localFiles = boardData?.ComponentLocalFiles ?? new List<ComponentLocalFileEntry>();
            var links = boardData?.ComponentLinks ?? new List<ComponentLinkEntry>();
            var componentEntries = boardData?.Components
                .Where(c => string.Equals(c.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase))
                .ToList() ?? new List<ComponentEntry>();

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

                popup.SetComponent(
                    boardLabel,
                    displayText,
                    componentEntries,
                    images,
                    localFiles,
                    links,
                    UserSettings.Region,
                    DataManager.DataRoot,
                    hasExplicitRegionComponents);

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
            this._singleComponentInfoWindow.SetComponent(
                boardLabel,
                displayText,
                componentEntries,
                images,
                localFiles,
                links,
                UserSettings.Region,
                DataManager.DataRoot,
                hasExplicitRegionComponents);

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

            this._blinkSelectedPhaseVisible = !this._blinkSelectedPhaseVisible;
            this.TabSchematicsControl.ApplyHighlightVisuals(
                hasComponentSelection,
                this.GetCurrentBlinkFactor(true));
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
        // Cancels "Add worklog" area-marking mode on Escape, closes the single component popup on
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
            if (e.Key == Key.F11)
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

            if (UserSettings.MultipleInstancesForComponentPopup)
                return;

            var popup = this._singleComponentInfoWindow;
            if (popup == null || !popup.IsVisible)
                return;

            popup.Close();
            e.Handled = true;
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
            if (this._currentBoardData == null)
                return;

            this.SetComponentHighlightRects(await Task.Run(() =>
                HighlightRectBuilder.BuildHighlightRects(this._currentBoardData, this._localRegion)));

            var previouslySelectedKeys = new HashSet<string>(
                this.ComponentFilterListBox.SelectedItems?.Cast<ComponentListItem>()
                    .Select(i => i.SelectionKey) ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;
            var componentItems = ComponentListBuilder.BuildComponentItems(this._currentBoardData, this._localRegion, activeCategories, searchTerm);

            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                // ItemsSource was just reassigned above; SelectAll can throw if Avalonia's
                // selection model has not yet caught up with the new item count.
                try { this.ComponentFilterListBox.SelectAll(); } catch { }
            }
            else
            {
                for (int i = 0; i < componentItems.Count; i++)
                {
                    if (previouslySelectedKeys.Contains(componentItems[i].SelectionKey))
                        this.ComponentFilterListBox.Selection.Select(i);
                }
            }
            this._suppressComponentHighlightUpdate = false;

            var survivingLabels = componentItems
                .Where(item => previouslySelectedKeys.Contains(item.SelectionKey))
                .Select(item => item.BoardLabel)
                .Where(l => !string.IsNullOrEmpty(l))
                .ToList();

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                survivingLabels = componentItems
                    .Select(item => item.BoardLabel)
                    .Where(l => !string.IsNullOrEmpty(l))
                    .ToList();
            }

            this.TabSchematicsControl.UpdateHighlightsForComponents(survivingLabels);
            this.TabContribute.LoadData(this._currentBoardData, this._localRegion);
            this.TabOverview.ApplyFilter(searchTerm);
        }

        // ###########################################################################################
        // Refreshes the component filter list based on search text, or switches to label-editor
        // search behavior while the component label editor is active.
        // ###########################################################################################
        public void OnComponentSearchTextChanged(object? sender, global::Avalonia.Controls.TextChangedEventArgs e)
        {
            if (this._suppressComponentSearchRefresh || this._currentBoardData == null || this._suppressCategoryFilterSave)
                return;

            string searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;

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
            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var componentItems = ComponentListBuilder.BuildComponentItems(this._currentBoardData!, this._localRegion, activeCategories, searchTerm);

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
        // Returns true when the current board has at least one component explicitly tagged as PAL or NTSC.
        // ###########################################################################################
        internal bool CurrentBoardHasExplicitRegionComponents()
        {
            return ComponentListBuilder.HasExplicitRegionComponents(this._currentBoardData);
        }

        // ###########################################################################################
        // Opens a maximized contribution editor window for the selected component.
        // ###########################################################################################
        internal void OpenComponentContributionWindow(string boardLabel)
        {
            if (this._currentBoardData == null || string.IsNullOrWhiteSpace(boardLabel))
            {
                return;
            }

            this.ShowComponentContributionWindow(window => window.LoadComponent(
                this._currentBoardData,
                DataManager.DataRoot,
                this.HardwareComboBox.SelectedItem as string ?? string.Empty,
                this.BoardComboBox.SelectedItem as string ?? string.Empty,
                this._localRegion,
                boardLabel,
                this.GetCurrentBoardEntry()?.ExcelDataFile ?? string.Empty));
        }

        // ###########################################################################################
        // Opens the contribution editor on a component that does not exist in the board data yet,
        // so a missing component can be suggested from scratch.
        // ###########################################################################################
        internal void OpenNewComponentContributionWindow()
        {
            if (this._currentBoardData == null)
            {
                return;
            }

            this.ShowComponentContributionWindow(window => window.LoadNewComponent(
                this._currentBoardData,
                DataManager.DataRoot,
                this.HardwareComboBox.SelectedItem as string ?? string.Empty,
                this.BoardComboBox.SelectedItem as string ?? string.Empty,
                this._localRegion,
                this.GetCurrentBoardEntry()?.ExcelDataFile ?? string.Empty));
        }

        // ###########################################################################################
        // Creates the contribution editor, lets the caller load it, and shows it maximized on the
        // screen the main window is on.
        // ###########################################################################################
        private void ShowComponentContributionWindow(Action<ComponentContributionWindow> loadContent)
        {
            var window = new ComponentContributionWindow();
            loadContent(window);

            this.PositionFullscreenWindowOnSameScreen(window);
            window.WindowState = Avalonia.Controls.WindowState.Maximized;
            window.Show(this);
            window.Focus();
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
        }

        // ###########################################################################################
        // Rebuilds runtime component lists and highlight caches after the schematic label editor
        // has modified the in-memory board data for the current session.
        // ###########################################################################################
        internal void RefreshRuntimeBoardStateAfterLabelEditorApply()
        {
            if (this._currentBoardData == null)
            {
                return;
            }

            var previouslySelectedKeys = new HashSet<string>(
                this.ComponentFilterListBox.SelectedItems?.Cast<ComponentListItem>()
                    .Select(i => i.SelectionKey) ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            string searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;

            this.SetComponentHighlightRects(
                HighlightRectBuilder.BuildHighlightRects(this._currentBoardData, this._localRegion));

            var componentItems = ComponentListBuilder.BuildComponentItems(this._currentBoardData, this._localRegion, activeCategories, searchTerm);

            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                // ItemsSource was just reassigned above; SelectAll can throw if Avalonia's
                // selection model has not yet caught up with the new item count.
                try
                {
                    this.ComponentFilterListBox.SelectAll();
                }
                catch
                {
                }
            }
            else
            {
                for (int i = 0; i < componentItems.Count; i++)
                {
                    if (previouslySelectedKeys.Contains(componentItems[i].SelectionKey))
                    {
                        this.ComponentFilterListBox.Selection.Select(i);
                    }
                }
            }

            this._suppressComponentHighlightUpdate = false;

            var survivingLabels = componentItems
                .Where(item => previouslySelectedKeys.Contains(item.SelectionKey))
                .Select(item => item.BoardLabel)
                .Where(label => !string.IsNullOrWhiteSpace(label))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                survivingLabels = componentItems
                    .Select(item => item.BoardLabel)
                    .Where(label => !string.IsNullOrWhiteSpace(label))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            this.TabSchematicsControl.UpdateHighlightsForComponents(survivingLabels);
            this.TabSchematicsControl.UpdateComponentLabels();
            this.TabOverview.LoadData(this._currentBoardData);
            this.TabContribute.LoadData(this._currentBoardData, this._localRegion);
        }
    }
}
