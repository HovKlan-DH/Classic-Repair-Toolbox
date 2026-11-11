using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE POPUP'S REGION BUTTONS AND ITS REGION LABEL (2026-11-09: a board may come in regions of its own) - in a part of
    // their own (code review, 2026-11-09: the window's own file had grown past 2,300 lines). One button per region of the
    // board, which button each is decided as Main's buttons decide it (RegionButtonGrid.ButtonsFor), each saying how many
    // of the component's images its region shows. The region picked here is the popup's own - CRT's stays as it was.
    //
    // Part of the ComponentInfoWindow partial class - see ComponentInfoWindow.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentInfoWindow
    {
        // The board's regions, in its order (BoardRegions.Of) - one button each; none without regions.
        private IReadOnlyList<string> _regions = [];

        // The board's own colors for its regions (owner request, 2026-11-11; RegionColors) - handed over by Main with the board,
        // before SetComponent. None: every region in the colors CRT gives it.
        internal IReadOnlyList<RegionColorEntry> RegionColorRows { get; set; } = [];

        private (string Background, string Text)? OwnColorsOf(string? region) => RegionColors.Of(this._regions, this.RegionColorRows, region);

        // The buttons of a board's own regions, made as a board needs them (2026-11-09).
        private readonly List<Button> _ownRegionButtons = new();

        // True while the buttons are set from code - a click then is none.
        private bool _suppressRegionToggle = false;

        // ###########################################################################################
        // ONE BUTTON PER REGION OF THE BOARD: "PAL" and "NTSC" the buttons they always were, where they always stood; a region
        // of the board's own a button made for it, in one of RegionChoice's colors - all stacked from the bottom of the footer
        // (RegionButtonLayout.PopupCellOf). Each button's Tag is its region.
        // ###########################################################################################
        private List<Button> RegionButtonsShown()
        {
            List<Button> buttons = RegionButtonGrid.ButtonsFor(
                this._regions,
                this.PalRegionButton,
                this.NtscRegionButton,
                this._ownRegionButtons,
                this.MakeOwnRegionButton,
                this.PopupRegionGrid,
                this.OwnColorsOf);

            // The grid's two Auto columns are its markup's and never change (RegionButtonLayout's header: Avalonia's Grid
            // threw on a column it had lost); a button of the second column stands 6px off the first by its margin.
            for (int index = 0; index < buttons.Count; index++)
            {
                (int row, int column) = RegionButtonLayout.PopupCellOf(index, buttons.Count);
                Grid.SetRow(buttons[index], row);
                Grid.SetColumn(buttons[index], column);
                buttons[index].Margin = new Thickness(column > 0 ? RegionButtonLayout.Gap : 0, 0, 0, 0);
            }

            return buttons;
        }

        // ###########################################################################################
        // A button for a region of the board's own, as PAL's and NTSC's look - 70px wide for a short name, wider for a long
        // one (up to BoardRegions.LongestName, its count after it) rather than cutting it off (code review, 2026-11-09:
        // "Japan (NTSC-J) (3)" in a fixed 70px). Its column is Auto; its row a fixed 28px, so it widens rather than wraps.
        // ###########################################################################################
        private Button MakeOwnRegionButton()
        {
            var made = new Button
            {
                MinWidth = 70,
                MinHeight = 24,
                FontSize = 11,
                Padding = new Thickness(6, 2),
                HorizontalContentAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                VerticalContentAlignment = Avalonia.Layout.VerticalAlignment.Center,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Bottom,
                Height = 24,
            };

            made.Content = RegionButtonGrid.NameBlock(made);
            made.Click += this.OnOwnRegionClick;

            return made;
        }

        // ###########################################################################################
        // Each button's caption: its region and how many of the component's images it shows - an image of no region counts
        // in every one.
        // ###########################################################################################
        private void UpdateRegionButtonCounters(IEnumerable<Button> buttons)
        {
            foreach (Button button in buttons)
            {
                string region = button.Tag as string ?? string.Empty;
                int count = ComponentImageQueries.CountImagesForRegion(this._allComponentImages, this._boardLabel, region);

                RegionButtonGrid.ShowText(button, $"{region} ({count})");
            }
        }

        private void OnOwnRegionClick(object? sender, RoutedEventArgs e)
        {
            if (!this._suppressRegionToggle && sender is Button { Tag: string region })
            {
                this.ShowRegionPicked(region);
            }
        }

        // ###########################################################################################
        // Switches the local region to PAL and reloads images without touching the global setting.
        // ###########################################################################################
        private void OnPalRegionClick(object? sender, RoutedEventArgs e)
        {
            if (!this._suppressRegionToggle)
            {
                this.ShowRegionPicked(BoardRegions.Pal);
            }
        }

        // ###########################################################################################
        // Switches the local region to NTSC and reloads images without touching the global setting.
        // ###########################################################################################
        private void OnNtscRegionClick(object? sender, RoutedEventArgs e)
        {
            if (!this._suppressRegionToggle)
            {
                this.ShowRegionPicked(BoardRegions.Ntsc);
            }
        }

        private void ShowRegionPicked(string region)
        {
            this._localRegion = region;
            this.UpdateRegionButtonsState();
            this.RefreshImages();
        }

        internal IReadOnlyList<string> RegionButtonsShownForTests =>
            this.RegionButtonsShown().Select(RegionButtonGrid.TextOf).ToList();

        // ###########################################################################################
        // Updates the region toggle and button states to match the current local region. Hides only the region buttons when
        // the board has no regions, while keeping the switches and Close button visible. The buttons are worked out ONCE
        // (code review, 2026-11-09: twice per refresh - for their counts, then for which is lit).
        // ###########################################################################################
        private void UpdateRegionButtonsState()
        {
            this._suppressRegionToggle = true;

            List<Button> buttons = this.RegionButtonsShown();

            this.UpdateRegionButtonCounters(buttons);

            if (this.FooterGrid is { } footerGrid && footerGrid.ColumnDefinitions.Count >= 5)
            {
                footerGrid.ColumnDefinitions[0].Width = this._hasExplicitRegionComponents
                    ? GridLength.Auto
                    : new GridLength(0, GridUnitType.Pixel);

                footerGrid.ColumnDefinitions[1].Width = this._hasExplicitRegionComponents
                    ? new GridLength(12, GridUnitType.Pixel)
                    : new GridLength(0, GridUnitType.Pixel);
            }

            foreach (Button button in buttons)
            {
                button.Classes.Set("active", string.Equals(button.Tag as string, this._localRegion, StringComparison.OrdinalIgnoreCase));
            }

            this._suppressRegionToggle = false;
            this.UpdateRegionLabel();
        }

        // ###########################################################################################
        // Updates the region label overlay in the top-left corner using the same color schema
        // as the region label in the Schematics tab, bound dynamically to the current local region.
        // ###########################################################################################
        private void UpdateRegionLabel()
        {
            string colorPrefix = RegionChoice.ColorsOf(this._regions, this._localRegion);

            this.InfoRegionText.Text = this._localRegion;
            this.InfoRegionBorder.IsVisible = this._hasExplicitRegionComponents;

            // The board's own colors where it gives the region any (2026-11-11).
            Handlers.Theming.RegionPaint.PaintLabel(this, this.InfoRegionBorder, this.InfoRegionText, colorPrefix, this.OwnColorsOf(this._localRegion));
        }
    }
}
