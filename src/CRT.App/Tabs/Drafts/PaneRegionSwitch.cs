using Avalonia.Controls;
using Avalonia.Layout;
using Handlers.DataHandling;
using System;

namespace CRT
{
    // ###########################################################################################
    // "SHOWING [PAL] [NTSC]" ON A COMPONENT'S PAGES (owner requests, 2026-11-01 and 2026-11-02: "please
    // show the two region buttons similarly as the Basics") - the switch under the component list, again:
    // picking here picks there (ComponentPane.OnPaneRegionPicked -> RegionPicked -> Main.PickRegion). Basics
    // and Pinout show it under their tick and only while it is ticked; the other pages whenever the board
    // has regions, as what they list is what the region shown shows.
    //
    // The buttons' looks are ComponentPane.axaml's `Button.PaneRegion` styles, in the colors of the buttons
    // under the component list.
    // ###########################################################################################
    public sealed class PaneRegionSwitch : StackPanel
    {
        public PaneRegionSwitch()
        {
            this.Orientation = Orientation.Horizontal;
            this.Spacing = Handlers.Theming.ButtonRow.Gap;
            this.Margin = new Avalonia.Thickness(0, 0, 0, 8);
            this.IsVisible = false;

            this.PalButton = this.AddRegionButton(BoardRegions.Pal, "pal");
            this.NtscButton = this.AddRegionButton(BoardRegions.Ntsc, "ntsc");

            this.Children.Insert(0, new TextBlock { Text = "Showing", VerticalAlignment = VerticalAlignment.Center, FontSize = 12 });
        }

        // A region picked that is not the one shown.
        public event EventHandler<string>? RegionPicked;

        internal Button PalButton { get; }

        internal Button NtscButton { get; }

        private Button AddRegionButton(string region, string colorClass)
        {
            var button = new Button
            {
                Content = region,
                MinWidth = 60,
                HorizontalContentAlignment = HorizontalAlignment.Center,
                FontSize = 12,
                Padding = new Avalonia.Thickness(4),
            };

            button.Classes.Add("PaneRegion");
            button.Classes.Add(colorClass);
            button.Click += (_, _) => this.RegionPicked?.Invoke(this, region);

            this.Children.Add(button);

            return button;
        }

        // Shown or not, and which region is the one shown.
        public void Show(bool visible, string? region)
        {
            this.IsVisible = visible;
            this.PalButton.Classes.Set("active", string.Equals(region, BoardRegions.Pal, StringComparison.OrdinalIgnoreCase));
            this.NtscButton.Classes.Set("active", string.Equals(region, BoardRegions.Ntsc, StringComparison.OrdinalIgnoreCase));
        }
    }
}
