using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Handlers.DataHandling;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // A SINGLE RESISTOR AS ITS PINOUT (owner requests, 2026-11-07: "automatically create such an image ... in component info
    // popup", and 2026-11-09: "the image/drawing of the resistor should be in the pinout instead (like the other components you
    // are drawing, including the explanation of colors"): its color bands drawn from its value (ResistorBandsDrawing), the
    // value they say - "560 Ω 5%" - its power rating when typed, and each band's color in words with a swatch. In the place
    // the case or a leaded part is drawn, in the component popup (so it detaches with the pinout) and on the Contribute tab's
    // "Pinout" page. ResistorColorCode decides; anything it cannot say - a resistor pack, a trimmer, a value no band stands
    // for - shows nothing.
    // ###########################################################################################
    public partial class ResistorPicture : UserControl
    {
        public ResistorPicture()
        {
            this.InitializeComponent();
            this.IsVisible = false;
        }

        // ###########################################################################################
        // `entry` drawn, or nothing for anything not a single resistor whose value the bands can say - true when shown.
        // ###########################################################################################
        public bool Show(ComponentEntry? entry)
        {
            IReadOnlyList<ResistorColorCode.ColorBand>? bands = entry is null ? null : ResistorColorCode.BandsOf(entry);

            this.ColorsPanel.Children.Clear();

            if (bands is null || !ResistorColorCode.TryRead(entry!.TechnicalNameOrValue, out ResistorColorCode.Reading? reading))
            {
                this.BandsDrawing.Bands = [];
                this.IsVisible = false;
                return false;
            }

            this.BandsDrawing.Bands = bands;
            this.ValueText.Text = ResistorColorCode.ValueText(reading!);
            this.PowerText.Text = reading!.Power.Length > 0 ? "Power rating: " + reading.Power : string.Empty;
            this.PowerText.IsVisible = reading.Power.Length > 0;

            IReadOnlyList<string> names = ResistorColorCode.ColorNames(bands);

            for (int i = 0; i < bands.Count; i++)
            {
                var line = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };

                line.Children.Add(new Border
                {
                    Width = 12,
                    Height = 12,
                    CornerRadius = new CornerRadius(2),
                    BorderThickness = new Thickness(1),
                    BorderBrush = new SolidColorBrush(Color.FromArgb(0x66, 0, 0, 0)),
                    Background = new SolidColorBrush(Color.Parse(ResistorColorCode.ColorsOf(bands[i].Color).Fill)),
                    VerticalAlignment = VerticalAlignment.Center,
                });

                line.Children.Add(new TextBlock { Text = names[i], FontSize = 12, VerticalAlignment = VerticalAlignment.Center });
                this.ColorsPanel.Children.Add(line);
            }

            this.IsVisible = true;
            return true;
        }

        // ---- for the tests ----------------------------------------------------------------------

        internal IReadOnlyList<ResistorColorCode.ColorBand> BandsShownForTests => this.BandsDrawing.Bands;

        internal string ValueShownForTests => this.ValueText.Text ?? string.Empty;

        internal string PowerShownForTests => this.PowerText.IsVisible ? this.PowerText.Text ?? string.Empty : string.Empty;

        // The band colors listed under the value, in words.
        internal IReadOnlyList<string> ColorsShownForTests =>
            this.ColorsPanel.Children
                .OfType<StackPanel>()
                .Select(line => line.Children.OfType<TextBlock>().First().Text ?? string.Empty)
                .ToList();
    }
}
