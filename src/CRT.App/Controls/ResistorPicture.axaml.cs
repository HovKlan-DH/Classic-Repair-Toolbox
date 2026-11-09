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
    //
    // AN AXIAL CERAMIC CAPACITOR TOO (owner request, 2026-11-09: "since that has color rings, would it then be possible to
    // show and translate those also, just alike a resistor?"): its bands as the owner's own parts carry them (CapacitorColorCode),
    // on the light-green body such parts have, its voltage where a resistor's power rating stands - and its fifth band, and a
    // tolerance not typed, outlined as not known, which a line under the colors explains.
    // ###########################################################################################
    public partial class ResistorPicture : UserControl
    {
        // An axial ceramic capacitor's body, as the board's own picture of one shows it.
        private const string CeramicBody = "#D3DC86";

        public ResistorPicture()
        {
            this.InitializeComponent();
            this.IsVisible = false;
        }

        // ###########################################################################################
        // `entry` drawn, or nothing for anything not a single resistor or an axial ceramic capacitor whose value the bands can
        // say - true when shown.
        // ###########################################################################################
        public bool Show(ComponentEntry? entry)
        {
            this.ColorsPanel.Children.Clear();

            if (entry is not null &&
                ResistorColorCode.TryBandsOf(entry, out IReadOnlyList<ResistorColorCode.ColorBand>? resistorBands, out ResistorColorCode.Reading? resistor))
            {
                this.ShowBands(resistorBands!, bodyColor: null, "Resistor value:", ResistorColorCode.ValueText(resistor!),
                    resistor!.Power.Length > 0 ? "Power rating: " + resistor.Power : string.Empty, note: string.Empty);
                return true;
            }

            if (entry is not null &&
                CapacitorColorCode.TryBandsOf(entry, out IReadOnlyList<ResistorColorCode.ColorBand>? capacitorBands, out CapacitorColorCode.Reading? capacitor))
            {
                this.ShowBands(capacitorBands!, CeramicBody, "Capacitor value:", CapacitorColorCode.ValueText(capacitor!),
                    capacitor!.Voltage.Length > 0 ? "Voltage rating: " + capacitor.Voltage : string.Empty,
                    CapacitorColorCode.NoteOf(capacitorBands!));
                return true;
            }

            this.BandsDrawing.Bands = [];
            this.IsVisible = false;
            return false;
        }

        private void ShowBands(IReadOnlyList<ResistorColorCode.ColorBand> bands, string? bodyColor, string heading, string value, string rating, string note)
        {
            this.BandsDrawing.Bands = bands;
            this.BandsDrawing.BodyColor = bodyColor;
            this.ValueHeading.Text = heading;
            this.ValueText.Text = value;
            this.PowerText.Text = rating;
            this.PowerText.IsVisible = rating.Length > 0;
            this.UnknownBandsNote.Text = note;
            this.UnknownBandsNote.IsVisible = note.Length > 0;

            IReadOnlyList<string> names = ResistorColorCode.ColorNames(bands);

            for (int i = 0; i < bands.Count; i++)
            {
                var line = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 6 };
                line.Children.Add(Swatch(bands[i].Color));
                line.Children.Add(new TextBlock { Text = names[i], FontSize = 12, VerticalAlignment = VerticalAlignment.Center });
                this.ColorsPanel.Children.Add(line);
            }

            this.IsVisible = true;
        }

        // A square of the band's color before its name - a band not known an empty one, outlined dashed, as it is drawn.
        private static Control Swatch(ResistorColorCode.Band color) => color == ResistorColorCode.Band.Unknown
            ? new Avalonia.Controls.Shapes.Rectangle
            {
                Width = 12,
                Height = 12,
                Stroke = new SolidColorBrush(Color.FromArgb(0xB3, 0x80, 0x80, 0x80)),
                StrokeThickness = 1,
                StrokeDashArray = [2, 2],
                VerticalAlignment = VerticalAlignment.Center,
            }
            : new Border
            {
                Width = 12,
                Height = 12,
                CornerRadius = new CornerRadius(2),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.FromArgb(0x66, 0, 0, 0)),
                Background = new SolidColorBrush(Color.Parse(ResistorColorCode.ColorsOf(color).Fill)),
                VerticalAlignment = VerticalAlignment.Center,
            };

        // ---- for the tests ----------------------------------------------------------------------

        internal IReadOnlyList<ResistorColorCode.ColorBand> BandsShownForTests => this.BandsDrawing.Bands;

        internal string ValueShownForTests => this.ValueText.Text ?? string.Empty;

        internal string PowerShownForTests => this.PowerText.IsVisible ? this.PowerText.Text ?? string.Empty : string.Empty;

        internal string HeadingShownForTests => this.ValueHeading.Text ?? string.Empty;

        internal string UnknownBandsNoteShownForTests => this.UnknownBandsNote.IsVisible ? this.UnknownBandsNote.Text ?? string.Empty : string.Empty;

        internal string? BodyColorForTests => this.BandsDrawing.BodyColor;

        // The band colors listed under the value, in words.
        internal IReadOnlyList<string> ColorsShownForTests =>
            this.ColorsPanel.Children
                .OfType<StackPanel>()
                .Select(line => line.Children.OfType<TextBlock>().First().Text ?? string.Empty)
                .ToList();
    }
}
