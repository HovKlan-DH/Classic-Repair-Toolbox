using Handlers.DataHandling;

namespace Handlers.Theming
{
    // ###########################################################################################
    // THE COLORS A CHIP DRAWN WITH ITS INSIDES IS PAINTED IN (owner request, 2026-11-16 - after the owner's picture: the pins
    // in red, yellow and blue; PinoutDrawing.Inside.cs paints). THE CASE IS THE PLAIN CASE'S, the theme's card ground and
    // text color (owner request, the same day: "it should have the same color as the normal drawn pinout IC"), and the
    // gates are outlined in the theme's text color - so what is drawn on it follows the theme too.
    //
    // A PIN'S DIRECTION IS ITS COLOR (owner request, 2026-11-17: "instead of coloring the pins depending on its expected
    // oscilloscope reading, then I would like to color it for its pin direction"; PinDirectionKinds.ByPin) - green an input,
    // red an output, blue an input/output (owner request, the same day: "change the "Input" color to a nice green'ish color,
    // and then have the "Input/output" as the blue'ish color"), grey anything else (not connected, ground, power, other,
    // unknown) - AND A
    // DISCREET ONE (owner request, 2026-11-16: "I do think this is 'too much color'. Can you make the colors more discrete?"):
    // per direction and theme four resources, named by the keys here - its INK, a muted shade for the edge of its square and
    // badges; its square's TINT, pale on the light theme and dark on the dark; the NUMBER on that tint; and its WIRES inside
    // the case, a color of their own so they can be quieter than the squares. The owner's pick (2026-11-16, of seven
    // variations): every number near-black - white on the dark theme - and every wire one quiet grey, so the color is in the
    // squares and badges alone. ThemeContrastTests holds the number to 4.5:1 on its tint, the ink to 3:1 on the case's card
    // and the window's ground and the wires to 3:1 on the card, in all three themes. No name is written inside the case
    // (owner request, 2026-11-16: "I do not think this text inside the IC is needed"): a pin's name stands outside it.
    // ###########################################################################################
    public static class PinDirectionColors
    {
        // The edge of a pin's square and badges - "PinoutChip_Input" and so on, by the kind's name.
        public static string InkKeyOf(PinDirectionKind kind) => "PinoutChip_" + kind.ToString();

        // The wires of a pin's net inside the case - a color of their own, so they can be quieter than the squares.
        public static string WireKeyOf(PinDirectionKind kind) => PinDirectionColors.InkKeyOf(kind) + "_Wire";

        // A pin's square and its badges at the gates.
        public static string FillKeyOf(PinDirectionKind kind) => PinDirectionColors.InkKeyOf(kind) + "_Bg";

        // The number on a pin's square and its badges.
        public static string NumberKeyOf(PinDirectionKind kind) => PinDirectionColors.InkKeyOf(kind) + "_Fg";

        // What the legend under the drawing calls each - the pin list's own words.
        public static string WordOf(PinDirectionKind kind) => kind switch
        {
            PinDirectionKind.Input => "Input",
            PinDirectionKind.Output => "Output",
            PinDirectionKind.InputOutput => "Input/output",
            _ => "Other",
        };
    }
}
