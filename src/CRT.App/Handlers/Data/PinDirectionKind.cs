using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // What a pin of a chip drawn with its insides is colored by: what goes into the chip, what comes out of it, both - and
    // everything else (not connected, ground, power, other, unknown), which shares one quiet color.
    public enum PinDirectionKind
    {
        Other,
        Input,
        Output,
        InputOutput,
    }

    // ###########################################################################################
    // A PIN IS COLORED BY ITS DIRECTION (owner request, 2026-11-17: "instead of coloring the pins depending on its expected
    // oscilloscope reading, then I would like to color it for its pin direction, as seen in the pinout; Not connected + Input +
    // Output + Input/Output etc. E.g. Not connected + Power + Other + Unknown can all have the same lightgrey'ish color"): the
    // pin list's word (ComponentVocabulary.PinDirections), read as CRT reads every word of it - capitals and the space round
    // it aside. A word the vocabulary does not have is no direction CRT colors.
    // ###########################################################################################
    public static class PinDirectionKinds
    {
        public static PinDirectionKind Of(string? direction) => (direction ?? string.Empty).Trim() switch
        {
            string word when word.Equals("Input", StringComparison.OrdinalIgnoreCase) => PinDirectionKind.Input,
            string word when word.Equals("Output", StringComparison.OrdinalIgnoreCase) => PinDirectionKind.Output,
            string word when word.Equals("Input/output", StringComparison.OrdinalIgnoreCase) => PinDirectionKind.InputOutput,
            _ => PinDirectionKind.Other,
        };

        // ###########################################################################################
        // Each of a case's `pinCount` pins as drawn: THE DIRECTION ITS PIN LIST GIVES IT (`pins`, the component's in the region
        // shown) - "Unknown" grey, as the list says it (owner request, 2026-11-17: the data is converted so the list says what
        // the drawing does - ConversionPinouts) - and for a pin the list has no direction for, or no row, the datasheet's of a
        // chip CRT knows (`chip`, LogicChip.DirectionOf), as its name is the datasheet's then; none for any other case.
        // ###########################################################################################
        public static IReadOnlyDictionary<int, PinDirectionKind> ByPin(int pinCount, IEnumerable<ComponentPinEntry> pins, LogicChip? chip)
        {
            ArgumentNullException.ThrowIfNull(pins);

            var said = new Dictionary<int, string>();

            foreach (ComponentPinEntry pin in pins)
            {
                if (ComponentVocabulary.PinNumberOf(pin.Pin) is int number && !string.IsNullOrWhiteSpace(pin.Direction))
                {
                    said.TryAdd(number, pin.Direction);
                }
            }

            var kinds = new Dictionary<int, PinDirectionKind>();

            for (int pin = 1; pin <= pinCount; pin++)
            {
                kinds[pin] = PinDirectionKinds.Of(said.TryGetValue(pin, out string? direction) ? direction : chip?.DirectionOf(pin));
            }

            return kinds;
        }
    }
}
