using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A component's pins from a board's loaded KiCad data (CRT 4.0 step 2 - "Fill from KiCad" and
    // the popup's "net on this board"): the KiCad types handed to KiCadComponentPins as plain values.
    //
    // The pads are the FIRST PCB file's - the render-time rule KiCadReferenceMatcher keeps (only it
    // lights up copper) - and the symbol pins every sheet's. KiCad's "no net" is no net.
    // ###########################################################################################
    internal static class KiCadPinTable
    {
        public static List<KiCadPin> PinsOf(KiCadProjectRoot? root, string? label)
        {
            string wanted = label?.Trim() ?? string.Empty;

            if (root is null || wanted.Length == 0)
            {
                return [];
            }

            bool IsOf(string? reference) => string.Equals(reference?.Trim(), wanted, StringComparison.OrdinalIgnoreCase);

            IEnumerable<KiCadPadInfo> pads = root.Pcb
                .Take(1)
                .SelectMany(pcb => pcb.Footprints)
                .Where(footprint => IsOf(footprint.Reference))
                .SelectMany(footprint => footprint.Pads.Select(pad => new KiCadPadInfo(
                    footprint.Reference ?? string.Empty,
                    pad.Number ?? string.Empty,
                    pad.PinFunction ?? string.Empty,
                    pad.PinType ?? string.Empty,
                    pad.Net is { IsNoNet: false } net ? net.Name ?? string.Empty : string.Empty)));

            IEnumerable<KiCadSymbolPinInfo> symbolPins = root.Schematics
                .SelectMany(schematic => schematic.Symbols)
                .Where(symbol => IsOf(symbol.Reference))
                .SelectMany(symbol => symbol.Pins.Select(pin => new KiCadSymbolPinInfo(symbol.Reference ?? string.Empty, pin.Number, pin.Name, pin.Type)));

            return KiCadComponentPins.Of(pads, symbolPins, wanted);
        }
    }
}
