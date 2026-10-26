using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A board's components from its loaded KiCad data (CRT 4.0 step 3 - "Components - 214 found in
    // KiCad [Add]"): the KiCad types handed to KiCadBoardComponents as plain values.
    //
    // The footprints are the FIRST PCB file's - the render-time rule KiCadReferenceMatcher keeps -
    // and the symbols every sheet's; each added component's pins are KiCadPinTable's, the ones "Fill
    // from KiCad" adds.
    // ###########################################################################################
    internal static class KiCadComponentTable
    {
        public static List<KiCadBoardComponent> Of(KiCadProjectRoot? root)
        {
            if (root is null)
            {
                return [];
            }

            List<KiCadFootprintInfo> footprints = root.Pcb
                .Take(1)
                .SelectMany(pcb => pcb.Footprints)
                .Select(footprint => new KiCadFootprintInfo(
                    footprint.Reference ?? string.Empty,
                    footprint.Value ?? string.Empty,
                    footprint.FootprintName ?? string.Empty,
                    footprint.Datasheet ?? string.Empty,
                    footprint.Pads
                        .Select(pad => new KiCadPadPlace(pad.Number ?? string.Empty, pad.AbsoluteCenter?.X ?? 0, pad.AbsoluteCenter?.Y ?? 0))
                        .ToList()))
                .ToList();

            List<KiCadSymbolInfo> symbols = root.Schematics
                .SelectMany(schematic => schematic.Symbols)
                .Select(symbol => new KiCadSymbolInfo(
                    symbol.Reference ?? string.Empty,
                    symbol.Value ?? string.Empty,
                    symbol.PropertiesDetailed
                        .FirstOrDefault(property => string.Equals(property.Name, "Datasheet", StringComparison.OrdinalIgnoreCase))
                        ?.Value ?? string.Empty))
                .ToList();

            return KiCadBoardComponents.Of(footprints, symbols);
        }

        // `board` with every component of `root` it lacks added (KiCadBoardComponents.WithComponents).
        public static BoardData WithComponents(BoardData board, KiCadProjectRoot root, out KiCadComponentsAdded added) =>
            KiCadBoardComponents.WithComponents(board, KiCadComponentTable.Of(root), label => KiCadPinTable.PinsOf(root, label), out added);
    }
}
