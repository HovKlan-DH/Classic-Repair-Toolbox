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
    //
    // In CRT.Data since 2026-11-06, with the KiCad reader, so the data tree's conversion makes each
    // board's pinouts from its KiCad data (ConversionPinouts). How many components each view SHOWS needs
    // the app's marking geometry, so that count is BoardKiCadFacts.CountsByView's.
    // ###########################################################################################
    internal static class KiCadComponentTable
    {
        public static List<KiCadBoardComponent> Of(KiCadProjectRoot? root) =>
            root is null ? [] : KiCadBoardComponents.Of(KiCadComponentTable.Footprints(root), KiCadComponentTable.Symbols(root));

        // ###########################################################################################
        // The components KiCad has a symbol for but no footprint (2026-11-01) - marked from the symbol on a
        // schematic sheet's picture, with nothing to mark them from on a side of the board
        // (KiCadBoardComponents.OnlyInSchematic).
        // ###########################################################################################
        public static List<string> OnlyInSchematic(KiCadProjectRoot? root) =>
            root is null ? [] : KiCadBoardComponents.OnlyInSchematic(KiCadComponentTable.Footprints(root), KiCadComponentTable.Symbols(root));

        private static List<KiCadFootprintInfo> Footprints(KiCadProjectRoot root) =>
            root.Pcb
                .Take(1)
                .SelectMany(pcb => pcb.Footprints)
                .Select(KiCadComponentTable.InfoOf)
                .ToList();

        // One footprint as KiCadBoardComponents takes it - BoardKiCadFacts.CountsByView counts by it too.
        internal static KiCadFootprintInfo InfoOf(KiCadPcbFootprint footprint) =>
            new(
                footprint.Reference ?? string.Empty,
                footprint.Value ?? string.Empty,
                footprint.FootprintName ?? string.Empty,
                footprint.Datasheet ?? string.Empty,
                footprint.Pads
                    .Select(pad => new KiCadPadPlace(pad.Number ?? string.Empty, pad.AbsoluteCenter?.X ?? 0, pad.AbsoluteCenter?.Y ?? 0))
                    .ToList());

        private static List<KiCadSymbolInfo> Symbols(KiCadProjectRoot root) =>
            root.Schematics
                .SelectMany(schematic => schematic.Symbols)
                .Select(symbol => new KiCadSymbolInfo(
                    symbol.Reference ?? string.Empty,
                    symbol.Value ?? string.Empty,
                    symbol.PropertiesDetailed
                        .FirstOrDefault(property => string.Equals(property.Name, "Datasheet", StringComparison.OrdinalIgnoreCase))
                        ?.Value ?? string.Empty))
                .ToList();

        // `board` with every component of `root` it lacks added (KiCadBoardComponents.WithComponents).
        public static BoardData WithComponents(BoardData board, KiCadProjectRoot root, out KiCadComponentsAdded added) =>
            KiCadBoardComponents.WithComponents(board, KiCadComponentTable.Of(root), label => KiCadPinTable.PinsOf(root, label), out added);
    }
}
