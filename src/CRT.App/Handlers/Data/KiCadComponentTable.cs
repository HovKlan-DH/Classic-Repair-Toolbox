using Handlers.Geometry;
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
        public static List<KiCadBoardComponent> Of(KiCadProjectRoot? root) =>
            root is null ? [] : KiCadBoardComponents.Of(KiCadComponentTable.Footprints(root), KiCadComponentTable.Symbols(root));

        // ###########################################################################################
        // The components KiCad knows but has no footprint for (2026-11-01) - nothing to mark them from,
        // so they are marked by hand (KiCadBoardComponents.OnlyInSchematic).
        // ###########################################################################################
        public static List<string> OnlyInSchematic(KiCadProjectRoot? root) =>
            root is null ? [] : KiCadBoardComponents.OnlyInSchematic(KiCadComponentTable.Footprints(root), KiCadComponentTable.Symbols(root));

        // ###########################################################################################
        // HOW MANY COMPONENTS EACH KiCad VIEW SHOWS, by the view's name, capitals ignored (owner request,
        // 2026-11-01: "per image I want to see status of how many components is available from KiCad"):
        // a side of the board, the components of its PCB file seen from that side - marking's own rule,
        // KiCadFootprintMarks.IsOnSide, so a through-hole part counts on both; a schematic sheet, the
        // parts drawn on it. Each once, however many units it has.
        // ###########################################################################################
        public static Dictionary<string, int> CountsByView(KiCadProjectRoot? root)
        {
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            if (root is null)
            {
                return counts;
            }

            foreach (KiCadProjectView view in root.Project.Views)
            {
                string name = view.DisplayName?.Trim() ?? string.Empty;

                if (name.Length > 0)
                {
                    counts[name] = KiCadComponentTable.CountOn(root, view);
                }
            }

            return counts;
        }

        private static int CountOn(KiCadProjectRoot root, KiCadProjectView view)
        {
            KiCadViewSide side = BoardKiCadFacts.SideOf(view);

            if (side != KiCadViewSide.Sheet)
            {
                if (view.SourceIndex < 0 || view.SourceIndex >= root.Pcb.Count)
                {
                    return 0;
                }

                return root.Pcb[view.SourceIndex].Footprints
                    .Where(footprint => KiCadBoardComponents.IsComponent(KiCadComponentTable.InfoOf(footprint)) &&
                                        KiCadFootprintMarks.IsOnSide(footprint, side == KiCadViewSide.Bottom ? "B.Cu" : "F.Cu"))
                    .Select(footprint => footprint.Reference?.Trim() ?? string.Empty)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count();
            }

            if (view.SourceIndex < 0 || view.SourceIndex >= root.Schematics.Count)
            {
                return 0;
            }

            return root.Schematics[view.SourceIndex].Symbols
                .Where(symbol => KiCadBoardComponents.IsComponentReference(symbol.Reference))
                .Select(symbol => symbol.Reference!.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count();
        }

        private static List<KiCadFootprintInfo> Footprints(KiCadProjectRoot root) =>
            root.Pcb
                .Take(1)
                .SelectMany(pcb => pcb.Footprints)
                .Select(KiCadComponentTable.InfoOf)
                .ToList();

        private static KiCadFootprintInfo InfoOf(KiCadPcbFootprint footprint) =>
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
