using Avalonia;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Identity keys for the KiCad overlay's per-net primitive cache, extracted from TabSchematics so
    // they can be tested without a display.
    //
    // The overlay caches each net's primitives so a hover change rebuilds only the nets it actually
    // affects instead of all of them. That cache is only as good as the keys: anything the primitives
    // depend on that is missing from a key means stale copper drawn on screen, which looks entirely
    // plausible and will not fail a build.
    //
    // The split between the two keys is the part worth understanding, because getting it wrong has
    // already happened once. State shared by every net belongs in the generation key, and changing it
    // drops the whole cache. State that varies per net belongs in the net key. The hovered component
    // was mistakenly treated as shared, so every hover cleared all 249 nets and the cache never once
    // hit - it only changes pin-1 marking on the pads of that one component, which is per-net.
    // ###########################################################################################
    public static class KiCadOverlayCacheKeys
    {
        // ###########################################################################################
        // Builds the key for state shared by every net. When this changes, every cached net is stale,
        // because all of these move geometry rather than merely recolour it.
        //
        // This is deliberately exhaustive rather than minimal: it is built once per rebuild, not once
        // per net, so an unnecessary field costs nothing while a missing one is a rendering bug.
        // ###########################################################################################
        public static string BuildGenerationKey(KiCadOverlaySharedState state)
        {
            if (state == null)
            {
                return string.Empty;
            }

            return string.Join(
                "|",
                // Board identity leads, because view ids such as "pcb:0:top" repeat across boards -
                // without it, switching board could hand one board's primitives to another.
                state.BoardScopeKey ?? string.Empty,
                state.SchematicName ?? string.Empty,
                state.ViewId ?? string.Empty,
                state.ViewSourceIndex.ToString(CultureInfo.InvariantCulture),
                state.PrimaryLayer ?? string.Empty,
                state.ContentRect.ToString(),
                state.WorldBounds.ToString(),
                KiCadOverlayCacheKeys.Number(state.CalibrationScaleX),
                KiCadOverlayCacheKeys.Number(state.CalibrationScaleY),
                KiCadOverlayCacheKeys.Number(state.CalibrationOffsetX),
                KiCadOverlayCacheKeys.Number(state.CalibrationOffsetY),
                state.CalibrationMirrorX.ToString(),
                state.CalibrationMirrorY.ToString(),
                state.OverlayColor.ToString(),
                state.OppositeTraceColor.ToString(),
                KiCadOverlayCacheKeys.Number(state.TranslatedOpacity),
                state.ShowOppositeSideTraces.ToString(),
                state.ShowZones.ToString(),
                state.IsCalibrationMode.ToString(),
                KiCadOverlayCacheKeys.References(state.ActiveReferences),
                KiCadOverlayCacheKeys.References(state.SelectedReferences));
        }

        // ###########################################################################################
        // Returns the hovered component's label if it has a pad on this net, and an empty string if
        // it does not.
        //
        // Hovering a component changes how its own pads are marked, so only the nets that component
        // sits on are affected. Feeding the label straight into the net key instead would invalidate
        // every net on every hover, which is exactly the bug this replaced.
        // ###########################################################################################
        public static string ResolveHoveredComponentForNet(
            string? hoveredComponentLabel,
            IEnumerable<string?> netPadReferences)
        {
            string hovered = hoveredComponentLabel?.Trim() ?? string.Empty;

            if (hovered.Length == 0 || netPadReferences == null)
            {
                return string.Empty;
            }

            foreach (string? reference in netPadReferences)
            {
                if (string.Equals(reference?.Trim(), hovered, StringComparison.OrdinalIgnoreCase))
                {
                    return hovered;
                }
            }

            return string.Empty;
        }

        // ###########################################################################################
        // Formats a double so two runs that produced the same value always produce the same text.
        // "R" round-trips, so values that differ in the last bits stay distinguishable.
        // ###########################################################################################
        private static string Number(double value)
        {
            return double.IsNaN(value) || double.IsInfinity(value)
                ? "x"
                : value.ToString("R", CultureInfo.InvariantCulture);
        }

        // ###########################################################################################
        // Formats a reference set order-independently. The same selection reached in a different order
        // is the same selection, and sorting keeps it from looking like a change.
        // ###########################################################################################
        private static string References(IReadOnlyCollection<string>? references)
        {
            if (references == null || references.Count == 0)
            {
                return string.Empty;
            }

            return string.Join(
                ",",
                references
                    .Where(reference => !string.IsNullOrWhiteSpace(reference))
                    .Select(reference => reference.Trim())
                    .OrderBy(reference => reference, StringComparer.OrdinalIgnoreCase));
        }
    }

    // ###########################################################################################
    // The overlay state that every net shares. Adding a field here and to BuildGenerationKey is how
    // a new global input gets covered; the tests assert that every field actually changes the key.
    // ###########################################################################################
    public sealed record KiCadOverlaySharedState
    {
        public string BoardScopeKey { get; init; } = string.Empty;
        public string SchematicName { get; init; } = string.Empty;
        public string ViewId { get; init; } = string.Empty;
        public int ViewSourceIndex { get; init; }
        public string PrimaryLayer { get; init; } = string.Empty;
        public Rect ContentRect { get; init; }
        public Rect WorldBounds { get; init; }
        public double CalibrationScaleX { get; init; } = 1.0;
        public double CalibrationScaleY { get; init; } = 1.0;
        public double CalibrationOffsetX { get; init; }
        public double CalibrationOffsetY { get; init; }
        public bool CalibrationMirrorX { get; init; }
        public bool CalibrationMirrorY { get; init; }
        public Color OverlayColor { get; init; }
        public Color OppositeTraceColor { get; init; }
        public double TranslatedOpacity { get; init; }
        public bool ShowOppositeSideTraces { get; init; }
        public bool ShowZones { get; init; }
        public bool IsCalibrationMode { get; init; }
        public IReadOnlyCollection<string> ActiveReferences { get; init; } = Array.Empty<string>();
        public IReadOnlyCollection<string> SelectedReferences { get; init; } = Array.Empty<string>();
    }
}
