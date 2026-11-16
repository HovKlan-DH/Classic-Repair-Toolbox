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
                KiCadOverlayCacheKeys.Number(state.OppositeTranslatedOpacity),
                state.ThisSideOpacityGiven.ToString(),
                state.OtherSideOpacityGiven.ToString(),
                state.ShowOppositeSideTraces.ToString(),
                state.ShowZones.ToString(),
                state.IsCalibrationMode.ToString(),
                KiCadOverlayCacheKeys.References(state.SelectedReferences));
        }

        // ###########################################################################################
        // What of the ACTIVE references - the selected components, and the hovered one while "show traces
        // on the selected component" is on - a net's drawing depends on: none for a net drawn whole (hovered,
        // locked, an important signal), "*" with no reference active at all (every pad a target), else
        // the active references with a pad on this net, sorted (2026-11-16, performance walkthrough).
        //
        // KiCadNetGraphBuilder.BuildKiCadPcbActiveDrawIds walks a net from the pads of active references and
        // stops at foreign pads - pads of THIS net - so a reference with no pad on the net changes nothing
        // drawn there. The active set stood in the generation key until then, so every component the pointer
        // entered or left over a PCB picture with a component selected rebuilt every net drawn
        // (KiCadOverlayCacheKeysTests holds the two together).
        //
        // `active` is ActiveSetOf the active references, made ONCE for every net of a rebuild (code review, 2026-11-16: made
        // again for each, a dense board's 249 nets copied the whole active set 249 times at every rebuild).
        // ###########################################################################################
        public static string ActiveReferencesOnNet(
            bool isExplicitHighlight,
            IReadOnlySet<string>? active,
            IEnumerable<string?>? netPadReferences)
        {
            if (isExplicitHighlight)
            {
                return string.Empty;
            }

            if (active is null)
            {
                return "*";
            }

            var onThisNet = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (string? reference in netPadReferences ?? [])
            {
                string trimmed = reference?.Trim() ?? string.Empty;

                if (trimmed.Length > 0 && active.Contains(trimmed))
                {
                    onThisNet.Add(trimmed.ToUpperInvariant());
                }
            }

            // "+" before the list: a set holding no reference of this net is not "*" (nothing active at all).
            return "+" + string.Join(",", onThisNet);
        }

        // The active references as ActiveReferencesOnNet takes them - trimmed, capitals no difference - or null with none at
        // all (every pad a target).
        public static IReadOnlySet<string>? ActiveSetOf(IReadOnlyCollection<string>? activeReferences) =>
            activeReferences == null || activeReferences.Count == 0
                ? null
                : new HashSet<string>(
                    activeReferences.Where(reference => !string.IsNullOrWhiteSpace(reference)).Select(reference => reference.Trim()),
                    StringComparer.OrdinalIgnoreCase);

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

        // The other side's traces at an opacity of their own (2026-10-30, TraceColours).
        public double OppositeTranslatedOpacity { get; init; }

        // Whether the picture's row gives them - the same opacity worked out or given draws a net under the
        // pointer differently (TraceColours; code review, 2026-11-04).
        public bool ThisSideOpacityGiven { get; init; }
        public bool OtherSideOpacityGiven { get; init; }
        public bool ShowOppositeSideTraces { get; init; }
        public bool ShowZones { get; init; }
        public bool IsCalibrationMode { get; init; }

        // The ACTIVE references are per net - KiCadOverlayCacheKeys.ActiveReferencesOnNet, in the net's appearance.
        public IReadOnlyCollection<string> SelectedReferences { get; init; } = Array.Empty<string>();
    }
}
