using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Per-net primitive cache for the KiCad overlay, extracted from TabSchematics so the part that
    // decides whether to reuse or rebuild can be tested without a display.
    //
    // The overlay draws every net of a board. A hover changes how two of them look, so rebuilding
    // all 249 wastes a fifth of a second; this keeps each net's primitives against the state they
    // were built for and hands them back when that state is unchanged.
    //
    // Serving an entry that should have been thrown away draws copper for a state that no longer
    // applies - a moved calibration box, another board, a toggled setting - and it looks entirely
    // plausible on screen. That is why the decision lives here rather than inline: it is the part
    // that can be proven, and it is generic over the primitive type so it never has to know what a
    // primitive is.
    // ###########################################################################################
    public sealed class KiCadOverlayNetCache<TPrimitive>
    {
        private readonly Dictionary<string, (KiCadNetAppearance Appearance, IReadOnlyList<TPrimitive> Primitives)>
            thisEntries = new(StringComparer.OrdinalIgnoreCase);

        // Null rather than empty, so the very first rebuild always registers its generation even if
        // that generation happens to be an empty string.
        private string? thisGeneration;

        public int Count => this.thisEntries.Count;

        // ###########################################################################################
        // Starts a rebuild against the given shared-state key, emptying the cache when that key has
        // changed. Returns true when it did, which callers can use for diagnostics.
        //
        // Everything in the generation key moves geometry rather than merely recolouring it, so a
        // change to any of it makes every entry stale - not just some.
        // ###########################################################################################
        public bool BeginRebuild(string? generationKey)
        {
            string key = generationKey ?? string.Empty;

            if (this.thisGeneration != null && string.Equals(this.thisGeneration, key, StringComparison.Ordinal))
            {
                return false;
            }

            this.thisEntries.Clear();
            this.thisGeneration = key;
            return true;
        }

        // ###########################################################################################
        // Returns a net's primitives when they were built for exactly this appearance.
        // ###########################################################################################
        public bool TryGet(
            string? netId,
            KiCadNetAppearance appearance,
            out IReadOnlyList<TPrimitive> primitives)
        {
            primitives = Array.Empty<TPrimitive>();

            if (string.IsNullOrWhiteSpace(netId))
            {
                return false;
            }

            if (!this.thisEntries.TryGetValue(netId, out var entry) || !entry.Appearance.Equals(appearance))
            {
                return false;
            }

            primitives = entry.Primitives;
            return true;
        }

        // ###########################################################################################
        // Records the primitives built for one net under the appearance they were built for.
        // ###########################################################################################
        public void Store(string? netId, KiCadNetAppearance appearance, IReadOnlyList<TPrimitive> primitives)
        {
            if (string.IsNullOrWhiteSpace(netId) || primitives == null)
            {
                return;
            }

            this.thisEntries[netId] = (appearance, primitives);
        }

        // ###########################################################################################
        // Drops every entry, leaving the current generation in place.
        // ###########################################################################################
        public void Clear()
        {
            this.thisEntries.Clear();
        }
    }

    // ###########################################################################################
    // The per-net state that decides what one net's primitives look like.
    //
    // Only state that genuinely varies between nets belongs here. Shared state belongs in the
    // generation key instead - putting the hovered component here as a bare label rather than "the
    // hovered component, if it sits on this net" would make every net differ on every hover and
    // defeat the cache entirely, which is how this went wrong the first time.
    // ###########################################################################################
    public readonly record struct KiCadNetAppearance
    {
        public KiCadNetAppearance(
            bool isExplicitHighlight,
            bool isHovered,
            bool shouldBlink,
            double opacity,
            string? hoveredComponentIfOnThisNet)
        {
            this.IsExplicitHighlight = isExplicitHighlight;
            this.IsHovered = isHovered;
            this.ShouldBlink = shouldBlink;

            // Rounded here rather than at the call site so it cannot be forgotten. Opacity comes from
            // a blink factor, and two visually identical values that differ in the last bits would
            // otherwise never compare equal - every net would miss and the cache would do nothing.
            //
            // The NaN mapping is normalisation, not a correctness fix, and it is worth being precise
            // about why: record equality compares through EqualityComparer<double>, which reports
            // NaN as equal to NaN. It is bare "==" that does not. So a NaN opacity would still match
            // itself here; it is folded to zero only so a malformed value cannot travel any further.
            this.Opacity = double.IsNaN(opacity) || double.IsInfinity(opacity)
                ? 0.0
                : Math.Round(opacity, 4);

            this.HoveredComponentIfOnThisNet = hoveredComponentIfOnThisNet?.Trim() ?? string.Empty;
        }

        public bool IsExplicitHighlight { get; }

        public bool IsHovered { get; }

        public bool ShouldBlink { get; }

        public double Opacity { get; }

        public string HoveredComponentIfOnThisNet { get; }
    }
}
