using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Tests for KiCadOverlayNetCache - the decision of whether a net's overlay primitives may be reused
// or must be rebuilt.
//
// Both ways of getting this wrong are silent, which is why it is worth testing rather than eyeing.
//
// Reusing an entry that should have been dropped draws copper built for a state that no longer
// applies - a moved calibration box, a different board, a toggled setting. Nothing throws and
// nothing goes red; the overlay is simply wrong in a way that looks plausible.
//
// Never reusing anything is just as quiet. It costs about a fifth of a second per rebuild and shows
// up only as the UI feeling sluggish. That is exactly what happened on the first attempt: a value
// that varies per net was treated as shared, so every hover cleared all 249 nets and the cache
// measured zero hits against 249 misses on every rebuild while appearing to work.
public sealed class KiCadOverlayNetCacheTests
{
    private static KiCadNetAppearance Appearance(
        bool explicitHighlight = false,
        bool hovered = false,
        bool blink = false,
        double opacity = 0.45,
        string? hoveredComponent = null) =>
        new(explicitHighlight, hovered, blink, opacity, hoveredComponent);

    private static KiCadOverlayNetCache<string> CacheAtGeneration(string generation = "gen-1")
    {
        var cache = new KiCadOverlayNetCache<string>();
        cache.BeginRebuild(generation);
        return cache;
    }

    // --------------------------------------------------------------------- reuse

    [Fact]
    public void A_net_stored_and_looked_up_unchanged_is_reused()
    {
        var cache = CacheAtGeneration();
        cache.Store("net-1", Appearance(), new[] { "a", "b" });

        Assert.True(cache.TryGet("net-1", Appearance(), out var primitives));
        Assert.Equal(new[] { "a", "b" }, primitives);
    }

    [Fact]
    public void The_stored_list_is_handed_back_rather_than_a_copy()
    {
        // The caller concatenates these into the frame's primitive list on every rebuild, so copying
        // would quietly reintroduce the allocation the cache exists to avoid.
        var cache = CacheAtGeneration();
        var stored = new[] { "a" };
        cache.Store("net-1", Appearance(), stored);

        cache.TryGet("net-1", Appearance(), out var primitives);

        Assert.Same(stored, primitives);
    }

    [Fact]
    public void A_net_that_was_never_stored_is_not_reused()
    {
        var cache = CacheAtGeneration();

        Assert.False(cache.TryGet("net-unknown", Appearance(), out var primitives));
        Assert.Empty(primitives);
    }

    [Fact]
    public void One_nets_primitives_are_never_served_for_another()
    {
        var cache = CacheAtGeneration();
        cache.Store("net-1", Appearance(), new[] { "one" });

        Assert.False(cache.TryGet("net-2", Appearance(), out _));
    }

    [Fact]
    public void Net_ids_match_regardless_of_case()
    {
        // Net ids come from contributed board files and are compared case-insensitively elsewhere;
        // differing here would turn every lookup into a miss on some boards and not others.
        var cache = CacheAtGeneration();
        cache.Store("Net-1", Appearance(), new[] { "a" });

        Assert.True(cache.TryGet("NET-1", Appearance(), out _));
    }

    // ------------------------------------------------------------- appearance changes

    [Theory]
    [InlineData("explicit highlight")]
    [InlineData("hovered")]
    [InlineData("blink")]
    [InlineData("opacity")]
    [InlineData("hovered component")]
    public void Any_change_of_appearance_forces_a_rebuild(string changedField)
    {
        var cache = CacheAtGeneration();
        cache.Store("net-1", Appearance(), new[] { "original" });

        KiCadNetAppearance changed = changedField switch
        {
            "explicit highlight" => Appearance(explicitHighlight: true),
            "hovered" => Appearance(hovered: true),
            "blink" => Appearance(blink: true),
            "opacity" => Appearance(opacity: 1.0),
            _ => Appearance(hoveredComponent: "U1"),
        };

        Assert.False(
            cache.TryGet("net-1", changed, out _),
            $"A change of {changedField} reused the cached primitives, so the net would draw with stale appearance.");
    }

    [Fact]
    public void A_net_untouched_by_the_change_still_reuses_while_its_neighbour_rebuilds()
    {
        // This is the whole point of caching per net rather than per rebuild: hovering a component
        // changes the nets it sits on and must leave the rest alone.
        var cache = CacheAtGeneration();
        cache.Store("net-hovered", Appearance(), new[] { "a" });
        cache.Store("net-elsewhere", Appearance(), new[] { "b" });

        Assert.False(cache.TryGet("net-hovered", Appearance(hoveredComponent: "U1"), out _));
        Assert.True(cache.TryGet("net-elsewhere", Appearance(), out _));
    }

    // ------------------------------------------------------------------- opacity

    [Fact]
    public void Opacities_differing_only_in_floating_point_noise_still_match()
    {
        // Opacity is derived from a blink factor, so the same visual state can arrive with different
        // low bits. Comparing them raw would make every net miss and the cache do nothing at all.
        var cache = CacheAtGeneration();
        cache.Store("net-1", Appearance(opacity: 0.45), new[] { "a" });

        Assert.True(cache.TryGet("net-1", Appearance(opacity: 0.45 + 1e-9), out _));
    }

    [Fact]
    public void Opacities_that_differ_visibly_do_not_match()
    {
        var cache = CacheAtGeneration();
        cache.Store("net-1", Appearance(opacity: 0.45), new[] { "a" });

        Assert.False(cache.TryGet("net-1", Appearance(opacity: 0.46), out _));
    }

    [Fact]
    public void A_malformed_opacity_still_matches_itself()
    {
        // Worth knowing which rule actually applies here, because the obvious one is wrong: record
        // equality compares doubles through EqualityComparer<double>, which reports NaN as equal to
        // NaN. Bare "==" does not. So this would hold even without the folding in the constructor -
        // the folding is there to stop a malformed value travelling further, not to make this pass.
        var cache = CacheAtGeneration();
        cache.Store("net-1", Appearance(opacity: double.NaN), new[] { "a" });

        Assert.True(cache.TryGet("net-1", Appearance(opacity: double.NaN), out _));
    }

    [Fact]
    public void A_hovered_component_label_is_matched_after_trimming()
    {
        var cache = CacheAtGeneration();
        cache.Store("net-1", Appearance(hoveredComponent: "U1"), new[] { "a" });

        Assert.True(cache.TryGet("net-1", Appearance(hoveredComponent: "  U1  "), out _));
    }

    [Fact]
    public void No_hovered_component_and_a_blank_one_are_the_same_state()
    {
        var cache = CacheAtGeneration();
        cache.Store("net-1", Appearance(hoveredComponent: null), new[] { "a" });

        Assert.True(cache.TryGet("net-1", Appearance(hoveredComponent: "   "), out _));
    }

    // ---------------------------------------------------------------- generation

    [Fact]
    public void A_new_generation_drops_every_net()
    {
        // Everything in the generation key moves geometry rather than recolouring it, so a change
        // makes all entries stale - not merely the ones that look different.
        var cache = CacheAtGeneration("gen-1");
        cache.Store("net-1", Appearance(), new[] { "a" });
        cache.Store("net-2", Appearance(), new[] { "b" });

        Assert.True(cache.BeginRebuild("gen-2"));

        Assert.Equal(0, cache.Count);
        Assert.False(cache.TryGet("net-1", Appearance(), out _));
        Assert.False(cache.TryGet("net-2", Appearance(), out _));
    }

    [Fact]
    public void An_unchanged_generation_keeps_every_net()
    {
        var cache = CacheAtGeneration("gen-1");
        cache.Store("net-1", Appearance(), new[] { "a" });

        Assert.False(cache.BeginRebuild("gen-1"));
        Assert.True(cache.TryGet("net-1", Appearance(), out _));
    }

    [Fact]
    public void Returning_to_a_previous_generation_does_not_resurrect_its_entries()
    {
        // Going back to an earlier view must not serve entries built before whatever happened in
        // between; they were dropped when the generation first changed and are gone for good.
        var cache = CacheAtGeneration("gen-1");
        cache.Store("net-1", Appearance(), new[] { "a" });

        cache.BeginRebuild("gen-2");
        cache.BeginRebuild("gen-1");

        Assert.False(cache.TryGet("net-1", Appearance(), out _));
    }

    [Fact]
    public void Storing_after_a_generation_change_survives_the_next_rebuild_at_that_generation()
    {
        var cache = CacheAtGeneration("gen-1");
        cache.BeginRebuild("gen-2");
        cache.Store("net-1", Appearance(), new[] { "fresh" });

        Assert.False(cache.BeginRebuild("gen-2"));
        Assert.True(cache.TryGet("net-1", Appearance(), out var primitives));
        Assert.Equal(new[] { "fresh" }, primitives);
    }

    [Fact]
    public void The_very_first_rebuild_registers_its_generation_even_when_it_is_empty()
    {
        // A blank generation is what an unresolvable view produces. If the cache started out already
        // believing it was at that generation, the first rebuild would not register and a later one
        // at a real generation would not clear.
        var cache = new KiCadOverlayNetCache<string>();

        Assert.True(cache.BeginRebuild(string.Empty));
        Assert.False(cache.BeginRebuild(string.Empty));
    }

    [Fact]
    public void A_null_generation_is_treated_as_blank_rather_than_throwing()
    {
        var cache = new KiCadOverlayNetCache<string>();

        Assert.True(cache.BeginRebuild(null));
        Assert.False(cache.BeginRebuild(string.Empty));
    }

    // --------------------------------------------------------------------- misuse

    [Fact]
    public void A_blank_net_id_is_neither_stored_nor_reused()
    {
        var cache = CacheAtGeneration();
        cache.Store("   ", Appearance(), new[] { "a" });

        Assert.Equal(0, cache.Count);
        Assert.False(cache.TryGet("   ", Appearance(), out _));
        Assert.False(cache.TryGet(null, Appearance(), out _));
    }

    [Fact]
    public void Clearing_empties_the_cache_without_disturbing_the_generation()
    {
        var cache = CacheAtGeneration("gen-1");
        cache.Store("net-1", Appearance(), new[] { "a" });

        cache.Clear();

        Assert.Equal(0, cache.Count);
        Assert.False(cache.BeginRebuild("gen-1"));
    }

    [Fact]
    public void Storing_a_net_twice_keeps_the_most_recent_appearance()
    {
        var cache = CacheAtGeneration();
        cache.Store("net-1", Appearance(), new[] { "old" });
        cache.Store("net-1", Appearance(hovered: true), new[] { "new" });

        Assert.False(cache.TryGet("net-1", Appearance(), out _));
        Assert.True(cache.TryGet("net-1", Appearance(hovered: true), out var primitives));
        Assert.Equal(new[] { "new" }, primitives);
    }
}
