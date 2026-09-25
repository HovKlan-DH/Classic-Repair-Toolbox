using Avalonia;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for HighlightSpatialIndex.GetIsDrafted - the parallel "is this rect a
// local draft edit" flag added for NewContributeStrategy.md Phase 2, session 2b, so
// SchematicHighlightsOverlay can tint drafted highlights differently. The spatial query logic
// itself (Query/AddToCells) needs no display and is Avalonia.Rect-only, but was untested before
// this change; this file covers only the new surface, not a retroactive full suite for the class.
public sealed class HighlightSpatialIndexTests
{
    private static readonly Rect[] TwoRects = { new(0, 0, 10, 10), new(100, 100, 10, 10) };

    [Fact]
    public void GetIsDrafted_is_false_for_every_rect_when_no_drafted_array_is_given()
    {
        var index = new HighlightSpatialIndex(TwoRects);

        Assert.False(index.GetIsDrafted(0));
        Assert.False(index.GetIsDrafted(1));
    }

    [Fact]
    public void GetIsDrafted_reflects_the_parallel_array_passed_at_construction()
    {
        var index = new HighlightSpatialIndex(TwoRects, isDrafted: new[] { true, false });

        Assert.True(index.GetIsDrafted(0));
        Assert.False(index.GetIsDrafted(1));
    }

    [Fact]
    public void GetIsDrafted_is_false_for_an_index_the_drafted_array_does_not_reach()
    {
        // A shorter isDrafted list than rects would be a caller bug, but this must fail safe
        // (never drafted) rather than throwing an IndexOutOfRangeException mid-render.
        var index = new HighlightSpatialIndex(TwoRects, isDrafted: new[] { true });

        Assert.True(index.GetIsDrafted(0));
        Assert.False(index.GetIsDrafted(1));
    }

    [Fact]
    public void Count_and_GetRect_are_unaffected_by_the_new_parameter()
    {
        var index = new HighlightSpatialIndex(TwoRects, isDrafted: new[] { true, true });

        Assert.Equal(2, index.Count);
        Assert.Equal(TwoRects[0], index.GetRect(0));
        Assert.Equal(TwoRects[1], index.GetRect(1));
    }

    [Fact]
    public void Query_results_are_unaffected_by_the_new_parameter()
    {
        var index = new HighlightSpatialIndex(TwoRects, isDrafted: new[] { true, false });
        var results = new List<int>();

        index.Query(new Rect(0, 0, 20, 20), results);

        Assert.Equal(new[] { 0 }, results);
    }
}
