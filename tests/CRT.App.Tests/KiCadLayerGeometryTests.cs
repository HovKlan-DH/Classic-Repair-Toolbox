using Avalonia;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Tests for KiCadLayerGeometry - board-side filtering, zone polygon selection and the KiCad
// naming rules. Previously private statics inside TabSchematics.
//
// Layer filtering is the difference between "these are the traces on the side you are looking
// at" and a plausible-looking overlay showing copper from the other side of the board.
public class KiCadLayerGeometryTests
{
    // ------------------------------------------------------- IsPointVisibleOnSide

    [Fact]
    public void Copper_on_the_requested_layer_is_visible()
    {
        Assert.True(KiCadLayerGeometry.IsPointVisibleOnSide(new[] { "F.Cu" }, "F.Cu"));
    }

    [Fact]
    public void Copper_on_the_other_side_is_not_visible()
    {
        Assert.False(KiCadLayerGeometry.IsPointVisibleOnSide(new[] { "B.Cu" }, "F.Cu"));
    }

    [Fact]
    public void Through_hole_copper_is_visible_from_both_sides()
    {
        // "*.Cu" means every copper layer - a through-hole pad must show on either view.
        Assert.True(KiCadLayerGeometry.IsPointVisibleOnSide(new[] { "*.Cu", "*.Mask" }, "F.Cu"));
        Assert.True(KiCadLayerGeometry.IsPointVisibleOnSide(new[] { "*.Cu", "*.Mask" }, "B.Cu"));
    }

    [Fact]
    public void Layer_matching_ignores_case_and_whitespace()
    {
        Assert.True(KiCadLayerGeometry.IsPointVisibleOnSide(new[] { "  f.cu  " }, "F.Cu"));
    }

    [Fact]
    public void An_item_with_no_layers_at_all_is_treated_as_visible()
    {
        // Some KiCad revisions omit the layer list; hiding that geometry would lose real copper.
        Assert.True(KiCadLayerGeometry.IsPointVisibleOnSide(Array.Empty<string>(), "F.Cu"));
    }

    [Fact]
    public void An_item_whose_layers_are_all_blank_is_not_visible()
    {
        // Blank entries are skipped, but the list is not empty, so the fallback does not apply.
        Assert.False(KiCadLayerGeometry.IsPointVisibleOnSide(new[] { "", "   " }, "F.Cu"));
    }

    [Fact]
    public void A_multi_layer_item_is_visible_when_any_layer_matches()
    {
        Assert.True(KiCadLayerGeometry.IsPointVisibleOnSide(new[] { "B.Cu", "In1.Cu", "F.Cu" }, "F.Cu"));
    }

    // ------------------------------------------------------------ zone visibility

    [Fact]
    public void A_zone_is_filtered_by_its_own_layers()
    {
        var zone = new KiCadPcbZone { Layers = { "B.Cu" } };

        Assert.True(KiCadLayerGeometry.IsZoneVisibleOnSide(zone, "B.Cu"));
        Assert.False(KiCadLayerGeometry.IsZoneVisibleOnSide(zone, "F.Cu"));
    }

    // ------------------------------------------------------ GetZoneWorldPolygons

    private static KiCadPcbZonePolygon Poly(params (double X, double Y)[] points)
    {
        var polygon = new KiCadPcbZonePolygon();
        foreach (var (x, y) in points)
        {
            polygon.Points.Add(new KiCadPoint2D { X = x, Y = y });
        }

        return polygon;
    }

    [Fact]
    public void Filled_polygons_are_preferred_over_the_outline()
    {
        // The filled polygons are the copper that actually got poured; the outline is only the
        // shape the designer drew.
        var zone = new KiCadPcbZone();
        zone.OutlinePolygons.Add(Poly((0, 0), (100, 0), (100, 100)));
        zone.FilledPolygons.Add(Poly((1, 1), (9, 1), (9, 9)));

        var polygons = KiCadLayerGeometry.GetZoneWorldPolygons(zone);

        Assert.Single(polygons);
        Assert.Equal(1, polygons[0][0].X);
    }

    [Fact]
    public void The_outline_is_used_when_nothing_was_poured()
    {
        var zone = new KiCadPcbZone();
        zone.OutlinePolygons.Add(Poly((0, 0), (10, 0), (10, 10)));

        var polygons = KiCadLayerGeometry.GetZoneWorldPolygons(zone);

        Assert.Single(polygons);
        Assert.Equal(3, polygons[0].Count);
    }

    [Fact]
    public void Polygons_with_fewer_than_three_points_are_dropped()
    {
        var zone = new KiCadPcbZone();
        zone.OutlinePolygons.Add(Poly((0, 0), (10, 0)));                 // a line, not an area
        zone.OutlinePolygons.Add(Poly((0, 0), (10, 0), (10, 10)));

        Assert.Single(KiCadLayerGeometry.GetZoneWorldPolygons(zone));
    }

    [Fact]
    public void A_zone_with_no_polygons_yields_none()
    {
        Assert.Empty(KiCadLayerGeometry.GetZoneWorldPolygons(new KiCadPcbZone()));
    }

    // ------------------------------------------------------- ComparePadDesignators

    [Fact]
    public void Numeric_pads_sort_numerically_not_as_text()
    {
        var pads = new[] { "10", "2", "1" };
        Array.Sort(pads, KiCadLayerGeometry.ComparePadDesignators);

        Assert.Equal(new[] { "1", "2", "10" }, pads);
    }

    [Fact]
    public void Numeric_pads_sort_before_lettered_pads()
    {
        // This is what lets a transistor footprint (B/C/E) pick a sensible primary pin while a
        // DIP still starts at 1.
        Assert.True(KiCadLayerGeometry.ComparePadDesignators("1", "B") < 0);
        Assert.True(KiCadLayerGeometry.ComparePadDesignators("B", "1") > 0);
    }

    [Fact]
    public void Lettered_pads_sort_alphabetically_ignoring_case()
    {
        var pads = new[] { "E", "b", "C" };
        Array.Sort(pads, KiCadLayerGeometry.ComparePadDesignators);

        Assert.Equal(new[] { "b", "C", "E" }, pads);
    }

    [Fact]
    public void Padding_and_nulls_are_tolerated()
    {
        Assert.Equal(0, KiCadLayerGeometry.ComparePadDesignators("  7  ", "7"));
        Assert.Equal(0, KiCadLayerGeometry.ComparePadDesignators(null, null));
    }

    // ------------------------------------------- TryExtractSchematicPageOrdinal

    [Theory]
    [InlineData("Schematics #1 of 2", 1)]
    [InlineData("Schematics #7 of 12", 7)]
    [InlineData("Anything #03 of 10", 3)]
    public void A_page_ordinal_is_read_from_the_schematic_name(string name, int expected)
    {
        Assert.True(KiCadLayerGeometry.TryExtractSchematicPageOrdinal(name, out int ordinal));
        Assert.Equal(expected, ordinal);
    }

    [Theory]
    [InlineData("Schematics")]              // no marker at all
    [InlineData("Schematics of 2")]         // no hash
    [InlineData("Schematics #1")]           // no " of "
    [InlineData("Schematics of 2 #1")]      // the wrong way round
    [InlineData("Schematics #0 of 2")]      // page 0 is not a page
    [InlineData("Schematics #x of 2")]      // no digits
    public void A_name_without_a_usable_ordinal_is_rejected(string name)
    {
        Assert.False(KiCadLayerGeometry.TryExtractSchematicPageOrdinal(name, out _));
    }

    // --------------------------------------------------- IsInternalSymbolReference

    [Theory]
    [InlineData("#PWR01")]
    [InlineData("#FLG0101")]
    public void Generated_power_and_flag_symbols_are_internal(string reference)
    {
        Assert.True(KiCadLayerGeometry.IsInternalSymbolReference(reference));
    }

    [Theory]
    [InlineData("U1")]
    [InlineData("R42")]
    [InlineData("")]
    public void Real_component_references_are_not_internal(string reference)
    {
        Assert.False(KiCadLayerGeometry.IsInternalSymbolReference(reference));
    }

    [Fact]
    public void A_null_reference_is_not_internal()
    {
        Assert.False(KiCadLayerGeometry.IsInternalSymbolReference(null!));
    }
}
