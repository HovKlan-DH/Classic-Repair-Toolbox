using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// TraceColours - how a picture's KiCad traces are drawn, this side's and the other side's (owner
// request, 2026-10-30: "be able to set the 'this side' trace color + the opacity for 'this side' and
// 'other side' traces"). The one rule the Schematics tab's copper and the Contribute tab's previews
// draw by - so a row that gives none of the new columns must draw exactly as before.
// ###########################################################################################
public sealed class TraceColoursTests
{
    // Nothing given: this side in the blank highlight colour, the other side DodgerBlue, both at the
    // highlight opacity (20 %) plus a quarter - what the Schematics tab always drew.
    [Fact]
    public void A_picture_with_no_colours_draws_as_before()
    {
        TraceColours none = TraceColours.Resolve(null);
        TraceColours blank = TraceColours.Resolve(new BoardSchematicEntry { SchematicName = "Top" });

        Assert.Equal(new TraceColours(Colors.DeepSkyBlue, 0.45, Colors.DodgerBlue, 0.45), none);
        Assert.Equal(none, blank);
    }

    // The highlight's colour and opacity carry this side's traces while their own columns are blank -
    // and both sides' opacity, a quarter above it, at most fully opaque.
    [Fact]
    public void Blank_trace_columns_follow_the_highlight()
    {
        TraceColours traces = TraceColours.Resolve(new BoardSchematicEntry
        {
            SchematicHighlightColor = "Red",
            SchematicHighlightOpacity = "90%",
            OppositeTraceHighlightColor = "Gold",
        });

        Assert.Equal(Colors.Red, traces.ThisSide);
        Assert.Equal(Colors.Gold, traces.OtherSide);
        Assert.Equal(1.0, traces.ThisSideOpacity, 3);
        Assert.Equal(1.0, traces.OtherSideOpacity, 3);
    }

    // Each column given is the picture's own - this side's colour apart from the highlight's, and the
    // two sides' opacities apart from each other.
    [Fact]
    public void Each_sides_colour_and_opacity_are_the_pictures_own_when_given()
    {
        TraceColours traces = TraceColours.Resolve(new BoardSchematicEntry
        {
            SchematicHighlightColor = "Red",
            SchematicHighlightOpacity = "30%",
            TraceHighlightColor = "#00FF00",
            TraceHighlightOpacity = "60%",
            OppositeTraceHighlightColor = "Orange",
            OppositeTraceHighlightOpacity = "15%",
        });

        Assert.Equal(Color.Parse("#00FF00"), traces.ThisSide);
        Assert.Equal(0.60, traces.ThisSideOpacity, 3);
        Assert.Equal(Colors.Orange, traces.OtherSide);
        Assert.Equal(0.15, traces.OtherSideOpacity, 3);
    }

    // Contributed text that cannot be read falls back as a blank column does - never a crash.
    [Fact]
    public void Unreadable_text_falls_back_as_blank()
    {
        TraceColours traces = TraceColours.Resolve(new BoardSchematicEntry
        {
            TraceHighlightColor = "not a colour",
            TraceHighlightOpacity = "lots",
            OppositeTraceHighlightOpacity = "%",
        });

        Assert.Equal(TraceColours.Resolve(null), traces);
    }

    // A new board's image: this side's traces in the Schematic image's color, and in whatever that is
    // changed to (owner request, 2026-11-03) - a blank row shows DeepSkyBlue beside an IndianRed highlight.
    [Fact]
    public void A_new_boards_image_has_its_traces_in_the_schematic_images_color()
    {
        BoardSchematicEntry row = new BoardSchematicEntry().WithColours(PictureColours.ForNewBoard);

        Assert.Equal(Colors.IndianRed, TraceColours.Resolve(row).ThisSide);
        Assert.Equal(Colors.Gold, TraceColours.Resolve(row.WithColours(PictureColours.ForNewBoard with { HighlightColor = "Gold" })).ThisSide);
    }

    // ###########################################################################################
    // *** AN OPACITY GIVEN HOLDS FOR THE NET UNDER THE POINTER TOO (owner report, 2026-11-03: "even if
    // I change the opacity down to 5% ... it uses a solid/100% opacity"). *** Hovering drew a net at 1.0
    // whatever the card said, and hovering is how most traces are ever seen.
    // ###########################################################################################
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void An_opacity_given_holds_for_the_net_under_the_pointer(bool hovered)
    {
        TraceColours traces = TraceColours.Resolve(new BoardSchematicEntry
        {
            TraceHighlightOpacity = "5%",
            OppositeTraceHighlightOpacity = "5%",
        });

        Assert.Equal(0.05, traces.ThisSideStroke(hovered), 3);
        Assert.Equal(0.05, traces.OtherSideStroke(hovered), 3);
    }

    // Each side by its own column: only the side the row says nothing of stands out under the pointer.
    [Fact]
    public void Each_side_goes_by_its_own_column_under_the_pointer()
    {
        TraceColours traces = TraceColours.Resolve(new BoardSchematicEntry { OppositeTraceHighlightOpacity = "5%" });

        Assert.Equal(1.0, traces.ThisSideStroke(hovered: true), 3);
        Assert.Equal(0.05, traces.OtherSideStroke(hovered: true), 3);
    }

    // Blank, nothing changes: a published board with no opacity of its own hovers as it always did.
    [Fact]
    public void With_no_opacity_given_the_net_under_the_pointer_stands_out_as_before()
    {
        TraceColours traces = TraceColours.Resolve(null);

        Assert.Equal(1.0, traces.ThisSideStroke(hovered: true), 3);
        Assert.Equal(1.0, traces.OtherSideStroke(hovered: true), 3);
        Assert.Equal(0.45, traces.ThisSideStroke(hovered: false), 3);
        Assert.Equal(0.45, traces.OtherSideStroke(hovered: false), 3);

        Assert.Equal(0.32, traces.ThisSideZoneFill(hovered: true), 3);
        Assert.Equal(Math.Clamp(0.45 * 0.65, 0.10, 0.38) + 0.12, traces.OtherSideZoneFill(hovered: true), 3);
        Assert.Equal(Math.Clamp(0.45 * 0.65, 0.10, 0.38), traces.ThisSideZoneFill(hovered: false), 3);
    }

    // A zone is never stronger than the copper it belongs to - its old 10 % floor was twice 5 % - and
    // keeps its old strength everywhere else.
    [Theory]
    [InlineData(0.0, 0.0)]
    [InlineData(0.05, 0.05)]
    [InlineData(0.10, 0.10)]
    [InlineData(0.20, 0.13)]
    [InlineData(0.45, 0.2925)]
    [InlineData(1.0, 0.38)]
    public void A_zone_is_lighter_than_its_copper(double opacity, double fill)
    {
        Assert.Equal(fill, TraceColours.ZoneFill(opacity), 4);
    }

    // A schematic sheet's wires have always been Orange while blank - and follow the card since
    // 2026-11-03, when "Traces on this side" changed nothing on a sheet.
    [Fact]
    public void A_schematic_sheet_keeps_its_own_blank_colour_and_follows_what_the_card_says()
    {
        Assert.Equal(Colors.Orange, TraceColours.Resolve(null, Colors.Orange).ThisSide);
        Assert.Equal(0.45, TraceColours.Resolve(null, Colors.Orange).ThisSideOpacity, 3);

        TraceColours given = TraceColours.Resolve(
            new BoardSchematicEntry { TraceHighlightColor = "Red", TraceHighlightOpacity = "5%" },
            Colors.Orange);

        Assert.Equal(Colors.Red, given.ThisSide);
        Assert.Equal(0.05, given.ThisSideOpacity, 3);
    }
}
