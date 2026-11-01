using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;

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
}
