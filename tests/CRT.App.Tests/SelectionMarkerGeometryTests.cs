using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// The eight corner/side marker segments that say "this rectangle can be resized". Two overlays
// draw them - the component label editor's selected highlight and a hovered worklog area - and
// they used to carry a verbatim copy of this maths each, so a change to one silently diverged from
// the other. Sharing it is what makes "the two look identical" true; these pin the layout.
public class SelectionMarkerGeometryTests
{
    private static readonly Rect Box = new(100, 100, 200, 120);

    // Four corners, each an L of two arms, plus two horizontal and two vertical side markers.
    [Fact]
    public void A_roomy_rectangle_gets_all_twelve_marker_segments()
    {
        var rects = SelectionMarkerGeometry.BuildSelectionMarkerRects(Box, 1.0);

        Assert.Equal(12, rects.Count);
    }

    // Every marker must touch the rectangle it belongs to. A marker floating away from the border
    // would advertise a grab point that is not where the hit rectangles actually are.
    [Fact]
    public void Every_marker_sits_on_the_rectangles_border()
    {
        var rects = SelectionMarkerGeometry.BuildSelectionMarkerRects(Box, 1.0);

        foreach (var marker in rects)
        {
            bool touchesVerticalEdge =
                Math.Abs(marker.Left - Box.Left) < 3.0 || Math.Abs(marker.Right - Box.Right) < 3.0;
            bool touchesHorizontalEdge =
                Math.Abs(marker.Top - Box.Top) < 3.0 || Math.Abs(marker.Bottom - Box.Bottom) < 3.0;

            Assert.True(
                touchesVerticalEdge || touchesHorizontalEdge,
                $"marker {marker} is not on the border of {Box}");
        }
    }

    // Side markers are dropped rather than allowed to run into the corner markers: two markers
    // meeting would read as one long edge handle, implying a resize the hit rects do not offer.
    [Fact]
    public void A_small_rectangle_drops_its_side_markers()
    {
        var rects = SelectionMarkerGeometry.BuildSelectionMarkerRects(new Rect(0, 0, 8, 8), 1.0);

        // The eight corner arms survive; the four side markers do not fit.
        Assert.Equal(8, rects.Count);
    }

    // Markers hold a constant SIZE on screen, so at higher zoom they must be smaller in board
    // coordinates - the same 1/scale rule the handle hit rects use.
    [Fact]
    public void Markers_shrink_in_board_space_as_the_board_is_zoomed_in()
    {
        var atUnit = SelectionMarkerGeometry.BuildSelectionMarkerRects(Box, 1.0)[0];
        var atFour = SelectionMarkerGeometry.BuildSelectionMarkerRects(Box, 4.0)[0];

        Assert.True(
            atFour.Width < atUnit.Width,
            $"expected a smaller marker at 4x zoom: {atFour.Width} vs {atUnit.Width}");
    }

    // A corner arm can never exceed half the rectangle, or the two corners on an edge would overlap
    // and the shape would read as a solid border rather than as corner handles.
    [Fact]
    public void Corner_markers_never_exceed_half_the_rectangle()
    {
        var narrow = new Rect(0, 0, 10, 400);
        var rects = SelectionMarkerGeometry.BuildSelectionMarkerRects(narrow, 1.0);

        foreach (var marker in rects)
        {
            Assert.True(
                marker.Width <= (narrow.Width / 2.0) + 3.0,
                $"marker {marker} is wider than half of {narrow}");
        }
    }

    // A degenerate scale must not divide by zero or produce infinities - the view matrix can report
    // one transiently during layout.
    [Fact]
    public void A_zero_scale_still_produces_finite_markers()
    {
        var rects = SelectionMarkerGeometry.BuildSelectionMarkerRects(Box, 0.0);

        Assert.NotEmpty(rects);
        foreach (var marker in rects)
        {
            Assert.True(double.IsFinite(marker.Width) && double.IsFinite(marker.Height), $"non-finite marker {marker}");
        }
    }
}
