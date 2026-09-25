using Avalonia;
using Handlers.Geometry;
using System.Linq;

namespace ClassicRepairToolbox.Tests;

// Where a worklog entry's "#N" badge lands on a schematic preview.
//
// The anchored-vs-parked split is the thing pinned here, and it is pinned because getting it
// backwards is a REPORTED bug, not a hypothetical one: an entry with "show marked area" ticked
// anchors its badge to the rectangle on the board, and one without has nothing to anchor to and
// parks in the corner. An earlier version anchored every badge regardless, so an entry meant to
// show only a parked pill still appeared pinned to wherever it happened to have been drawn.
//
// WorkbooksBoardPreviewTests still covers this end to end through a real headless layout pass; what
// these add is the rule itself, in milliseconds and without a window.
//
// No collection attribute: pure geometry, no statics, no controls, no filesystem.
public class WorklogBadgeLayoutTests
{
    private static readonly PixelSize Bitmap = new(100, 100);

    // A 100x100 content rect against a 100x100 bitmap, so one bitmap pixel is one local unit and
    // an anchor rect's expected position can be read straight off the input.
    private static readonly Rect Content = new(0, 0, 100, 100);

    private static WorklogBadgeLayout.BadgePlacementRequest Anchored(Rect area, double w = 20, double h = 10) =>
        new(new Size(w, h), area);

    private static WorklogBadgeLayout.BadgePlacementRequest Parked(double w = 20, double h = 10) =>
        new(new Size(w, h), null);

    // Straddling the area's TOP-LEFT corner, half the badge either side of it - which is what
    // BadgeGeometry.GetCenterScaledCentreOffset yields, and what the Schematics tab's own anchored
    // pills do. The point is that the position is derived from the area at all; a parked badge's is
    // not.
    [Fact]
    public void An_anchored_badge_straddles_its_marked_areas_corner()
    {
        var area = new Rect(40, 50, 20, 20);

        var positions = WorklogBadgeLayout.ArrangeBadges(
            new[] { Anchored(area, w: 20, h: 10) }, Content, Bitmap, 10.0, 6.0);

        Assert.Single(positions);
        Assert.Equal(40 - 20 / 2.0, positions[0].X, 3);
        Assert.Equal(50 - 10 / 2.0, positions[0].Y, 3);
    }

    // The bug, stated as an assertion: a parked badge must NOT sit on the area it was drawn at. Its
    // position depends only on the content rect and on how many parked badges precede it.
    [Fact]
    public void A_parked_badge_goes_to_the_top_right_rather_than_to_its_area()
    {
        var positions = WorklogBadgeLayout.ArrangeBadges(
            new[] { Parked(w: 20, h: 10) }, Content, Bitmap, 10.0, 6.0);

        Assert.Single(positions);

        // Inset from the right edge by the margin: 100 - 10 - 20.
        Assert.Equal(70.0, positions[0].X, 3);
        Assert.Equal(10.0, positions[0].Y, 3);
    }

    // Parked badges must not pile up on one another - several unticked entries on one schematic is
    // the ordinary case, and overlapping pills are unreadable and unclickable. Two of them go one
    // above the other (ParkedBadgeGeometry keeps a single column up to two badges).
    [Fact]
    public void Parked_badges_stack_instead_of_overlapping()
    {
        var positions = WorklogBadgeLayout.ArrangeBadges(
            new[] { Parked(h: 10), Parked(h: 10) }, Content, Bitmap, 10.0, 6.0);

        Assert.Equal(2, positions.Count);
        Assert.Equal(10.0, positions[0].Y, 3);
        Assert.Equal(10.0 + 10.0 + 6.0, positions[1].Y, 3);
        Assert.Equal(positions[0].X, positions[1].X, 3);
    }

    // Beyond two they wrap into a grid rather than running off the bottom of a tall single column -
    // ParkedBadgeGeometry's own rule, exercised here through the layout this pane actually calls so
    // a change to that rule shows up as a failure here too. All four distinct, none overlapping.
    [Fact]
    public void More_than_two_parked_badges_wrap_into_a_grid()
    {
        var positions = WorklogBadgeLayout.ArrangeBadges(
            new[] { Parked(), Parked(), Parked(), Parked() }, Content, Bitmap, 10.0, 6.0);

        Assert.Equal(4, positions.Count);
        Assert.Equal(4, positions.Distinct().Count());

        // Two columns, two rows: 0 and 1 share a row, 0 and 2 share a column.
        Assert.Equal(positions[0].Y, positions[1].Y, 3);
        Assert.Equal(positions[0].X, positions[2].X, 3);
        Assert.NotEqual(positions[0].X, positions[1].X);
        Assert.NotEqual(positions[0].Y, positions[2].Y);
    }

    // The two kinds share one canvas, and the result has to line up with the INPUT order so the
    // caller can zip positions straight back onto its controls. A mis-ordered result would put each
    // badge at another badge's position - silently, and only when a preview mixed the two kinds.
    [Fact]
    public void A_mixed_set_keeps_every_badge_in_its_input_position()
    {
        var area = new Rect(40, 50, 20, 20);

        var positions = WorklogBadgeLayout.ArrangeBadges(
            new[]
            {
                Parked(w: 20, h: 10),
                Anchored(area, w: 20, h: 10),
                Parked(w: 20, h: 10),
            },
            Content, Bitmap, 10.0, 6.0);

        Assert.Equal(3, positions.Count);

        // Index 1 is the anchored one and must be on its area, not in the corner stack.
        Assert.Equal(40 - 20 / 2.0, positions[1].X, 3);
        Assert.Equal(50 - 10 / 2.0, positions[1].Y, 3);

        // Indexes 0 and 2 are the parked ones - the two of them, ignoring the anchored one between,
        // so they stack as a pair rather than being spread as if there were three parked badges.
        Assert.Equal(70.0, positions[0].X, 3);
        Assert.Equal(10.0, positions[0].Y, 3);
        Assert.Equal(70.0, positions[2].X, 3);
        Assert.Equal(10.0 + 10.0 + 6.0, positions[2].Y, 3);
    }

    // Parked badges are placed against the drawn CONTENT rect, not the control's bounds. The two
    // differ whenever a Uniform-stretched image is letterboxed inside its control - which the
    // previews' MaxHeight cap makes reachable - and using the bounds put parked badges out in the
    // letterbox, outside the outline marking where the image actually is, while anchored badges on
    // the same canvas stayed correct.
    [Fact]
    public void Parked_badges_stay_inside_a_letterboxed_images_drawn_area()
    {
        // A control 200 wide whose image only draws 100 of them.
        var letterboxed = new Rect(0, 0, 100, 80);

        var positions = WorklogBadgeLayout.ArrangeBadges(
            new[] { Parked(w: 20, h: 10) }, letterboxed, Bitmap, 10.0, 6.0);

        // 100 - 10 - 20, i.e. against the CONTENT's right edge. Against a 200-wide control it would
        // have been 170, well outside the drawn image.
        Assert.Equal(70.0, positions[0].X, 3);
        Assert.True(positions[0].X + 20 <= letterboxed.Width, "parked badge overflowed the drawn image");
    }

    [Fact]
    public void No_badges_yields_no_positions()
    {
        Assert.Empty(WorklogBadgeLayout.ArrangeBadges(
            System.Array.Empty<WorklogBadgeLayout.BadgePlacementRequest>(), Content, Bitmap, 10.0, 6.0));

        Assert.Empty(WorklogBadgeLayout.ArrangeBadges(null!, Content, Bitmap, 10.0, 6.0));
    }
}
