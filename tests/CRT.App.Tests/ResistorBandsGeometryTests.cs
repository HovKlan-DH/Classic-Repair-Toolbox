using Avalonia;
using Handlers.Geometry;
using System.Collections.Generic;
using System.Linq;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Where the resistor picture's parts stand (ResistorBandsGeometry; owner request, 2026-11-07, after the picture the
// owner makes by hand) - the bands across the body as a resistor is printed, the pills under them never touching and
// never outside the picture.
// ###########################################################################################
public sealed class ResistorBandsGeometryTests
{
    // The digits and the multiplier from the left cap evenly to the middle, the tolerance apart near the right end.
    [Fact]
    public void The_digits_and_multiplier_stand_evenly_and_the_tolerance_apart_near_the_end()
    {
        Assert.Equal([0.115, 0.26, 0.405, 0.55, 0.88], ResistorBandsGeometry.BandShares(5, withTolerance: true).Select(share => System.Math.Round(share, 3)));
        Assert.Equal([0.115, 0.3325, 0.55, 0.88], ResistorBandsGeometry.BandShares(4, withTolerance: true).Select(share => System.Math.Round(share, 4)));
        Assert.Equal([0.115, 0.3325, 0.55], ResistorBandsGeometry.BandShares(3, withTolerance: false).Select(share => System.Math.Round(share, 4)));
        Assert.Empty(ResistorBandsGeometry.BandShares(0, withTolerance: false));
    }

    [Fact]
    public void Every_band_stands_on_the_body_and_the_leads_at_its_ends()
    {
        ResistorBandsGeometry.Layout layout = ResistorBandsGeometry.For(300, 4, withTolerance: true);

        Assert.Equal(4, layout.Bands.Count);
        Assert.All(layout.Bands, band =>
        {
            Assert.True(band.Left >= layout.Body.Left && band.Right <= layout.Body.Right);
            Assert.Equal(layout.Body.Height, band.Height);
        });

        Assert.Equal(0, layout.LeftLead.Left);
        Assert.Equal(300, layout.RightLead.Right, 3);
        Assert.True(layout.PillTop > layout.Body.Bottom);
        Assert.Equal(layout.PillTop + ResistorBandsGeometry.PillHeight + 1, layout.Height);
    }

    // A stem starts at the body's foot where its band is - a cap's foot lower than the middle's.
    [Fact]
    public void A_stem_starts_at_the_body_foot_under_its_band()
    {
        ResistorBandsGeometry.Layout layout = ResistorBandsGeometry.For(300, 4, withTolerance: true);

        Assert.Equal(layout.Body.Bottom, ResistorBandsGeometry.StemTop(layout, layout.Bands[0]));
        Assert.Equal(layout.Body.Bottom - layout.MiddleInset, ResistorBandsGeometry.StemTop(layout, layout.Bands[2]));
        Assert.Equal(layout.Body.Bottom, ResistorBandsGeometry.StemTop(layout, layout.Bands[3]));
    }

    // ###########################################################################################
    // The pills are centred under their bands while there is room, pushed right when a wide one would touch its
    // neighbour, and the row moved left as a whole when it runs out of the picture - never past its left edge.
    // ###########################################################################################
    [Fact]
    public void The_pills_stand_under_their_bands_apart_and_inside_the_picture()
    {
        IReadOnlyList<double> roomy = ResistorBandsGeometry.PillLefts([50, 100, 200], [20, 20, 40], 300);
        Assert.Equal([40, 90, 180], roomy);

        // A wide multiplier pill pushes the next one right, never onto it.
        double[] widths = [20, 60, 20];
        IReadOnlyList<double> crowded = ResistorBandsGeometry.PillLefts([50, 70, 90], widths, 300);
        Assert.False(ResistorBandsGeometry.Overlap(crowded, widths));
        Assert.Equal(40, crowded[0]);
        Assert.True(crowded[1] >= crowded[0] + widths[0] + ResistorBandsGeometry.PillGap);

        // Running past the right edge, the row moves left.
        IReadOnlyList<double> atTheEnd = ResistorBandsGeometry.PillLefts([150, 290], [40, 40], 300);
        Assert.Equal(260, atTheEnd[1] + 0, 3);
        Assert.True(atTheEnd[1] + 40 <= 300);
        Assert.Equal(120, atTheEnd[0], 3);
    }

    [Fact]
    public void A_row_wider_than_the_picture_starts_at_its_left_edge()
    {
        IReadOnlyList<double> lefts = ResistorBandsGeometry.PillLefts([10, 20], [80, 80], 100);

        Assert.Equal(0, lefts[0]);
        Assert.False(ResistorBandsGeometry.Overlap(lefts, [80, 80]));
    }
}
