using Avalonia;
using Handlers.Geometry;
using System;
using Xunit;

namespace Classic_Repair_Toolbox.Tests;

// ###########################################################################################
// KiCadCalibrationTransform turns a change of calibration into a matrix, so a calibration drag
// can transform already-built overlay geometry instead of rebuilding all of it.
//
// The rule these tests exist to defend: applying the delta matrix to geometry built at
// calibration A must land in EXACTLY the same place as building that geometry at calibration B
// in the first place. If it does not, the traces drawn while dragging are simply wrong - and
// wrong in a way that looks entirely plausible on screen, because the traces still form a sane
// board, just misaligned against the image underneath. That is the whole point of calibration,
// so it would be silently defeated.
//
// ReferenceMapWorldToLocal below is a deliberate, literal re-implementation of
// TabSchematics.MapKiCadWorldToLocal (which is private to a UserControl and so unreachable from
// here). It is the oracle: the matrix is asserted against it rather than against hand-computed
// numbers, so an error in the derivation cannot be copied into the expectation. If that method
// ever changes, this copy must change with it or these tests will start defending the wrong map.
// ###########################################################################################
public class KiCadCalibrationTransformTests
{
    // The mapping TabSchematics actually performs, copied verbatim in behaviour.
    private static Point ReferenceMapWorldToLocal(
        double worldX,
        double worldY,
        Rect worldBounds,
        Rect contentRect,
        KiCadCalibrationValues calibration,
        double pixelWidth,
        double pixelHeight)
    {
        if (worldBounds.Width <= 0 || worldBounds.Height <= 0)
        {
            return new Point(contentRect.X, contentRect.Y);
        }

        double nx = (worldX - worldBounds.X) / worldBounds.Width;
        double ny = (worldY - worldBounds.Y) / worldBounds.Height;

        if (calibration.MirrorX)
        {
            nx = 1.0 - nx;
        }

        if (calibration.MirrorY)
        {
            ny = 1.0 - ny;
        }

        nx *= calibration.ScaleX;
        ny *= calibration.ScaleY;

        double localX = contentRect.X + (nx * contentRect.Width);
        double localY = contentRect.Y + (ny * contentRect.Height);

        if (pixelWidth > 0)
        {
            localX += calibration.OffsetX * (contentRect.Width / pixelWidth);
        }
        else
        {
            localX += calibration.OffsetX;
        }

        if (pixelHeight > 0)
        {
            localY += calibration.OffsetY * (contentRect.Height / pixelHeight);
        }
        else
        {
            localY += calibration.OffsetY;
        }

        return new Point(localX, localY);
    }

    // A board-shaped world box and viewport, deliberately with a non-zero origin on both: an
    // origin of zero hides an entire class of sign error in the constant term.
    private static readonly Rect WorldBounds = new(30.5, -12.25, 160.0, 95.0);

    private static readonly Rect ContentRect = new(17.0, 23.0, 1280.0, 760.0);

    private const double PixelWidth = 4220.0;

    private const double PixelHeight = 2941.0;

    // Spread across the board rather than clustered, so a transform that is right in the middle
    // and wrong at the edges (the signature of a bad mirror term) cannot pass.
    public static TheoryData<double, double> SamplePoints() => new()
    {
        { 30.5, -12.25 },
        { 190.5, 82.75 },
        { 110.0, 35.0 },
        { 45.75, 70.5 },
        { 175.25, -3.5 },
    };

    private static void AssertDeltaMatchesDirectBuild(
        KiCadCalibrationValues reference,
        KiCadCalibrationValues current,
        double worldX,
        double worldY)
    {
        bool built = KiCadCalibrationTransform.TryBuildDeltaTransform(
            reference,
            current,
            WorldBounds,
            ContentRect,
            PixelWidth,
            PixelHeight,
            out var transform);

        Assert.True(built);

        Point builtAtReference = ReferenceMapWorldToLocal(
            worldX, worldY, WorldBounds, ContentRect, reference, PixelWidth, PixelHeight);

        Point expected = ReferenceMapWorldToLocal(
            worldX, worldY, WorldBounds, ContentRect, current, PixelWidth, PixelHeight);

        Point actual = builtAtReference.Transform(transform);

        // Sub-thousandth of a pixel. Anything looser would let a real misalignment through, since
        // these are already in on-screen local units.
        Assert.Equal(expected.X, actual.X, 6);
        Assert.Equal(expected.Y, actual.Y, 6);
    }

    [Theory]
    [MemberData(nameof(SamplePoints))]
    public void A_pure_offset_change_lands_where_a_direct_rebuild_would(double worldX, double worldY)
    {
        // Dragging the box bodily across the image - the Move handle - changes only the offsets.
        var reference = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, false, false);
        var current = new KiCadCalibrationValues(0.82, 0.77, 505.0, -90.0, false, false);

        AssertDeltaMatchesDirectBuild(reference, current, worldX, worldY);
    }

    [Theory]
    [MemberData(nameof(SamplePoints))]
    public void A_scale_change_lands_where_a_direct_rebuild_would(double worldX, double worldY)
    {
        // Dragging a corner handle - the box grows or shrinks about its stored edges.
        var reference = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, false, false);
        var current = new KiCadCalibrationValues(0.61, 0.94, 120.0, 260.0, false, false);

        AssertDeltaMatchesDirectBuild(reference, current, worldX, worldY);
    }

    [Theory]
    [MemberData(nameof(SamplePoints))]
    public void A_combined_scale_and_offset_change_lands_where_a_direct_rebuild_would(
        double worldX,
        double worldY)
    {
        // The realistic case: dragging one corner moves an edge AND resizes at the same time.
        var reference = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, false, false);
        var current = new KiCadCalibrationValues(1.13, 0.58, -45.0, 415.0, false, false);

        AssertDeltaMatchesDirectBuild(reference, current, worldX, worldY);
    }

    [Theory]
    [MemberData(nameof(SamplePoints))]
    public void Acquiring_a_horizontal_mirror_lands_where_a_direct_rebuild_would(
        double worldX,
        double worldY)
    {
        // Mirroring is how a flipped board is represented, and it is reached BY dragging an edge
        // across its opposite - so a drag really does cross this boundary mid-gesture.
        var reference = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, false, false);
        var current = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, true, false);

        AssertDeltaMatchesDirectBuild(reference, current, worldX, worldY);
    }

    [Theory]
    [MemberData(nameof(SamplePoints))]
    public void Acquiring_a_vertical_mirror_lands_where_a_direct_rebuild_would(
        double worldX,
        double worldY)
    {
        var reference = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, false, false);
        var current = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, false, true);

        AssertDeltaMatchesDirectBuild(reference, current, worldX, worldY);
    }

    [Theory]
    [MemberData(nameof(SamplePoints))]
    public void Losing_a_mirror_lands_where_a_direct_rebuild_would(double worldX, double worldY)
    {
        // The reverse direction: geometry built mirrored, dragged back to unmirrored. Asserted
        // separately because the delta is not symmetric - it divides by the REFERENCE coefficient.
        var reference = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, true, true);
        var current = new KiCadCalibrationValues(0.9, 0.7, 60.0, 310.0, false, false);

        AssertDeltaMatchesDirectBuild(reference, current, worldX, worldY);
    }

    [Theory]
    [MemberData(nameof(SamplePoints))]
    public void A_doubly_mirrored_change_lands_where_a_direct_rebuild_would(double worldX, double worldY)
    {
        var reference = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, true, true);
        var current = new KiCadCalibrationValues(0.55, 1.2, 480.0, -25.0, true, true);

        AssertDeltaMatchesDirectBuild(reference, current, worldX, worldY);
    }

    [Fact]
    public void An_unchanged_calibration_produces_the_identity_transform()
    {
        // Not a special case in the code - it falls out of the algebra - but it is the property
        // that makes a drag returning to its start draw exactly what it started with.
        var calibration = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, false, false);

        Assert.True(KiCadCalibrationTransform.TryBuildDeltaTransform(
            calibration,
            calibration,
            WorldBounds,
            ContentRect,
            PixelWidth,
            PixelHeight,
            out var transform));

        Assert.Equal(1.0, transform.M11, 9);
        Assert.Equal(1.0, transform.M22, 9);
        Assert.Equal(0.0, transform.M31, 9);
        Assert.Equal(0.0, transform.M32, 9);
    }

    [Fact]
    public void A_mirror_only_change_is_a_genuine_flip_rather_than_a_no_op()
    {
        // Guards against the delta collapsing to identity whenever the scales happen to match:
        // a mirror must invert the axis, so the X scale is negative and the Y scale is not.
        var reference = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, false, false);
        var current = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, true, false);

        Assert.True(KiCadCalibrationTransform.TryBuildDeltaTransform(
            reference,
            current,
            WorldBounds,
            ContentRect,
            PixelWidth,
            PixelHeight,
            out var transform));

        Assert.Equal(-1.0, transform.M11, 9);
        Assert.Equal(1.0, transform.M22, 9);
    }

    [Fact]
    public void A_degenerate_reference_calibration_refuses_to_build_a_transform()
    {
        // A zero scale collapses the board to a line, so there is nothing to divide by and no
        // delta exists. The caller must rebuild instead of transforming, so this must say so
        // rather than handing back an infinity that poisons every coordinate silently.
        var reference = new KiCadCalibrationValues(0.0, 0.77, 120.0, 260.0, false, false);
        var current = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, false, false);

        Assert.False(KiCadCalibrationTransform.TryBuildDeltaTransform(
            reference,
            current,
            WorldBounds,
            ContentRect,
            PixelWidth,
            PixelHeight,
            out var transform));

        Assert.Equal(Matrix.Identity, transform);
    }

    [Fact]
    public void An_empty_world_or_content_rect_refuses_to_build_a_transform()
    {
        var calibration = new KiCadCalibrationValues(0.82, 0.77, 120.0, 260.0, false, false);

        Assert.False(KiCadCalibrationTransform.TryBuildDeltaTransform(
            calibration,
            calibration,
            new Rect(0, 0, 0, 0),
            ContentRect,
            PixelWidth,
            PixelHeight,
            out _));

        Assert.False(KiCadCalibrationTransform.TryBuildDeltaTransform(
            calibration,
            calibration,
            WorldBounds,
            new Rect(0, 0, 0, 0),
            PixelWidth,
            PixelHeight,
            out _));
    }

    [Fact]
    public void A_missing_image_size_falls_back_to_an_unscaled_offset()
    {
        // MapKiCadWorldToLocal uses the offset unscaled when there is no bitmap to measure
        // against, and the axis map has to agree with it or geometry built in that state jumps.
        var map = KiCadCalibrationTransform.BuildAxisMap(
            calibrationScale: 1.0,
            calibrationOffset: 40.0,
            isMirrored: false,
            worldBoundsOrigin: 0.0,
            worldBoundsSize: 100.0,
            contentOrigin: 0.0,
            contentSize: 200.0,
            imagePixelSize: 0.0);

        // World 0 maps to the constant alone, which here is just the raw offset.
        Assert.Equal(40.0, map.Constant, 9);
    }
}
