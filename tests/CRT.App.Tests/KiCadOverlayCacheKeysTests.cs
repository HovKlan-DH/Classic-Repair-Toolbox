using Avalonia;
using Avalonia.Media;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Tests for KiCadOverlayCacheKeys - the identity keys behind the KiCad overlay's per-net primitive
// cache.
//
// This is a cache over what gets drawn, so the failure mode is not a crash or a red test: it is
// stale copper on screen that looks entirely plausible. Two mistakes are possible and these tests
// guard both.
//
// Leaving a shared input out of the generation key means the overlay keeps drawing geometry built
// for a state that no longer applies - a moved calibration box, a different board, a toggled
// setting. Every_field_changes_the_key exists so that dropping a field from the key is a failing
// test rather than a rendering bug found weeks later.
//
// Putting per-net state into the shared key is the opposite mistake and it has already happened
// here: the hovered component label went into the generation key, so every hover cleared all 249
// nets and the cache measured 0 hits against 249 misses on every single rebuild. Hovering only
// changes pin-1 marking on that component's own pads, so it belongs to the nets it touches.
public sealed class KiCadOverlayCacheKeysTests
{
    private static KiCadOverlaySharedState Baseline() => new()
    {
        BoardScopeKey = "board-a",
        SchematicName = "Top (replica)",
        ViewId = "pcb:0:top",
        ViewSourceIndex = 0,
        PrimaryLayer = "F.Cu",
        ContentRect = new Rect(0, 0, 800, 600),
        WorldBounds = new Rect(0, 0, 389, 135),
        CalibrationScaleX = 1.0,
        CalibrationScaleY = 1.0,
        CalibrationOffsetX = 0,
        CalibrationOffsetY = 0,
        CalibrationMirrorX = false,
        CalibrationMirrorY = false,
        OverlayColor = Colors.DeepSkyBlue,
        OppositeTraceColor = Colors.DodgerBlue,
        TranslatedOpacity = 0.45,
        ShowOppositeSideTraces = true,
        ShowZones = true,
        IsCalibrationMode = false,
        ActiveReferences = new[] { "U1", "U2" },
        SelectedReferences = new[] { "U1", "U2" },
    };

    private static string Key(KiCadOverlaySharedState state) =>
        KiCadOverlayCacheKeys.BuildGenerationKey(state);

    // ------------------------------------------------------------------- stability

    [Fact]
    public void The_same_state_always_produces_the_same_key()
    {
        // Without this the cache would never hit, and every rebuild would redo all the work.
        Assert.Equal(Key(Baseline()), Key(Baseline()));
    }

    [Fact]
    public void A_null_state_is_handled_rather_than_throwing()
    {
        Assert.Equal(string.Empty, KiCadOverlayCacheKeys.BuildGenerationKey(null!));
    }

    // ------------------------------------------------- every shared input must invalidate

    public static TheoryData<string, KiCadOverlaySharedState> ChangedStates()
    {
        var b = Baseline();

        return new TheoryData<string, KiCadOverlaySharedState>
        {
            { "BoardScopeKey", b with { BoardScopeKey = "board-b" } },
            { "SchematicName", b with { SchematicName = "Bottom (replica)" } },
            { "ViewId", b with { ViewId = "pcb:0:bottom" } },
            { "ViewSourceIndex", b with { ViewSourceIndex = 1 } },
            { "PrimaryLayer", b with { PrimaryLayer = "B.Cu" } },
            { "ContentRect", b with { ContentRect = new Rect(0, 0, 801, 600) } },
            { "WorldBounds", b with { WorldBounds = new Rect(0, 0, 390, 135) } },
            { "CalibrationScaleX", b with { CalibrationScaleX = 1.01 } },
            { "CalibrationScaleY", b with { CalibrationScaleY = 1.01 } },
            { "CalibrationOffsetX", b with { CalibrationOffsetX = 5 } },
            { "CalibrationOffsetY", b with { CalibrationOffsetY = 5 } },
            { "CalibrationMirrorX", b with { CalibrationMirrorX = true } },
            { "CalibrationMirrorY", b with { CalibrationMirrorY = true } },
            { "OverlayColor", b with { OverlayColor = Colors.Red } },
            { "OppositeTraceColor", b with { OppositeTraceColor = Colors.Red } },
            { "TranslatedOpacity", b with { TranslatedOpacity = 0.5 } },
            { "ShowOppositeSideTraces", b with { ShowOppositeSideTraces = false } },
            { "ShowZones", b with { ShowZones = false } },
            { "IsCalibrationMode", b with { IsCalibrationMode = true } },
            { "ActiveReferences", b with { ActiveReferences = new[] { "U1" } } },
            { "SelectedReferences", b with { SelectedReferences = new[] { "U1" } } },
        };
    }

    [Theory]
    [MemberData(nameof(ChangedStates))]
    public void Every_shared_input_changes_the_key(string fieldName, KiCadOverlaySharedState changed)
    {
        // If this fails, the named field is no longer part of the generation key - which means the
        // overlay will keep drawing geometry built before that field changed.
        Assert.True(
            Key(Baseline()) != Key(changed),
            $"Changing {fieldName} did not change the cache key, so the overlay would draw stale geometry.");
    }

    // --------------------------------------------------------------- reference sets

    [Fact]
    public void Reference_order_does_not_change_the_key()
    {
        // The same selection reached in a different order is the same selection. Without sorting,
        // re-selecting the same components would needlessly drop the entire cache.
        var forwards = Baseline() with { ActiveReferences = new[] { "U1", "U2", "U3" } };
        var backwards = Baseline() with { ActiveReferences = new[] { "U3", "U2", "U1" } };

        Assert.Equal(Key(forwards), Key(backwards));
    }

    [Fact]
    public void Blank_references_are_ignored()
    {
        var withBlanks = Baseline() with { ActiveReferences = new[] { "U1", "  ", "U2", "" } };
        var without = Baseline() with { ActiveReferences = new[] { "U1", "U2" } };

        Assert.Equal(Key(without), Key(withBlanks));
    }

    [Fact]
    public void An_empty_reference_set_is_not_the_same_as_one_with_a_reference()
    {
        // Nothing selected draws every net; one component selected does not. They must not share a key.
        var none = Baseline() with { ActiveReferences = Array.Empty<string>() };

        Assert.NotEqual(Key(Baseline()), Key(none));
    }

    [Fact]
    public void A_malformed_calibration_value_does_not_throw()
    {
        // Contributed board data is not always well formed, and a calibration box can be degenerate.
        var broken = Baseline() with { CalibrationScaleX = double.NaN, CalibrationOffsetY = double.PositiveInfinity };

        Assert.False(string.IsNullOrEmpty(Key(broken)));
    }

    // ------------------------------------------------------ hovered component is per-net

    [Fact]
    public void A_net_carrying_the_hovered_components_pad_reports_it()
    {
        Assert.Equal(
            "U1",
            KiCadOverlayCacheKeys.ResolveHoveredComponentForNet("U1", new[] { "R4", "U1", "C7" }));
    }

    [Fact]
    public void A_net_that_does_not_touch_the_hovered_component_reports_nothing()
    {
        // This is the whole point: nets away from the hovered component keep their key and stay
        // cached. Returning the label here instead would invalidate every net on every hover.
        Assert.Equal(
            string.Empty,
            KiCadOverlayCacheKeys.ResolveHoveredComponentForNet("U1", new[] { "R4", "C7" }));
    }

    [Fact]
    public void Hovering_nothing_reports_nothing()
    {
        Assert.Equal(string.Empty, KiCadOverlayCacheKeys.ResolveHoveredComponentForNet(null, new[] { "U1" }));
        Assert.Equal(string.Empty, KiCadOverlayCacheKeys.ResolveHoveredComponentForNet("   ", new[] { "U1" }));
    }

    [Fact]
    public void Pad_references_are_matched_ignoring_case_and_padding()
    {
        // Board data is contributed by hand, so a reference can arrive spaced or in another case.
        Assert.Equal(
            "u1",
            KiCadOverlayCacheKeys.ResolveHoveredComponentForNet("u1", new[] { " U1 " }));
    }

    [Fact]
    public void A_net_with_no_pads_reports_nothing()
    {
        Assert.Equal(string.Empty, KiCadOverlayCacheKeys.ResolveHoveredComponentForNet("U1", Array.Empty<string?>()));
        Assert.Equal(string.Empty, KiCadOverlayCacheKeys.ResolveHoveredComponentForNet("U1", null!));
    }

    [Fact]
    public void A_null_pad_reference_is_skipped_rather_than_matching()
    {
        Assert.Equal(
            string.Empty,
            KiCadOverlayCacheKeys.ResolveHoveredComponentForNet("U1", new string?[] { null, "R4" }));
    }
}
