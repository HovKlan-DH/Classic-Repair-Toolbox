using Avalonia.Input;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// PanGesture - CRT's one way of moving a picture, the Schematics tab's (owner request, 2026-11-06: "I do think we need to
// unify how pan works, and then have the same pan functionality as present in the 'Schematics' tab"): the right or the
// middle button pans, never the left, which selects, draws or moves what is under it.
// ###########################################################################################
public sealed class PanGestureTests
{
    [Theory]
    [InlineData(RawInputModifiers.RightMouseButton, true)]
    [InlineData(RawInputModifiers.MiddleMouseButton, true)]
    [InlineData(RawInputModifiers.LeftMouseButton, false)]
    [InlineData(RawInputModifiers.None, false)]
    public void Only_the_right_or_the_middle_button_pans(RawInputModifiers buttons, bool pans)
    {
        Assert.Equal(pans, PanGesture.IsPanPress(new PointerPointProperties(buttons, PointerUpdateKind.Other)));
    }

    [Fact]
    public void The_tip_names_both_buttons_and_the_wheel()
    {
        Assert.Contains("right or middle mouse button", PanGesture.Hint, StringComparison.Ordinal);
        Assert.Contains("scroll to zoom", PanGesture.Hint, StringComparison.Ordinal);
    }
}
