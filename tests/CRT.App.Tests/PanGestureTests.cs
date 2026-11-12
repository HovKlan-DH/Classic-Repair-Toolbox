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
        foreach (string tip in (string[])[PanGesture.Hint, PanGesture.SchematicImageHint])
        {
            Assert.Contains("right or middle mouse button", tip, StringComparison.Ordinal);
            Assert.Contains("scroll to zoom", tip, StringComparison.Ordinal);
        }
    }

    // The Schematics tab's tip says "schematic image", never "picture" (owner request, 2026-11-11: "Always refer to the
    // 'picture' in the 'Schematics' tab as a 'schematic image'"); the photo viewer's moves a picture.
    [Fact]
    public void The_Schematics_tab_s_tip_moves_the_schematic_image()
    {
        Assert.Contains("move the schematic image", PanGesture.SchematicImageHint, StringComparison.Ordinal);
        Assert.DoesNotContain("picture", PanGesture.SchematicImageHint, StringComparison.OrdinalIgnoreCase);
    }
}
