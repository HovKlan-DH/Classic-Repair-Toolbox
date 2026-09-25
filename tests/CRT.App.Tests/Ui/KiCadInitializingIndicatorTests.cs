using Avalonia.Controls;
using Avalonia.Layout;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The "KiCad data initializing..." indicator in the bottom-right corner of the schematics view.
//
// A board's KiCad project loads in the background (TabSchematics.LoadKiCadProjectForCurrentBoardAsync)
// so the schematic image can be shown immediately. For the few seconds that load takes, hovering a
// trace, clicking a pad or picking a net does nothing at all, which is indistinguishable from a
// broken overlay. The indicator exists to make that wait visible, so what matters here is that it
// is hidden by default, that it can be turned on and off, and that it sits where the user was told
// to look for it without stealing pointer input from the schematic underneath.
//
// The load method itself is not driven here: it needs a Main window, real board data and a real
// KiCad project on disk. What is covered is the switch that method flips, plus the markup it
// flips it on.
// ###########################################################################################
[Collection("HeadlessUi")]
public class KiCadInitializingIndicatorTests
{
    [Fact]
    public void The_indicator_is_hidden_until_a_load_asks_for_it()
    {
        // Most boards have no KiCad data at all, so an indicator that defaulted to visible would
        // sit there forever on those boards.
        UiTest.Run(() => Assert.False(IndicatorOf(new TabSchematics()).IsVisible));
    }

    [Fact]
    public void Showing_the_indicator_puts_the_initializing_text_on_screen()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = new();

            tab.SetKiCadInitializingIndicatorVisible(true);

            Assert.True(IndicatorOf(tab).IsVisible);
            Assert.Equal("KiCad data initializing...", IndicatorTextOf(tab).Text);
        });
    }

    [Fact]
    public void Hiding_the_indicator_takes_it_back_off_the_schematic()
    {
        // This is the "once ready, remove the highlight" half - the load method calls it on
        // completion, on failure, and when the board is switched out from under a running load.
        UiTest.Run(() =>
        {
            TabSchematics tab = new();

            tab.SetKiCadInitializingIndicatorVisible(true);
            tab.SetKiCadInitializingIndicatorVisible(false);

            Assert.False(IndicatorOf(tab).IsVisible);
        });
    }

    [Fact]
    public void Showing_the_indicator_twice_leaves_it_showing()
    {
        // Both Main (when the board image goes up) and the load method itself switch it on, so
        // the two calls overlap by design and the second must not toggle the first back off.
        UiTest.Run(() =>
        {
            TabSchematics tab = new();

            tab.SetKiCadInitializingIndicatorVisible(true);
            tab.SetKiCadInitializingIndicatorVisible(true);

            Assert.True(IndicatorOf(tab).IsVisible);
        });
    }

    [Fact]
    public void The_indicator_sits_in_the_bottom_right_corner_and_ignores_the_pointer()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = new();

            // The corner comes from the shared bottom-right stack it lives in, above the traces
            // panel, so the two cannot land on top of each other when a board has saved traces.
            StackPanel stack = tab.FindControl<StackPanel>("SchematicsBottomRightOverlayStack")!;

            Assert.Equal(HorizontalAlignment.Right, stack.HorizontalAlignment);
            Assert.Equal(VerticalAlignment.Bottom, stack.VerticalAlignment);
            Assert.Same(stack, IndicatorOf(tab).Parent);

            // It is a status message, not a panel: wheel-zoom and click-drag over that corner
            // must keep reaching the schematic underneath.
            Assert.False(IndicatorOf(tab).IsHitTestVisible);
        });
    }

    private static Border IndicatorOf(TabSchematics tab) =>
        tab.FindControl<Border>("KiCadInitializingIndicator")!;

    private static TextBlock IndicatorTextOf(TabSchematics tab) =>
        tab.FindControl<TextBlock>("KiCadInitializingIndicatorText")!;
}
