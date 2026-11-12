using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The "Loading KiCad data..." badge in the bottom-right corner of the schematics view.
//
// A board's KiCad project loads in the background (TabSchematics.LoadKiCadProjectForCurrentBoardAsync)
// so the schematic image can be shown immediately. For the few seconds that load takes, hovering a
// trace, clicking a pad or picking a net does nothing at all, which is indistinguishable from a
// broken overlay. The badge exists to make that wait visible, so what matters here is that it
// is hidden by default, that it can be turned on and off, and that it sits where the user was told
// to look for it without stealing pointer input from the schematic underneath.
//
// It stands where "Has KiCad data" stands, beside "Draw trace" (owner request, 2026-11-12) - that the
// two take turns there is KiCadCalibrationInteractionTests', which has a schematic image with traces.
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
    public void Showing_the_indicator_puts_the_loading_text_on_screen()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = new();

            tab.SetKiCadInitializingIndicatorVisible(true);

            Assert.True(IndicatorOf(tab).IsVisible);
            Assert.Equal("Loading KiCad data...", IndicatorTextOf(tab).Text);
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
    public void The_indicator_sits_beside_Draw_trace_and_ignores_the_pointer()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = new();

            StackPanel stack = tab.FindControl<StackPanel>("SchematicsBottomRightOverlayStack")!;
            StackPanel row = tab.FindControl<StackPanel>("SchematicsViewToolsRow")!;

            Assert.Equal(HorizontalAlignment.Right, stack.HorizontalAlignment);
            Assert.Equal(VerticalAlignment.Bottom, stack.VerticalAlignment);

            // One place for the two badges, before "Draw trace" (owner request, 2026-11-12: "in the same position").
            Assert.Same(row, IndicatorOf(tab).Parent);
            Assert.Equal(
                ["KiCadInitializingIndicator", "KiCadDataBadge", "DrawTraceButton"],
                row.Children.Select(child => child.Name).ToArray());

            // It is a status message, not a panel: wheel-zoom and click-drag over that corner
            // must keep reaching the schematic underneath - the "Has KiCad data" badge's too
            // (code review, 2026-11-12).
            Assert.False(IndicatorOf(tab).IsHitTestVisible);
            Assert.False(tab.KiCadDataBadge.IsHitTestVisible);
        });
    }

    // "Draw trace" and its badges stand LEFT of the "Traces visible" panel, never on top of it (owner request,
    // 2026-11-12: "Do have the "Draw trace" button and the "KiCad data" badge to the left of that - not on top of it").
    [Fact]
    public void Draw_trace_and_its_badges_stand_left_of_the_traces_panel()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = new();

            StackPanel stack = tab.FindControl<StackPanel>("SchematicsBottomRightOverlayStack")!;
            StackPanel row = tab.FindControl<StackPanel>("SchematicsViewToolsRow")!;
            Border traces = tab.FindControl<Border>("TracesPanel")!;

            Assert.Equal(Orientation.Horizontal, stack.Orientation);
            Assert.True(stack.Children.IndexOf(row) < stack.Children.IndexOf(traces));
            Assert.Equal(VerticalAlignment.Bottom, row.VerticalAlignment);
            Assert.Equal(VerticalAlignment.Bottom, traces.VerticalAlignment);
        });
    }

    // "Draw trace" shows the hand, as it can be pressed (owner request, 2026-11-12).
    [Fact]
    public void Draw_trace_shows_the_hand()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = new();

            Assert.Equal(new Cursor(StandardCursorType.Hand).ToString(), tab.DrawTraceButton.Cursor?.ToString());
        });
    }

    private static Border IndicatorOf(TabSchematics tab) =>
        tab.FindControl<Border>("KiCadInitializingIndicator")!;

    private static TextBlock IndicatorTextOf(TabSchematics tab) =>
        tab.FindControl<TextBlock>("KiCadInitializingIndicatorText")!;
}
