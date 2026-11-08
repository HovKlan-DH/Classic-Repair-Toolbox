using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;
using Avalonia.VisualTree;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// The "Time spent (hours)" readback in the Add/Edit work done dialog - the line under the field
// that says 1,25 back as "1 hour and 15 minutes" while it is typed.
//
// WorklogDurationFormatterTests already pins the arithmetic. What is only testable here is the
// WIRING: that the ValueChanged hook is actually attached (a mis-typed handler name in the markup
// fails the XAML parse, but a handler attached to the wrong control does not), that the numbers
// come out BOLD and the words plain, and that the line collapses rather than blanking - an empty
// but visible TextBlock still occupies its height and would make the dialog jump as soon as a
// value is typed.
//
// The bold/plain split also means the block carries its content in Inlines with Text null, so a
// reader that looks only at Text sees this line as blank - the same trap the Workbooks summary
// strip and TextLinkRenderer both document.
[Collection("HeadlessUi")]
public sealed class WorkDoneHoursReadbackTests
{
    private static TextBlock Readback(Window window) =>
        window.GetVisualDescendants()
            .OfType<TextBlock>()
            .Single(block => block.Name == "HoursReadbackText");

    private static NumericUpDown HoursField(Window window) =>
        window.GetVisualDescendants()
            .OfType<NumericUpDown>()
            .Single(field => field.Name == "HoursNumericUpDown");

    // Show() so the visual tree is built and the named controls can be found at all.
    private static WorklogAddWorkDoneWindow BuildShownWindow()
    {
        var window = new WorklogAddWorkDoneWindow();
        window.Show();
        return window;
    }

    private static string VisibleText(TextBlock block) =>
        block.Inlines is { Count: > 0 }
            ? string.Concat(block.Inlines.OfType<Run>().Select(run => run.Text))
            : block.Text ?? string.Empty;

    // The whole feature, end to end through the control: type 1.25, read back the words.
    [Fact]
    public void Setting_the_hours_updates_the_readback_live()
    {
        UiTest.Run(() =>
        {
            var window = BuildShownWindow();

            HoursField(window).Value = 1.25m;

            var readback = Readback(window);

            Assert.True(readback.IsVisible);
            Assert.Equal("1 hour and 15 minutes", VisibleText(readback));

            window.Close();
        });
    }

    // Only the NUMBERS are bold. The finished string cannot show the difference, so this walks the
    // runs - the same reason the summary strip's own weight test does.
    [Fact]
    public void Only_the_numbers_are_bold()
    {
        UiTest.Run(() =>
        {
            var window = BuildShownWindow();

            HoursField(window).Value = 1.25m;

            var runs = Readback(window).Inlines!.OfType<Run>().ToList();

            Assert.Equal(4, runs.Count);
            Assert.Equal(FontWeight.Bold, runs[0].FontWeight);   // "1"
            Assert.NotEqual(FontWeight.Bold, runs[1].FontWeight); // " hour and "
            Assert.Equal(FontWeight.Bold, runs[2].FontWeight);   // "15"
            Assert.NotEqual(FontWeight.Bold, runs[3].FontWeight); // " minutes"

            window.Close();
        });
    }

    // An untouched field has nothing to say, and the line is COLLAPSED rather than left visible and
    // empty - a visible empty TextBlock keeps its height, so the buttons below it would shift down
    // the moment a value was typed and back up when it was cleared.
    [Fact]
    public void The_readback_is_collapsed_when_the_field_is_untouched()
    {
        UiTest.Run(() =>
        {
            var window = BuildShownWindow();

            Assert.False(Readback(window).IsVisible);

            window.Close();
        });
    }

    // Clearing the field takes the line away again rather than leaving the previous answer standing
    // under a now-empty box.
    [Fact]
    public void Clearing_the_hours_hides_the_readback_again()
    {
        UiTest.Run(() =>
        {
            var window = BuildShownWindow();
            var field = HoursField(window);

            field.Value = 2.5m;
            Assert.True(Readback(window).IsVisible);

            field.Value = 0m;

            var readback = Readback(window);
            Assert.False(readback.IsVisible);

            // Emptied, not merely hidden - a hidden block still holding last time's runs would show
            // the stale answer the instant anything made it visible again.
            Assert.Equal(string.Empty, VisibleText(readback));

            window.Close();
        });
    }

    // Successive edits REPLACE the previous answer rather than appending to it - the Inlines
    // collection is cleared on every pass.
    [Fact]
    public void Retyping_replaces_the_previous_readback_rather_than_appending()
    {
        UiTest.Run(() =>
        {
            var window = BuildShownWindow();
            var field = HoursField(window);

            field.Value = 1.25m;
            field.Value = 3m;

            Assert.Equal("3 hours", VisibleText(Readback(window)));

            window.Close();
        });
    }

    // Opening the dialog on an EXISTING row shows the readback straight away, without the user
    // having to touch the field first.
    [Fact]
    public void Opening_in_edit_mode_shows_the_readback_for_the_existing_value()
    {
        UiTest.Run(() =>
        {
            var window = new WorklogAddWorkDoneWindow();
            window.InitializeForEdit("Replaced CIA with another 6526", 1.25, 175);
            window.Show();

            var readback = Readback(window);

            Assert.True(readback.IsVisible);
            Assert.Equal("1 hour and 15 minutes", VisibleText(readback));

            window.Close();
        });
    }

    // Editing a row recorded as ZERO hours assigns 0 over the field's existing 0, which raises no
    // ValueChanged at all - so the readback state has to be settled explicitly. It is collapsed,
    // which is right, but reaching that by luck rather than by the explicit call would break the
    // moment the field's default changed.
    [Fact]
    public void Opening_in_edit_mode_on_zero_hours_leaves_the_readback_collapsed()
    {
        UiTest.Run(() =>
        {
            var window = new WorklogAddWorkDoneWindow();
            window.InitializeForEdit("Bench notes only", 0.0, 0.0);
            window.Show();

            Assert.False(Readback(window).IsVisible);

            window.Close();
        });
    }
}
