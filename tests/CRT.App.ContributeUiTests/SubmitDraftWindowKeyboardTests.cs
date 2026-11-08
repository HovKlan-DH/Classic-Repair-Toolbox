using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// WHAT ENTER DOES IN THE SUBMIT DIALOG (owner question, 2026-10-13: "how else to insert a
// new-line?").
//
// Submit is the dialog's DEFAULT button (code review, 2026-10-13), so Enter sends once the form is
// filled in. *** EXCEPT inside "What did you change?" ***, which is a multi-line box
// (AcceptsReturn): there Enter must start a new line and send nothing. Avalonia's TextBox marks
// the Enter it turns into a new line as handled, and a default button only answers an Enter
// nobody handled - so this holds only while both stay as they are. A real key press through the
// headless input stack is what lets these tests fail.
//
// The dialog is deliberately NOT initialised: with no board handed to it, OnSubmitClick returns
// at its first line, so even a failing run never starts a real submission (CLAUDE.md test rule
// 6). The button's Click event still fires, and that is what is counted.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SubmitDraftWindowKeyboardTests
{
    private static SubmitDraftWindow BuildFilledInWindow()
    {
        var window = new SubmitDraftWindow();

        // Show() so the visual tree is built and focus can land in a box.
        window.Show();

        // Both filled in, so Submit is enabled - a disabled default button would make every
        // assertion below pass for the wrong reason.
        window.GetControl<TextBox>("EmailTextBox").Text = "contributor@example.com";
        window.GetControl<TextBox>("SummaryTextBox").Text = "Corrected R12";

        // TextChanged - what enables Submit - is raised from the dispatcher queue, not at once.
        Dispatcher.UIThread.RunJobs();

        Assert.True(window.GetControl<Button>("SubmitButton").IsEnabled);

        return window;
    }

    private static void PressEnter(Window window) =>
        window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);

    [Fact]
    public void Enter_in_the_summary_box_starts_a_new_line_and_sends_nothing()
    {
        UiTest.Run(() =>
        {
            SubmitDraftWindow window = BuildFilledInWindow();
            TextBox summary = window.GetControl<TextBox>("SummaryTextBox");

            int submitClicks = 0;
            window.GetControl<Button>("SubmitButton").Click += (_, _) => submitClicks++;

            summary.Focus();
            summary.CaretIndex = summary.Text!.Length;
            PressEnter(window);

            Assert.Equal(0, submitClicks);
            Assert.Contains('\n', summary.Text);

            window.Close();
        });
    }

    // The other half: without it, the test above would also pass if Enter did nothing anywhere.
    [Fact]
    public void Enter_in_the_email_box_sends_once_the_form_is_filled_in()
    {
        UiTest.Run(() =>
        {
            SubmitDraftWindow window = BuildFilledInWindow();

            int submitClicks = 0;
            window.GetControl<Button>("SubmitButton").Click += (_, _) => submitClicks++;

            window.GetControl<TextBox>("EmailTextBox").Focus();
            PressEnter(window);

            Assert.Equal(1, submitClicks);

            window.Close();
        });
    }
}
