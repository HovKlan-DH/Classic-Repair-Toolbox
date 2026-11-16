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
// filled in. A box of several lines would have to take Enter as a new line instead - "What did you
// change?" and then the credit's "Main contribution" were one, until both left the dialog on
// 2026-11-15 (owner requests): the email address is all there is to type. A real key press through
// the headless input stack is what lets this test fail.
//
// The dialog is deliberately NOT initialised: with no board handed to it, OnSubmitClick returns at
// its first line, so even a failing run never starts a real submission (CLAUDE.md test rule 6). The
// button's Click event still fires, and that is what is counted.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SubmitDraftWindowKeyboardTests
{
    private static void PressEnter(Window window) =>
        window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);

    [Fact]
    public void Enter_in_the_email_box_sends_once_the_form_is_filled_in()
    {
        UiTest.Run(() =>
        {
            var window = new SubmitDraftWindow();

            // Show() so the visual tree is built and focus can land in a box.
            window.Show();

            try
            {
                window.GetControl<TextBox>("EmailTextBox").Text = "contributor@example.com";

                // TextChanged - what enables Submit - is raised from the dispatcher queue, not at once.
                Dispatcher.UIThread.RunJobs();

                // Filled in, so Submit is enabled - a disabled default button would make the assertion below pass for the wrong
                // reason.
                Assert.True(window.GetControl<Button>("SubmitButton").IsEnabled);

                int submitClicks = 0;
                window.GetControl<Button>("SubmitButton").Click += (_, _) => submitClicks++;

                window.GetControl<TextBox>("EmailTextBox").Focus();
                PressEnter(window);

                Assert.Equal(1, submitClicks);
            }
            finally
            {
                window.Close();
            }
        });
    }
}
