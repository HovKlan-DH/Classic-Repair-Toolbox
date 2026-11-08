using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The confirmation for deleting a board completely (owner request, 2026-10-03: "with a
// confirmation box"), shown.
//
// *** THE POINT OF THE DIALOG IS THE LIST. *** A delete cannot be undone, so what goes from each
// place - and whose open work goes with it - is on screen before anything happens. The reason is
// asked for only when a contributor is mailed, and then it is required.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class DeleteBoardWindowTests
{
    private static readonly DateTimeOffset Sent = new(2026, 10, 2, 9, 0, 0, TimeSpan.Zero);

    private static BoardDeletePlanAnswer Plan(int open = 0) =>
        new("Commodore/C64/999999", "Commodore", "C64", "999999", "f",
            12, 11, true, true, true, 5, 2, 0,
            [.. Enumerable.Range(1, open).Select(index =>
                new BoardDeleteOpenSubmission(index, "pending", $"c{index}@example.com", "Corrected U8.", DeleteBoardWindowTests.Sent))]);

    private static DeleteBoardWindow Shown(BoardDeletePlanAnswer plan)
    {
        var window = new DeleteBoardWindow();
        window.Initialize(plan);
        window.Show();

        return window;
    }

    private static IReadOnlyList<string> Lines(Window window, string panel) =>
        window.FindControl<StackPanel>(panel)!.Children.OfType<TextBlock>().Select(block => block.Text ?? string.Empty).ToList();

    [Fact]
    public void It_lists_what_goes_and_names_every_open_submission()
    {
        UiTest.Run(() =>
        {
            BoardDeletePlanAnswer plan = DeleteBoardWindowTests.Plan(open: 2);
            DeleteBoardWindow window = DeleteBoardWindowTests.Shown(plan);

            Assert.Equal(BoardDeletionWording.Headline(plan), window.FindControl<TextBlock>("HeadlineText")!.Text);
            Assert.Equal(BoardDeletionWording.WhatGoes(plan), DeleteBoardWindowTests.Lines(window, "WhatGoesPanel"));
            Assert.Equal(plan.OpenSubmissions.Select(BoardDeletionWording.OpenLine), DeleteBoardWindowTests.Lines(window, "OpenListPanel"));
            Assert.True(window.FindControl<StackPanel>("OpenPanel")!.IsVisible);
            Assert.Equal(BoardDeletionWording.ConfirmButton, window.FindControl<Button>("ConfirmButton")!.Content);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE REASON IS REQUIRED WHILE A CONTRIBUTOR IS MAILED *** - it is the mail's only content.
    // Delete stays off until something is typed, and what comes back is trimmed.
    // ###########################################################################################
    [Fact]
    public void With_open_submissions_Delete_waits_for_a_reason_and_hands_it_back()
    {
        UiTest.Run(() =>
        {
            DeleteBoardWindow window = DeleteBoardWindowTests.Shown(DeleteBoardWindowTests.Plan(open: 1));
            Button confirm = window.FindControl<Button>("ConfirmButton")!;

            Assert.False(confirm.IsEnabled);

            window.FindControl<TextBox>("ReasonTextBox")!.Text = "   ";
            Dispatcher.UIThread.RunJobs();
            Assert.False(confirm.IsEnabled);

            window.FindControl<TextBox>("ReasonTextBox")!.Text = "  It was a test board.  ";
            Dispatcher.UIThread.RunJobs();
            Assert.True(confirm.IsEnabled);

            confirm.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

            Assert.True(window.WasConfirmed);
            Assert.Equal("It was a test board.", window.Reason);
        });
    }

    // Nobody to tell: no reason box, and Delete is on at once.
    [Fact]
    public void Without_open_submissions_no_reason_is_asked_for()
    {
        UiTest.Run(() =>
        {
            DeleteBoardWindow window = DeleteBoardWindowTests.Shown(DeleteBoardWindowTests.Plan(open: 0));
            Button confirm = window.FindControl<Button>("ConfirmButton")!;

            Assert.False(window.FindControl<StackPanel>("OpenPanel")!.IsVisible);
            Assert.True(confirm.IsEnabled);

            confirm.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

            Assert.True(window.WasConfirmed);
            Assert.Equal(string.Empty, window.Reason);
        });
    }

    // ###########################################################################################
    // *** ENTER AND ESCAPE BOTH CANCEL - including with Delete FOCUSED. *** A focused Button consumes
    // Enter on the bubbling route and fires its Click, so a reflexive keypress would delete a board.
    // Asserted on the button's OWN Click, which is what fails against a bubbling handler - the window
    // closes either way. Same rule as RollBackBetaWindow and DeleteWorkbookWindow.
    // ###########################################################################################
    [Fact]
    public void Enter_cancels_even_with_Delete_focused()
    {
        UiTest.Run(() =>
        {
            DeleteBoardWindow window = DeleteBoardWindowTests.Shown(DeleteBoardWindowTests.Plan(open: 0));

            Button confirm = window.FindControl<Button>("ConfirmButton")!;
            confirm.Focus();

            bool clicked = false;
            confirm.Click += (_, _) => clicked = true;

            window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);

            Assert.False(clicked);
            Assert.False(window.WasConfirmed);
        });
    }

    [Fact]
    public void Escape_cancels()
    {
        UiTest.Run(() =>
        {
            DeleteBoardWindow window = DeleteBoardWindowTests.Shown(DeleteBoardWindowTests.Plan(open: 1));

            window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, keySymbol: null);

            Assert.False(window.WasConfirmed);
        });
    }

    // A second line of the reason can be typed: Enter in the box is a new line, not a cancel.
    [Fact]
    public void Enter_inside_the_reason_box_does_not_cancel()
    {
        UiTest.Run(() =>
        {
            DeleteBoardWindow window = DeleteBoardWindowTests.Shown(DeleteBoardWindowTests.Plan(open: 1));

            TextBox reason = window.FindControl<TextBox>("ReasonTextBox")!;
            reason.Text = "First line.";
            Dispatcher.UIThread.RunJobs();
            reason.Focus();

            window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);

            Assert.True(window.IsVisible);
            Assert.False(window.WasConfirmed);
            Assert.Contains("First line.", reason.Text, StringComparison.Ordinal);

            window.Close();
        });
    }
}
