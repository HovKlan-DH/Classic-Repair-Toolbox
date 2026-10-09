using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using CRT;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The reason asked before a Boards-screen change is published to BETA (owner request, 2026-10-03:
// "do ask for a change reason when clicking the 'Save changes' button, so this can go along with the
// change, just like any normal contribution"), shown.
//
// The change goes straight to BETA once this is confirmed, so: the files it removes are named
// first, the reason is required and handed back trimmed, and no keypress publishes by reflex.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class PublishBoardChangeWindowTests
{
    private const string BoardId = "Commodore/C64/250407";

    private static PublishBoardChangeWindow Shown(IReadOnlyList<string>? removals = null, string? reason = null)
    {
        var window = new PublishBoardChangeWindow();
        window.Initialize(PublishBoardChangeWindowTests.BoardId, removals ?? [], reason);
        window.Show();

        return window;
    }

    private static IReadOnlyList<string> Removals(Window window) =>
        window.FindControl<StackPanel>("RemovalsPanel")!.Children.OfType<TextBlock>().Select(block => block.Text ?? string.Empty).ToList();

    [Fact]
    public void It_says_where_the_change_goes_and_asks_for_a_reason()
    {
        UiTest.Run(() =>
        {
            PublishBoardChangeWindow window = PublishBoardChangeWindowTests.Shown();

            Assert.Equal(BoardSections.ReasonTitle, window.Title);
            Assert.Equal(BoardSections.ReasonHeadline(PublishBoardChangeWindowTests.BoardId), window.FindControl<TextBlock>("HeadlineText")!.Text);
            Assert.Equal(BoardSections.ReasonExplanation, window.FindControl<TextBlock>("ExplanationText")!.Text);
            Assert.Equal(BoardSections.ReasonLabel, window.FindControl<TextBlock>("ReasonLabelText")!.Text);
            Assert.Equal(BoardSections.PublishButton, window.FindControl<Button>("PublishButton")!.Content);

            // A change that removes nothing shows no list at all.
            Assert.False(window.FindControl<StackPanel>("RemovalsPanel")!.IsVisible);

            window.Close();
        });
    }

    // What the publish removes is on screen BEFORE it happens - the heading, then each file.
    [Fact]
    public void The_files_the_change_removes_are_named_before_it_is_published()
    {
        UiTest.Run(() =>
        {
            PublishBoardChangeWindow window = PublishBoardChangeWindowTests.Shown(["Commodore/C64/250407/manual.pdf"]);

            Assert.True(window.FindControl<StackPanel>("RemovalsPanel")!.IsVisible);
            Assert.Equal(
                [BoardSections.RemovalsHeading(["x"])!, "Commodore/C64/250407/manual.pdf"],
                PublishBoardChangeWindowTests.Removals(window));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE REASON IS REQUIRED *** - it goes with the change, like a contribution's description.
    // Publish stays off until something is typed, and what comes back is trimmed.
    // ###########################################################################################
    [Fact]
    public void Publish_waits_for_a_reason_and_hands_it_back_trimmed()
    {
        UiTest.Run(() =>
        {
            PublishBoardChangeWindow window = PublishBoardChangeWindowTests.Shown();
            Button publish = window.FindControl<Button>("PublishButton")!;
            TextBox box = window.FindControl<TextBox>("ReasonTextBox")!;

            Assert.False(publish.IsEnabled);

            box.Text = "   ";
            Dispatcher.UIThread.RunJobs();
            Assert.False(publish.IsEnabled);

            box.Text = "  Corrected U8's part number.  ";
            Dispatcher.UIThread.RunJobs();
            Assert.True(publish.IsEnabled);

            publish.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

            Assert.Equal("Corrected U8's part number.", window.Reason);
        });
    }

    // A reason given earlier is kept when the dialog opens again, and Publish is on at once.
    [Fact]
    public void A_reason_given_earlier_is_kept()
    {
        UiTest.Run(() =>
        {
            PublishBoardChangeWindow window = PublishBoardChangeWindowTests.Shown(reason: "Corrected U8.");

            Assert.Equal("Corrected U8.", window.FindControl<TextBox>("ReasonTextBox")!.Text);
            Assert.True(window.FindControl<Button>("PublishButton")!.IsEnabled);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** ENTER AND ESCAPE BOTH CANCEL - including with Publish FOCUSED. *** A focused Button consumes
    // Enter on the bubbling route and fires its Click, so a reflexive keypress would publish to BETA.
    // Asserted on the button's OWN Click, which is what fails against a bubbling handler - the window
    // closes either way. Same rule as RollBackBetaWindow and DeleteBoardWindow.
    // ###########################################################################################
    [Fact]
    public void Enter_cancels_even_with_Publish_focused()
    {
        UiTest.Run(() =>
        {
            PublishBoardChangeWindow window = PublishBoardChangeWindowTests.Shown(reason: "Corrected U8.");

            Button publish = window.FindControl<Button>("PublishButton")!;
            publish.Focus();

            bool clicked = false;
            publish.Click += (_, _) => clicked = true;

            window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);

            Assert.False(clicked);
            Assert.Null(window.Reason);
        });
    }

    [Fact]
    public void Escape_cancels()
    {
        UiTest.Run(() =>
        {
            PublishBoardChangeWindow window = PublishBoardChangeWindowTests.Shown(reason: "Corrected U8.");

            window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, keySymbol: null);

            Assert.Null(window.Reason);
        });
    }

    // A second line of the reason can be typed: Enter in the box is a new line, not a cancel.
    [Fact]
    public void Enter_inside_the_reason_box_does_not_cancel()
    {
        UiTest.Run(() =>
        {
            PublishBoardChangeWindow window = PublishBoardChangeWindowTests.Shown(reason: "First line.");

            TextBox box = window.FindControl<TextBox>("ReasonTextBox")!;
            box.Focus();

            window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);

            Assert.True(window.IsVisible);
            Assert.Null(window.Reason);
            Assert.Contains("First line.", box.Text, StringComparison.Ordinal);

            window.Close();
        });
    }
}
