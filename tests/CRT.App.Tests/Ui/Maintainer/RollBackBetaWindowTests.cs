using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The confirmation for pushing a BETA board back to the queue (owner decision, 2026-09-27).
//
// *** THE POINT OF THE DIALOG IS THE NAMES. *** A rollback is per BOARD, so it takes back EVERY
// submission merged since the last promotion; one contributor's work cannot be picked out. The
// dialog exists so a maintainer sees whose work that is before deciding, rather than discovering
// afterwards that two other people's accepted contributions went with it.
//
// Unlike TabMaintainer this window has no OnOpened side effects, so it can
// be shown - which is what a real key press needs.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class RollBackBetaWindowTests
{
    private static readonly DateTimeOffset Decided = new(2026, 9, 24, 8, 30, 0, TimeSpan.Zero);

    private static BetaRollbackPlanView Plan(
        BetaRollbackKind kind = BetaRollbackKind.RestoreFromProduction,
        int returning = 1) =>
        new(
            "Commodore/C64/250407",
            kind,
            ["board/sheet.png"],
            ["board/new.png"],
            Enumerable.Range(0, returning)
                .Select(index => new CarriedSubmission(index + 1, $"c{index}@example.com", "Corrected U8.", RollBackBetaWindowTests.Decided))
                .ToList());

    private static RollBackBetaWindow Shown(BetaRollbackPlanView plan)
    {
        var window = new RollBackBetaWindow();
        window.Initialize(plan);
        window.Show();

        return window;
    }

    // ###########################################################################################
    // The same window confirms Beta > Prod's "Reject" (owner request, 2026-09-28) - titled and
    // worded as a rejection, still naming everyone it takes out.
    // ###########################################################################################
    [Fact]
    public void As_a_rejection_it_says_reject_and_still_names_everyone()
    {
        UiTest.Run(() =>
        {
            var window = new RollBackBetaWindow();
            window.Initialize(RollBackBetaWindowTests.Plan(returning: 2), reject: true);
            window.Show();

            Assert.Equal("Reject", window.Title);
            Assert.Equal("Reject all 2", window.FindControl<Button>("ConfirmButton")!.Content);
            Assert.Contains("are rejected", window.FindControl<TextBlock>("ExplanationText")!.Text, StringComparison.Ordinal);
            Assert.Equal(2, window.FindControl<StackPanel>("ReturningPanel")!.Children.Count);

            window.Close();
        });
    }

    [Fact]
    public void The_window_builds_and_names_every_contributor_it_takes_back()
    {
        UiTest.Run(() =>
        {
            RollBackBetaWindow window = RollBackBetaWindowTests.Shown(RollBackBetaWindowTests.Plan(returning: 3));

            List<string> lines = window.FindControl<StackPanel>("ReturningPanel")!
                .Children.OfType<TextBlock>()
                .Select(TabMaintainer.TextOf)
                .ToList();

            Assert.Equal(3, lines.Count);
            Assert.All(lines, line => Assert.Contains("@example.com", line, StringComparison.Ordinal));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE COMMENT IS REQUIRED. *** It is the contributor's only feedback - contributing needs
    // no account, so there is no inbox and no thread. The server refuses a blank one too; this is
    // the half that stops the maintainer reaching a refusal in the first place.
    // ###########################################################################################
    [Fact]
    public void The_confirm_button_stays_off_until_a_reason_is_typed()
    {
        UiTest.Run(() =>
        {
            RollBackBetaWindow window = RollBackBetaWindowTests.Shown(RollBackBetaWindowTests.Plan());

            Button confirm = window.FindControl<Button>("ConfirmButton")!;
            TextBox comment = window.FindControl<TextBox>("CommentTextBox")!;

            Assert.False(confirm.IsEnabled);

            // TextChanged is raised through the dispatcher, so it is pumped the way a real
            // keystroke would pump it rather than asserting on a handler that has not run yet.
            comment.Text = "The U8 pinout is wrong.";
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();
            Assert.True(confirm.IsEnabled);

            // Whitespace is not a reason.
            comment.Text = "   ";
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();
            Assert.False(confirm.IsEnabled);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** ENTER AND ESCAPE BOTH CANCEL - including with the confirm button FOCUSED. *** A focused
    // Button consumes Enter on the bubbling route and fires its Click, so a reflexive keypress
    // would take a board out of BETA. The fix is the Tunnel route, and this test fails against a
    // bubbling handler. Same rule as DeleteWorkbookWindow and DeleteWorklogWindow.
    // ###########################################################################################
    [Fact]
    public void Enter_cancels_even_with_the_confirm_button_focused()
    {
        UiTest.Run(() =>
        {
            RollBackBetaWindow window = RollBackBetaWindowTests.Shown(RollBackBetaWindowTests.Plan());

            TextBox comment = window.FindControl<TextBox>("CommentTextBox")!;
            comment.Text = "A reason.";
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();

            Button confirm = window.FindControl<Button>("ConfirmButton")!;
            confirm.Focus();

            // ###########################################################################################
            // Watching the button's OWN Click is what makes this able to fail. Asserting only that
            // the window closed passes either way - Escape and the cancel path close it too. The
            // difference is whether OnConfirmClick ALSO ran and pushed the board back.
            // ###########################################################################################
            bool confirmed = false;
            confirm.Click += (_, _) => confirmed = true;

            window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);

            Assert.False(confirmed);
            Assert.False(window.WasConfirmed);
        });
    }

    [Fact]
    public void Escape_cancels()
    {
        UiTest.Run(() =>
        {
            RollBackBetaWindow window = RollBackBetaWindowTests.Shown(RollBackBetaWindowTests.Plan());

            window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, keySymbol: null);

            Assert.False(window.WasConfirmed);
        });
    }

    // ###########################################################################################
    // Enter must still work INSIDE the comment box - it accepts returns, so a maintainer writing a
    // second line of explanation must not have the dialog close under them.
    // ###########################################################################################
    [Fact]
    public void Enter_inside_the_comment_box_does_not_cancel()
    {
        UiTest.Run(() =>
        {
            RollBackBetaWindow window = RollBackBetaWindowTests.Shown(RollBackBetaWindowTests.Plan());

            TextBox comment = window.FindControl<TextBox>("CommentTextBox")!;
            comment.Text = "First line.";
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();
            comment.Focus();

            window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);

            // Still open and still holding what was typed.
            Assert.False(window.WasConfirmed);
            Assert.Contains("First line.", comment.Text, StringComparison.Ordinal);

            window.Close();
        });
    }

    // Confirming hands back the trimmed reason, which is what reaches the contributor.
    [Fact]
    public void Confirming_hands_back_the_reason()
    {
        UiTest.Run(() =>
        {
            RollBackBetaWindow window = RollBackBetaWindowTests.Shown(RollBackBetaWindowTests.Plan());

            window.FindControl<TextBox>("CommentTextBox")!.Text = "  The U8 pinout is wrong.  ";

            RollBackBetaWindowTests.Click(window.FindControl<Button>("ConfirmButton")!);

            Assert.True(window.WasConfirmed);
            Assert.Equal("The U8 pinout is wrong.", window.Comment);
        });
    }

    private static void Click(Button button) =>
        button.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

    // ###########################################################################################
    // *** THE SHARED FILES ARE LISTED, BY NAME (code review, 2026-09-27). *** They reach every
    // board citing them, so the maintainer sees exactly which before confirming - and the panel is
    // hidden when there are none.
    // ###########################################################################################
    [Fact]
    public void The_confirmation_lists_the_shared_files_it_puts_back()
    {
        UiTest.Run(() =>
        {
            RollBackBetaWindow window = RollBackBetaWindowTests.Shown(
                RollBackBetaWindowTests.Plan() with { SharedRestored = ["Commodore/Shared files/6526.png"] });

            StackPanel shared = window.FindControl<StackPanel>("SharedPanel")!;
            List<string> lines = shared.Children.OfType<TextBlock>().Select(TabMaintainer.TextOf).ToList();

            Assert.True(shared.IsVisible);
            Assert.StartsWith("Shared files, used by every board that cites them", lines[0], StringComparison.Ordinal);
            Assert.Contains("Commodore/Shared files/6526.png", lines);

            window.Close();
        });
    }

    [Fact]
    public void With_no_shared_files_the_shared_panel_is_hidden()
    {
        UiTest.Run(() =>
        {
            RollBackBetaWindow window = RollBackBetaWindowTests.Shown(RollBackBetaWindowTests.Plan());

            Assert.False(window.FindControl<StackPanel>("SharedPanel")!.IsVisible);

            window.Close();
        });
    }

    // A board never promoted says it leaves BETA, rather than describing a restore.
    [Fact]
    public void A_board_never_promoted_says_it_leaves_beta()
    {
        UiTest.Run(() =>
        {
            RollBackBetaWindow window = RollBackBetaWindowTests.Shown(
                RollBackBetaWindowTests.Plan(BetaRollbackKind.RemoveFromBeta));

            Assert.Contains(
                "Remove this board from BETA",
                window.FindControl<TextBlock>("HeadlineText")!.Text ?? string.Empty,
                StringComparison.Ordinal);

            Assert.Equal("Remove from BETA", window.FindControl<Button>("ConfirmButton")!.Content);

            window.Close();
        });
    }
}
