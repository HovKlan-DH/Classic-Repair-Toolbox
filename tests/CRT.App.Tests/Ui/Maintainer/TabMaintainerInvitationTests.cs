using System.Reflection;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// "I HAVE AN INVITATION" ON THE SIGN-IN SCREEN (2026-09-27; TabMaintainer.Invitation.cs).
//
// *** THE INVITATION IS SHOWN INSTEAD OF THE SIGN-IN FORM, NOT UNDER IT (owner request, 2026-10-04:
// "it should show only the information relevant to that - it should not show the top login, as
// this can be confusing. There should be a cancel button also, so it can return to the previous
// view"). *** It used to open under the email and password boxes and the "Sign in" button, which
// have nothing to do with accepting an invitation. Cancel goes back to the form as it was left.
// (Moved here from TabMaintainerModesTests with that change.)
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerInvitationTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static void Press(TabMaintainer main, string button) =>
        main.FindControl<Button>(button)!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    private static void FillInvitation(TabMaintainer main, string code, string name, string password)
    {
        main.FindControl<TextBox>("InvitationCodeTextBox")!.Text = code;
        main.FindControl<TextBox>("InvitationNameTextBox")!.Text = name;
        main.FindControl<TextBox>("InvitationPasswordTextBox")!.Text = password;
    }

    // Every box, button and line of text actually drawn on the sign-in screen - read off a shown
    // window, so a control moved out of the panel that hides it would be caught.
    private static List<Control> Drawn(TabMaintainer main) =>
        main.FindControl<StackPanel>("SignInPanel")!
            .GetVisualDescendants()
            .OfType<Control>()
            .Where(control => control is TextBox or Button or TextBlock && control.IsEffectivelyVisible)
            .ToList();

    private static bool IsInside(Control control, Control panel) =>
        control.GetVisualAncestors().Contains(panel);

    [Fact]
    public void I_have_an_invitation_shows_the_invitation_alone_with_nothing_of_the_sign_in_form()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            var window = new Window { Content = main, Width = 900, Height = 900 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            StackPanel invitation = main.FindControl<StackPanel>("InvitationPanel")!;

            // Before: the form, and nothing of the invitation.
            Assert.True(main.FindControl<TextBox>("EmailTextBox")!.IsEffectivelyVisible);
            Assert.All(Drawn(main), control => Assert.False(IsInside(control, invitation), control.Name ?? control.GetType().Name));

            Press(main, "HaveInvitationButton");
            Dispatcher.UIThread.RunJobs();

            // After: the invitation and nothing else - no email box, no "Sign in", no explanation of
            // the tab, no "I forgot my password".
            List<Control> drawn = Drawn(main);

            Assert.True(invitation.IsEffectivelyVisible);
            Assert.Contains(main.FindControl<TextBox>("InvitationCodeTextBox")!, drawn);
            Assert.Contains(main.FindControl<Button>("AcceptInvitationButton")!, drawn);
            Assert.Contains(main.FindControl<Button>("CancelInvitationButton")!, drawn);
            Assert.All(drawn, control => Assert.True(IsInside(control, invitation), control.Name ?? control.GetType().Name));

            window.Close();
        });
    }

    // Cancel: back to the form exactly as it was left (what was typed, an open reset panel), and the
    // invitation's boxes and message emptied - a password typed there must not sit in a hidden box,
    // and the next "I have an invitation" starts blank.
    [Fact]
    public async Task Cancel_goes_back_to_the_sign_in_form_as_it_was_and_empties_the_invitation()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            int asked = 0;

            main.AcceptInvitationOverrideForTests = _ =>
            {
                asked++;
                return Task.FromResult(ReviewApiResult<AcceptInvitationAnswer>.Failed(ReviewApiFailure.Refused, "No."));
            };

            main.FindControl<TextBox>("EmailTextBox")!.Text = "anna@example.com";
            typeof(TabMaintainer).GetMethod("ShowResetPanel", Any)!.Invoke(main, null);

            Press(main, "HaveInvitationButton");
            FillInvitation(main, string.Empty, "Anna", "correct horse battery staple");

            // An empty code is answered on the spot, in the invitation's own message line.
            await main.AcceptInvitationAsync();
            Assert.Equal(0, asked);
            Assert.True(main.FindControl<TextBlock>("InvitationMessageText")!.IsVisible);

            Press(main, "CancelInvitationButton");

            Assert.True(main.FindControl<StackPanel>("SignInFormPanel")!.IsVisible);
            Assert.False(main.FindControl<StackPanel>("InvitationPanel")!.IsVisible);
            Assert.Equal("anna@example.com", main.FindControl<TextBox>("EmailTextBox")!.Text);
            Assert.True(main.FindControl<StackPanel>("ResetPanel")!.IsVisible);

            Assert.Equal(string.Empty, main.FindControl<TextBox>("InvitationNameTextBox")!.Text);
            Assert.Equal(string.Empty, main.FindControl<TextBox>("InvitationPasswordTextBox")!.Text);

            // Opened again, it starts blank, with no message left from before.
            Press(main, "HaveInvitationButton");

            Assert.True(main.FindControl<StackPanel>("InvitationPanel")!.IsVisible);
            Assert.False(main.FindControl<StackPanel>("SignInFormPanel")!.IsVisible);
            Assert.False(main.FindControl<TextBlock>("InvitationMessageText")!.IsVisible);
            Assert.Equal(string.Empty, main.FindControl<TextBlock>("InvitationMessageText")!.Text);
        });
    }

    // ###########################################################################################
    // Accepting: the code, a name and a password go to the server; on success the sign-in form
    // comes back with its fields filled with the invited address and the password - the password
    // reset's rule, since accepting opens no session - saying the account is ready.
    // ###########################################################################################
    [Fact]
    public async Task Accepting_an_invitation_sends_the_code_name_and_password_and_fills_the_sign_in_fields()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            var sent = new List<AcceptInvitationRequest>();

            main.AcceptInvitationOverrideForTests = request =>
            {
                sent.Add(request);
                return Task.FromResult(ReviewApiResult<AcceptInvitationAnswer>.Ok(
                    new AcceptInvitationAnswer("anna@example.com", ["Commodore/C64/250407"], "Your account is ready.")));
            };

            Press(main, "HaveInvitationButton");
            Assert.True(main.FindControl<StackPanel>("InvitationPanel")!.IsVisible);

            FillInvitation(main, "  the-code  ", "Anna", "correct horse battery staple");

            await main.AcceptInvitationAsync();

            Assert.Equal([new AcceptInvitationRequest("the-code", "Anna", "correct horse battery staple")], sent);
            Assert.Equal("anna@example.com", main.FindControl<TextBox>("EmailTextBox")!.Text);
            Assert.Equal("correct horse battery staple", main.FindControl<TextBox>("PasswordTextBox")!.Text);
            Assert.False(main.FindControl<StackPanel>("InvitationPanel")!.IsVisible);
            Assert.True(main.FindControl<StackPanel>("SignInFormPanel")!.IsVisible);
            Assert.Equal("Your account is ready.", main.FindControl<TextBlock>("SignInMessageText")!.Text);
            Assert.True(main.FindControl<TextBlock>("SignInMessageText")!.IsVisible);
            Assert.Equal(string.Empty, main.FindControl<TextBox>("InvitationCodeTextBox")!.Text);
            Assert.Equal(string.Empty, main.FindControl<TextBox>("InvitationPasswordTextBox")!.Text);
        });
    }

    // ###########################################################################################
    // *** NO ANSWER IN TWO MINUTES GOES BACK TO THE SIGN-IN FORM (code review, 2026-10-04). *** The
    // account may exist by now, so signing in is the next step - and the message saying so stayed in
    // the invitation view, where no sign-in form was on screen, while Cancel (the only way back)
    // emptied the password just chosen. Now the form comes back with that password filled in and the
    // message under it.
    // ###########################################################################################
    [Fact]
    public async Task No_answer_to_an_invitation_goes_back_to_the_sign_in_form_with_the_password_filled_in()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();

            main.AcceptInvitationOverrideForTests = _ => Task.FromResult(
                ReviewApiResult<AcceptInvitationAnswer>.Failed(ReviewApiFailure.TimedOut, "No answer."));

            Press(main, "HaveInvitationButton");
            FillInvitation(main, "the-code", "Anna", "correct horse battery staple");

            await main.AcceptInvitationAsync();

            Assert.True(main.FindControl<StackPanel>("SignInFormPanel")!.IsVisible);
            Assert.False(main.FindControl<StackPanel>("InvitationPanel")!.IsVisible);
            Assert.Equal("correct horse battery staple", main.FindControl<TextBox>("PasswordTextBox")!.Text);

            TextBlock message = main.FindControl<TextBlock>("SignInMessageText")!;
            Assert.True(message.IsVisible);
            Assert.Equal(MaintainerWaitWording.InvitationNoAnswer, message.Text);

            // Nothing of it left in the hidden invitation's boxes.
            Assert.Equal(string.Empty, main.FindControl<TextBox>("InvitationPasswordTextBox")!.Text);
        });
    }

    // A refusal keeps everything typed and shows the server's words - in the invitation's own line,
    // since the form's is hidden with the form; an empty box asks nobody.
    [Fact]
    public async Task A_refused_or_incomplete_invitation_keeps_the_panel_and_says_why()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();
            int asked = 0;

            main.AcceptInvitationOverrideForTests = _ =>
            {
                asked++;
                return Task.FromResult(ReviewApiResult<AcceptInvitationAnswer>.Failed(
                    ReviewApiFailure.Refused, "That invitation has expired. Ask the administrator for a new one."));
            };

            Press(main, "HaveInvitationButton");

            TextBlock message = main.FindControl<TextBlock>("InvitationMessageText")!;

            await main.AcceptInvitationAsync();
            Assert.Equal(0, asked);
            Assert.Equal("Paste the code from the invitation email first.", message.Text);
            Assert.True(message.IsVisible);

            FillInvitation(main, "old-code", "Anna", "correct horse battery staple");

            await main.AcceptInvitationAsync();

            Assert.Equal(1, asked);
            Assert.Equal("That invitation has expired. Ask the administrator for a new one.", message.Text);
            Assert.True(main.FindControl<StackPanel>("InvitationPanel")!.IsVisible);
            Assert.False(main.FindControl<StackPanel>("SignInFormPanel")!.IsVisible);
            Assert.Equal("old-code", main.FindControl<TextBox>("InvitationCodeTextBox")!.Text);

            // Nothing was said in the hidden form's line, so going back finds it as it was.
            Assert.False(main.FindControl<TextBlock>("SignInMessageText")!.IsVisible);
        });
    }

    // A remembered sign-in can be restored while the invitation is on screen. Signing in leaves it,
    // so the next sign-out lands on the sign-in form, not on a half-filled invitation.
    [Fact]
    public void Signing_in_leaves_the_invitation_so_the_next_sign_out_finds_the_form()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();

            Press(main, "HaveInvitationButton");
            FillInvitation(main, "the-code", "Anna", "correct horse battery staple");

            typeof(TabMaintainer).GetMethod("ShowQueuePanel", Any)!.Invoke(main, null);
            typeof(TabMaintainer).GetMethod("StopQueueChecks", Any)!.Invoke(main, null);

            Assert.False(main.FindControl<StackPanel>("InvitationPanel")!.IsVisible);
            Assert.True(main.FindControl<StackPanel>("SignInFormPanel")!.IsVisible);
            Assert.Equal(string.Empty, main.FindControl<TextBox>("InvitationCodeTextBox")!.Text);
            Assert.Equal(string.Empty, main.FindControl<TextBox>("InvitationPasswordTextBox")!.Text);
        });
    }
}
