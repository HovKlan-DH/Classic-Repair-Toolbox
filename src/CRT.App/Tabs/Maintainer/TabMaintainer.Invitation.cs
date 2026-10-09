using System;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // "I HAVE AN INVITATION" ON THE SIGN-IN SCREEN (2026-09-27). The administrator invites an
    // address to maintain a board (the Boards screen); the mail carries a code, and this is where
    // it is redeemed - the code, a name and a password - which makes the account and puts it in the
    // pool of every board that address was invited to. See CRT.Server's MaintainerInvitationFlows.
    //
    // *** ON SUCCESS THE SIGN-IN FIELDS ARE FILLED, NOT A SESSION OPENED *** - the password reset's
    // rule, for the same reason: accepting is not signing in, the server issues no session for it,
    // and holding the password to sign in behind the person's back would be inventing one.
    //
    // *** THE INVITATION IS SHOWN INSTEAD OF THE SIGN-IN FORM, NOT UNDER IT *** (owner request,
    // 2026-10-04). Under it, the invitation sat below an email and password box and a "Sign in"
    // button that have nothing to do with accepting one, which was confusing. Cancel goes back to
    // the form exactly as it was left - what was typed there, a reset code panel, its message - and
    // empties the invitation's own boxes, so a password typed there does not sit in a hidden box.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private void OnHaveInvitationClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            this.ShowInvitationView(true);
            this.FindControl<TextBox>("InvitationCodeTextBox")?.Focus();
        }

        private void OnCancelInvitationClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            this.CloseInvitationView();

        // Back to the sign-in form, the invitation's boxes and message emptied. Also where signing
        // in leaves it (ShowQueuePanel), so the next sign-out never lands on a half-filled invitation.
        private void CloseInvitationView()
        {
            foreach (string name in new[] { "InvitationCodeTextBox", "InvitationNameTextBox", "InvitationPasswordTextBox" })
            {
                if (this.FindControl<TextBox>(name) is TextBox box)
                    box.Text = string.Empty;
            }

            this.ShowInvitationMessage(null);
            this.ShowInvitationView(false);
        }

        // One of the two views of the sign-in screen, never both.
        private void ShowInvitationView(bool show)
        {
            if (this.FindControl<StackPanel>("InvitationPanel") is StackPanel invitation)
                invitation.IsVisible = show;

            if (this.FindControl<StackPanel>("SignInFormPanel") is StackPanel form)
                form.IsVisible = !show;
        }

        // The invitation view's own message line - the form's is hidden with the form. Always a
        // failure: success goes back to the form and says so there.
        private void ShowInvitationMessage(string? message)
        {
            if (this.FindControl<TextBlock>("InvitationMessageText") is not TextBlock text)
                return;

            text.Text = message ?? string.Empty;
            text.IsVisible = !string.IsNullOrWhiteSpace(message);
        }

        // Enter in the password box accepts, as it signs in in the box above.
        private async void OnInvitationPasswordKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                await this.AcceptInvitationAsync();
            }
        }

        private async void OnAcceptInvitationClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) =>
            await this.AcceptInvitationAsync();

        internal async Task AcceptInvitationAsync()
        {
            var code = this.FindControl<TextBox>("InvitationCodeTextBox");
            var name = this.FindControl<TextBox>("InvitationNameTextBox");
            var password = this.FindControl<TextBox>("InvitationPasswordTextBox");
            var button = this.FindControl<Button>("AcceptInvitationButton");

            if (code is null || name is null || password is null || button is null)
                return;

            string enteredCode = code.Text?.Trim() ?? string.Empty;
            string enteredName = name.Text?.Trim() ?? string.Empty;
            string enteredPassword = password.Text ?? string.Empty;

            // Checked here because the round trip could say nothing more useful about an empty box.
            if (enteredCode.Length == 0)
            {
                this.ShowInvitationMessage("Paste the code from the invitation email first.");
                return;
            }

            if (enteredName.Length == 0)
            {
                this.ShowInvitationMessage("Type the name others will see.");
                return;
            }

            if (enteredPassword.Length == 0)
            {
                this.ShowInvitationMessage("Type the password you want to use.");
                return;
            }

            button.IsEnabled = false;
            this.ShowInvitationMessage(null);

            try
            {
                ReviewApiResult<AcceptInvitationAnswer> result = await ServerWait.CallAsync(
                    this,
                    MaintainerWaitWording.AcceptingInvitation,
                    token => this.SendAcceptanceAsync(enteredCode, enteredName, enteredPassword, token));

                // ###########################################################################
                // No answer in two minutes: the account may exist now, and the code may be spent -
                // signing in is the way to find out, never "use the code again" first. So it goes
                // back to the sign-in form with the password chosen filled in, and says so there
                // (code review, 2026-10-04): said in the invitation view, it told the invitee to
                // sign in where no sign-in form was on screen, and Cancel - the only way back -
                // emptied the password just chosen. The address is the one the invitation was
                // mailed to, which only the invitee knows here.
                // ###########################################################################
                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    this.CloseInvitationView();
                    this.ShowSignInMessage(MaintainerWaitWording.InvitationNoAnswer, isError: true);

                    if (this.FindControl<TextBox>("PasswordTextBox") is TextBox chosen)
                        chosen.Text = enteredPassword;

                    this.FindControl<TextBox>("EmailTextBox")?.Focus();
                    return;
                }

                if (!result.IsOk)
                {
                    this.ShowInvitationMessage(result.Message);
                    return;
                }

                AcceptInvitationAnswer answer = result.Value!;

                // The code is spent, so its boxes are emptied - a second press would only be told
                // so - and the sign-in form comes back, saying the account is ready.
                this.CloseInvitationView();
                this.ShowSignInMessage(answer.Message, isError: false);

                if (this.FindControl<TextBox>("EmailTextBox") is TextBox email)
                    email.Text = answer.Email;

                if (this.FindControl<TextBox>("PasswordTextBox") is TextBox signInPassword)
                {
                    signInPassword.Text = enteredPassword;
                    signInPassword.Focus();
                }
            }
            finally
            {
                button.IsEnabled = true;
            }
        }

        private async Task<ReviewApiResult<AcceptInvitationAnswer>> SendAcceptanceAsync(string code, string name, string password, CancellationToken token)
        {
            if (this.AcceptInvitationOverrideForTests is not null)
                return await this.AcceptInvitationOverrideForTests(new AcceptInvitationRequest(code, name, password));

            // Its own client, like the password reset's: the shared one is made by signing in.
            using var client = new ReviewApiClient(ReviewApiRoutes.DefaultBaseAddress);

            return await client.AcceptInvitationAsync(code, name, password, token);
        }

        // Answers the acceptance instead of the server, and sees what was sent - for tests.
        internal Func<AcceptInvitationRequest, Task<ReviewApiResult<AcceptInvitationAnswer>>>? AcceptInvitationOverrideForTests { get; set; }
    }
}
