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
    // address to maintain a system (the Systems screen); the mail carries a code, and this is where
    // it is redeemed - the code, a name and a password - which makes the account and puts it in the
    // pool of every system that address was invited to. See CRT.Server's MaintainerInvitationFlows.
    //
    // *** ON SUCCESS THE SIGN-IN FIELDS ARE FILLED, NOT A SESSION OPENED *** - the password reset's
    // rule, for the same reason: accepting is not signing in, the server issues no session for it,
    // and holding the password to sign in behind the person's back would be inventing one.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private void OnHaveInvitationClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (this.FindControl<StackPanel>("InvitationPanel") is not StackPanel panel)
                return;

            panel.IsVisible = true;
            this.FindControl<TextBox>("InvitationCodeTextBox")?.Focus();
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
                this.ShowSignInMessage("Paste the code from the invitation email first.");
                return;
            }

            if (enteredName.Length == 0)
            {
                this.ShowSignInMessage("Type the name others will see.");
                return;
            }

            if (enteredPassword.Length == 0)
            {
                this.ShowSignInMessage("Type the password you want to use.");
                return;
            }

            button.IsEnabled = false;

            try
            {
                ReviewApiResult<AcceptInvitationAnswer> result = await ServerWait.CallAsync(
                    this,
                    MaintainerWaitWording.AcceptingInvitation,
                    token => this.SendAcceptanceAsync(enteredCode, enteredName, enteredPassword, token));

                // No answer in two minutes: the account may exist now, and the code may be spent -
                // signing in is the way to find out, never "use the code again" first.
                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    this.ShowSignInMessage(MaintainerWaitWording.InvitationNoAnswer);
                    return;
                }

                if (!result.IsOk)
                {
                    this.ShowSignInMessage(result.Message);
                    return;
                }

                AcceptInvitationAnswer answer = result.Value!;

                this.ShowSignInMessage(answer.Message, isError: false);

                // The code is spent; a second press would only be told so.
                code.Text = string.Empty;
                name.Text = string.Empty;
                password.Text = string.Empty;

                if (this.FindControl<StackPanel>("InvitationPanel") is StackPanel panel)
                    panel.IsVisible = false;

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
