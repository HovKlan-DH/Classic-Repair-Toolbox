using System;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // "I HAVE AN INVITATION" in CRT's account window (2026-09-27; the "Maintainer" tab's sign-in screen's until 2026-11-25):
    // the code from the administrator's mail, a name and a password make the account for the address invited and put it in
    // the pool of every board that address was invited to - THE ONE WAY AN ACCOUNT IS MADE (owner decision, 2026-11-28:
    // "Entirely remove the new "Account" possibility"; "Create an account" was here from 2026-11-25). On success the sign-in
    // view comes back filled in, one press from signed in - as for a new password, no session is invented.
    // ###########################################################################################
    public partial class AccountWindow
    {
        // -----------------------------------------------------------------------------------
        // I have an invitation
        // -----------------------------------------------------------------------------------

        private void OnHaveInvitationClick(object? sender, RoutedEventArgs e) => this.ShowView(AccountView.Invitation);

        // Enter in the password box accepts, as it signs in in the box above.
        private async void OnInvitationPasswordKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
                return;

            e.Handled = true;
            await HandlerGuard.RunAsync(this.AcceptInvitationAsync, this.ShowInvitationMessage);
        }

        private async void OnAcceptInvitationClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.AcceptInvitationAsync, this.ShowInvitationMessage);

        internal async Task AcceptInvitationAsync()
        {
            TextBox code = this.Box("InvitationCodeTextBox");
            TextBox name = this.Box("InvitationNameTextBox");
            TextBox password = this.Box("InvitationPasswordTextBox");
            Button button = this.FindControl<Button>("AcceptInvitationButton")!;

            string enteredCode = code.Text?.Trim() ?? string.Empty;
            string enteredName = name.Text?.Trim() ?? string.Empty;
            string enteredPassword = password.Text ?? string.Empty;

            // Checked here because the round trip could say nothing more useful about an empty box - or a name the rule refuses.
            if (enteredCode.Length == 0)
            {
                this.ShowInvitationMessage("Paste the code from the invitation email first.");
                return;
            }

            if (DisplayNames.Problem(enteredName, sayRule: false) is string nameProblem)
            {
                this.ShowInvitationMessage(nameProblem);
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

                // No answer in two minutes: the account may exist now and the code may be spent - signing in is the way to
                // find out, never "use the code again" first. So the sign-in view comes back with the password chosen
                // filled in, and says so there; the address is the one invited, which only the invitee knows here.
                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    this.ClearInvitation();
                    this.ShowView(AccountView.SignIn);
                    this.ShowSignInMessage(MaintainerWaitWording.InvitationNoAnswer, isError: true);
                    this.Box("PasswordTextBox").Text = enteredPassword;
                    return;
                }

                if (!result.IsOk)
                {
                    this.ShowInvitationMessage(result.Message);
                    return;
                }

                AcceptInvitationAnswer answer = result.Value!;

                // The code is spent: its boxes emptied, and the sign-in view back, saying the account is ready.
                this.ClearInvitation();
                this.ShowView(AccountView.SignIn);
                this.ShowSignInMessage(answer.Message, isError: false);
                this.Box("EmailTextBox").Text = answer.Email;
                this.Box("PasswordTextBox").Text = enteredPassword;
                this.Box("PasswordTextBox").Focus();
            }
            finally
            {
                button.IsEnabled = true;
            }
        }

        private void ClearInvitation()
        {
            foreach (string box in new[] { "InvitationCodeTextBox", "InvitationNameTextBox", "InvitationPasswordTextBox" })
                this.Box(box).Text = string.Empty;

            this.ShowInvitationMessage(null);
        }

        // The invitation view's own line - always a refusal: success goes back to the sign-in view and says so there.
        private void ShowInvitationMessage(string? message)
        {
            TextBlock text = this.FindControl<TextBlock>("InvitationMessageText")!;

            text.Text = message ?? string.Empty;
            text.IsVisible = !string.IsNullOrWhiteSpace(message);
        }

        // Answers the acceptance instead of the server, and sees what was sent - for a test of what no fake server can give
        // quickly: two minutes with no answer.
        internal Func<AcceptInvitationRequest, Task<ReviewApiResult<AcceptInvitationAnswer>>>? AcceptInvitationOverrideForTests { get; set; }

        private async Task<ReviewApiResult<AcceptInvitationAnswer>> SendAcceptanceAsync(string code, string name, string password, CancellationToken token)
        {
            if (this.AcceptInvitationOverrideForTests is not null)
                return await this.AcceptInvitationOverrideForTests(new AcceptInvitationRequest(code, name, password));

            if (this.thisHost is not AccountHost host)
                return ReviewApiResult<AcceptInvitationAnswer>.Failed(ReviewApiFailure.Unreachable, "Nothing to send it with.");

            using ReviewApiClient client = host.NewClient();

            return await client.AcceptInvitationAsync(code, name, password, token);
        }
    }
}
