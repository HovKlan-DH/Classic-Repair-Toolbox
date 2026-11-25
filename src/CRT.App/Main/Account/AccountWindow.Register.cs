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
    // "CREATE AN ACCOUNT" AND "I HAVE AN INVITATION" in CRT's account window.
    //
    // CREATE AN ACCOUNT (2026-11-25, accounts for everybody): a display name, an address and a password, sent to POST
    // /api/accounts/register. The name is held to CRT.Data's DisplayNames AS IT IS TYPED - the server's own rule, said under
    // the box in its words - and whether another account has it is the server's to say. The answer is the server's NEUTRAL
    // one, the same whether or not the address had an account ("If that address can receive mail ..." - AccountFlows' header),
    // so the sign-in view comes back with the address filled in and that sentence: the person confirms the address from the
    // mail, then signs in. Making an account is not signing in - the server issues no session for it.
    //
    // I HAVE AN INVITATION (2026-09-27; the "Maintainer" tab's sign-in screen's until 2026-11-25): the code from the
    // administrator's mail, a name and a password make the account for the address invited and put it in the pool of every
    // board that address was invited to. On success the sign-in view comes back filled in, one press from signed in - as for
    // a new password, no session is invented.
    // ###########################################################################################
    public partial class AccountWindow
    {
        // -----------------------------------------------------------------------------------
        // Create an account
        // -----------------------------------------------------------------------------------

        private void OnCreateAccountClick(object? sender, RoutedEventArgs e)
        {
            TextBox email = this.Box("RegisterEmailTextBox");

            if (string.IsNullOrWhiteSpace(email.Text))
                email.Text = this.Box("EmailTextBox").Text?.Trim();

            this.ShowView(AccountView.Register);
        }

        // The rule's sentence under the box as the name is typed - nothing for an empty box, which "Create an account" asks for.
        private void OnRegisterNameChanged(object? sender, TextChangedEventArgs e)
        {
            string? typed = this.Box("RegisterNameTextBox").Text;
            string? problem = string.IsNullOrWhiteSpace(typed) ? null : DisplayNames.Problem(typed, sayRule: false);

            TextBlock line = this.FindControl<TextBlock>("RegisterNameProblemText")!;
            line.Text = problem ?? string.Empty;
            line.IsVisible = problem is not null;
        }

        private async void OnRegisterClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.RegisterAsync, message => this.ShowRegisterMessage(message, isError: true));

        private async void OnRegisterPasswordKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
                return;

            e.Handled = true;
            await HandlerGuard.RunAsync(this.RegisterAsync, message => this.ShowRegisterMessage(message, isError: true));
        }

        internal async Task RegisterAsync()
        {
            if (this.thisHost is not AccountHost host)
                return;

            string name = this.Box("RegisterNameTextBox").Text?.Trim() ?? string.Empty;
            string email = this.Box("RegisterEmailTextBox").Text?.Trim() ?? string.Empty;
            TextBox password = this.Box("RegisterPasswordTextBox");
            string typedPassword = password.Text ?? string.Empty;

            // What the round trip could say no better: the name's rule (the server's own), an empty address or password.
            if (DisplayNames.Problem(name, sayRule: false) is string nameProblem)
            {
                this.ShowRegisterMessage(nameProblem, isError: true);
                return;
            }

            if (email.Length == 0)
            {
                this.ShowRegisterMessage("Type your email address.", isError: true);
                return;
            }

            if (typedPassword.Length == 0)
            {
                this.ShowRegisterMessage("Type the password you want to use.", isError: true);
                return;
            }

            Button button = this.FindControl<Button>("RegisterButton")!;
            button.IsEnabled = false;
            this.ShowRegisterMessage(null, isError: false);

            try
            {
                using ReviewApiClient client = host.NewClient();

                ReviewApiResult<string> result = await ServerWait.CallAsync(
                    this, MaintainerWaitWording.CreatingAccount, token => client.RegisterAsync(email, typedPassword, name, token));

                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    this.ShowRegisterMessage(MaintainerWaitWording.CreateAccountNoAnswer, isError: true);
                    return;
                }

                if (!result.IsOk)
                {
                    this.ShowRegisterMessage(result.Message, isError: true);
                    return;
                }

                // The server's neutral sentence, on the sign-in view with the address filled in - the password is not kept:
                // the account may not be this person's (the address had one already), and signing in waits for the mail.
                password.Text = string.Empty;
                this.Box("EmailTextBox").Text = email;
                this.ShowView(AccountView.SignIn);
                this.ShowSignInMessage(AccountWindow.AfterRegistering(result.Value), isError: false);
            }
            finally
            {
                button.IsEnabled = true;
            }
        }

        // What the sign-in view says once an account was asked for: the server's neutral sentence, then what to do next.
        internal static string AfterRegistering(string? serverSentence) =>
            (string.IsNullOrWhiteSpace(serverSentence) ? "If that address can receive mail, a message is on its way to it." : serverSentence.Trim()) +
            " Open the link in it to confirm your address, then sign in here with your address and password.";

        private void ShowRegisterMessage(string? message, bool isError) =>
            WindowMessage.ShowOutcome(this.FindControl<TextBlock>("RegisterMessageText"), message, isError);

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
