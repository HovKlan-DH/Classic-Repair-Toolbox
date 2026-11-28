using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // SIGNING IN AND "I FORGOT MY PASSWORD" in CRT's account window (2026-11-25 - the "Maintainer" tab's sign-in screen until
    // then, moved here whole when accounts opened to everybody; its rules came with it).
    //
    // *** A SIGN-IN IS HANDED TO THE MAINTAINER TAB, WHICH OWNS THE SESSION *** (AccountHost.UseSignIn) - remembered there
    // (ReviewSessionStore.Remember, on a successful sign-in only: never a token that has not been proved to work), shared with
    // the rest of CRT from there. The window then closes.
    //
    // *** "I FORGOT MY PASSWORD" ANSWERS THE SAME FOR EVERY ADDRESS. *** The server's neutral 202, so this view never tells
    // whether an address has an account - asking for a code shows the same sentence whatever it is. The code is typed into
    // the same view, as on the account web page: a code mailed some other way needs no request first.
    //
    // *** RESETTING A PASSWORD IS NOT SIGNING IN. *** The server issues no session for it, so the sign-in view comes back with
    // the address and the new password filled in - one press from signed in, with no session invented behind the person's back.
    // ###########################################################################################
    public partial class AccountWindow
    {
        // -----------------------------------------------------------------------------------
        // Sign in
        // -----------------------------------------------------------------------------------

        private async void OnSignInClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.SignInAsync, message => this.ShowSignInMessage(message));

        // Enter in the password box signs in - a safe, repeatable action.
        private async void OnPasswordKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
                return;

            e.Handled = true;
            await HandlerGuard.RunAsync(this.SignInAsync, message => this.ShowSignInMessage(message));
        }

        internal async Task SignInAsync()
        {
            if (this.thisHost is not AccountHost host)
                return;

            Button button = this.FindControl<Button>("SignInButton")!;
            TextBox password = this.Box("PasswordTextBox");

            string email = this.Box("EmailTextBox").Text?.Trim() ?? string.Empty;
            string typed = password.Text ?? string.Empty;

            // Disabled for the duration, so a second press cannot race a second session in.
            button.IsEnabled = false;
            this.ShowSignInMessage(null);

            ReviewApiClient client = host.NewClient();
            bool handedOver = false;

            try
            {
                // A timeout is an ordinary failure here: signing in changes nothing worth checking.
                ReviewApiResult<ReviewSession> result = await ServerWait.CallAsync(
                    this, MaintainerWaitWording.SigningIn, token => client.LoginAsync(email, typed, token));

                if (!result.IsOk)
                {
                    this.ShowSignInMessage(result.Message);
                    return;
                }

                // The password goes the moment it is no longer needed - a filled password box is a credential in plain sight.
                password.Text = string.Empty;

                // The window closes first: the tab reads its lists under CRT's own "please wait", over CRT's window.
                handedOver = true;
                this.Close();

                await host.UseSignIn(result.Value!, client);
            }
            finally
            {
                if (!handedOver)
                    client.Dispose();

                button.IsEnabled = true;
            }
        }

        // The sign-in view's one message line - red for a refusal, green for good news ("Your account is made").
        private void ShowSignInMessage(string? message, bool isError = true) =>
            WindowMessage.ShowOutcome(this.FindControl<TextBlock>("SignInMessageText"), message, isError);

        // -----------------------------------------------------------------------------------
        // I forgot my password
        // -----------------------------------------------------------------------------------

        // The view, with the address typed for signing in carried into step 1 - no request yet.
        private void OnForgotPasswordClick(object? sender, RoutedEventArgs e)
        {
            TextBox forgotEmail = this.Box("ForgotEmailTextBox");

            if (string.IsNullOrWhiteSpace(forgotEmail.Text))
                forgotEmail.Text = this.Box("EmailTextBox").Text?.Trim();

            this.ShowView(AccountView.Forgot);
        }

        private async void OnSendResetCodeClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.SendResetCodeAsync, message => this.ShowForgotMessage(message, isError: true));

        internal async Task SendResetCodeAsync()
        {
            if (this.thisHost is not AccountHost host)
                return;

            string email = this.Box("ForgotEmailTextBox").Text?.Trim() ?? string.Empty;

            // The one thing worth checking here - the server cannot mail nothing.
            if (email.Length == 0)
            {
                this.ShowForgotMessage("Type your email address first.", isError: true);
                return;
            }

            Button button = this.FindControl<Button>("SendResetCodeButton")!;
            button.IsEnabled = false;
            this.ShowForgotMessage(null, isError: false);

            try
            {
                using ReviewApiClient client = host.NewClient();

                ReviewApiResult<string> result = await ServerWait.CallAsync(
                    this, MaintainerWaitWording.AskingForResetCode, token => client.ForgotPasswordAsync(email, token));

                // A code may still be mailed after the window stopped waiting.
                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    this.ShowForgotMessage(MaintainerWaitWording.ResetCodeNoAnswer, isError: false);
                    return;
                }

                // The server's neutral sentence, the same for every address.
                this.ShowForgotMessage(result.IsOk ? result.Value : result.Message, isError: !result.IsOk);

                if (result.IsOk)
                    this.Box("ResetCodeTextBox").Focus();
            }
            finally
            {
                button.IsEnabled = true;
            }
        }

        private async void OnSetPasswordClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.SetPasswordAsync, message => this.ShowResetMessage(message, isError: true));

        private async void OnNewPasswordKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
                return;

            e.Handled = true;
            await HandlerGuard.RunAsync(this.SetPasswordAsync, message => this.ShowResetMessage(message, isError: true));
        }

        // ###########################################################################################
        // The code from the mail and the new password. Set: back to the sign-in view, the address and the new password filled
        // in and the good news said there - the code is spent, so its boxes are emptied (a second press would only be told
        // "already used").
        // ###########################################################################################
        internal async Task SetPasswordAsync()
        {
            if (this.thisHost is not AccountHost host)
                return;

            TextBox code = this.Box("ResetCodeTextBox");
            TextBox password = this.Box("NewPasswordTextBox");

            string typedCode = code.Text?.Trim() ?? string.Empty;
            string typedPassword = password.Text ?? string.Empty;

            if (typedCode.Length == 0)
            {
                this.ShowResetMessage("Paste the code from the email first.", isError: true);
                return;
            }

            if (typedPassword.Length == 0)
            {
                this.ShowResetMessage("Type the password you want to use.", isError: true);
                return;
            }

            Button button = this.FindControl<Button>("SetPasswordButton")!;
            button.IsEnabled = false;
            this.ShowResetMessage(null, isError: false);

            try
            {
                using ReviewApiClient client = host.NewClient();

                ReviewApiResult<string> result = await ServerWait.CallAsync(
                    this, MaintainerWaitWording.SettingPassword, token => client.ResetPasswordAsync(typedCode, typedPassword, token));

                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    this.ShowResetMessage(MaintainerWaitWording.PasswordNoAnswer, isError: true);
                    return;
                }

                if (!result.IsOk)
                {
                    this.ShowResetMessage(result.Message, isError: true);
                    return;
                }

                code.Text = string.Empty;
                password.Text = string.Empty;

                string? email = this.Box("ForgotEmailTextBox").Text?.Trim();

                if (!string.IsNullOrEmpty(email))
                    this.Box("EmailTextBox").Text = email;

                this.Box("PasswordTextBox").Text = typedPassword;
                this.ShowView(AccountView.SignIn);
                this.ShowSignInMessage(result.Value, isError: false);
                this.Box("PasswordTextBox").Focus();
            }
            finally
            {
                button.IsEnabled = true;
            }
        }

        private void ShowForgotMessage(string? message, bool isError) =>
            WindowMessage.ShowOutcome(this.FindControl<TextBlock>("ForgotMessageText"), message, isError);

        private void ShowResetMessage(string? message, bool isError) =>
            WindowMessage.ShowOutcome(this.FindControl<TextBlock>("ResetMessageText"), message, isError);
    }
}
