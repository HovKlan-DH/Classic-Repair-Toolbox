using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // "MY ACCOUNT" (owner requests, 2026-10-03: the maintainer "can edit his/her own email address
    // and name"; and 2026-10-04: moved from the "Your account" window, opened from the bottom-left
    // corner, to the first entry under the "Account" screen, Sign out with it). Shown by
    // TabMaintainer.Account.cs; the server's side is AccountSelfServiceFlows.
    //
    //   NAME     - one request ("Name or handle").
    //   ADDRESS  - two: "Send code" mails a code to the new address, and the code typed in after
    //              "I have received a code" makes the change ("Change address"). Another screen can
    //              be shown, or CRT closed, in between (owner wording, 2026-10-03).
    //   PASSWORD - the new one twice.
    //   SIGN OUT - handed to the tab (SignOutRequested), which asks about unsaved table changes
    //              first and owns everything signing out clears.
    //
    // *** NO CURRENT PASSWORD ANYWHERE (owner decision, 2026-10-03: "no need for that, as I see it
    // as you are already logged in"). *** The session is enough - an accepted risk, see the server's
    // AccountSelfServiceFlows. An address or password change signs out every OTHER computer.
    //
    // *** EVERY ACCOUNT THE SERVER HANDS BACK IS PASSED ON AT ONCE *** (AccountChanged), so the
    // tab's "Logged in as" line, the remembered sign-in and - through the tab's SignedInChanged - the
    // Feedback tab and the Submit dialog follow the change. A change arriving the other way (the
    // name and address read again at launch) comes in through UseAccount.
    //
    // *** A TIMEOUT IS NEVER "IT FAILED" *** (WaitWording's rule). The name and the address are
    // read back from the server and the sentence says what it holds; a code or a password cannot be
    // read back, and the sentence says what may have happened.
    //
    // The new password is cleared from its boxes once it is set: nothing here needs it again, and a
    // filled password box left on screen is a credential in plain sight (the sign-in screen's rule).
    // For the same reason everything typed here goes with the session: Initialize(null, null) on
    // signing out empties every box, so the next account on this computer finds nothing waiting.
    // ###########################################################################################
    public partial class MyAccountView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        public MyAccountView()
        {
            this.InitializeComponent();
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        // Every account the server hands back, as it arrives - the tab's ApplyAccount.
        internal Action<AccountAnswer>? AccountChanged { get; set; }

        // "Sign out" pressed - the tab's own sign-out, which asks about unsaved changes first.
        internal Func<Task>? SignOutRequested { get; set; }

        // The session as it stands after whatever was changed here - the token never changes.
        internal ReviewSession? Session => this.thisSession;

        // ###########################################################################################
        // The account signed in on the tab - or none, on signing out. Either way everything typed for
        // the previous one goes: a half-typed password or a code left in a box would be waiting for
        // whoever signs in next on this computer.
        // ###########################################################################################
        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;

            foreach (string box in new[] { "NewEmailTextBox", "CodeTextBox", "NewPasswordTextBox", "RepeatPasswordTextBox" })
                this.Box(box).Text = string.Empty;

            foreach (string message in new[] { "NameMessageText", "EmailMessageText", "PasswordMessageText" })
                this.ShowMessage(message, null);

            this.Panel("CodePanel").IsVisible = false;
            this.Box("NameTextBox").Text = session?.DisplayName ?? string.Empty;

            this.ShowLoggedInAs();
            this.RefreshButtons();
        }

        // ###########################################################################################
        // The account as changed somewhere else - read again at launch, or by this view itself on
        // its way round through the tab. The name box follows unless something else is being typed
        // into it.
        // ###########################################################################################
        internal void UseAccount(ReviewSession session)
        {
            ArgumentNullException.ThrowIfNull(session);

            if (this.thisSession is not { } before)
                return;

            TextBox name = this.Box("NameTextBox");

            if (string.Equals(name.Text?.Trim() ?? string.Empty, before.DisplayName, StringComparison.Ordinal))
                name.Text = session.DisplayName;

            this.thisSession = session;
            this.ShowLoggedInAs();
            this.RefreshButtons();
        }

        // -----------------------------------------------------------------------------------
        // Name
        // -----------------------------------------------------------------------------------

        private async void OnSaveNameClick(object? sender, RoutedEventArgs e) => await this.SaveNameAsync();

        internal async Task SaveNameAsync()
        {
            if (this.thisClient is not { } client || this.thisSession is not { } session)
                return;

            string wanted = this.Box("NameTextBox").Text?.Trim() ?? string.Empty;

            if (!MyAccountRules.CanSaveName(session.DisplayName, wanted))
                return;

            this.ShowMessage("NameMessageText", null);

            ReviewApiResult<AccountChangeAnswer> result = await ServerWait.CallAsync(
                this, MaintainerWaitWording.SavingName, token => client.ChangeNameAsync(session, wanted, token));

            if (result.Failure == ReviewApiFailure.TimedOut)
            {
                AccountAnswer? now = await this.ReadAccountAgainAsync();
                bool done = now is not null && string.Equals(now.DisplayName, wanted, StringComparison.Ordinal);

                this.ShowMessage("NameMessageText", MaintainerWaitWording.NameAfterTimeout(wanted, now), isError: !done);
                return;
            }

            if (!result.IsOk)
            {
                this.ShowMessage("NameMessageText", result.Message);
                return;
            }

            this.Apply(result.Value!.Account);
            this.ShowMessage("NameMessageText", result.Value.Message, isError: false);
        }

        // -----------------------------------------------------------------------------------
        // Email address
        // -----------------------------------------------------------------------------------

        private async void OnSendCodeClick(object? sender, RoutedEventArgs e) => await this.SendCodeAsync();

        internal async Task SendCodeAsync()
        {
            if (this.thisClient is not { } client || this.thisSession is not { } session)
                return;

            string wanted = this.Box("NewEmailTextBox").Text?.Trim() ?? string.Empty;

            if (!MyAccountRules.CanAskForCode(session.Email, wanted))
                return;

            this.ShowMessage("EmailMessageText", null);

            ReviewApiResult<AccountChangeAnswer> result = await ServerWait.CallAsync(
                this, MaintainerWaitWording.SendingEmailCode, token => client.RequestEmailChangeAsync(session, wanted, token));

            if (result.Failure == ReviewApiFailure.TimedOut)
            {
                this.ShowMessage("EmailMessageText", MaintainerWaitWording.EmailCodeNoAnswer, isError: false);
                return;
            }

            if (!result.IsOk)
            {
                this.ShowMessage("EmailMessageText", result.Message);
                return;
            }

            AccountChangeAnswer answer = result.Value!;

            // The code is typed in after "I have received a code" (owner wording, 2026-10-03), which
            // the section's text names - so the box stays put away until then. Only the capitals
            // changing is done already, with no code.
            if (!answer.CodeSent)
            {
                this.Apply(answer.Account);
                this.Box("NewEmailTextBox").Text = string.Empty;
            }

            this.ShowMessage("EmailMessageText", answer.Message, isError: false);
        }

        private void OnHaveCodeClick(object? sender, RoutedEventArgs e) => this.ShowCodePanel();

        private async void OnConfirmEmailClick(object? sender, RoutedEventArgs e) => await this.ConfirmCodeAsync();

        internal async Task ConfirmCodeAsync()
        {
            if (this.thisClient is not { } client || this.thisSession is not { } session)
                return;

            string code = this.Box("CodeTextBox").Text?.Trim() ?? string.Empty;

            if (!MyAccountRules.CanConfirmCode(code))
                return;

            string before = session.Email;

            this.ShowMessage("EmailMessageText", null);

            ReviewApiResult<AccountChangeAnswer> result = await ServerWait.CallAsync(
                this, MaintainerWaitWording.ChangingEmail, token => client.ConfirmEmailChangeAsync(session, code, token));

            if (result.Failure == ReviewApiFailure.TimedOut)
            {
                AccountAnswer? now = await this.ReadAccountAgainAsync();
                bool done = now is not null && !string.Equals(now.Email, before, StringComparison.Ordinal);

                this.ShowMessage("EmailMessageText", MaintainerWaitWording.EmailAfterTimeout(before, now), isError: !done);
                return;
            }

            if (!result.IsOk)
            {
                this.ShowMessage("EmailMessageText", result.Message);
                return;
            }

            this.Apply(result.Value!.Account);

            this.Box("CodeTextBox").Text = string.Empty;
            this.Box("NewEmailTextBox").Text = string.Empty;
            this.Panel("CodePanel").IsVisible = false;

            this.ShowMessage("EmailMessageText", result.Value.Message, isError: false);
        }

        private void ShowCodePanel()
        {
            this.Panel("CodePanel").IsVisible = true;
            this.Box("CodeTextBox").Focus();
        }

        // -----------------------------------------------------------------------------------
        // Password
        // -----------------------------------------------------------------------------------

        private async void OnChangePasswordClick(object? sender, RoutedEventArgs e) => await this.ChangePasswordAsync();

        internal async Task ChangePasswordAsync()
        {
            if (this.thisClient is not { } client || this.thisSession is not { } session)
                return;

            string wanted = this.Box("NewPasswordTextBox").Text ?? string.Empty;

            if (!MyAccountRules.CanChangePassword(wanted, this.Box("RepeatPasswordTextBox").Text))
                return;

            this.ShowMessage("PasswordMessageText", null);

            ReviewApiResult<AccountChangeAnswer> result = await ServerWait.CallAsync(
                this, MaintainerWaitWording.ChangingPassword, token => client.ChangePasswordAsync(session, wanted, token));

            // The new password is kept only when the server refused it, so a rule broken can be
            // fixed without typing both again.
            if (result.Failure == ReviewApiFailure.TimedOut)
            {
                this.ClearNewPasswords();
                this.ShowMessage("PasswordMessageText", MaintainerWaitWording.NewPasswordNoAnswer);
                return;
            }

            if (!result.IsOk)
            {
                this.ShowMessage("PasswordMessageText", result.Message);
                return;
            }

            this.ClearNewPasswords();
            this.Apply(result.Value!.Account);
            this.ShowMessage("PasswordMessageText", result.Value.Message, isError: false);
        }

        private void ClearNewPasswords()
        {
            this.Box("NewPasswordTextBox").Text = string.Empty;
            this.Box("RepeatPasswordTextBox").Text = string.Empty;
        }

        // -----------------------------------------------------------------------------------
        // Sign out
        // -----------------------------------------------------------------------------------

        private async void OnSignOutClick(object? sender, RoutedEventArgs e)
        {
            if (this.SignOutRequested is { } signOut)
                await signOut();
        }

        // -----------------------------------------------------------------------------------
        // Shared
        // -----------------------------------------------------------------------------------

        // The account the server now holds: this view's session follows it, and so does everything
        // outside the view, at once.
        private void Apply(AccountAnswer? account)
        {
            if (account is null || this.thisSession is not { } session)
                return;

            this.thisSession = session.WithAccount(account);
            this.Box("NameTextBox").Text = this.thisSession.DisplayName;
            this.ShowLoggedInAs();
            this.RefreshButtons();

            this.AccountChanged?.Invoke(account);
        }

        // After a timeout on a change that can be read back: what the server holds now, applied, or
        // null when that look failed too.
        private async Task<AccountAnswer?> ReadAccountAgainAsync()
        {
            if (this.thisClient is not { } client || this.thisSession is not { } session)
                return null;

            ReviewApiResult<AccountAnswer> read = await ServerWait.CallAsync(
                this, WaitWording.Checking, token => client.GetAccountAsync(session, token));

            if (!read.IsOk)
                return null;

            this.Apply(read.Value);
            return read.Value;
        }

        private void OnTextChanged(object? sender, TextChangedEventArgs e) => this.RefreshButtons();

        private void RefreshButtons()
        {
            if (this.thisSession is not { } session)
            {
                foreach (string button in new[] { "SaveNameButton", "SendCodeButton", "ConfirmEmailButton", "ChangePasswordButton" })
                    this.Button(button).IsEnabled = false;

                return;
            }

            this.Button("SaveNameButton").IsEnabled = MyAccountRules.CanSaveName(session.DisplayName, this.Box("NameTextBox").Text);

            this.Button("SendCodeButton").IsEnabled = MyAccountRules.CanAskForCode(session.Email, this.Box("NewEmailTextBox").Text);

            this.Button("ConfirmEmailButton").IsEnabled = MyAccountRules.CanConfirmCode(this.Box("CodeTextBox").Text);

            string? newPassword = this.Box("NewPasswordTextBox").Text;
            string? repeated = this.Box("RepeatPasswordTextBox").Text;

            this.Button("ChangePasswordButton").IsEnabled = MyAccountRules.CanChangePassword(newPassword, repeated);

            // The two new passwords disagreeing is said where the button is, but only until the
            // server has said something of its own there.
            if (MyAccountRules.PasswordProblem(newPassword, repeated) is string problem)
                this.ShowMessage("PasswordMessageText", problem);
            else if (this.FindControl<TextBlock>("PasswordMessageText")?.Text == MyAccountRules.PasswordsDiffer)
                this.ShowMessage("PasswordMessageText", null);
        }

        // "Logged in as Dennis (dh@example.com)", the name bold - the tab's own line, on one line.
        private void ShowLoggedInAs()
        {
            if (this.FindControl<TextBlock>("LoggedInAsText") is not TextBlock text)
                return;

            if (this.thisSession is not { } session)
            {
                text.Inlines?.Clear();
                text.Text = string.Empty;
                return;
            }

            TabMaintainer.ShowCounts(text, [new ReviewNoteRun("Logged in as ", false), .. MyAccountRules.LoggedInAs(session.DisplayName, session.Email)]);
        }

        // Red for a refusal, green for a change made - the sign-in screen's two colours.
        private void ShowMessage(string name, string? message, bool isError = true)
        {
            if (this.FindControl<TextBlock>(name) is not TextBlock text)
                return;

            text.Text = message ?? string.Empty;
            text.IsVisible = !string.IsNullOrWhiteSpace(message);
            text.Foreground = isError ? Brushes.IndianRed : Brushes.SeaGreen;
        }

        private TextBox Box(string name) => this.FindControl<TextBox>(name)!;

        private Button Button(string name) => this.FindControl<Button>(name)!;

        private Control Panel(string name) => this.FindControl<Control>(name)!;
    }
}
