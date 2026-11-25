using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // CRT'S ACCOUNT WINDOW (owner requests, 2026-11-25: accounts for everybody - "This user account can then be used for
    // "Maintainer" tab or this collection thing, but also anyone else can register an account"; "Allow registration both
    // places ... it should be clear that it gives access to both"). Opened by the account icon in the top right corner of CRT's
    // window (Main.Account.cs). Until then the only way in was the "Maintainer" tab's sign-in screen, and an account was made
    // only by the administrator's invitation.
    //
    // *** THE SESSION STAYS THE MAINTAINER TAB'S. *** The tab holds it (TabMaintainer.Session.cs - remembered, restored at
    // launch, slid as the server slides it, forgotten on a 401) and hands it to the rest of CRT (Main.ShareMaintainerSignIn);
    // this window only signs in, and hands what it got over through AccountHost.UseSignIn - Main gives it to the tab. Its "My
    // account" works on the tab's own client and session, and every account the server hands back goes to the tab
    // (AccountHost.AccountChanged), which owns what follows from it. So one session, one remembered file, as before.
    //
    // *** SIGNED IN, THE WINDOW CLOSES *** - the account icon then shows the sign-in; opened again, "My account". Signing out
    // closes it too, and the tab signs out (it asks about a table's unsaved changes first, over CRT's window).
    //
    // FILE MAP - this file (the views, opening, "My account" and the address not confirmed), AccountWindow.SignIn.cs (signing
    // in, "I forgot my password") and AccountWindow.Register.cs ("Create an account", "I have an invitation").
    // ###########################################################################################
    public partial class AccountWindow : Window
    {
        private AccountHost? thisHost;

        public AccountWindow()
        {
            this.InitializeComponent();
            Handlers.Theming.ScreenSizedDialog.Use(this);
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        // ###########################################################################################
        // What the window works with - Main's (the Maintainer tab's session, a client for a request before anybody is signed
        // in, and where a sign-in, an account changed and a sign-out go), or a test's. Shows the view for where the account
        // stands now: "My account" when signed in, else the sign-in.
        // ###########################################################################################
        internal void Initialize(AccountHost host)
        {
            ArgumentNullException.ThrowIfNull(host);

            this.thisHost = host;

            this.ShowNotRememberedNote(host.SignInIsRemembered);
            this.ShowTestServerNote(host.UsesTestServer);

            MyAccountView account = this.MyAccountDetail;
            account.AccountChanged = this.OnAccountChanged;
            account.SignOutRequested = this.SignOutAsync;

            if (host.SignedIn() is ReviewSession session && host.ClientInUse() is ReviewApiClient client)
            {
                account.Initialize(client, session);
                this.ShowView(AccountView.SignedIn);
            }
            else
            {
                account.Initialize(null, null);
                this.ShowView(AccountView.SignIn);
            }
        }

        // The account as the server holds it, read once the window is shown signed in - whether the address is confirmed.
        protected override async void OnOpened(EventArgs e)
        {
            base.OnOpened(e);

            await HandlerGuard.RunAsync(this.ReadAccountAsync, message => this.ShowAccountMessage(message, isError: true));
        }

        private MyAccountView MyAccountDetail => this.FindControl<MyAccountView>("MyAccountPanel")!;

        // -----------------------------------------------------------------------------------
        // The five views
        // -----------------------------------------------------------------------------------

        internal enum AccountView
        {
            SignIn,
            Register,
            Forgot,
            Invitation,
            SignedIn
        }

        internal AccountView ViewShown { get; private set; }

        // One view at a time, never two - the other views' boxes stay as they were, so Cancel goes back to the sign-in as it
        // was left. The first box of the view shown takes the keyboard.
        internal void ShowView(AccountView view)
        {
            this.ViewShown = view;

            this.Panel("SignInFormPanel").IsVisible = view == AccountView.SignIn;
            this.Panel("RegisterPanel").IsVisible = view == AccountView.Register;
            this.Panel("ForgotPanel").IsVisible = view == AccountView.Forgot;
            this.Panel("InvitationPanel").IsVisible = view == AccountView.Invitation;
            this.Panel("SignedInPanel").IsVisible = view == AccountView.SignedIn;

            this.Title = view switch
            {
                AccountView.Register => AccountWording.CreateAccountButton,
                AccountView.SignedIn => "My account",
                _ => "Account"
            };

            string? first = view switch
            {
                AccountView.SignIn => "EmailTextBox",
                AccountView.Register => "RegisterNameTextBox",
                AccountView.Forgot => "ForgotEmailTextBox",
                AccountView.Invitation => "InvitationCodeTextBox",
                _ => null
            };

            if (first is not null)
                this.FindControl<TextBox>(first)?.Focus();
        }

        // ###########################################################################################
        // Cancel - back to the sign-in view as it was left. *** A PASSWORD TYPED IN THE VIEW LEFT GOES WITH IT *** (the
        // invitation's rule since 2026-10-04): a password must not sit in a hidden box. The invitation is emptied whole, its
        // code and message too, so the next "I have an invitation" starts blank; the other views keep the name, the address
        // and a code typed, should the person come back.
        // ###########################################################################################
        private void OnBackToSignInClick(object? sender, RoutedEventArgs e)
        {
            switch (this.ViewShown)
            {
                case AccountView.Invitation:
                    this.ClearInvitation();
                    break;

                case AccountView.Register:
                    this.Box("RegisterPasswordTextBox").Text = string.Empty;
                    break;

                case AccountView.Forgot:
                    this.Box("NewPasswordTextBox").Text = string.Empty;
                    break;
            }

            this.ShowView(AccountView.SignIn);
        }

        private void OnCloseClick(object? sender, RoutedEventArgs e) => this.Close();

        // -----------------------------------------------------------------------------------
        // Signed in: "My account", and the address not confirmed yet
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // GET /api/accounts/me, quietly - nobody pressed anything, so no "please wait": the window is usable while it is
        // asked. Shows the notice for an address not confirmed, and hands the account on (the name or address may have
        // changed elsewhere). A failure says nothing - "My account" still works on the session held.
        // ###########################################################################################
        private async Task ReadAccountAsync()
        {
            if (this.ViewShown != AccountView.SignedIn || this.thisHost?.ClientInUse() is not ReviewApiClient client || this.thisHost.SignedIn() is not ReviewSession session)
                return;

            ReviewApiResult<AccountAnswer> read = await client.GetAccountAsync(session);

            if (read.IsOk && read.Value is AccountAnswer account)
                this.OnAccountChanged(account);
        }

        // Every account the server hands back - "My account"'s changes and the read above: to the tab, which owns the session,
        // and to the notice.
        private void OnAccountChanged(AccountAnswer account)
        {
            this.thisHost?.AccountChanged(account);

            if (this.thisHost?.SignedIn() is ReviewSession session)
                this.MyAccountDetail.UseAccount(session);

            this.ShowNotConfirmed(!account.IsVerified);
        }

        private void ShowNotConfirmed(bool show) => this.Panel("NotConfirmedPanel").IsVisible = show;

        private async void OnSendMailAgainClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.SendMailAgainAsync, message => this.ShowAccountMessage(message, isError: true));

        // A new link to the account's own address - or, confirmed meanwhile, the server says so and the notice goes.
        internal async Task SendMailAgainAsync()
        {
            if (this.thisHost?.ClientInUse() is not ReviewApiClient client || this.thisHost.SignedIn() is not ReviewSession session)
                return;

            Button button = this.FindControl<Button>("SendMailAgainButton")!;
            button.IsEnabled = false;
            this.ShowAccountMessage(null, isError: false);

            try
            {
                ReviewApiResult<AccountChangeAnswer> result = await ServerWait.CallAsync(
                    this, MaintainerWaitWording.SendingVerificationAgain, token => client.SendVerificationAgainAsync(session, token));

                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    this.ShowAccountMessage(MaintainerWaitWording.VerificationAgainNoAnswer, isError: false);
                    return;
                }

                if (!result.IsOk)
                {
                    this.ShowAccountMessage(result.Message, isError: true);
                    return;
                }

                if (result.Value!.Account is AccountAnswer account)
                    this.OnAccountChanged(account);

                this.ShowAccountMessage(result.Value.Message, isError: false);
            }
            finally
            {
                button.IsEnabled = true;
            }
        }

        // "Sign out" under "My account": the window closes, then the tab signs out - it asks about a table's unsaved changes
        // first, over CRT's window, and forgets the remembered sign-in.
        private async Task SignOutAsync()
        {
            this.Close();

            if (this.thisHost is AccountHost host)
                await host.SignOut();
        }

        private void ShowAccountMessage(string? message, bool isError) =>
            WindowMessage.ShowOutcome(this.FindControl<TextBlock>("AccountMessageText"), message, isError);

        // -----------------------------------------------------------------------------------
        // Shared
        // -----------------------------------------------------------------------------------

        // The line under "Sign in" saying a sign-in is not remembered on this system (ReviewSessionStore.NotRememberedNote).
        private void ShowNotRememberedNote(bool isSupported)
        {
            TextBlock text = this.FindControl<TextBlock>("NotRememberedText")!;
            string? note = ReviewSessionStore.NotRememberedNote(isSupported);

            text.Text = note ?? string.Empty;
            text.IsVisible = note is not null;
        }

        // The line over the sign-in saying it is the TEST instance's (ContributionServer.SignInNote).
        private void ShowTestServerNote(bool usesTestServer)
        {
            TextBlock text = this.FindControl<TextBlock>("TestServerSignInText")!;

            text.Text = usesTestServer ? Handlers.Online.ContributionServer.SignInNote : string.Empty;
            text.IsVisible = usesTestServer;
        }

        private Control Panel(string name) => this.FindControl<Control>(name)!;

        private TextBox Box(string name) => this.FindControl<TextBox>(name)!;
    }

    // ###########################################################################################
    // What the account window works with - Main's, or a test's:
    //
    //   SignedIn            the session in use (the Maintainer tab's), or null
    //   ClientInUse         the client its requests go through - "My account" uses it, so the server's sliding expiry is
    //                       followed as for every other request of the session
    //   NewClient           a client for a request before anybody is signed in (sign in, a code, an account, an invitation)
    //   UseSignIn           signed in here: the session and the client it was got with, handed to the tab
    //   AccountChanged      the account as the server now holds it, to the tab
    //   SignOut             the tab's own sign-out
    //   SignInIsRemembered  whether this system can remember a sign-in (ReviewSessionStore.IsSupported)
    //   UsesTestServer      whether CRT talks to the test instance (AppConfig.UsesTestServer)
    // ###########################################################################################
    internal sealed record AccountHost(
        Func<ReviewSession?> SignedIn,
        Func<ReviewApiClient?> ClientInUse,
        Func<ReviewApiClient> NewClient,
        Func<ReviewSession, ReviewApiClient, Task> UseSignIn,
        Action<AccountAnswer> AccountChanged,
        Func<Task> SignOut,
        bool SignInIsRemembered,
        bool UsesTestServer);
}
