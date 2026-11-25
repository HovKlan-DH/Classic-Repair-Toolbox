using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE ACCOUNT ICON IN THE WINDOW'S TOP RIGHT CORNER (owner requests, 2026-11-25: accounts for everybody - "we now need a
    // possibility to have user accounts in CRT. This user account can then be used for "Maintainer" tab or this collection
    // thing, but also anyone else can register an account"). Right of the theme's icon, drawn as it is and as large, 22 x 22:
    // a person, or a person with a check while signed in - its tooltip saying which. A click opens the account window
    // (AccountWindow): signing in, "Create an account", "I forgot my password", "I have an invitation", or "My account".
    //
    // *** THE SESSION IS THE MAINTAINER TAB'S *** (TabMaintainer.Session.cs), as it always was: this part hands the window the
    // tab's session and client, and hands the tab what the window did (AccountHost) - a sign-in, an account changed, a sign-out.
    // The icon follows every change of it, from wherever it came (ShareMaintainerSignIn).
    // ###########################################################################################
    public partial class Main
    {
        // Font Awesome's user and user-check.
        internal const string SignedOutGlyph = "";
        internal const string SignedInGlyph = "";

        internal const string SignedOutTip = "Sign in or create an account";

        internal static string SignedInTip(ReviewSession session)
        {
            string name = string.IsNullOrWhiteSpace(session.DisplayName) ? session.Email : session.DisplayName.Trim();

            return $"Signed in as {name} - your account";
        }

        // The window shown instead of as a dialog in a test, and the client its requests go through - never the live server.
        internal Func<AccountWindow, Task>? ShowAccountWindowOverrideForTests { get; set; }

        internal Func<ReviewApiClient>? AccountClientOverrideForTests { get; set; }

        private void WireAccountIcon()
        {
            this.ShowAccountIcon();
            this.TabMaintainer.SignInWanted += () => _ = HandlerGuard.RunAsync(this.OpenAccountWindowAsync, _ => { });
        }

        private async void OnAccountButtonClick(object? sender, RoutedEventArgs e) =>
            await HandlerGuard.RunAsync(this.OpenAccountWindowAsync, _ => { });

        // ###########################################################################################
        // The account window, over CRT's - on the view for where the account stands: "My account" signed in, else the sign-in.
        // ###########################################################################################
        internal async Task OpenAccountWindowAsync()
        {
            var window = new AccountWindow();
            window.Initialize(this.AccountHostForWindow());

            if (this.ShowAccountWindowOverrideForTests is Func<AccountWindow, Task> show)
            {
                await show(window);
                return;
            }

            await window.ShowDialog(this);
        }

        internal AccountHost AccountHostForWindow() =>
            new(
                SignedIn: () => this.TabMaintainer.SignedIn,
                ClientInUse: () => this.TabMaintainer.ClientInUse,
                NewClient: () => this.AccountClientOverrideForTests?.Invoke() ?? new ReviewApiClient(ReviewApiRoutes.DefaultBaseAddress),
                UseSignIn: (session, client) => this.TabMaintainer.UseSignInAsync(session, client, readLists: UserSettings.EnableMaintainerTab),
                AccountChanged: this.TabMaintainer.ApplyAccount,
                SignOut: this.SignOutFromAccountWindowAsync,
                SessionEnded: this.TabMaintainer.EndSessionTheServerRefused,
                SignInIsRemembered: ReviewSessionStore.IsSupported,
                UsesTestServer: AppConfig.UsesTestServer);

        // ###########################################################################################
        // "Sign out" in the account window: the tab's own sign-out - which first asks about a table of its with changes not
        // saved, so the tab is brought forward first when it holds one (as quitting does, ConfirmLeavingMaintainerTableAsync):
        // the question is about what is on it.
        // ###########################################################################################
        private async Task SignOutFromAccountWindowAsync()
        {
            if (this.TabMaintainer.CommitCellsAndHasUnsavedTableEdits() && this.MaintainerTabItem != null && this.MainTabControl != null)
                this.MainTabControl.SelectedItem = this.MaintainerTabItem;

            await this.TabMaintainer.SignOutAsync();
        }

        // The person, with a check while signed in, and the tooltip saying who.
        private void ShowAccountIcon()
        {
            ReviewSession? session = this.TabMaintainer.SignedIn;

            this.AccountGlyph.Text = session is null ? SignedOutGlyph : SignedInGlyph;
            ToolTip.SetTip(this.AccountButton, session is null ? SignedOutTip : Main.SignedInTip(session));
        }

        // The icon as drawn - for tests.
        internal (string Glyph, string? Tip) AccountIconForTests =>
            (this.AccountGlyph.Text ?? string.Empty, ToolTip.GetTip(this.AccountButton) as string);
    }
}
