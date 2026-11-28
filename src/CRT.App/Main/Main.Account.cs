using System;
using System.Threading.Tasks;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE ACCOUNT WINDOW, OPENED BY THE "MAINTAINER" TAB'S "Sign in" (2026-11-25; AccountWindow): signing in, "I forgot my
    // password", "I have an invitation", or "My account". An account is a MAINTAINER'S, made by an invitation (owner decision,
    // 2026-11-28: "Entirely remove the new "Account" possibility ... This should of course continue to work as-is for the
    // "Maintainer" tab") - the account icon in the window's top right corner, which opened it for everybody from 2026-11-25,
    // went then, with "Create an account".
    //
    // *** THE SESSION IS THE MAINTAINER TAB'S *** (TabMaintainer.Session.cs), as it always was: this part hands the window the
    // tab's session and client, and hands the tab what the window did (AccountHost) - a sign-in, an account changed, a sign-out.
    // ###########################################################################################
    public partial class Main
    {
        // The window shown instead of as a dialog in a test, and the client its requests go through - never the live server.
        internal Func<AccountWindow, Task>? ShowAccountWindowOverrideForTests { get; set; }

        internal Func<ReviewApiClient>? AccountClientOverrideForTests { get; set; }

        private void WireAccountWindow() =>
            this.TabMaintainer.SignInWanted += () => _ = HandlerGuard.RunAsync(this.OpenAccountWindowAsync, _ => { });

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
    }
}
