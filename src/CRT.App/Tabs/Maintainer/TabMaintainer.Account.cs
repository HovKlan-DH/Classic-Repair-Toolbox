using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE "ACCOUNT" SCREEN (owner request, 2026-10-04: rename "Administrator activities" to
    // "Account", "make this tab available to all maintainers", and move the account functionality
    // from the bottom-left corner to a new entry "My account" on it). The entries on the left, the
    // chosen one on the right. It was "Admin" (2026-09-27), offered only to an administrator.
    //
    // *** EVERY MAINTAINER'S: "MY ACCOUNT" AND "SERVER VERSION". *** The administrator's own entries
    // are a screen of their own since 2026-11-07, "MANAGEMENT" (owner request: 'a new tab named
    // "Management" (shown before "Account") - it should only be visible by me ... Then move all the
    // menu items from "Account" into there ... no need to have a "lock" icon on each individual menu
    // item'): its tab - padlocked - is shown only while the server says this account is an
    // administrator (ShowManagementScreen, from the server's isAdministrator). The server refuses
    // everybody else at every one of those routes anyway; hiding them is so nobody is offered a
    // button that could only ever be refused. Until then each entry stood under Account, padlocked,
    // and was taken out of the list for anybody else.
    //
    // *** AN ENTRY WITH SOMETHING WAITING SAYS SO (owner request, 2026-11-07: "when I then navigate
    // to the 'Account' tab, then I cannot see where to go next - nothing indicates which menu I
    // should choose"). *** "Oscilloscope models" carries the count of models waiting, as the screen's
    // tab does; and the screen opened with nothing chosen opens on it while any wait.
    //
    // "MY ACCOUNT" (MyAccountView), the first entry and what the screen opens on: the name, the
    // email address, the password - the "Your account" window of 2026-10-03 - and "Sign out", both
    // of which were buttons under the lists until 2026-10-04. Every account the server hands back
    // goes through ApplyAccount: the session takes the new name and address, the remembered sign-in
    // is written again, and SignedInChanged reaches the Feedback tab and the Submit dialog, which
    // use the account's address while signed in (Main.ShareMaintainerSignIn).
    //
    // *** THE REMEMBERED NAME AND ADDRESS ARE READ AGAIN ONCE A LAUNCH. *** The stored session
    // keeps them as they were at sign-in, and they can change elsewhere - on another computer, or
    // by the server. ReadAccountQuietlyAsync asks GET /api/accounts/me once, quietly (nobody
    // pressed anything), and only on the paths that already talk to the server - never from
    // RestoreSignInQuietly, which by design asks nothing. A refusal changes nothing here: a 401 is
    // the queue's own check's to act on, which ends the session properly.
    //
    // "SERVER VERSION" (owner requests, 2026-10-04: "I would like to see the server version listed,
    // so it is clear to me what has been deployed", then as an entry of its own, and the same day
    // every maintainer's): since 2026-11-24 a panel each for the server deployed, its automatic
    // updates and mail it could not send, and the server's own changelog under them
    // (ServerVersionView, worded by ServerVersionDisplay - owner requests: "make the "Account" page a
    // little prettier ... so it is more easy to overview", and 'have a "Changelog" like there is in
    // "About"'). GET /api/health, asked with no session, and GET /api/review/server-status with it,
    // every time the entry is shown - chosen, or Account shown again with it chosen - so a deploy
    // shows the next time; quietly, with no "please wait", since nobody is waiting for them.
    //
    // THE ADMINISTRATOR'S ENTRIES - the Management screen's - and what each reads:
    //
    //   "Maintainers" (MaintainerPoolView) - here since 2026-10-04 (owner request: "'Send
    //   invitation' and 'Add as maintainer' gets moved to the 'Admin' tab ... as this is something
    //   only the admin should be able to do"), on the Boards screen from 2026-09-27. Its boards
    //   and accounts are read on choosing it.
    //
    //   "Order of boards" (BoardOrderView, 2026-10-04) reads BETA's drop-down list on choosing it -
    //   unless moves are waiting to be saved there.
    //
    //   "Unused files", whose list reads every workbook in the tree and takes seconds, is read on
    //   choosing it, on switching tree, and on its Refresh button - never again merely on coming
    //   back to Account.
    //
    //   "Rebuild checksum manifests" (2026-10-01) reads NOTHING on choosing it: a rebuild writes, so
    //   it waits for its own button to be pressed.
    //
    //   "Delete a board" (2026-10-03) reads every board when chosen - the list its Delete buttons
    //   sit in - and nothing is deleted before its confirmation (BoardDeletionView).
    //
    //   "API usage" (2026-10-04: "how about tracking the API end-points, to see if it is possible to
    //   retire any") reads which CRT versions called each route when chosen, and on its own Refresh
    //   and days box (ApiUsageView).
    //
    //   "Oscilloscope models" (CRT 4.0 step 5, 2026-10-29) reads the models sent in from CRT and waiting,
    //   when chosen, when Account is shown again with it chosen, on its own Refresh, and when the
    //   count the queue answers differs from what it shows; each is published or turned down there
    //   (OscilloscopeModelsView).
    //
    //   "Reset contribution data", THE LAST ENTRY (2026-10-04: a clean start at go-live), reads the
    //   counts when chosen; nothing is deleted before RESET is typed and the button pressed, and
    //   only while the server's switch is on (DataResetView).
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private bool thisAccountRead;

        private MyAccountView MyAccountDetail => this.FindControl<MyAccountView>("MyAccountPanel")!;

        // ###########################################################################################
        // Once, from the constructor: "My account" hands its changes and its "Sign out" to the tab,
        // and the Management screen starts hidden - it is shown only when the server says so, never in
        // the moment between signing in and its first answer.
        // ###########################################################################################
        private void WireAccountScreen()
        {
            this.MyAccountDetail.AccountChanged = this.ApplyAccount;
            this.MyAccountDetail.SignOutRequested = this.SignOutAsync;
            this.MyAccountDetail.SessionEnded = this.EndSessionTheServerRefused;

            // What the models screen reads is what the badges count (TabMaintainer.Modes.cs).
            this.OscilloscopeModelsAdmin.ModelsWaitingRead = this.UseOscilloscopeModelsWaiting;

            this.ShowManagementScreen(false);
        }

        private async void OnAccountSelectionChanged(object? sender, SelectionChangedEventArgs e) =>
            await HandlerGuard.RunAsync(this.AccountSelectionChangedAsync, this.ShowHandlerFailure);

        private async Task AccountSelectionChangedAsync()
        {
            this.ApplyModeVisibility();

            object? chosen = this.FindControl<ListBox>("AccountList")?.SelectedItem;

            // "My account" reads nothing: it shows the session the tab already holds.
            if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("ServerVersionItem")))
                await this.ReadServerVersionAsync();
        }

        private async void OnManagementSelectionChanged(object? sender, SelectionChangedEventArgs e) =>
            await HandlerGuard.RunAsync(this.ManagementSelectionChangedAsync, this.ShowHandlerFailure);

        // Each of the administrator's entries reads what it shows when chosen - see the header.
        private async Task ManagementSelectionChangedAsync()
        {
            this.ApplyModeVisibility();

            object? chosen = this.FindControl<ListBox>("ManagementList")?.SelectedItem;

            if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("MaintainersItem")))
                await this.MaintainerPoolAdmin.LoadAsync();
            else if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("BoardOrderItem")))
                await this.BoardOrderAdmin.LoadAsync();
            else if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("UnusedFilesItem")))
                await this.UnusedFilesAdmin.LoadAsync();
            else if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("DeleteBoardItem")))
                await this.BoardDeletionAdmin.LoadAsync();
            else if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("ApiUsageItem")))
                await this.ApiUsageAdmin.LoadAsync();
            else if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("OscilloscopeModelsItem")))
                await this.OscilloscopeModelsAdmin.LoadAsync();
            else if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("ResetDataItem")))
                await this.DataResetAdmin.LoadAsync();
        }

        // On showing Account: the first time, "My account" is chosen; after that, whatever was chosen
        // stays as it was - see the header. "Server version" is asked again, so a deploy since shows.
        private async Task EnterAccountAsync()
        {
            if (this.FindControl<ListBox>("AccountList") is not ListBox list)
                return;

            if (list.SelectedItem is null)
                list.SelectedItem = this.FindControl<ListBoxItem>("MyAccountItem");
            else if (ReferenceEquals(list.SelectedItem, this.FindControl<ListBoxItem>("ServerVersionItem")))
                await this.ReadServerVersionAsync();
        }

        // ###########################################################################################
        // On showing Management: with nothing chosen yet, the entry something waits on - "Oscilloscope models"
        // while models wait (owner request, 2026-11-07: "nothing indicates which menu I should choose") - else
        // "Maintainers", the first; after that, whatever was chosen stays. "Oscilloscope models" chosen is read
        // again (code review, 2026-11-02): its badge on the screen's button is why it was pressed, and the panel
        // must not say what it read before that.
        // ###########################################################################################
        private async Task EnterManagementAsync()
        {
            if (this.FindControl<ListBox>("ManagementList") is not ListBox list)
                return;

            if (list.SelectedItem is null)
                list.SelectedItem = this.FindControl<ListBoxItem>(this.thisOscilloscopeModelsWaiting > 0 ? "OscilloscopeModelsItem" : "MaintainersItem");
            else if (ReferenceEquals(list.SelectedItem, this.FindControl<ListBoxItem>("OscilloscopeModelsItem")))
                await this.OscilloscopeModelsAdmin.LoadAsync();
        }

        // ###########################################################################################
        // The models on show follow the count the queue just answered, quietly, and at every check while
        // on screen (code review, 2026-11-02; OscilloscopeModelsView.FollowWaitingCountAsync) - only that
        // entry on the Management screen. Anywhere else nothing shows them, and showing them reads them.
        // ###########################################################################################
        private Task FollowOscilloscopeModelsAsync(bool onScreen) =>
            this.thisIsAdministrator &&
            this.thisMode == MaintainerMode.Management &&
            ReferenceEquals(this.FindControl<ListBox>("ManagementList")?.SelectedItem, this.FindControl<ListBoxItem>("OscilloscopeModelsItem"))
                ? this.OscilloscopeModelsAdmin.FollowWaitingCountAsync(this.thisOscilloscopeModelsWaiting, onScreen)
                : Task.CompletedTask;

        // ###########################################################################################
        // The Management screen's tab shown or hidden - from the server's word on whether this account is an
        // administrator (SetAdministrator). Hidden while it is on show - an account that stops being an administrator -
        // the tab goes to "Account", "My account" FIRST, so nothing of the administrator's stays on show.
        // ###########################################################################################
        private void ShowManagementScreen(bool isAdministrator)
        {
            if (this.FindControl<Button>("ManagementModeButton") is Button button)
                button.IsVisible = isAdministrator;

            if (!isAdministrator && this.thisMode == MaintainerMode.Management)
            {
                this.thisMode = MaintainerMode.Account;

                if (this.FindControl<ListBox>("AccountList") is ListBox list)
                    list.SelectedItem = this.FindControl<ListBoxItem>("MyAccountItem");
            }
        }

        // Whether the Management screen's tab is offered - for tests.
        internal bool ManagementScreenOfferedForTests => this.FindControl<Button>("ManagementModeButton")?.IsVisible == true;

        // The Management screen's entries, in order - for tests.
        internal IReadOnlyList<string> ManagementEntriesForTests =>
            this.FindControl<ListBox>("ManagementList")!.Items.OfType<ListBoxItem>()
                .Select(item => item.Name!)
                .ToList();

        // The entries this account is offered, in order - for tests.
        internal IReadOnlyList<string> AccountEntriesShownForTests =>
            this.FindControl<ListBox>("AccountList")!.Items.OfType<ListBoxItem>()
                .Select(item => item.Name!)
                .ToList();

        // -----------------------------------------------------------------------------------
        // My account
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // The account as the server now holds it. Nothing happens when it changes nothing - the
        // usual case at launch - so the Feedback tab is not touched for no reason.
        // ###########################################################################################
        internal void ApplyAccount(AccountAnswer account)
        {
            if (this.thisSession is not { } session)
                return;

            ReviewSession updated = session.WithAccount(account);
            bool changed = !Equals(updated, session);

            if (changed)
            {
                this.UseSession(updated);

                // The stored copy too, or the next launch shows the old name and address until it has
                // read them again. Declines on its own where it cannot store safely.
                ReviewSessionStore.Remember(updated);

                this.ShowSignedInAs();
                this.MyAccountDetail.UseAccount(updated);
            }

            // Asked whether or not anything else changed: whether the account maintains anything moves without its name or
            // address. AFTER the session takes the new name (code review, 2026-11-25), as the "not a maintainer" line names who
            // is signed in - built before, it said the name before; and a new name redraws that line when it stays.
            this.ApplyStanding(account);

            if (changed && this.thisMaintainsNothing)
                this.ShowSignInPanelState(maintainsNothing: true);
        }

        // ###########################################################################################
        // *** AN ACCOUNT THAT MAINTAINS NO BOARD IS SAID SO (2026-11-25, accounts for everybody). *** Until then every account
        // was a maintainer's - only an invitation made one. Now anybody's account can sign in, and with "Enable Maintainer tab"
        // checked it would meet lists the server refuses it. So, from the server's own answer (GET /api/accounts/me: no board's
        // pool and no administrator), the sign-in panel says so instead of the queue - and the queue comes back, its lists read,
        // the moment the server says otherwise (an administrator made it a maintainer, read again at the next launch or sign-in).
        // ###########################################################################################
        private bool thisMaintainsNothing;

        private void ApplyStanding(AccountAnswer account)
        {
            bool maintainsNothing = !account.IsAdministrator && account.MaintainerOf.Count == 0;

            if (maintainsNothing == this.thisMaintainsNothing)
                return;

            this.thisMaintainsNothing = maintainsNothing;

            if (this.FindControl<Control>("SignInPanel") is not Control signIn || this.FindControl<Grid>("QueuePanel") is not Grid queue)
                return;

            if (maintainsNothing)
            {
                this.StopQueueChecks();
                queue.IsVisible = false;
                signIn.IsVisible = true;
                this.ShowSignInPanelState(maintainsNothing: true);
                return;
            }

            this.ShowSignInPanelState(maintainsNothing: false);
            this.ShowQueuePanel();

            // *** ONLY ON SCREEN *** (code review, 2026-11-25): the lists are read under the "please wait" of the window the tab
            // stands in - off screen that is CRT's window found some other way, dimmed behind the account window for a tab not
            // even shown. Off screen they wait for the tab to be shown, as a sign-in's do (UseSignInAsync, thisListsWanted).
            if (TopLevel.GetTopLevel(this) is not null)
                _ = HandlerGuard.RunAsync(this.ReadListsAsync, this.ShowHandlerFailure);
            else
                this.thisListsWanted = true;
        }

        // ###########################################################################################
        // *** A 401 MET OUTSIDE THE TAB'S OWN CHECKS *** - "My account", here or in CRT's account window (code review,
        // 2026-11-25). The queue's 401 path, for a session the server has ended (signed out elsewhere, a password reset, the
        // account locked): with a table holding unsaved changes the tab keeps them and says so (KeepTablesAfterSessionEnded),
        // otherwise the sign-in screen, the remembered sign-in forgotten (ShowSignInPanel) - and SignedInChanged tells the
        // rest of CRT, the account icon included.
        // ###########################################################################################
        internal void EndSessionTheServerRefused()
        {
            if (this.thisSession is null)
                return;

            if (this.WouldLoseAnyTableChanges)
            {
                this.KeepTablesAfterSessionEnded(expired: false);
                return;
            }

            this.ShowSignInPanel();
            this.ShowSignInMessage(TabMaintainer.SessionEndedMessage);
        }

        internal const string SessionEndedMessage = "You are no longer signed in - signed out on another computer, or the sign-in ended. Sign in again.";

        // Whether the sign-in panel says the account maintains no board - for tests.
        internal bool SaysAccountMaintainsNothingForTests => this.FindControl<TextBlock>("NotMaintainerText")?.IsVisible == true;

        // ###########################################################################################
        // *** READ ONCE IT HAS BEEN READ, NOT ONCE IT HAS BEEN ASKED *** (code review, 2026-11-25). Marked read before the
        // request, a /me that failed at sign-in was never asked again, and an account that maintains nothing stayed on the
        // queue panel, refused every minute. A failure leaves it to be asked again - by the queue's 403 too (RefreshQueueAsync),
        // the one answer that says the account's standing may not be what the tab thinks.
        // ###########################################################################################
        private async Task ReadAccountQuietlyAsync()
        {
            if (this.thisAccountRead || this.thisClient is not { } client || this.thisSession is not { } session)
                return;

            ReviewApiResult<AccountAnswer> read = await client.GetAccountAsync(session);

            // Only onto the session it was asked for - a sign-out or another sign-in meanwhile wins
            // (its expiry slid meanwhile is still that session).
            if (!read.IsOk || !this.IsSessionInUse(session))
                return;

            this.thisAccountRead = true;
            this.ApplyAccount(read.Value!);
        }

        // ###########################################################################################
        // "Signed in as:" and, under it, the name in bold and the address (owner request,
        // 2026-10-04). Named, because somebody with both a maintainer and an administrator account
        // needs to know which one they are acting as before they publish anything.
        // ###########################################################################################
        private void ShowSignedInAs()
        {
            if (this.thisSession is { } session && this.FindControl<TextBlock>("SignedInAsText") is TextBlock text)
                TabMaintainer.ShowCounts(text, MyAccountRules.SignedInAs(session.DisplayName, session.Email));
        }

        // -----------------------------------------------------------------------------------
        // Server version
        // -----------------------------------------------------------------------------------

        private ServerVersionView ServerVersionPage => this.FindControl<ServerVersionView>("ServerVersionView")!;

        // ###########################################################################################
        // What the server reports - the panels on screen stay until the next answer replaces them, and the
        // changelog is built again only when its text changed (ServerVersionView). "Asking" is said only on a
        // page with nothing on it yet.
        // ###########################################################################################
        private async Task ReadServerVersionAsync()
        {
            if (this.thisClient is not ReviewApiClient client)
                return;

            if (!this.ServerVersionPage.HasPanels)
                this.ServerVersionPage.ShowPanels(ServerVersionDisplay.Panels(null, null, ClientVersionContract.ApiRevision, asking: true));

            ReviewApiResult<HealthStatus> answer = await client.GetServerVersionAsync();

            // The mails it could not send (2026-10-13), how it keeps itself up to date (2026-11-09) and its
            // changelog (2026-11-24) - asked with the session, from a server that answered at all; an older
            // server's 404 simply says nothing about any of them.
            ServerStatusAnswer? serverStatus = null;

            if (answer.IsOk && this.thisSession is ReviewSession session)
            {
                ReviewApiResult<ServerStatusAnswer> status = await client.GetServerStatusAsync(session);
                serverStatus = status.IsOk ? status.Value : null;
            }

            this.ServerVersionPage.ShowPanels(ServerVersionDisplay.Panels(
                answer.IsOk ? answer.Value?.Version : null,
                answer.IsOk ? answer.Value?.ApiRevision : null,
                ClientVersionContract.ApiRevision,
                unsentMails: serverStatus?.UnsentMails,
                updates: serverStatus?.Updates));

            // The changelog only from a server that said how it stands; one that answered with none is older.
            this.ServerVersionPage.ShowChangelog(serverStatus?.Changelog, answered: serverStatus is not null);
        }

        // The page's cards as lines, one per line - null while the page is not shown. For tests.
        internal string? ServerVersionLineForTests =>
            this.FindControl<Control>("ServerVersionView") is { IsVisible: true }
                ? string.Join("\n", this.ServerVersionPage.PanelLinesForTests)
                : null;

        // The page itself, for a test reading its changelog.
        internal ServerVersionView ServerVersionPageForTests => this.ServerVersionPage;

        // A client for a test, as signing in would give the tab - never the real server.
        internal void UseClientForTests(ReviewApiClient client) => this.UseClient(client);

        // -----------------------------------------------------------------------------------
        // The administrator's entries
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // After a pool changed (Management > Maintainers): the queue (assigning somebody may have been
        // prompted by a submission on screen, and the queue is filtered by pools) and the boards
        // (they count maintainers, and the open board's Maintainer view lists them).
        // ###########################################################################################
        private async Task AfterPoolChangeAsync()
        {
            await this.RefreshQueueAsync(background: true);
            await this.RefreshBoardsAsync(background: true);
        }

        // ###########################################################################################
        // After a board was deleted: its submissions are gone from the queue, it may have been on
        // the BETA list, and it is gone from the boards - all three read again, quietly. The same
        // after the contribution data was reset (2026-10-04): every submission gone from the queue
        // and the BETA list, every board without its maintainers and history.
        // ###########################################################################################
        private async Task AfterBoardDeletedAsync()
        {
            await this.RefreshQueueAsync(background: true);
            await this.RefreshBetaAsync(background: true);
            await this.RefreshBoardsAsync(background: true);
        }

        private void ClearAccountScreen()
        {
            if (this.FindControl<ListBox>("AccountList") is ListBox list)
                list.SelectedItem = null;

            if (this.FindControl<ListBox>("ManagementList") is ListBox management)
                management.SelectedItem = null;

            this.ServerVersionPage.Clear();
            this.MaintainerPoolAdmin.Clear();
            this.BoardOrderAdmin.Clear();
            this.UnusedFilesAdmin.Clear();
            this.RebuildManifestsAdmin.Clear();
            this.BoardDeletionAdmin.Clear();
            this.ApiUsageAdmin.Clear();
            this.OscilloscopeModelsAdmin.Clear();
            this.DataResetAdmin.Clear();
        }
    }
}
