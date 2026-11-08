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
    // every maintainer's): the server's version and API revision and this CRT's revision, three
    // lines ServerVersionDisplay words. GET /api/health, asked with no session every time the entry
    // is shown - chosen, or Account shown again with it chosen - so a deploy shows the next time;
    // quietly, with no "please wait", since nobody is waiting for those lines.
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

            if (Equals(updated, session))
                return;

            this.UseSession(updated);

            // The stored copy too, or the next launch shows the old name and address until it has
            // read them again. Declines on its own where it cannot store safely.
            ReviewSessionStore.Remember(updated);

            this.ShowSignedInAs();
            this.MyAccountDetail.UseAccount(updated);
        }

        private async Task ReadAccountQuietlyAsync()
        {
            if (this.thisAccountRead || this.thisClient is not { } client || this.thisSession is not { } session)
                return;

            this.thisAccountRead = true;

            ReviewApiResult<AccountAnswer> read = await client.GetAccountAsync(session);

            // Only onto the session it was asked for - a sign-out or another sign-in meanwhile wins
            // (its expiry slid meanwhile is still that session).
            if (read.IsOk && this.IsSessionInUse(session))
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

        // What the server reports - the lines on screen stay until the next answer replaces them.
        private async Task ReadServerVersionAsync()
        {
            if (this.thisClient is not ReviewApiClient client)
                return;

            if (this.FindControl<StackPanel>("ServerVersionLines") is { Children.Count: 0 })
                this.ShowServerVersionLines(ServerVersionDisplay.Lines(null, null, ClientVersionContract.ApiRevision, asking: true));

            ReviewApiResult<HealthStatus> answer = await client.GetServerVersionAsync();

            // The mails it could not send (2026-10-13) - asked with the session, from a server that
            // answered at all; an older server's 404 simply says nothing about mail.
            int? unsentMails = null;

            if (answer.IsOk && this.thisSession is ReviewSession session)
            {
                ReviewApiResult<ServerStatusAnswer> status = await client.GetServerStatusAsync(session);
                unsentMails = status.IsOk ? status.Value?.UnsentMails : null;
            }

            this.ShowServerVersionLines(ServerVersionDisplay.Lines(
                answer.IsOk ? answer.Value?.Version : null,
                answer.IsOk ? answer.Value?.ApiRevision : null,
                ClientVersionContract.ApiRevision,
                unsentMails: unsentMails));
        }

        // One TextBlock a line, its value bold; none clears the panel.
        private void ShowServerVersionLines(IReadOnlyList<IReadOnlyList<ReviewNoteRun>>? lines)
        {
            if (this.FindControl<StackPanel>("ServerVersionLines") is not StackPanel panel)
                return;

            panel.Children.Clear();

            foreach (IReadOnlyList<ReviewNoteRun> line in lines ?? [])
            {
                var block = new TextBlock { TextWrapping = Avalonia.Media.TextWrapping.Wrap };
                TabMaintainer.ShowCounts(block, line);
                panel.Children.Add(block);
            }
        }

        // The lines as shown, one per line - null while the panel is not shown. For tests.
        internal string? ServerVersionLineForTests =>
            this.FindControl<Control>("ServerVersionView") is { IsVisible: true } &&
            this.FindControl<StackPanel>("ServerVersionLines") is StackPanel panel
                ? string.Join("\n", panel.Children.OfType<TextBlock>().Select(TabMaintainer.TextOf))
                : null;

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

            this.ShowServerVersionLines(null);
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
