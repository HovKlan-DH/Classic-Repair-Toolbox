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
    // *** EVERY MAINTAINER'S: "MY ACCOUNT" AND "SERVER VERSION". *** The rest is the administrator's
    // own, and stays so: each such entry carries the class AdminOnly in the markup, is TAKEN OUT of
    // the list for everybody else (ShowAdministratorEntries, from the server's isAdministrator) and
    // is marked with a padlock after its name ("No title/helper text on this icon"). The server
    // refuses everybody else at every one of those routes anyway; hiding them is so nobody is
    // offered a button that could only ever be refused.
    //
    // *** OUT OF THE LIST, NOT IsVisible = false. *** The list's VirtualizingStackPanel sets
    // IsVisible back to true on every container it realises, so hidden entries came back in a real
    // window while a tab that was never shown - every test of it - kept them hidden (seen in a
    // render, 2026-10-04; In_a_shown_window_a_maintainer_is_drawn_only_the_two_entries_that_are_theirs).
    // Out of the list they also cannot be reached with the arrow keys.
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
    // THE ADMINISTRATOR'S ENTRIES, and what each reads:
    //
    //   "Maintainers" (MaintainerPoolView) - here since 2026-10-04 (owner request: "'Send
    //   invitation' and 'Add as maintainer' gets moved to the 'Admin' tab ... as this is something
    //   only the admin should be able to do"), on the Systems screen from 2026-09-27. Its systems
    //   and accounts are read on choosing it.
    //
    //   "Order of systems" (SystemOrderView, 2026-10-04) reads BETA's drop-down list on choosing it -
    //   unless moves are waiting to be saved there.
    //
    //   "Unused files", whose list reads every workbook in the tree and takes seconds, is read on
    //   choosing it, on switching tree, and on its Refresh button - never again merely on coming
    //   back to Account.
    //
    //   "Rebuild checksum manifests" (2026-10-01) reads NOTHING on choosing it: a rebuild writes, so
    //   it waits for its own button to be pressed.
    //
    //   "Delete a system" (2026-10-03) reads every system when chosen - the list its Delete buttons
    //   sit in - and nothing is deleted before its confirmation (SystemDeletionView).
    //
    //   "API usage" (2026-10-04: "how about tracking the API end-points, to see if it is possible to
    //   retire any") reads which CRT versions called each route when chosen, and on its own Refresh
    //   and days box (ApiUsageView).
    //
    //   "Reset contribution data", THE LAST ENTRY (2026-10-04: a clean start at go-live), reads the
    //   counts when chosen; nothing is deleted before RESET is typed and the button pressed, and
    //   only while the server's switch is on (DataResetView).
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private bool thisAccountRead;

        // The administrator's entries in the markup's order - read off the list the first time, and
        // kept here while they are out of it.
        private List<ListBoxItem>? thisAdministratorEntries;

        private MyAccountView MyAccountDetail => this.FindControl<MyAccountView>("MyAccountPanel")!;

        // ###########################################################################################
        // Once, from the constructor: "My account" hands its changes and its "Sign out" to the tab,
        // and the administrator's entries start OUT of the list - they come in only when the server
        // says so, never in the moment between signing in and its first answer.
        // ###########################################################################################
        private void WireAccountScreen()
        {
            this.MyAccountDetail.AccountChanged = this.ApplyAccount;
            this.MyAccountDetail.SignOutRequested = this.SignOutAsync;

            this.ShowAdministratorEntries(false);
        }

        private async void OnAccountSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            this.ApplyModeVisibility();

            object? chosen = this.FindControl<ListBox>("AccountList")?.SelectedItem;

            // "My account" reads nothing: it shows the session the tab already holds.
            if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("ServerVersionItem")))
                await this.ReadServerVersionAsync();
            else if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("MaintainersItem")))
                await this.MaintainerPoolAdmin.LoadAsync();
            else if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("SystemOrderItem")))
                await this.SystemOrderAdmin.LoadAsync();
            else if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("UnusedFilesItem")))
                await this.UnusedFilesAdmin.LoadAsync();
            else if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("DeleteSystemItem")))
                await this.SystemDeletionAdmin.LoadAsync();
            else if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("ApiUsageItem")))
                await this.ApiUsageAdmin.LoadAsync();
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
        // The administrator's entries put in the list or taken out of it - from the server's word on
        // whether this account is one (SetAdministrator). Taken out while one of them is chosen - an
        // account that stops being an administrator - the screen goes back to "My account" FIRST, so
        // nothing of the administrator's stays on show and the list is never left with no entry.
        //
        // They all follow the two every maintainer has, so putting them back at the end, in the
        // markup's order, gives the markup's list again.
        // ###########################################################################################
        private void ShowAdministratorEntries(bool isAdministrator)
        {
            if (this.FindControl<ListBox>("AccountList") is not ListBox list)
                return;

            this.thisAdministratorEntries ??= list.Items.OfType<ListBoxItem>().Where(item => item.Classes.Contains("AdminOnly")).ToList();

            if (!isAdministrator && list.SelectedItem is ListBoxItem selected && this.thisAdministratorEntries.Contains(selected))
                list.SelectedItem = this.FindControl<ListBoxItem>("MyAccountItem");

            foreach (ListBoxItem item in this.thisAdministratorEntries)
            {
                bool inList = list.Items.Contains(item);

                if (isAdministrator && !inList)
                    list.Items.Add(item);
                else if (!isAdministrator && inList)
                    list.Items.Remove(item);
            }
        }

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

            // Only onto the session it was asked for - a sign-out or another sign-in meanwhile wins.
            if (read.IsOk && ReferenceEquals(this.thisSession, session))
                this.ApplyAccount(read.Value!);
        }

        // ###########################################################################################
        // "Logged in as:" and, under it, the name in bold and the address (owner request,
        // 2026-10-04). Named, because somebody with both a maintainer and an administrator account
        // needs to know which one they are acting as before they publish anything.
        // ###########################################################################################
        private void ShowSignedInAs()
        {
            if (this.thisSession is { } session && this.FindControl<TextBlock>("SignedInAsText") is TextBlock text)
                TabMaintainer.ShowCounts(text, MyAccountRules.LoggedInAs(session.DisplayName, session.Email));
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

            this.ShowServerVersionLines(ServerVersionDisplay.Lines(
                answer.IsOk ? answer.Value?.Version : null,
                answer.IsOk ? answer.Value?.ApiRevision : null,
                ClientVersionContract.ApiRevision));
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
        internal void UseClientForTests(ReviewApiClient client) => this.thisClient = client;

        // -----------------------------------------------------------------------------------
        // The administrator's entries
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // After a pool changed (Account > Maintainers): the queue (assigning somebody may have been
        // prompted by a submission on screen, and the queue is filtered by pools) and the systems
        // (they count maintainers, and the open system's Maintainer view lists them).
        // ###########################################################################################
        private async Task AfterPoolChangeAsync()
        {
            await this.RefreshQueueAsync(background: true);
            await this.RefreshSystemsAsync(background: true);
        }

        // ###########################################################################################
        // After a system was deleted: its submissions are gone from the queue, it may have been on
        // the BETA list, and it is gone from the systems - all three read again, quietly. The same
        // after the contribution data was reset (2026-10-04): every submission gone from the queue
        // and the BETA list, every system without its maintainers and history.
        // ###########################################################################################
        private async Task AfterSystemDeletedAsync()
        {
            await this.RefreshQueueAsync(background: true);
            await this.RefreshBetaAsync(background: true);
            await this.RefreshSystemsAsync(background: true);
        }

        private void ClearAccountScreen()
        {
            if (this.FindControl<ListBox>("AccountList") is ListBox list)
                list.SelectedItem = null;

            this.ShowServerVersionLines(null);
            this.MaintainerPoolAdmin.Clear();
            this.SystemOrderAdmin.Clear();
            this.UnusedFilesAdmin.Clear();
            this.RebuildManifestsAdmin.Clear();
            this.SystemDeletionAdmin.Clear();
            this.ApiUsageAdmin.Clear();
            this.DataResetAdmin.Clear();
        }
    }
}
