using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // SETTING A SYSTEM'S MAINTAINERS, ON THE SYSTEMS SCREEN (owner request, 2026-09-27: "As admin I
    // should be allowed to set a system maintainer in the 'Systems' list - so that should be moved
    // from 'Admin' section. I should be able to either select an existing maintainer or invite a new
    // maintainer via email.").
    //
    // For an ADMINISTRATOR the maintainers section grows: a Remove button on each maintainer, the
    // invitations nobody has accepted yet with a Withdraw button each, and below them the two ways
    // to add somebody - an account chosen from the list, or an address invited by email (the server
    // mails it a code; see CRT.Server's MaintainerInvitationFlows). Everybody else sees the section
    // as before, read-only.
    //
    // *** NO AUTHORITY LIVES HERE. *** The controls are shown to an administrator as a courtesy; the
    // server refuses anybody else regardless, and its sentence lands in the message line.
    //
    // After every change the system is READ AGAIN from the server rather than patched locally, so
    // the screen is always what the server holds - the rule the old "Set maintainers" window kept.
    // ###########################################################################################
    public partial class SystemView
    {
        private bool thisIsAdministrator;

        // The detail on screen, so the section can be drawn again when the administrator flag changes.
        private SystemDetailAnswer? thisDetail;

        // Every account, for the "choose somebody" list - read when an administrator opens a system.
        private readonly List<ReviewAccountRow> thisAccounts = [];

        // The accounts the list offers, in its order.
        private readonly List<ReviewAccountRow> thisAccountChoices = [];

        // Told after a pool changed: the queue and the systems list are read again (the queue is
        // filtered by pools, and the list counts maintainers).
        public Func<Task>? AfterPoolChange { get; set; }

        // The server's word on whether this account is an administrator (TabMaintainer).
        public void SetAdministrator(bool isAdministrator)
        {
            this.thisIsAdministrator = isAdministrator;

            if (!isAdministrator)
            {
                this.thisAccounts.Clear();
                this.ShowPoolMessage(null, isError: false);
            }

            this.ShowMaintainers(this.thisDetail);
        }

        // ###########################################################################################
        // The maintainers section. The heading and lines for everybody; the buttons, the invitations
        // and the two ways to add somebody for an administrator.
        // ###########################################################################################
        private void ShowMaintainers(SystemDetailAnswer? detail)
        {
            var section = this.FindControl<StackPanel>("MaintainersSection");
            var admin = this.FindControl<StackPanel>("MaintainerAdminPanel");

            if (section is null || admin is null)
                return;

            section.Children.Clear();
            admin.IsVisible = detail is not null && this.thisIsAdministrator;

            if (detail is null)
                return;

            section.Children.Add(SystemView.Heading(SystemsDisplay.MaintainersHeading(detail.Maintainers.Count), detail.Maintainers.Count == 0));

            foreach (PoolMaintainerEntry maintainer in detail.Maintainers)
            {
                TextBlock line = SystemView.Line(SystemsDisplay.MaintainerLine(maintainer));

                if (!this.thisIsAdministrator)
                {
                    section.Children.Add(line);
                    continue;
                }

                string systemId = detail.System.SystemId;
                long accountId = maintainer.AccountId;
                string name = maintainer.DisplayName;

                section.Children.Add(SystemView.WithButton(
                    line,
                    "Remove",
                    "RemoveMaintainer",
                    async () => await this.ChangePoolAsync(
                        new PoolAction(PoolActionKind.Remove, systemId, AccountId: accountId, Who: name),
                        $"{name} no longer maintains {systemId}. It takes effect on their next request.")));
            }

            if (this.thisIsAdministrator)
            {
                foreach (MaintainerInvitationEntry invitation in detail.Invitations ?? [])
                {
                    long invitationId = invitation.Id;
                string invited = invitation.Email;

                    section.Children.Add(SystemView.WithButton(
                        SystemView.TwoLines(SystemsDisplay.InvitationLine(invitation), SystemsDisplay.InvitationFooter(invitation), note: null),
                        "Withdraw",
                        "WithdrawInvitation",
                        async () => await this.ChangePoolAsync(new PoolAction(PoolActionKind.Withdraw, detail.System.SystemId, InvitationId: invitationId, Who: invited), null)));
                }
            }

            this.ShowAccountChoices(detail);
        }

        // The accounts that can be offered: everybody not already maintaining this system. The line
        // says why one cannot be granted (MaintainerAssignmentDisplay), as the old window did.
        private void ShowAccountChoices(SystemDetailAnswer detail)
        {
            if (this.FindControl<ComboBox>("AddAccountCombo") is not ComboBox combo)
                return;

            HashSet<long> already = detail.Maintainers.Select(maintainer => maintainer.AccountId).ToHashSet();

            this.thisAccountChoices.Clear();
            this.thisAccountChoices.AddRange(this.thisAccounts.Where(account => !already.Contains(account.Id)));

            combo.ItemsSource = this.thisAccountChoices.Select(MaintainerAssignmentDisplay.AccountChoice).ToList();
            combo.SelectedIndex = -1;
        }

        // Reads every account, for the list - an administrator's system only.
        private async Task ReadAccountsAsync()
        {
            if (!this.thisIsAdministrator || this.thisClient is null || this.thisSession is null)
                return;

            ReviewApiResult<ReviewAccountsResponse> accounts = await this.thisClient.GetAccountsAsync(this.thisSession);

            if (!accounts.IsOk)
            {
                this.ShowPoolMessage(accounts.Message, isError: true);
                return;
            }

            this.thisAccounts.Clear();
            this.thisAccounts.AddRange(accounts.Value!.Accounts);

            if (this.thisDetail is not null)
                this.ShowAccountChoices(this.thisDetail);
        }

        private async void OnAddAccountClick(object? sender, RoutedEventArgs e) => await this.AddChosenAccountAsync();

        internal async Task AddChosenAccountAsync()
        {
            if (this.thisDetail is null || this.FindControl<ComboBox>("AddAccountCombo") is not ComboBox combo)
                return;

            int index = combo.SelectedIndex;

            if (index < 0 || index >= this.thisAccountChoices.Count)
            {
                this.ShowPoolMessage("Choose an account first.", isError: true);
                return;
            }

            ReviewAccountRow account = this.thisAccountChoices[index];

            // Said here, before the round trip, in the words the list already shows. The server
            // refuses it regardless.
            if (MaintainerAssignmentDisplay.WhyNotGrantable(account) is string why)
            {
                this.ShowPoolMessage($"{account.Email} cannot be a maintainer: {why}.", isError: true);
                return;
            }

            await this.ChangePoolAsync(
                new PoolAction(PoolActionKind.Add, this.thisDetail.System.SystemId, AccountId: account.Id, Who: account.DisplayName),
                $"{account.DisplayName} now maintains {this.thisDetail.System.SystemId}.");
        }

        private async void OnInviteClick(object? sender, RoutedEventArgs e) => await this.InviteTypedAddressAsync();

        internal async Task InviteTypedAddressAsync()
        {
            if (this.thisDetail is null || this.FindControl<TextBox>("InviteEmailBox") is not TextBox box)
                return;

            string email = box.Text?.Trim() ?? string.Empty;

            if (email.Length == 0)
            {
                this.ShowPoolMessage("Type the email address of the person to invite.", isError: true);
                return;
            }

            // An address that already has an account is chosen from the list - the server says so
            // too, but this costs no round trip and names the way that works.
            if (this.thisAccounts.FirstOrDefault(account => string.Equals(account.Email, email, StringComparison.OrdinalIgnoreCase)) is ReviewAccountRow existing)
            {
                this.ShowPoolMessage($"{existing.Email} already has an account - choose it in the list above instead.", isError: true);
                return;
            }

            if (await this.ChangePoolAsync(new PoolAction(PoolActionKind.Invite, this.thisDetail.System.SystemId, Email: email, Who: email), null))
                box.Text = string.Empty;
        }

        // ###########################################################################################
        // Sends one change, says what happened, and reads the system (and the accounts) again. The
        // success sentence is the server's when it sends one (an invitation's), else `done`.
        //
        // *** UNDER THE OVERLAY, AND CHECKED AFTER A TIMEOUT (2026-09-28). *** No answer in two
        // minutes does not mean nothing changed: the system is read again and the pool itself says
        // whether the change landed (PoolAction.HasLanded).
        // ###########################################################################################
        private async Task<bool> ChangePoolAsync(PoolAction action, string? done)
        {
            bool changed = false;

            await BusyOverlay.HoldAsync(this, action.Waiting, async () =>
            {
                ReviewApiResult<string> result = await ServerWait.CallAsync(this, action.Waiting, token => this.SendAsync(action, token));

                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    bool? landed = await this.HasLandedAsync(action);

                    this.ShowPoolMessage(
                        MaintainerWaitWording.MaintainersAfterTimeout(landed, action.DoneClause, action.NotDoneClause),
                        isError: landed != true);

                    changed = landed == true;
                }
                else if (!result.IsOk)
                {
                    this.ShowPoolMessage(result.Message, isError: true);
                    return;
                }
                else
                {
                    this.ShowPoolMessage(done ?? result.Value ?? "Done.", isError: false);
                    changed = true;
                }

                await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingSystem(action.SystemId), async () =>
                {
                    if (this.ShownSystem is SystemOverviewEntry shown)
                        await this.ShowSystemAsync(shown, keepShown: true);

                    if (this.AfterPoolChange is not null)
                        await this.AfterPoolChange();
                });
            });

            return changed;
        }

        // After a timeout: the system read again, and whether the change is in it. Null when it
        // could not be read.
        private async Task<bool?> HasLandedAsync(PoolAction action)
        {
            if (this.PoolDetailOverrideForTests is not null)
                return action.HasLanded(await this.PoolDetailOverrideForTests(action.SystemId));

            if (this.thisClient is not ReviewApiClient client || this.thisSession is not ReviewSession session)
                return null;

            ReviewApiResult<SystemDetailAnswer> detail = await ServerWait.CallAsync(
                this, WaitWording.Checking, token => client.GetSystemDetailAsync(session, action.SystemId, token));

            return detail.IsOk ? action.HasLanded(detail.Value!) : null;
        }

        private Task<ReviewApiResult<string>> SendAsync(PoolAction action, CancellationToken token)
        {
            if (this.PoolActionOverrideForTests is not null)
                return this.PoolActionOverrideForTests(action);

            if (this.thisClient is null || this.thisSession is null)
                return Task.FromResult(ReviewApiResult<string>.Failed(ReviewApiFailure.NotSignedIn, "Sign in first."));

            return action.Kind switch
            {
                PoolActionKind.Add => this.thisClient.AddMaintainerAsync(this.thisSession, action.SystemId, action.AccountId, token),
                PoolActionKind.Remove => this.thisClient.RemoveMaintainerAsync(this.thisSession, action.SystemId, action.AccountId, token),
                PoolActionKind.Invite => this.thisClient.InviteMaintainerAsync(this.thisSession, action.SystemId, action.Email, token),
                _ => this.thisClient.WithdrawInvitationAsync(this.thisSession, action.InvitationId, token)
            };
        }

        // A line with a small button at its right.
        private static Grid WithButton(Control content, string label, string buttonClass, Func<Task> onClick)
        {
            var row = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };

            var button = new Button
            {
                Content = label,
                Margin = new Avalonia.Thickness(8, 0, 0, 0),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                Classes = { buttonClass }
            };

            button.Click += async (_, _) => await onClick();

            Grid.SetColumn(button, 1);
            row.Children.Add(content);
            row.Children.Add(button);

            return row;
        }

        private void ShowPoolMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("PoolMessageText"), message, isError);

        // -----------------------------------------------------------------------------------
        // For tests - no server.
        // -----------------------------------------------------------------------------------

        // Answers every change instead of the server, and sees what was sent.
        internal Func<PoolAction, Task<ReviewApiResult<string>>>? PoolActionOverrideForTests { get; set; }

        // Answers the look at the system after a timeout, instead of the server.
        internal Func<string, Task<SystemDetailAnswer>>? PoolDetailOverrideForTests { get; set; }

        internal void UseAccountsForTests(IReadOnlyList<ReviewAccountRow> accounts)
        {
            this.thisAccounts.Clear();
            this.thisAccounts.AddRange(accounts);

            if (this.thisDetail is not null)
                this.ShowAccountChoices(this.thisDetail);
        }
    }
}
