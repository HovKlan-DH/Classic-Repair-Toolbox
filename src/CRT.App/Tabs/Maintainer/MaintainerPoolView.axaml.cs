using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // Account > Maintainers - see the markup for what it is. The pool controls the Systems screen's
    // Maintainer view carried from 2026-09-27 to 2026-10-04, moved here unchanged in what they send
    // (owner request, 2026-10-04: "this is something only the admin should be able to do").
    //
    // *** WHAT IS READ, AND WHEN. *** Choosing "Maintainers" on the Account screen reads every system
    // (with its maintainers, for the count beside each name) and every account (for the "choose
    // somebody" list). Choosing a system reads its detail - its maintainers and the invitations
    // nobody has accepted, which only the detail carries. Nothing is read on the minute check.
    //
    // *** NO AUTHORITY LIVES HERE. *** The screen is offered to an administrator only; the server
    // refuses anybody else regardless, and its sentence lands in the message line.
    //
    // *** AFTER EVERY CHANGE THE SYSTEM IS READ AGAIN FROM THE SERVER *** rather than patched
    // locally, so the screen is always what the server holds - and the tab is told
    // (AfterPoolChange), since the queue is filtered by pools and the Systems list counts maintainers.
    //
    // *** UNDER THE OVERLAY, AND CHECKED AFTER A TIMEOUT (2026-09-28). *** No answer in two minutes
    // does not mean nothing changed: the system is read again and the pool itself says whether the
    // change landed (PoolAction.HasLanded).
    // ###########################################################################################
    public partial class MaintainerPoolView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        // Every system, in the combo's order.
        private readonly List<ReviewSystemRow> thisSystems = [];

        // Every account, for the "choose somebody" list.
        private readonly List<ReviewAccountRow> thisAccounts = [];

        // The accounts the list offers - everybody not already maintaining the system - in its order.
        private readonly List<ReviewAccountRow> thisAccountChoices = [];

        // The chosen system's detail, as last read.
        private SystemDetailAnswer? thisDetail;

        // Filling the combo from code is not a choice.
        private bool thisIsFilling;

        // A counter that drops an answer the selection has moved past.
        private int thisRequest;

        public MaintainerPoolView()
        {
            this.InitializeComponent();
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
        }

        // Told after a pool changed: the tab reads its queue and its systems again.
        public Func<Task>? AfterPoolChange { get; set; }

        // The system on show, or null.
        internal string? ShownSystemId => this.thisDetail?.System.SystemId;

        // Signed out: nothing of the previous account's systems, people or messages left.
        public void Clear()
        {
            this.thisRequest++;
            this.thisSystems.Clear();
            this.thisAccounts.Clear();
            this.thisAccountChoices.Clear();
            this.thisDetail = null;

            this.FillSystems(keep: null);
            this.ShowPool(null);
            this.ShowMessage(null, isError: false);
            this.ShowPoolMessage(null, isError: false);
        }

        // ###########################################################################################
        // Reads every system and every account, under CRT's overlay, keeping the system chosen when
        // it is still there (and reading it again). A list that cannot be read leaves the previous
        // one, with the reason - "could not ask" must never read as "there are none".
        // ###########################################################################################
        public async Task LoadAsync()
        {
            if (this.thisClient is not ReviewApiClient client || this.thisSession is not ReviewSession session)
                return;

            ReviewApiResult<ReviewSystemsResponse> systems = await ServerWait.CallAsync(
                this, MaintainerWaitWording.ReadingSystems, token => client.GetSystemsAsync(session, token));

            if (!systems.IsOk)
            {
                this.ShowMessage(systems.Message, isError: true);
                return;
            }

            ReviewApiResult<ReviewAccountsResponse> accounts = await ServerWait.CallAsync(
                this, MaintainerWaitWording.ReadingAccounts, token => client.GetAccountsAsync(session, token));

            if (!accounts.IsOk)
            {
                this.ShowMessage(accounts.Message, isError: true);
                return;
            }

            this.ShowMessage(null, isError: false);
            this.UseAccounts(accounts.Value!.Accounts);
            await this.UseSystemsAsync(systems.Value!.Systems);
        }

        // ###########################################################################################
        // The systems in the combo, by name, each with how many maintain it. The one chosen stays
        // chosen when it is still there - and is read again, since its pool may have changed.
        // ###########################################################################################
        private async Task UseSystemsAsync(IReadOnlyList<ReviewSystemRow> systems)
        {
            string? chosen = this.ShownSystemId;

            this.thisSystems.Clear();
            this.thisSystems.AddRange(systems
                .Where(system => !string.IsNullOrWhiteSpace(system.SystemId))
                .OrderBy(system => MaintainerAssignmentDisplay.SystemLine(system), StringComparer.OrdinalIgnoreCase));

            this.FillSystems(keep: chosen);

            if (chosen is not null && this.thisSystems.Any(system => string.Equals(system.SystemId, chosen, StringComparison.Ordinal)))
                await this.ShowSystemAsync(chosen);
            else
                this.ShowPool(null);
        }

        private void FillSystems(string? keep)
        {
            if (this.FindControl<ComboBox>("SystemCombo") is not ComboBox combo)
                return;

            this.thisIsFilling = true;

            try
            {
                combo.ItemsSource = this.thisSystems.Select(MaintainerAssignmentDisplay.SystemLine).ToList();
                combo.SelectedIndex = keep is null
                    ? -1
                    : this.thisSystems.FindIndex(system => string.Equals(system.SystemId, keep, StringComparison.Ordinal));
            }
            finally
            {
                this.thisIsFilling = false;
            }
        }

        private async void OnSystemChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisIsFilling || sender is not ComboBox combo)
                return;

            int index = combo.SelectedIndex;

            if (index < 0 || index >= this.thisSystems.Count)
                return;

            this.ShowPoolMessage(null, isError: false);
            await this.ShowSystemAsync(this.thisSystems[index].SystemId);
        }

        // Reads one system's pool and shows it.
        private async Task ShowSystemAsync(string systemId)
        {
            int request = ++this.thisRequest;

            SystemDetailAnswer? detail = await this.ReadDetailAsync(systemId, MaintainerWaitWording.ReadingSystem(systemId));

            if (request != this.thisRequest)
                return;

            if (detail is not null)
                this.ShowPool(detail);
        }

        // The system's detail - from a test's answer, or the server's under the overlay. Null, with
        // the reason shown, when it cannot be read.
        private async Task<SystemDetailAnswer?> ReadDetailAsync(string systemId, string waiting)
        {
            if (this.PoolDetailOverrideForTests is not null)
                return await this.PoolDetailOverrideForTests(systemId);

            if (this.thisClient is not ReviewApiClient client || this.thisSession is not ReviewSession session)
                return null;

            ReviewApiResult<SystemDetailAnswer> detail = await ServerWait.CallAsync(
                this, waiting, token => client.GetSystemDetailAsync(session, systemId, token));

            if (!detail.IsOk)
            {
                this.ShowMessage(detail.Message, isError: true);
                return null;
            }

            return detail.Value;
        }

        // ###########################################################################################
        // The chosen system's pool: the heading, a line and a Remove button per maintainer, a line
        // and a Withdraw button per open invitation, and the accounts the list can offer. Null hides
        // it all.
        // ###########################################################################################
        private void ShowPool(SystemDetailAnswer? detail)
        {
            this.thisDetail = detail;

            if (this.FindControl<StackPanel>("PoolPanel") is StackPanel pool)
                pool.IsVisible = detail is not null;

            if (this.FindControl<StackPanel>("MaintainersSection") is not StackPanel section)
                return;

            section.Children.Clear();

            if (detail is null)
            {
                this.ShowAccountChoices();
                return;
            }

            int count = detail.Maintainers.Count;

            section.Children.Add(new TextBlock
            {
                Text = SystemsDisplay.MaintainersHeading(count),
                FontSize = 14,
                FontWeight = count == 0 ? Avalonia.Media.FontWeight.Normal : Avalonia.Media.FontWeight.SemiBold,
                Opacity = count == 0 ? 0.7 : 1,
                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                Margin = new Avalonia.Thickness(0, 0, 0, 2)
            });

            string systemId = detail.System.SystemId;

            foreach (PoolMaintainerEntry maintainer in detail.Maintainers)
            {
                long accountId = maintainer.AccountId;
                string name = maintainer.DisplayName;

                section.Children.Add(MaintainerPoolView.WithButton(
                    new TextBlock { Text = SystemsDisplay.MaintainerLine(maintainer), TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                    "Remove",
                    "RemoveMaintainer",
                    async () => await this.ChangePoolAsync(
                        new PoolAction(PoolActionKind.Remove, systemId, AccountId: accountId, Who: name),
                        $"{name} no longer maintains {systemId}. It takes effect on their next request.")));
            }

            foreach (MaintainerInvitationEntry invitation in detail.Invitations ?? [])
            {
                long invitationId = invitation.Id;
                string invited = invitation.Email;

                var lines = new StackPanel { Spacing = 1 };
                lines.Children.Add(new TextBlock { Text = SystemsDisplay.InvitationLine(invitation), TextWrapping = Avalonia.Media.TextWrapping.Wrap });
                lines.Children.Add(new TextBlock { Text = SystemsDisplay.InvitationFooter(invitation), FontSize = 11, Opacity = 0.7, TextWrapping = Avalonia.Media.TextWrapping.Wrap });

                section.Children.Add(MaintainerPoolView.WithButton(
                    lines,
                    "Withdraw",
                    "WithdrawInvitation",
                    async () => await this.ChangePoolAsync(new PoolAction(PoolActionKind.Withdraw, systemId, InvitationId: invitationId, Who: invited), null)));
            }

            this.ShowAccountChoices();
        }

        // The accounts that can be offered: everybody not already maintaining the chosen system. The
        // line says why one cannot be granted (MaintainerAssignmentDisplay).
        private void ShowAccountChoices()
        {
            if (this.FindControl<ComboBox>("AddAccountCombo") is not ComboBox combo)
                return;

            HashSet<long> already = (this.thisDetail?.Maintainers ?? []).Select(maintainer => maintainer.AccountId).ToHashSet();

            this.thisAccountChoices.Clear();

            if (this.thisDetail is not null)
                this.thisAccountChoices.AddRange(this.thisAccounts.Where(account => !already.Contains(account.Id)));

            combo.ItemsSource = this.thisAccountChoices.Select(MaintainerAssignmentDisplay.AccountChoice).ToList();
            combo.SelectedIndex = -1;
        }

        private void UseAccounts(IReadOnlyList<ReviewAccountRow> accounts)
        {
            this.thisAccounts.Clear();
            this.thisAccounts.AddRange(accounts);
            this.ShowAccountChoices();
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
        // Sends one change, says what happened, and reads the system and the list of systems again.
        // The success sentence is the server's when it sends one (an invitation's), else `done`.
        // ###########################################################################################
        private async Task<bool> ChangePoolAsync(PoolAction action, string? done)
        {
            bool changed = false;

            await BusyOverlay.HoldAsync(this, action.Waiting, async () =>
            {
                ReviewApiResult<string> result = await ServerWait.CallAsync(this, action.Waiting, token => this.SendAsync(action, token));

                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    SystemDetailAnswer? now = await this.ReadDetailAsync(action.SystemId, WaitWording.Checking);
                    bool? landed = now is null ? null : action.HasLanded(now);

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
                    if (this.PoolDetailOverrideForTests is null)
                    {
                        await this.ReloadSystemsQuietlyAsync();
                    }
                    else if (this.ShownSystemId is string shown)
                    {
                        await this.ShowSystemAsync(shown);
                    }

                    if (this.AfterPoolChange is not null)
                        await this.AfterPoolChange();
                });
            });

            return changed;
        }

        // The systems again (their counts moved) and the chosen one with them - no message cleared.
        private async Task ReloadSystemsQuietlyAsync()
        {
            if (this.thisClient is not ReviewApiClient client || this.thisSession is not ReviewSession session)
                return;

            ReviewApiResult<ReviewSystemsResponse> systems = await client.GetSystemsAsync(session);

            if (systems.IsOk)
                await this.UseSystemsAsync(systems.Value!.Systems);
            else if (this.ShownSystemId is string shown)
                await this.ShowSystemAsync(shown);
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

        private void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);

        private void ShowPoolMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("PoolMessageText"), message, isError);

        // -----------------------------------------------------------------------------------
        // For tests - no server.
        // -----------------------------------------------------------------------------------

        // Answers every change instead of the server, and sees what was sent.
        internal Func<PoolAction, Task<ReviewApiResult<string>>>? PoolActionOverrideForTests { get; set; }

        // Answers every read of a system instead of the server.
        internal Func<string, Task<SystemDetailAnswer>>? PoolDetailOverrideForTests { get; set; }

        // The lists as LoadAsync would have read them, and a system as choosing it would show it.
        internal Task UseListsForTests(IReadOnlyList<ReviewSystemRow> systems, IReadOnlyList<ReviewAccountRow> accounts)
        {
            this.UseAccounts(accounts);
            return this.UseSystemsAsync(systems);
        }

        internal Task ChooseSystemForTests(string systemId)
        {
            int index = this.thisSystems.FindIndex(system => string.Equals(system.SystemId, systemId, StringComparison.Ordinal));

            if (this.FindControl<ComboBox>("SystemCombo") is ComboBox combo)
            {
                this.thisIsFilling = true;
                combo.SelectedIndex = index;
                this.thisIsFilling = false;
            }

            return index < 0 ? Task.CompletedTask : this.ShowSystemAsync(systemId);
        }

        // The texts of the pool, top to bottom.
        internal IReadOnlyList<string> PoolTextsForTests() =>
            this.FindControl<StackPanel>("MaintainersSection")!.Children
                .SelectMany(MaintainerPoolView.TextsOf)
                .ToList();

        private static IEnumerable<string> TextsOf(Control control) => control switch
        {
            TextBlock block => [block.Text ?? string.Empty],
            Panel panel => panel.Children.Where(child => child is not Button).SelectMany(MaintainerPoolView.TextsOf),
            _ => []
        };
    }
}
