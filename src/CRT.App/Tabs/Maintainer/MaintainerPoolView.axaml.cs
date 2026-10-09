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
    // Account > Maintainers - see the markup for what it is. The pool controls the Boards screen's
    // Maintainer view carried from 2026-09-27 to 2026-10-04, moved here unchanged in what they send
    // (owner request, 2026-10-04: "this is something only the admin should be able to do").
    //
    // *** WHAT IS READ, AND WHEN. *** Choosing "Maintainers" on the Account screen reads every board
    // (with its maintainers, for the count beside each name) and every account (for the "choose
    // somebody" list). Choosing a board reads its detail - its maintainers and the invitations
    // nobody has accepted, which only the detail carries. Nothing is read on the minute check.
    //
    // *** NO AUTHORITY LIVES HERE. *** The screen is offered to an administrator only; the server
    // refuses anybody else regardless, and its sentence lands in the message line.
    //
    // *** AFTER EVERY CHANGE THE BOARD IS READ AGAIN FROM THE SERVER *** rather than patched
    // locally, so the screen is always what the server holds - and the tab is told
    // (AfterPoolChange), since the queue is filtered by pools and the Boards list counts maintainers.
    //
    // *** UNDER THE OVERLAY, AND CHECKED AFTER A TIMEOUT (2026-09-28). *** No answer in two minutes
    // does not mean nothing changed: the board is read again and the pool itself says whether the
    // change landed (PoolAction.HasLanded).
    // ###########################################################################################
    public partial class MaintainerPoolView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        // Every board, in the combo's order.
        private readonly List<ReviewBoardRow> thisBoards = [];

        // Every account, for the "choose somebody" list.
        private readonly List<ReviewAccountRow> thisAccounts = [];

        // The accounts the list offers - everybody not already maintaining the board - in its order.
        private readonly List<ReviewAccountRow> thisAccountChoices = [];

        // The chosen board's detail, as last read.
        private BoardDetailAnswer? thisDetail;

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

        // Told after a pool changed: the tab reads its queue and its boards again.
        public Func<Task>? AfterPoolChange { get; set; }

        // The Boards screen's board ids, top to bottom - the order the combo lists the boards in
        // (BoardsDisplay.InBoardsListOrder). Unset or empty: by name.
        public Func<IReadOnlyList<string>>? BoardsListOrder { get; set; }

        // The board on show, or null.
        internal string? ShownBoardId => this.thisDetail?.Board.BoardId;

        // Signed out: nothing of the previous account's boards, people or messages left.
        public void Clear()
        {
            this.thisRequest++;
            this.thisBoards.Clear();
            this.thisAccounts.Clear();
            this.thisAccountChoices.Clear();
            this.thisDetail = null;

            this.FillBoards(keep: null);
            this.ShowPool(null);
            this.ShowMessage(null, isError: false);
            this.ShowPoolMessage(null, isError: false);
        }

        // ###########################################################################################
        // Reads every board and every account, under CRT's overlay, keeping the board chosen when
        // it is still there (and reading it again). A list that cannot be read leaves the previous
        // one, with the reason - "could not ask" must never read as "there are none".
        // ###########################################################################################
        public async Task LoadAsync()
        {
            if (this.thisClient is not ReviewApiClient client || this.thisSession is not ReviewSession session)
                return;

            ReviewApiResult<ReviewBoardsResponse> boards = await ServerWait.CallAsync(
                this, MaintainerWaitWording.ReadingBoards, token => client.GetBoardsAsync(session, token));

            if (!boards.IsOk)
            {
                this.ShowMessage(boards.Message, isError: true);
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
            await this.UseBoardsAsync(boards.Value!.Boards);
        }

        // ###########################################################################################
        // The boards in the combo, in the Boards screen's order, each with how many maintain it. The
        // one chosen stays chosen when it is still there - and is read again, since its pool may
        // have changed.
        // ###########################################################################################
        private async Task UseBoardsAsync(IReadOnlyList<ReviewBoardRow> boards)
        {
            string? chosen = this.ShownBoardId;

            this.thisBoards.Clear();
            this.thisBoards.AddRange(BoardsDisplay.InBoardsListOrder(
                boards.Where(board => !string.IsNullOrWhiteSpace(board.BoardId)),
                board => board.BoardId,
                MaintainerAssignmentDisplay.BoardLine,
                this.BoardsListOrder?.Invoke()));

            this.FillBoards(keep: chosen);

            if (chosen is not null && this.thisBoards.Any(board => string.Equals(board.BoardId, chosen, StringComparison.Ordinal)))
                await this.ShowBoardAsync(chosen);
            else
                this.ShowPool(null);
        }

        private void FillBoards(string? keep)
        {
            if (this.FindControl<ComboBox>("BoardCombo") is not ComboBox combo)
                return;

            this.thisIsFilling = true;

            try
            {
                combo.ItemsSource = this.thisBoards.Select(MaintainerAssignmentDisplay.BoardLine).ToList();
                combo.SelectedIndex = keep is null
                    ? -1
                    : this.thisBoards.FindIndex(board => string.Equals(board.BoardId, keep, StringComparison.Ordinal));
            }
            finally
            {
                this.thisIsFilling = false;
            }
        }

        private async void OnBoardChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisIsFilling || sender is not ComboBox combo)
                return;

            int index = combo.SelectedIndex;

            if (index < 0 || index >= this.thisBoards.Count)
                return;

            this.ShowPoolMessage(null, isError: false);
            await this.ShowBoardAsync(this.thisBoards[index].BoardId);
        }

        // Reads one board's pool and shows it.
        private async Task ShowBoardAsync(string boardId)
        {
            int request = ++this.thisRequest;

            BoardDetailAnswer? detail = await this.ReadDetailAsync(boardId, MaintainerWaitWording.ReadingBoard(boardId));

            if (request != this.thisRequest)
                return;

            if (detail is not null)
                this.ShowPool(detail);
        }

        // The board's detail - from a test's answer, or the server's under the overlay. Null, with
        // the reason shown, when it cannot be read.
        private async Task<BoardDetailAnswer?> ReadDetailAsync(string boardId, string waiting)
        {
            if (this.PoolDetailOverrideForTests is not null)
                return await this.PoolDetailOverrideForTests(boardId);

            if (this.thisClient is not ReviewApiClient client || this.thisSession is not ReviewSession session)
                return null;

            ReviewApiResult<BoardDetailAnswer> detail = await ServerWait.CallAsync(
                this, waiting, token => client.GetBoardDetailAsync(session, boardId, token));

            if (!detail.IsOk)
            {
                this.ShowMessage(detail.Message, isError: true);
                return null;
            }

            return detail.Value;
        }

        // ###########################################################################################
        // The chosen board's pool: the heading, a line and a Remove button per maintainer, a line
        // and a Withdraw button per open invitation, and the accounts the list can offer. Null hides
        // it all.
        // ###########################################################################################
        private void ShowPool(BoardDetailAnswer? detail)
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
                Text = BoardsDisplay.MaintainersHeading(count),
                FontSize = 14,
                FontWeight = count == 0 ? Avalonia.Media.FontWeight.Normal : Avalonia.Media.FontWeight.SemiBold,
                Opacity = count == 0 ? 0.7 : 1,
                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                Margin = new Avalonia.Thickness(0, 0, 0, 2)
            });

            string boardId = detail.Board.BoardId;

            foreach (PoolMaintainerEntry maintainer in detail.Maintainers)
            {
                long accountId = maintainer.AccountId;
                string name = maintainer.DisplayName;

                // The person in bold (owner request, 2026-10-09).
                var line = new TextBlock { TextWrapping = Avalonia.Media.TextWrapping.Wrap };
                TabMaintainer.ShowCounts(line, BoardsDisplay.MaintainerRuns(maintainer));

                section.Children.Add(MaintainerPoolView.WithButton(
                    line,
                    "Remove",
                    "RemoveMaintainer",
                    async () => await this.ChangePoolAsync(
                        new PoolAction(PoolActionKind.Remove, boardId, AccountId: accountId, Who: name),
                        $"{name} no longer maintains {boardId}. It takes effect on their next request.")));
            }

            foreach (MaintainerInvitationEntry invitation in detail.Invitations ?? [])
            {
                long invitationId = invitation.Id;
                string invited = invitation.Email;

                var lines = new StackPanel { Spacing = 1 };
                var invitedLine = new TextBlock { TextWrapping = Avalonia.Media.TextWrapping.Wrap };
                TabMaintainer.ShowCounts(invitedLine, BoardsDisplay.InvitationRuns(invitation));
                lines.Children.Add(invitedLine);
                lines.Children.Add(new TextBlock { Text = BoardsDisplay.InvitationFooter(invitation), FontSize = 11, Opacity = 0.7, TextWrapping = Avalonia.Media.TextWrapping.Wrap });

                section.Children.Add(MaintainerPoolView.WithButton(
                    lines,
                    "Withdraw",
                    "WithdrawInvitation",
                    async () => await this.ChangePoolAsync(new PoolAction(PoolActionKind.Withdraw, boardId, InvitationId: invitationId, Who: invited), null)));
            }

            this.ShowAccountChoices();
        }

        // The accounts that can be offered: everybody not already maintaining the chosen board. The
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
                new PoolAction(PoolActionKind.Add, this.thisDetail.Board.BoardId, AccountId: account.Id, Who: account.DisplayName),
                $"{account.DisplayName} now maintains {this.thisDetail.Board.BoardId}.");
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

            if (await this.ChangePoolAsync(new PoolAction(PoolActionKind.Invite, this.thisDetail.Board.BoardId, Email: email, Who: email), null))
                box.Text = string.Empty;
        }

        // ###########################################################################################
        // Sends one change, says what happened, and reads the board and the list of boards again.
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
                    BoardDetailAnswer? now = await this.ReadDetailAsync(action.BoardId, WaitWording.Checking);
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

                await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingBoard(action.BoardId), async () =>
                {
                    if (this.PoolDetailOverrideForTests is null)
                    {
                        await this.ReloadBoardsQuietlyAsync();
                    }
                    else if (this.ShownBoardId is string shown)
                    {
                        await this.ShowBoardAsync(shown);
                    }

                    if (this.AfterPoolChange is not null)
                        await this.AfterPoolChange();
                });
            });

            return changed;
        }

        // The boards again (their counts moved) and the chosen one with them - no message cleared.
        private async Task ReloadBoardsQuietlyAsync()
        {
            if (this.thisClient is not ReviewApiClient client || this.thisSession is not ReviewSession session)
                return;

            ReviewApiResult<ReviewBoardsResponse> boards = await client.GetBoardsAsync(session);

            if (boards.IsOk)
                await this.UseBoardsAsync(boards.Value!.Boards);
            else if (this.ShownBoardId is string shown)
                await this.ShowBoardAsync(shown);
        }

        private Task<ReviewApiResult<string>> SendAsync(PoolAction action, CancellationToken token)
        {
            if (this.PoolActionOverrideForTests is not null)
                return this.PoolActionOverrideForTests(action);

            if (this.thisClient is null || this.thisSession is null)
                return Task.FromResult(ReviewApiResult<string>.Failed(ReviewApiFailure.NotSignedIn, "Sign in first."));

            return action.Kind switch
            {
                PoolActionKind.Add => this.thisClient.AddMaintainerAsync(this.thisSession, action.BoardId, action.AccountId, token),
                PoolActionKind.Remove => this.thisClient.RemoveMaintainerAsync(this.thisSession, action.BoardId, action.AccountId, token),
                PoolActionKind.Invite => this.thisClient.InviteMaintainerAsync(this.thisSession, action.BoardId, action.Email, token),
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

        // Answers every read of a board instead of the server.
        internal Func<string, Task<BoardDetailAnswer>>? PoolDetailOverrideForTests { get; set; }

        // The lists as LoadAsync would have read them, and a board as choosing it would show it.
        internal Task UseListsForTests(IReadOnlyList<ReviewBoardRow> boards, IReadOnlyList<ReviewAccountRow> accounts)
        {
            this.UseAccounts(accounts);
            return this.UseBoardsAsync(boards);
        }

        internal Task ChooseBoardForTests(string boardId)
        {
            int index = this.thisBoards.FindIndex(board => string.Equals(board.BoardId, boardId, StringComparison.Ordinal));

            if (this.FindControl<ComboBox>("BoardCombo") is ComboBox combo)
            {
                this.thisIsFilling = true;
                combo.SelectedIndex = index;
                this.thisIsFilling = false;
            }

            return index < 0 ? Task.CompletedTask : this.ShowBoardAsync(boardId);
        }

        // The texts of the pool, top to bottom.
        internal IReadOnlyList<string> PoolTextsForTests() =>
            this.FindControl<StackPanel>("MaintainersSection")!.Children
                .SelectMany(MaintainerPoolView.TextsOf)
                .ToList();

        private static IEnumerable<string> TextsOf(Control control) => control switch
        {
            TextBlock block => [TabMaintainer.TextOf(block)],
            Panel panel => panel.Children.Where(child => child is not Button).SelectMany(MaintainerPoolView.TextsOf),
            _ => []
        };
    }
}
