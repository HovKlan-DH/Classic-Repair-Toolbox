using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // "Delete a board" on the Account screen (owner request, 2026-10-03). See the markup for what it
    // is; the server's BoardDeletionFlow for what a delete does and refuses.
    //
    // *** PLAN, CONFIRM, DELETE - in that order, every time. *** Delete first asks the server what
    // would go (BoardDeletePlanAnswer) and shows it in DeleteBoardWindow; the delete then sends
    // the plan's fingerprint back, so the server deletes nothing the administrator was not shown. A
    // second press after a delete that stopped part-way plans afresh, which is what finishes it.
    //
    // *** NOTHING IS READ UNTIL IT IS CHOSEN. *** The boards are read when this item is chosen on
    // the Account screen and after each delete - not on the minute check, which this screen does not
    // need.
    // ###########################################################################################
    public partial class BoardDeletionView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        public BoardDeletionView()
        {
            this.InitializeComponent();

            this.SetText("HeadingText", BoardDeletionWording.ListHeading);
            this.SetText("ExplanationText", BoardDeletionWording.Explanation);
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        // After a delete went through: the tab reads its queue, BETA list and boards again.
        public Func<Task>? AfterDelete { get; set; }

        // The Boards screen's board ids, top to bottom - the order the boards are listed in here
        // (BoardsDisplay.InBoardsListOrder). Unset or empty: by name.
        public Func<IReadOnlyList<string>>? BoardsListOrder { get; set; }

        // ###########################################################################################
        // The confirmation, for tests: ShowDialog blocks headlessly, so a test answers in its place.
        // The shipped path builds DeleteBoardWindow through the same Initialize a test of the window
        // drives.
        // ###########################################################################################
        internal Func<BoardDeletePlanAnswer, Task<(bool Confirmed, string Reason)>>? ConfirmOverrideForTests { get; set; }

        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
        }

        // Signed out: nothing of the previous account's list or message left on screen.
        public void Clear()
        {
            this.FindControl<StackPanel>("BoardsPanel")?.Children.Clear();
            this.ShowMessage(null, isError: false);
        }

        // ###########################################################################################
        // Reads every board, under CRT's overlay. A list that cannot be read leaves the previous one
        // in place, with the reason - "could not ask" must never read as "there are none".
        // ###########################################################################################
        public async Task LoadAsync()
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;

            ReviewApiResult<BoardOverviewAnswer> result = await ServerWait.CallAsync(
                this,
                MaintainerWaitWording.ReadingBoards,
                token => client.GetBoardOverviewAsync(session, token));

            if (!result.IsOk)
            {
                this.ShowMessage(result.Message, isError: true);
                return;
            }

            this.ShowBoards(result.Value!.Boards);
        }

        // ###########################################################################################
        // One row per board, by name: its name, its grey line - the Boards screen's own words - and
        // a red Delete button. Internal so a test can lay a list out without a server.
        // ###########################################################################################
        internal void ShowBoards(IReadOnlyList<BoardOverviewEntry> boards)
        {
            ArgumentNullException.ThrowIfNull(boards);

            if (this.FindControl<StackPanel>("BoardsPanel") is not StackPanel panel)
                return;

            panel.Children.Clear();

            foreach (BoardOverviewEntry board in BoardsDisplay.InBoardsListOrder(boards, board => board.BoardId, BoardsDisplay.Name, this.BoardsListOrder?.Invoke()))
                panel.Children.Add(this.BoardRow(board));

            if (boards.Count == 0)
                panel.Children.Add(new TextBlock { Text = BoardDeletionWording.NoBoards, Opacity = 0.7 });
        }

        private Control BoardRow(BoardOverviewEntry board)
        {
            var texts = new StackPanel { Spacing = 1, VerticalAlignment = VerticalAlignment.Center };

            texts.Children.Add(new TextBlock { Text = BoardsDisplay.Name(board), FontSize = 13, TextWrapping = TextWrapping.Wrap });

            string line = BoardsDisplay.ListLine(board);

            if (line.Length > 0)
                texts.Children.Add(new TextBlock { Text = line, FontSize = 11, Opacity = 0.7, TextWrapping = TextWrapping.Wrap });

            var delete = new Button
            {
                Content = BoardDeletionWording.DeleteButton,
                Tag = board,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(12, 0, 0, 0)
            };

            delete.Classes.Add("DeleteBoardButton");
            delete.Bind(Button.ForegroundProperty, delete.GetResourceObservable("Button_Cancel_Fg"));
            delete.Bind(Button.BackgroundProperty, delete.GetResourceObservable("Button_Cancel_Bg"));
            delete.Bind(Button.BorderBrushProperty, delete.GetResourceObservable("Button_Cancel_Border"));
            delete.Click += async (_, _) => await this.DeleteAsync(board);

            var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            grid.Children.Add(texts);
            grid.Children.Add(delete);
            Grid.SetColumn(delete, 1);

            var row = new Border
            {
                Padding = new Thickness(0, 6),
                BorderThickness = new Thickness(0, 0, 0, 1),
                Child = grid
            };

            row.Bind(Border.BorderBrushProperty, row.GetResourceObservable("Queue_Divider"));

            return row;
        }

        // ###########################################################################################
        // Plan, confirm, delete - see the header. Every Delete button is off for the duration, so no
        // second delete starts while one is being confirmed or carried out. Internal so a test can
        // await what a click starts.
        // ###########################################################################################
        internal async Task DeleteAsync(BoardOverviewEntry board)
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;

            this.SetDeleteButtonsEnabled(false);
            this.ShowMessage(null, isError: false);

            try
            {
                ReviewApiResult<BoardDeletePlanAnswer> planned = await ServerWait.CallAsync(
                    this,
                    MaintainerWaitWording.ReadingDeletionPlan(board.BoardId),
                    token => client.GetBoardDeletePlanAsync(session, board.BoardId, token));

                if (!planned.IsOk)
                {
                    this.ShowMessage(planned.Message, isError: true);
                    return;
                }

                BoardDeletePlanAnswer plan = planned.Value!;

                // Cannot go at all: said here, in the server's words, and nothing to confirm.
                if (!string.IsNullOrWhiteSpace(plan.BlockedBecause))
                {
                    this.ShowMessage(plan.BlockedBecause, isError: true);
                    return;
                }

                (bool confirmed, string reason) = await this.ConfirmAsync(plan);

                if (!confirmed)
                    return;

                await this.CarryOutAsync(client, session, plan, reason);
            }
            finally
            {
                this.SetDeleteButtonsEnabled(true);
            }
        }

        private async Task<(bool Confirmed, string Reason)> ConfirmAsync(BoardDeletePlanAnswer plan)
        {
            if (this.ConfirmOverrideForTests is not null)
                return await this.ConfirmOverrideForTests(plan);

            if (TopLevel.GetTopLevel(this) is not Window owner)
                return (false, string.Empty);

            var confirm = new DeleteBoardWindow();
            confirm.Initialize(plan);

            await confirm.ShowDialog(owner);

            return (confirm.WasConfirmed, confirm.Reason);
        }

        // ###########################################################################################
        // The delete, held under the overlay until the lists are read again too (BusyOverlay.HoldAsync
        // - one wait, no flicker between the steps). No answer in two minutes is never "it failed":
        // whether the board is still listed says whether it went (MaintainerWaitWording).
        // ###########################################################################################
        private async Task CarryOutAsync(ReviewApiClient client, ReviewSession session, BoardDeletePlanAnswer plan, string reason)
        {
            string waiting = MaintainerWaitWording.DeletingBoard(plan.BoardId);
            string? outcome = null;
            bool failed = true;

            await BusyOverlay.HoldAsync(this, waiting, async () =>
            {
                ReviewApiResult<BoardDeleteAnswer> result = await ServerWait.CallAsync(
                    this,
                    waiting,
                    token => client.DeleteBoardAsync(session, plan.BoardId, plan.Fingerprint, reason, token));

                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    bool? stillListed = await this.IsStillListedAsync(client, session, plan.BoardId);

                    outcome = MaintainerWaitWording.DeleteAfterTimeout(plan.BoardId, stillListed);
                    failed = stillListed != false;
                    await this.AfterChangeAsync();
                    return;
                }

                if (!result.IsOk)
                {
                    outcome = result.Message;

                    // A delete refused part-way has still changed something - the lists show what is left.
                    await this.LoadAsync();
                    return;
                }

                outcome = BoardDeletionWording.Done(result.Value!);
                failed = false;
                await this.AfterChangeAsync();
            });

            this.ShowMessage(outcome ?? "The delete did not complete.", isError: failed);
        }

        // The list again, and the tab's own lists. Each reads quietly; a failed read leaves its list.
        private async Task AfterChangeAsync()
        {
            await this.LoadAsync();

            if (this.AfterDelete is not null)
                await this.AfterDelete();
        }

        // Null when that look failed too.
        private async Task<bool?> IsStillListedAsync(ReviewApiClient client, ReviewSession session, string boardId)
        {
            ReviewApiResult<BoardOverviewAnswer> list = await ServerWait.CallAsync(
                this,
                MaintainerWaitWording.ReadingBoards,
                token => client.GetBoardOverviewAsync(session, token));

            return list.IsOk
                ? list.Value!.Boards.Any(entry => string.Equals(entry.BoardId, boardId, StringComparison.Ordinal))
                : null;
        }

        private void SetDeleteButtonsEnabled(bool enabled)
        {
            foreach (Button button in this.DeleteButtons())
                button.IsEnabled = enabled;
        }

        // Every row's Delete button, in the order shown.
        internal IReadOnlyList<Button> DeleteButtons() =>
            this.FindControl<StackPanel>("BoardsPanel")?.Children
                .OfType<Border>()
                .Select(row => row.Child)
                .OfType<Grid>()
                .SelectMany(grid => grid.Children.OfType<Button>())
                .ToList() ?? [];

        // Every board's name, in the order shown - for tests.
        internal IReadOnlyList<string> NamesForTests() =>
            this.FindControl<StackPanel>("BoardsPanel")?.Children
                .OfType<Border>()
                .Select(row => row.Child)
                .OfType<Grid>()
                .SelectMany(grid => grid.Children.OfType<StackPanel>())
                .Select(texts => texts.Children.OfType<TextBlock>().First().Text ?? string.Empty)
                .ToList() ?? [];

        internal string? MessageForTests => this.FindControl<TextBlock>("MessageText") is { IsVisible: true } message ? message.Text : null;

        private void SetText(string name, string text)
        {
            if (this.FindControl<TextBlock>(name) is TextBlock block)
                block.Text = text;
        }

        private void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);
    }
}
