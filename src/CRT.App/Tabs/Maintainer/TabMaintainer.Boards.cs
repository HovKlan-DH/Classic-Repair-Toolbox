using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Media;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // THE "BOARDS" SCREEN'S LIST (owner request, 2026-09-27): every board - the `boards` rows and
    // the boards in the BETA tree - each as its name and one grey line. Choosing one shows its
    // maintainers, contributors and submissions in BoardDetailView on the right.
    //
    // For EVERY maintainer, every board (owner decision: "Everything for everyone") - the server's
    // BoardOverviewFlow says why. The Boards button's discreet badge is the count.
    //
    // A check nobody asked for re-reads the chosen board's detail only when its row changed, and
    // then without blanking the panel first - the same rule the queue and the BETA list keep.
    //
    // *** THE DROP-DOWN LISTS ARE READ WITH IT (owner request, 2026-09-27). *** A new board must be
    // placed in CRT's drop-down lists before it can be published to BETA, and it is placed here. So
    // the list is in CRT's drop-down order with the boards waiting for a place FIRST and marked,
    // the Boards button's badge turns into an attention badge counting the ones this account can
    // place, and the chosen board's panel carries BoardPlacementView. A listing that could not be
    // read leaves the previous one in place - never "not in the lists" on a failed request.
    //
    // *** A BOARD'S TABLE CAN HOLD A CHANGE NOT PUBLISHED YET (2026-10-03). *** Choosing another
    // board asks first, as choosing another submission does. A change published to BETA reads the
    // BETA > Stable list and the boards again (AfterBoardChangePublishedAsync), staying on this
    // screen; one the server made into a submission but could not publish opens it under Contributor
    // Submissions, where it waits (OpenSentSubmissionAsync).
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private readonly List<BoardOverviewEntry> thisBoards = [];

        // BETA's drop-down lists and the boards not in them, as last read.
        private BoardListingAnswer? thisListing;

        // Whether the list has been read since signing in - the badge says nothing until it has.
        private bool thisBoardsKnown;

        private bool thisSuppressBoardsSelection;

        private async Task RefreshBoardsAsync(bool background)
        {
            if (!this.CanAskServer)
                return;

            ReviewApiResult<BoardOverviewAnswer> result = await this.thisClient!.GetBoardOverviewAsync(this.thisSession!);

            if (!result.IsOk)
            {
                // Left as it was: "could not ask" must not read as "there are none".
                this.ShowBoardsListMessage(result.Message, isError: true);
                return;
            }

            ReviewApiResult<BoardListingAnswer> listing = await this.thisClient!.GetBoardListingAsync(this.thisSession!);

            await this.ApplyBoardsListAsync(result.Value!, background, listing.IsOk ? listing.Value : null);
        }

        // After a placement was saved: the lists again, the chosen board kept.
        private Task AfterPlacementSavedAsync() => this.RefreshBoardsAsync(background: true);

        // ###########################################################################################
        // The minute check away from the Boards screen (code review, 2026-09-27): the drop-down
        // listing only, which the Approve gate and the Boards badge read - not the overview, whose
        // answer walks both data trees on the server. The list on the Boards screen is brought up
        // to date when that screen is shown again.
        // ###########################################################################################
        private async Task RefreshListingAsync()
        {
            if (!this.CanAskServer)
                return;

            ReviewApiResult<BoardListingAnswer> listing = await this.thisClient!.GetBoardListingAsync(this.thisSession!);

            // Left as it was: "could not ask" must not read as "nothing needs a place".
            if (!listing.IsOk || listing.Value is null)
                return;

            this.thisListing = listing.Value;
            this.UpdateModeBadges();
            this.ReapplyApprovalGate();
        }

        // The Boards screen's board ids, top to bottom - Account's Maintainers and "Delete a board"
        // list theirs in the same order (BoardsDisplay.InBoardsListOrder).
        private IReadOnlyList<string> BoardsListOrder() => this.thisBoards.Select(board => board.BoardId).ToList();

        // `listing` null keeps the one already known (a failed read, or a test that has none).
        internal async Task ApplyBoardsListAsync(BoardOverviewAnswer answer, bool background, BoardListingAnswer? listing = null)
        {
            ArgumentNullException.ThrowIfNull(answer);

            BoardOverviewEntry? before = this.SelectedBoard;

            if (listing is not null)
                this.thisListing = listing;

            this.thisBoards.Clear();
            this.thisBoards.AddRange(BoardPlacementDisplay.InListOrder(answer.Boards, this.thisListing));
            this.thisBoardsKnown = true;

            BoardOverviewEntry? kept = before is null
                ? null
                : this.thisBoards.FirstOrDefault(board => string.Equals(board.BoardId, before.BoardId, StringComparison.Ordinal));

            if (this.FindControl<ListBox>("BoardsList") is ListBox list)
            {
                this.thisSuppressBoardsSelection = true;

                try
                {
                    List<ListBoxItem> items = this.thisBoards
                        .Select(board => TabMaintainer.BoardItem(board, BoardPlacementDisplay.ListMark(this.thisListing, board.BoardId)))
                        .ToList();
                    list.ItemsSource = items;
                    list.SelectedItem = kept is null ? null : items.FirstOrDefault(item => ReferenceEquals(item.Tag, kept));
                }
                finally
                {
                    this.thisSuppressBoardsSelection = false;
                }
            }

            this.ShowBoardsListMessage(this.thisBoards.Count == 0 ? "There are no boards yet." : null, isError: false);
            this.UpdateModeBadges();

            // Before the detail: a board just placed into BETA's lists shows as listed at once.
            this.BoardDetail.UseListing(this.thisListing);

            // A new board placed since changes whether the open submission may be approved.
            this.ReapplyApprovalGate();

            if (before is null)
                return;

            if (kept is null)
            {
                // Gone from the list while its table holds a change not sent: the change stays on
                // screen - emptying the panel would throw it away unasked.
                if (!this.BoardDetail.HasUnsavedTableEdits)
                    await this.BoardDetail.ShowBoardAsync(null);

                return;
            }

            if (!background || QueueRefreshRules.BoardChanged(before, kept))
                await this.BoardDetail.ShowBoardAsync(kept, keepShown: background);
        }

        private async void OnBoardsSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisSuppressBoardsSelection)
                return;

            BoardOverviewEntry? board = this.SelectedBoard;

            // *** LEAVING A BOARD'S TABLE WITH A CHANGE NOT SENT ASKS FIRST (2026-10-03). *** The
            // panel shows one board, so choosing another replaces the table. Cancel - or a Save the
            // server refused - puts the selection back on the board it holds.
            if (!string.Equals(board?.BoardId, this.BoardDetail.ShownBoard?.BoardId, StringComparison.Ordinal) &&
                !await this.BoardDetail.MayLeaveTableAsync())
            {
                this.ReselectShownBoard();
                return;
            }

            if (board is null)
            {
                await this.BoardDetail.ShowBoardAsync(null);
                return;
            }

            // The board the tab opens on next time nothing waits, after a restart too
            // (TabMaintainer.OpenOnEntry.cs).
            this.RememberBoard(board.BoardId);

            // Chosen by the maintainer, so waited for under the overlay (2026-09-28) - the minute
            // check re-reads a changed board without it.
            if (!await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingBoard(board.BoardId), () => this.BoardDetail.ShowBoardAsync(board)))
                this.BoardDetail.ShowMessage(WaitWording.NoAnswer, isError: true);
        }

        // Puts the list's selection back on the board the panel holds, after the maintainer chose to
        // stay in its table. Suppressed, so it does not count as a new choice.
        private void ReselectShownBoard()
        {
            if (this.FindControl<ListBox>("BoardsList") is not ListBox list || list.ItemsSource is not IEnumerable<ListBoxItem> items)
                return;

            string? shown = this.BoardDetail.ShownBoard?.BoardId;

            this.thisSuppressBoardsSelection = true;

            try
            {
                list.SelectedItem = items.FirstOrDefault(item =>
                    item.Tag is BoardOverviewEntry entry && string.Equals(entry.BoardId, shown, StringComparison.Ordinal));
            }
            finally
            {
                this.thisSuppressBoardsSelection = false;
            }
        }

        // ###########################################################################################
        // A change published from a board's table is in BETA (owner decision, 2026-10-03: "it should
        // go directly to the next queue, 'BETA > Stable'"): that list and the boards are read again
        // at once, so the BETA > Stable badge counts it and the board's line and revisions say so -
        // without leaving the Boards screen, where the maintainer may look at the next board. The
        // minute check would get there too, a minute later.
        // ###########################################################################################
        private async Task AfterBoardChangePublishedAsync()
        {
            await this.RefreshBetaAsync(background: true);
            await this.RefreshBoardsAsync(background: true);
        }

        // ###########################################################################################
        // A change from a board's table that the server made into a submission but could NOT publish
        // (something changed between its check and the publish): it is opened under Contributor
        // Submissions, where it waits and can be approved. Remembered first, so the screen opens on it
        // when nothing else is open there; chosen outright when something else is, which asks about
        // that table's unsaved changes as any click in the queue would.
        // ###########################################################################################
        internal async Task OpenSentSubmissionAsync(long submissionId)
        {
            this.RememberSubmission(submissionId);

            await this.ShowModeAsync(MaintainerMode.Review);

            if (this.thisSelectedId != submissionId &&
                this.thisQueueEntries.TryGetValue(submissionId, out QueueEntry? entry) &&
                this.FindControl<ListBox>("QueueList") is ListBox queue)
            {
                queue.SelectedItem = entry.Item;
            }
        }

        private BoardOverviewEntry? SelectedBoard =>
            (this.FindControl<ListBox>("BoardsList")?.SelectedItem as ListBoxItem)?.Tag as BoardOverviewEntry;

        internal BoardOverviewEntry? SelectedBoardForTests => this.SelectedBoard;

        internal BoardDetailView BoardDetailForTests => this.BoardDetail;

        internal IReadOnlyList<string> BoardsTextsForTests() =>
            this.FindControl<ListBox>("BoardsList")?.ItemsSource is IEnumerable<ListBoxItem> items
                ? items.SelectMany(item => item.Content is Control content ? TabMaintainer.VisibleTexts(content) : []).ToList()
                : [];

        private void ClearBoards()
        {
            this.thisBoards.Clear();
            this.thisListing = null;
            this.thisBoardsKnown = false;

            if (this.FindControl<ListBox>("BoardsList") is ListBox list)
            {
                this.thisSuppressBoardsSelection = true;
                list.ItemsSource = null;
                this.thisSuppressBoardsSelection = false;
            }

            this.ShowBoardsListMessage(null, isError: false);
            this.BoardDetail.Clear();
        }

        private void ShowBoardsListMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("BoardsListMessageText"), message, isError);

        // One board: its name, and where its data is, how many maintain it, and - when it is so -
        // that it is closed to contributions. A board waiting for a place in the drop-down lists
        // carries `mark` under it, in the queue's new-board colours.
        private static ListBoxItem BoardItem(BoardOverviewEntry board, string? mark)
        {
            var content = new StackPanel { Spacing = 1 };

            content.Children.Add(new TextBlock
            {
                Text = BoardsDisplay.Name(board),
                FontSize = 13,
                TextWrapping = TextWrapping.Wrap
            });

            // Where its data is and who maintains it, what is still to be done in bold.
            var line = new TextBlock
            {
                FontSize = 11,
                Opacity = 0.7,
                TextWrapping = TextWrapping.Wrap
            };

            TabMaintainer.ShowParts(line, BoardsDisplay.ListLineParts(board));
            content.Children.Add(line);

            // How its submissions went (2026-10-09) - nothing from a server that does not say.
            if (BoardsDisplay.SubmissionCountsLine(board.SubmissionCounts) is string counts)
            {
                content.Children.Add(new TextBlock
                {
                    Text = counts,
                    FontSize = 11,
                    Opacity = 0.7,
                    TextWrapping = TextWrapping.Wrap
                });
            }

            if (mark is not null)
            {
                var text = new TextBlock { Text = mark, FontSize = 11, TextWrapping = TextWrapping.Wrap };

                var badge = new Border
                {
                    Classes = { "PlacementMark" },
                    BorderThickness = new Avalonia.Thickness(1),
                    CornerRadius = new Avalonia.CornerRadius(3),
                    Padding = new Avalonia.Thickness(6, 1),
                    Margin = new Avalonia.Thickness(0, 3, 0, 0),
                    HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Left,
                    Child = text
                };

                badge.Bind(Border.BackgroundProperty, badge.GetResourceObservable("Badge_NewBoard_Bg"));
                badge.Bind(Border.BorderBrushProperty, badge.GetResourceObservable("Badge_NewBoard_Border"));
                text.Bind(TextBlock.ForegroundProperty, text.GetResourceObservable("Badge_NewBoard_Fg"));

                content.Children.Add(badge);
            }

            return new ListBoxItem { Tag = board, Classes = { "ListEntry" }, Content = content };
        }
    }
}
