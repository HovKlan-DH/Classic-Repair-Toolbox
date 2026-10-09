using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // "CONTRIBUTOR SUBMISSIONS" AND "BETA > PROD" OPEN ON SOMETHING (owner request, 2026-09-30:
    // "When selecting one of the button "Contributor submission" or "Beta > Prod" then it should
    // select either whatever the user viewed last time (if set) or it should show the first entry,
    // instead of the maintainer needing to press it").
    //
    // Which entry is MaintainerModes.EntryToOpen: the one looked at last while it is still listed,
    // else the first. "Last time" survives a restart - Main hands the two ids over from UserSettings
    // (UseRememberedSelections), the way it hands "Show changes only".
    //
    // WHEN: choosing either screen's button, the lists read at sign-in (or when the tab is first shown
    // on a remembered session), and coming back to the tab. Choosing an entry is the ordinary
    // selection - the same wait, the same table or plan - as if it had been clicked.
    //
    // *** NEVER OVER SOMETHING OPEN, AND NEVER OFF SCREEN. *** Only when nothing is open on that
    // screen: a submission decided elsewhere stays on screen with its table (TabMaintainer.QueueRefresh
    // .cs), and choosing another there would ask about its unsaved changes out of the blue. And never
    // from the badge's own check while the tab is not shown - that would open a submission, under
    // CRT's "please wait", while the maintainer is working on a schematic.
    //
    // *** WITH NOTHING WAITING, THE TAB OPENS ON BOARDS (owner request, 2026-10-04: "if there is no
    // queue awaiting, when opening the "Maintainer" tab, then go to "Systems" and show the last
    // selected system"). *** Opening is showing the tab, or signing in on it. The screen is decided
    // ONCE per opening, by MaintainerModes.ScreenOnOpening, as soon as it is KNOWN whether anything
    // waits (WaitingIsKnown): something waiting in either list settles it at once, "nothing" only
    // once both lists are read (BadgeKnown). Until then nothing is opened, or the queue's first
    // submission would be open before it is known that the tab goes to Boards. On Boards it opens on the board
    // looked at last (UserSettings, like the queues' entries), else the first, once that list is
    // read. Choosing a screen's tab meanwhile settles it: the maintainer chose.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private long? thisRememberedSubmissionId;
        private string? thisRememberedBetaBoardId;
        private string? thisRememberedBoardId;
        private Action<long?>? thisRememberSubmission;
        private Action<string?>? thisRememberBetaBoard;
        private Action<string?>? thisRememberBoard;

        // The tab was opened and its screen is not decided yet - see the header.
        private bool thisOpeningScreenPending;

        // The tab opened on Boards, and the board it opens on is not chosen yet - waiting for the
        // list, which the first look after a launch reads only once the screen is Boards.
        private bool thisOpeningBoardPending;

        // Called by Main with UserSettings' three ids; a tab a test builds on its own remembers only
        // for as long as it lives.
        internal void UseRememberedSelections(
            long? submissionId,
            string? betaBoardId,
            Action<long?> rememberSubmission,
            Action<string?> rememberBetaBoard,
            string? boardId = null,
            Action<string?>? rememberBoard = null)
        {
            this.thisRememberedSubmissionId = submissionId;
            this.thisRememberedBetaBoardId = betaBoardId;
            this.thisRememberedBoardId = boardId;
            this.thisRememberSubmission = rememberSubmission ?? throw new ArgumentNullException(nameof(rememberSubmission));
            this.thisRememberBetaBoard = rememberBetaBoard ?? throw new ArgumentNullException(nameof(rememberBetaBoard));
            this.thisRememberBoard = rememberBoard;
        }

        private void RememberBoard(string boardId)
        {
            this.thisRememberedBoardId = boardId;
            this.thisRememberBoard?.Invoke(boardId);
        }

        // ###########################################################################################
        // The tab is opened - shown, or signed in on: its screen is decided anew. Closed - the tab
        // detached, or a screen chosen by hand: nothing is decided for it any more.
        // ###########################################################################################
        private void BeginOpening() => this.thisOpeningScreenPending = true;

        private void EndOpening()
        {
            this.thisOpeningScreenPending = false;
            this.thisOpeningBoardPending = false;
        }

        // Whether something is open on the screen on show - what keeps a queue screen on show.
        private bool IsSomethingOpenOnShownScreen => this.thisMode switch
        {
            MaintainerMode.Review => this.thisSelectedId is not null || this.IsTableOpen,
            MaintainerMode.Beta => this.SelectedBetaRow is not null,
            _ => false
        };

        // ###########################################################################################
        // Whether it is known if anything waits for this account: something counted in a list
        // already read says yes - one list is enough for that - and "no" needs both lists read
        // (BadgeKnown). A BETA list that cannot be read then never holds back a queue that has work.
        // ###########################################################################################
        private bool WaitingIsKnown => this.TabAttention > 0 || this.BadgeKnown;

        // ###########################################################################################
        // Decides the screen of an opening, once WaitingIsKnown. Moved here without reading - the
        // Boards list is read by the check that follows an opening, or was read already; choosing
        // the board waits for it (SelectOnEntry).
        // ###########################################################################################
        private void DecideOpeningScreen()
        {
            if (!this.thisOpeningScreenPending || !this.WaitingIsKnown)
                return;

            this.thisOpeningScreenPending = false;

            MaintainerMode screen = MaintainerModes.ScreenOnOpening(this.thisMode, this.TabAttention, this.IsSomethingOpenOnShownScreen);

            if (screen != MaintainerMode.Boards)
                return;

            this.thisOpeningBoardPending = true;

            if (this.thisMode != screen)
            {
                this.thisMode = screen;
                this.ApplyModeVisibility();
            }
        }

        // The screen the tab would open on now, were it opened - for the read-ahead, which reads a
        // submission only for an opening that will show it (TabMaintainer.Prefetch.cs).
        private MaintainerMode ScreenOnNextOpening =>
            this.WaitingIsKnown
                ? MaintainerModes.ScreenOnOpening(this.thisMode, this.TabAttention, this.IsSomethingOpenOnShownScreen)
                : this.thisMode;

        private void RememberSubmission(long id)
        {
            this.thisRememberedSubmissionId = id;
            this.thisRememberSubmission?.Invoke(id);
        }

        private void RememberBetaBoard(string boardId)
        {
            this.thisRememberedBetaBoardId = boardId;
            this.thisRememberBetaBoard?.Invoke(boardId);
        }

        // ###########################################################################################
        // Chooses the entry the screen on show opens on - see the header. On an opening, the screen
        // first, and nothing at all until it is decided. Nothing on Admin, nothing on Boards except
        // for an opening, and nothing when something is already open. Signed out, the lists are empty.
        // ###########################################################################################
        private void SelectOnEntry()
        {
            this.DecideOpeningScreen();

            if (this.thisOpeningScreenPending)
                return;

            switch (this.thisMode)
            {
                // The board looked at last, else the first - once the list is there to choose from.
                // A board still on screen (gone from the list with a change not published, say) is
                // never replaced.
                case MaintainerMode.Boards when this.thisOpeningBoardPending && this.thisBoardsKnown:
                    this.thisOpeningBoardPending = false;

                    if (this.SelectedBoard is null &&
                        this.BoardDetail.ShownBoard is null &&
                        this.FindControl<ListBox>("BoardsList") is ListBox boards &&
                        boards.ItemsSource is IEnumerable<ListBoxItem> boardItems)
                    {
                        List<ListBoxItem> boardRows = boardItems.Where(item => item.Tag is BoardOverviewEntry).ToList();

                        string? boardId = MaintainerModes.EntryToOpen(
                            boardRows.Select(item => ((BoardOverviewEntry)item.Tag!).BoardId).ToList(),
                            this.thisRememberedBoardId);

                        boards.SelectedItem = boardRows.FirstOrDefault(item => ((BoardOverviewEntry)item.Tag!).BoardId == boardId);
                    }

                    break;

                // The same entry the tab reads ahead while it is away (TabMaintainer.Prefetch.cs).
                case MaintainerMode.Review when this.thisSelectedId is null && !this.IsTableOpen:
                    if (this.FindControl<ListBox>("QueueList") is ListBox queue &&
                        this.EntryToOpenOnReview() is ReviewQueueRow row &&
                        this.thisQueueEntries.TryGetValue(row.Id, out QueueEntry? entry))
                    {
                        queue.SelectedItem = entry.Item;
                    }

                    break;

                case MaintainerMode.Beta when this.SelectedBetaRow is null:
                    if (this.FindControl<ListBox>("BetaList") is ListBox beta && beta.ItemsSource is IEnumerable<ListBoxItem> items)
                    {
                        List<ListBoxItem> rows = items.Where(item => item.Tag is ProductionBoardRow).ToList();

                        string? boardId = MaintainerModes.EntryToOpen(
                            rows.Select(item => ((ProductionBoardRow)item.Tag!).BoardId).ToList(),
                            this.thisRememberedBetaBoardId);

                        beta.SelectedItem = rows.FirstOrDefault(item => ((ProductionBoardRow)item.Tag!).BoardId == boardId);
                    }

                    break;
            }
        }

        // The submissions as the queue shows them, top to bottom - under their boards' headings.
        private List<long> QueueIdsInListOrder(ListBox queue) =>
            queue.ItemsSource is IEnumerable<ListBoxItem> items
                ? items.Select(item => item.Tag).OfType<ReviewQueueRow>().Select(row => row.Id).ToList()
                : [];

        // Runs the rule as the tab does on entry - for tests.
        internal void SelectOnEntryForTests() => this.SelectOnEntry();

        // The tab opened as showing or signing in opens it, then the rule run - for tests.
        internal void OpenForTests()
        {
            this.BeginOpening();
            this.SelectOnEntry();
        }
    }
}
