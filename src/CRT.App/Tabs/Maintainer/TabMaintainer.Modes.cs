using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE FOUR SCREENS BEHIND THE BUTTONS AT THE TOP LEFT (owner request, 2026-09-27): Review,
    // BETA, Boards and Account (Admin until 2026-10-04, the administrator's alone). Each changes
    // the list on the left AND the panel on the right; the "Production", "Maintainers" and "Unused
    // files" windows they replace opened over the queue.
    //
    // WHAT THIS PART OWNS: which screen is shown, the buttons' badges, handing the panels the
    // session, and reading the other screens' lists alongside the queue. The screens themselves are
    // TabMaintainer.Beta.cs, .Boards.cs and .Account.cs; what a badge counts is MaintainerModes.
    //
    // *** SWITCHING SCREEN HIDES, IT NEVER CLOSES. *** Each screen keeps its list, its selection and
    // its right-hand panel while another is shown - so an open submission's table, unsaved changes
    // and all, is exactly where it was on coming back to Review, and nothing needs asking about.
    // Signing out and closing the window still ask, as they did.
    //
    // *** THE BADGES ARE KEPT CURRENT BY THE QUEUE'S OWN CHECK. *** Every minute while the window is
    // in front (TabMaintainer.QueueRefresh.cs) the BETA list and the boards are read with the
    // queue, so the counts are right on whichever screen is shown. Choosing a screen reads its list
    // once more, so what it shows is current the moment it is chosen. The Review and BETA counts
    // added up are the tab's own badge in CRT's row of tabs (2026-09-30), kept current off screen too.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private MaintainerMode thisMode = MaintainerMode.Review;

        // From the server's queue answer - never worked out here. It shows the administrator's own
        // entries on the Account screen (TabMaintainer.Account.cs).
        private bool thisIsAdministrator;

        // The screen on show.
        internal MaintainerMode ShownMode => this.thisMode;

        // Each screen's button, its list on the left and its panel on the right. Account's right-hand
        // side is whichever of its panels is chosen - ApplyModeVisibility.
        private static readonly (MaintainerMode Mode, string Button, string List, string? Panel)[] ModeScreens =
        [
            (MaintainerMode.Review, "ReviewModeButton", "ReviewList", "ReviewPanel"),
            (MaintainerMode.Beta, "BetaModeButton", "BetaListPanel", "BetaDetailView"),
            (MaintainerMode.Boards, "BoardsModeButton", "BoardsListPanel", "BoardDetailPane"),
            (MaintainerMode.Account, "AccountModeButton", "AccountList", null)
        ];

        private async void OnModeClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (sender is not Button button)
                return;

            foreach ((MaintainerMode mode, string name, _, _) in TabMaintainer.ModeScreens)
            {
                if (string.Equals(button.Name, name, StringComparison.Ordinal))
                {
                    await this.ShowModeAsync(mode);
                    return;
                }
            }
        }

        // ###########################################################################################
        // Shows one screen, and reads its list again - also when it is the screen already shown, so
        // pressing its button is the way to ask "anything new?" without a Refresh button of its own.
        // Every screen is every maintainer's since 2026-10-04 - Account included; only its
        // administrator's entries are not (TabMaintainer.Account.cs).
        // ###########################################################################################
        internal async Task ShowModeAsync(MaintainerMode mode)
        {
            // A screen chosen - by the maintainer, or for something to show - settles an opening
            // still undecided (TabMaintainer.OpenOnEntry.cs).
            this.EndOpening();

            this.thisMode = mode;
            this.ApplyModeVisibility();

            switch (mode)
            {
                // Both open on the entry looked at last, or the first (TabMaintainer.OpenOnEntry.cs).
                case MaintainerMode.Review:
                    await this.RefreshQueueAsync(background: true);
                    this.SelectOnEntry();
                    break;

                case MaintainerMode.Beta:
                    await this.RefreshBetaAsync(background: true);
                    this.SelectOnEntry();
                    break;

                case MaintainerMode.Boards:
                    await this.RefreshBoardsAsync(background: true);
                    break;

                case MaintainerMode.Account:
                    await this.EnterAccountAsync();
                    break;
            }
        }

        // Which list, which panel and which button look chosen - from thisMode alone.
        private void ApplyModeVisibility()
        {
            foreach ((MaintainerMode mode, string button, string list, string? panel) in TabMaintainer.ModeScreens)
            {
                bool shown = mode == this.thisMode;

                if (this.FindControl<Button>(button) is Button modeButton)
                    modeButton.Classes.Set("Selected", shown);

                this.SetShown(list, shown);

                if (panel is not null)
                    this.SetShown(panel, shown);
            }

            object? accountItem = this.FindControl<ListBox>("AccountList")?.SelectedItem;
            bool account = this.thisMode == MaintainerMode.Account;

            foreach ((string panel, string item) in TabMaintainer.AccountPanels)
                this.SetShown(panel, account && ReferenceEquals(accountItem, this.FindControl<ListBoxItem>(item)));
        }

        // The Account screen's panels on the right, each with its entry on the left.
        private static readonly (string Panel, string Item)[] AccountPanels =
        [
            ("MyAccountPanel", "MyAccountItem"),
            ("ServerVersionView", "ServerVersionItem"),
            ("MaintainerPoolAdminView", "MaintainersItem"),
            ("BoardOrderAdminView", "BoardOrderItem"),
            ("UnusedFilesAdminView", "UnusedFilesItem"),
            ("RebuildManifestsAdminView", "RebuildManifestsItem"),
            ("BoardDeletionAdminView", "DeleteBoardItem"),
            ("ApiUsageAdminView", "ApiUsageItem"),
            ("DataResetAdminView", "ResetDataItem")
        ];

        private void SetShown(string name, bool shown)
        {
            if (this.FindControl<Control>(name) is Control control)
                control.IsVisible = shown;
        }

        // ###########################################################################################
        // The server's word on whether this account is an administrator. The Account screen is every
        // maintainer's; only its administrator's entries follow this - shown for one, hidden for
        // everybody else, and an account that stops being one while such an entry is chosen goes
        // back to "My account" (TabMaintainer.Account.cs).
        // ###########################################################################################
        private void SetAdministrator(bool isAdministrator)
        {
            this.thisIsAdministrator = isAdministrator;
            this.ShowAdministratorEntries(isAdministrator);
            this.ApplyModeVisibility();
        }

        // ###########################################################################################
        // The three counts on the buttons - see MaintainerModes. The Review count follows the queue
        // entries as they stand, so an opened submission's own answer ("with the other approver")
        // is counted the way its row shows it.
        // ###########################################################################################
        private void UpdateModeBadges()
        {
            int review = MaintainerModes.ReviewAttention(this.thisQueueEntries.Values.Select(entry => entry.Row));
            int? beta = this.thisBetaKnown ? MaintainerModes.BetaAttention(this.thisBeta) : null;
            int? boards = this.thisBoardsKnown ? this.thisBoards.Count : null;
            int needingPlace = BoardPlacementDisplay.NeedingPlace(this.thisListing);

            this.SetBadge("ReviewBadge", "ReviewBadgeText", MaintainerModes.AttentionBadge(review));
            this.SetBadge("BetaBadge", "BetaBadgeText", beta is int waiting ? MaintainerModes.AttentionBadge(waiting) : null);

            // Discreet (a count of all boards) - until one waits for a place this account can give
            // it, when it becomes an attention badge counting those (2026-09-27).
            this.SetBadge(
                "BoardsBadge",
                "BoardsBadgeText",
                needingPlace > 0 ? MaintainerModes.AttentionBadge(needingPlace) : MaintainerModes.CountBadge(boards));

            if (this.FindControl<Border>("BoardsBadge") is Border boardsBadge)
            {
                boardsBadge.Classes.Set("Attention", needingPlace > 0);
                boardsBadge.Classes.Set("Count", needingPlace <= 0);
            }

            // *** NO TOOLTIPS ON THE FOUR BUTTONS (owner request, 2026-09-30: "please remove the
            // title - no need to see it"). *** They said what each screen is and what its badge
            // counts; the labels and the badges say it well enough.

            // The tab's own badge in CRT's row of tabs: the two attention badges added up.
            this.TabAttention = review + (beta ?? 0);
            this.thisShowTabBadge?.Invoke(this.TabAttention);
        }

        // ###########################################################################################
        // THE TAB'S BADGE IS MAIN'S TO DRAW (owner request, 2026-09-30) - it sits in CRT's row of
        // tabs, which this control is not part of. Main hands over how to show it, and whether it can
        // be seen at all (the tab turned on, CRT's window not minimised), which decides whether the
        // minute check runs while the tab is off screen - QueueRefreshRules.MinuteCheck.
        //
        // Called by Main, like UseRememberedChoices; a tab a test builds on its own has no badge and
        // checks nothing off screen.
        // ###########################################################################################
        internal void UseTabBadge(Func<bool> canBeSeen, Action<int> show)
        {
            this.thisTabBadgeCanBeSeen = canBeSeen ?? throw new ArgumentNullException(nameof(canBeSeen));
            this.thisShowTabBadge = show ?? throw new ArgumentNullException(nameof(show));

            show(this.TabAttention);
        }

        private Func<bool>? thisTabBadgeCanBeSeen;
        private Action<int>? thisShowTabBadge;

        // What the tab's badge counts now - MaintainerModes.TabAttention, as the buttons show it.
        internal int TabAttention { get; private set; }

        // ###########################################################################################
        // WHETHER THE BADGE'S NUMBER IS A REAL ANSWER (code review, 2026-10-01). Read by Main for
        // "Hide the Maintainer tab while no work is waiting" (MaintainerModes.TabIsShown).
        //
        // *** ZERO MEANS "NOTHING WAITS" ONLY ONCE BOTH LISTS HAVE BEEN READ. *** TabAttention starts
        // at zero, and stays at zero when the server cannot be reached - so a launch that restores
        // the remembered sign-in and then fails to read the queue would have read as "no work" and
        // hidden the tab, with the only way back being the Configuration tab. The number is trusted
        // only while somebody is signed in AND the queue and the BETA list have each been read at
        // least once since; a sign-in screen (including the 401 that empties the badge) is never
        // "no work", so the way back in stays on offer.
        // ###########################################################################################
        internal bool BadgeKnown => this.thisSession is not null && this.thisQueueKnown && this.thisBetaKnown;

        private bool thisQueueKnown;

        private void SetBadge(string badge, string text, string? value)
        {
            if (this.FindControl<TextBlock>(text) is TextBlock block)
                block.Text = value ?? string.Empty;

            this.SetShown(badge, value is not null);
        }

        // The badge a button shows, or null when it shows none - for tests.
        internal string? ModeBadgeForTests(MaintainerMode mode)
        {
            string name = mode switch
            {
                MaintainerMode.Review => "ReviewBadge",
                MaintainerMode.Beta => "BetaBadge",
                MaintainerMode.Boards => "BoardsBadge",
                _ => string.Empty
            };

            return this.FindControl<Border>(name) is { IsVisible: true } badge && badge.Child is TextBlock text
                ? text.Text
                : null;
        }

        // ###########################################################################################
        // The panels act for the account signed in here, with this window's own client - handed to
        // them on every sign-in, so a second account never inherits the first one's session.
        // ###########################################################################################
        private void InitialiseScreens()
        {
            this.BetaDetail.Initialize(this.thisClient, this.thisSession);
            this.BetaDetail.AfterChange = this.AfterBetaChangeAsync;

            this.BoardDetail.Initialize(this.thisClient, this.thisSession);
            this.BoardDetail.AfterPlacementSaved = this.AfterPlacementSavedAsync;
            this.BoardDetail.AfterSent = this.OpenSentSubmissionAsync;
            this.BoardDetail.AfterPublished = this.AfterBoardChangePublishedAsync;

            this.MyAccountDetail.Initialize(this.thisClient, this.thisSession);

            this.MaintainerPoolAdmin.Initialize(this.thisClient, this.thisSession);
            this.MaintainerPoolAdmin.AfterPoolChange = this.AfterPoolChangeAsync;
            this.MaintainerPoolAdmin.BoardsListOrder = this.BoardsListOrder;
            this.BoardOrderAdmin.Initialize(this.thisClient, this.thisSession);
            this.UnusedFilesAdmin.Initialize(this.thisClient, this.thisSession);
            this.RebuildManifestsAdmin.Initialize(this.thisClient, this.thisSession);
            this.BoardDeletionAdmin.Initialize(this.thisClient, this.thisSession);
            this.BoardDeletionAdmin.AfterDelete = this.AfterBoardDeletedAsync;
            this.BoardDeletionAdmin.BoardsListOrder = this.BoardsListOrder;
            this.ApiUsageAdmin.Initialize(this.thisClient, this.thisSession);
            this.DataResetAdmin.Initialize(this.thisClient, this.thisSession);
            this.DataResetAdmin.AfterReset = this.AfterBoardDeletedAsync;
        }

        // ###########################################################################################
        // Signed out: every screen's list and panel goes with the session - "My account" with every
        // box emptied - and the next sign-in, as whoever it is, starts on Review with nothing of the
        // previous account's on screen, and the administrator's entries hidden until the server says
        // otherwise.
        // ###########################################################################################
        private void ResetScreens()
        {
            this.ResetSubmissionViews();
            this.ClearBeta();
            this.ClearBoards();
            this.ClearAccountScreen();

            this.BetaDetail.Initialize(null, null);
            this.BoardDetail.Initialize(null, null);
            this.MyAccountDetail.Initialize(null, null);
            this.MaintainerPoolAdmin.Initialize(null, null);
            this.BoardOrderAdmin.Initialize(null, null);
            this.UnusedFilesAdmin.Initialize(null, null);
            this.RebuildManifestsAdmin.Initialize(null, null);
            this.BoardDeletionAdmin.Initialize(null, null);
            this.ApiUsageAdmin.Initialize(null, null);
            this.DataResetAdmin.Initialize(null, null);

            this.thisMode = MaintainerMode.Review;
            this.SetAdministrator(false);
            this.ApplyModeVisibility();
            this.UpdateModeBadges();
        }

        // The BETA list and the boards, read alongside the queue. Background: nothing on screen is
        // reloaded that has not changed. The minute check reads the Boards overview only while its
        // screen is shown - QueueRefreshRules.ReadsBoardsOverview.
        private async Task RefreshOtherListsAsync(bool minuteCheck = false)
        {
            await this.RefreshBetaAsync(background: true);

            if (QueueRefreshRules.ReadsBoardsOverview(minuteCheck, this.thisMode, this.thisBoardsKnown))
                await this.RefreshBoardsAsync(background: true);
            else
                await this.RefreshListingAsync();
        }

        // Whether a request can be made now - signed in, with a session that has not expired. The
        // queue's own check is what notices an expiry and says so; the other lists stay quiet.
        private bool CanAskServer =>
            this.thisClient is not null &&
            this.thisSession is not null &&
            this.thisSession.IsUsableAt(DateTimeOffset.UtcNow);

        private BetaView BetaDetail => this.FindControl<BetaView>("BetaDetailView")!;

        private BoardDetailView BoardDetail => this.FindControl<BoardDetailView>("BoardDetailPane")!;


        private MaintainerPoolView MaintainerPoolAdmin => this.FindControl<MaintainerPoolView>("MaintainerPoolAdminView")!;

        private BoardOrderView BoardOrderAdmin => this.FindControl<BoardOrderView>("BoardOrderAdminView")!;

        private UnusedFilesView UnusedFilesAdmin => this.FindControl<UnusedFilesView>("UnusedFilesAdminView")!;

        private RebuildManifestsView RebuildManifestsAdmin => this.FindControl<RebuildManifestsView>("RebuildManifestsAdminView")!;

        private BoardDeletionView BoardDeletionAdmin => this.FindControl<BoardDeletionView>("BoardDeletionAdminView")!;

        private ApiUsageView ApiUsageAdmin => this.FindControl<ApiUsageView>("ApiUsageAdminView")!;

        private DataResetView DataResetAdmin => this.FindControl<DataResetView>("DataResetAdminView")!;
    }
}
