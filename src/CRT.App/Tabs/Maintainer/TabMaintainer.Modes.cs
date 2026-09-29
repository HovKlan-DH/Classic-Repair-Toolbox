using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE FOUR SCREENS BEHIND THE BUTTONS AT THE TOP LEFT (owner request, 2026-09-27): Review,
    // BETA, Systems and Admin. Each changes the list on the left AND the panel on the right; the
    // "Production", "Maintainers" and "Unused files" windows they replace opened over the queue.
    //
    // WHAT THIS PART OWNS: which screen is shown, the buttons' badges, handing the panels the
    // session, and reading the other screens' lists alongside the queue. The screens themselves are
    // TabMaintainer.Beta.cs, .Systems.cs and .Admin.cs; what a badge counts is MaintainerModes.
    //
    // *** SWITCHING SCREEN HIDES, IT NEVER CLOSES. *** Each screen keeps its list, its selection and
    // its right-hand panel while another is shown - so an open submission's table, unsaved changes
    // and all, is exactly where it was on coming back to Review, and nothing needs asking about.
    // Signing out and closing the window still ask, as they did.
    //
    // *** THE BADGES ARE KEPT CURRENT BY THE QUEUE'S OWN CHECK. *** Every minute while the window is
    // in front (TabMaintainer.QueueRefresh.cs) the BETA list and the systems are read with the
    // queue, so the counts are right on whichever screen is shown. Choosing a screen reads its list
    // once more, so what it shows is current the moment it is chosen.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private MaintainerMode thisMode = MaintainerMode.Review;

        // From the server's queue answer - never worked out here. It shows the Admin button.
        private bool thisIsAdministrator;

        // The screen on show.
        internal MaintainerMode ShownMode => this.thisMode;

        // Each screen's button, its list on the left and its panel on the right. Admin's right-hand
        // side is whichever of its two panels is chosen - ApplyModeVisibility.
        private static readonly (MaintainerMode Mode, string Button, string List, string? Panel)[] ModeScreens =
        [
            (MaintainerMode.Review, "ReviewModeButton", "ReviewList", "ReviewPanel"),
            (MaintainerMode.Beta, "BetaModeButton", "BetaListPanel", "BetaDetailView"),
            (MaintainerMode.Systems, "SystemsModeButton", "SystemsListPanel", "SystemDetailView"),
            (MaintainerMode.Admin, "AdminModeButton", "AdminList", null)
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
        // Admin is refused to anybody the server did not say is an administrator.
        // ###########################################################################################
        internal async Task ShowModeAsync(MaintainerMode mode)
        {
            if (mode == MaintainerMode.Admin && !this.thisIsAdministrator)
                return;

            this.thisMode = mode;
            this.ApplyModeVisibility();

            switch (mode)
            {
                case MaintainerMode.Review:
                    await this.RefreshQueueAsync(background: true);
                    break;

                case MaintainerMode.Beta:
                    await this.RefreshBetaAsync(background: true);
                    break;

                case MaintainerMode.Systems:
                    await this.RefreshSystemsAsync(background: true);
                    break;

                case MaintainerMode.Admin:
                    await this.EnterAdminAsync();
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

            object? adminItem = this.FindControl<ListBox>("AdminList")?.SelectedItem;
            bool admin = this.thisMode == MaintainerMode.Admin;

            this.SetShown("UnusedFilesAdminView", admin && ReferenceEquals(adminItem, this.FindControl<ListBoxItem>("UnusedFilesItem")));
        }

        private void SetShown(string name, bool shown)
        {
            if (this.FindControl<Control>(name) is Control control)
                control.IsVisible = shown;
        }

        // ###########################################################################################
        // The server's word on whether this account is an administrator. The Admin button shows only
        // for one; an account that stops being one while on the Admin screen goes back to Review.
        // ###########################################################################################
        private void SetAdministrator(bool isAdministrator)
        {
            this.thisIsAdministrator = isAdministrator;
            this.SetShown("AdminModeButton", isAdministrator);

            // The administrator sets maintainers on the Systems screen (2026-09-27).
            this.SystemDetail.SetAdministrator(isAdministrator);

            if (!isAdministrator && this.thisMode == MaintainerMode.Admin)
            {
                this.thisMode = MaintainerMode.Review;
                this.ApplyModeVisibility();
            }
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
            int? systems = this.thisSystemsKnown ? this.thisSystems.Count : null;
            int needingPlace = SystemPlacementDisplay.NeedingPlace(this.thisListing);

            this.SetBadge("ReviewBadge", "ReviewBadgeText", MaintainerModes.AttentionBadge(review));
            this.SetBadge("BetaBadge", "BetaBadgeText", beta is int waiting ? MaintainerModes.AttentionBadge(waiting) : null);

            // Discreet (a count of all systems) - until one waits for a place this account can give
            // it, when it becomes an attention badge counting those (2026-09-27).
            this.SetBadge(
                "SystemsBadge",
                "SystemsBadgeText",
                needingPlace > 0 ? MaintainerModes.AttentionBadge(needingPlace) : MaintainerModes.CountBadge(systems));

            if (this.FindControl<Border>("SystemsBadge") is Border systemsBadge)
            {
                systemsBadge.Classes.Set("Attention", needingPlace > 0);
                systemsBadge.Classes.Set("Count", needingPlace <= 0);
            }

            this.SetTip("ReviewModeButton", MaintainerModes.Tooltip(MaintainerMode.Review, this.thisSession is null ? null : review));
            this.SetTip("BetaModeButton", MaintainerModes.Tooltip(MaintainerMode.Beta, beta));
            this.SetTip("SystemsModeButton", MaintainerModes.Tooltip(MaintainerMode.Systems, systems, needingPlace));
            this.SetTip("AdminModeButton", MaintainerModes.Tooltip(MaintainerMode.Admin, null));
        }

        private void SetBadge(string badge, string text, string? value)
        {
            if (this.FindControl<TextBlock>(text) is TextBlock block)
                block.Text = value ?? string.Empty;

            this.SetShown(badge, value is not null);
        }

        private void SetTip(string button, string tip)
        {
            if (this.FindControl<Button>(button) is Button control)
                ToolTip.SetTip(control, tip);
        }

        // The badge a button shows, or null when it shows none - for tests.
        internal string? ModeBadgeForTests(MaintainerMode mode)
        {
            string name = mode switch
            {
                MaintainerMode.Review => "ReviewBadge",
                MaintainerMode.Beta => "BetaBadge",
                MaintainerMode.Systems => "SystemsBadge",
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

            this.SystemDetail.Initialize(this.thisClient, this.thisSession);
            this.SystemDetail.AfterPlacementSaved = this.AfterPlacementSavedAsync;
            this.SystemDetail.AfterPoolChange = this.AfterPoolChangeAsync;

            this.UnusedFilesAdmin.Initialize(this.thisClient, this.thisSession);
        }

        // ###########################################################################################
        // Signed out: every screen's list and panel goes with the session, and the next sign-in - as
        // whoever it is - starts on Review with nothing of the previous account's on screen.
        // ###########################################################################################
        private void ResetScreens()
        {
            this.CloseFilesWindow();
            this.ClearBeta();
            this.ClearSystems();
            this.ClearAdmin();

            this.BetaDetail.Initialize(null, null);
            this.SystemDetail.Initialize(null, null);
            this.UnusedFilesAdmin.Initialize(null, null);

            this.thisMode = MaintainerMode.Review;
            this.SetAdministrator(false);
            this.ApplyModeVisibility();
            this.UpdateModeBadges();
        }

        // The BETA list and the systems, read alongside the queue. Background: nothing on screen is
        // reloaded that has not changed. The minute check reads the Systems overview only while its
        // screen is shown - QueueRefreshRules.ReadsSystemsOverview.
        private async Task RefreshOtherListsAsync(bool minuteCheck = false)
        {
            await this.RefreshBetaAsync(background: true);

            if (QueueRefreshRules.ReadsSystemsOverview(minuteCheck, this.thisMode, this.thisSystemsKnown))
                await this.RefreshSystemsAsync(background: true);
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

        private SystemView SystemDetail => this.FindControl<SystemView>("SystemDetailView")!;


        private UnusedFilesView UnusedFilesAdmin => this.FindControl<UnusedFilesView>("UnusedFilesAdminView")!;
    }
}
