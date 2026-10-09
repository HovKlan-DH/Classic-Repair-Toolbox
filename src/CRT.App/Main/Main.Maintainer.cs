using System.Threading.Tasks;
using Avalonia.Controls;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE MAINTAINER TAB, as the window sees it (2026-09-29: the separate CRT Maintainer application
    // became a tab in CRT - owner decision).
    //
    // WHAT THIS PART OWNS: whether the tab is shown (the Configuration tab's "Enable Maintainer tab",
    // UserSettings.EnableMaintainerTab), the window's LAYOUT while it is selected, the tab's
    // BADGE in the row of tabs (2026-09-30), and handing its SIGN-IN to the tabs that ask for an
    // email address (2026-10-01). The tab itself - sign-in, the four screens, the table, and what
    // the badge counts - is TabMaintainer.
    //
    // *** WHILE THE TAB IS SELECTED, THE SIDEBAR AND THE WORKLOG BAR GO. *** Neither has anything to
    // do with reviewing: the hardware/board/component list and the worklog bar act on the board on
    // screen, and the maintainer's screens are two panels of their own that need the width. So
    // entering the tab collapses columns 0 and 1 of RootGrid (the sidebar and its splitter) and
    // hides the worklog bar; leaving it puts back the width that was there. The banners above the
    // tabs stay - an update or a sync is app-wide.
    //
    // *** THE SIDEBAR'S WIDTH IS NEVER SAVED WHILE IT IS COLLAPSED. *** OnMainSplitterPointerReleased
    // saves LeftPanel's width, which is 0 then, and the next launch would open with no sidebar. It
    // asks IsMaintainerLayoutActive first.
    //
    // *** HIDING THE TAB NEVER SIGNS OUT OR CLOSES ITS TABLE. *** Unticking the setting only hides
    // it, like switching screen inside it - "hide, never close". Quitting CRT still asks about
    // unsaved table changes (OnWindowClosing).
    // ###########################################################################################
    public partial class Main
    {
        // The sidebar's width when the tab was entered, put back on leaving it.
        private GridLength? thisWidthBeforeMaintainerLayout;

        internal bool IsMaintainerLayoutActive => this.thisWidthBeforeMaintainerLayout is not null;

        // ###########################################################################################
        // Shows or hides the tab from the setting - the same shape as ApplyOscilloscopeTabVisibility.
        // Hidden while selected, the selection moves to the first visible tab, and that selection
        // change is what puts the layout back.
        // ###########################################################################################
        public void ApplyMaintainerTabVisibility()
        {
            if (this.MaintainerTabItem == null || this.MainTabControl == null)
                return;

            bool isEnabled = UserSettings.EnableMaintainerTab;

            // Whether it is SHOWN is MaintainerModes.TabIsShown's - "Enable Maintainer tab" narrowed
            // by "only while work is waiting" (owner request, 2026-10-01). The badge's number is
            // what "work" means, and the tab currently selected - or holding unsaved table edits -
            // is never taken away.
            bool isShown = MaintainerModes.TabIsShown(
                isEnabled,
                UserSettings.ShowMaintainerTabOnlyWhenWorkWaiting,
                this.TabMaintainer.BadgeKnown,
                this.TabMaintainer.TabAttention,
                ReferenceEquals(this.MainTabControl.SelectedItem, this.MaintainerTabItem),
                this.TabMaintainer.HasUnsavedTableEdits);

            this.MaintainerTabItem.IsVisible = isShown;

            // Neither turning the tab off nor hiding it touches its sign-in (ShareMaintainerSignIn).

            if (isEnabled)
            {
                // Ticked after the launch: the badge starts now. Before it (the constructor) this
                // does nothing - StartAsync starts it once the window is up. Started even when the
                // tab is HIDDEN for want of work: the badge check is the only thing that can
                // discover there is work, and so bring the tab back.
                if (this.thisMaintainerBadgeMayStart)
                    this.StartMaintainerBadge();

                // Turned on: does the server still serve this CRT? (Main.UpdateRequired.cs)
                this.AskApiRevisionOnceInUse();
            }

            if (!isShown)
                this.MoveSelectionOffHiddenTab(this.MaintainerTabItem);
        }

        // ###########################################################################################
        // *** THE MAINTAINER'S SIGN-IN, FOR THE TABS THAT ASK FOR AN EMAIL ADDRESS (owner request,
        // 2026-10-01: "When I am a maintainer, and I have logged in, then I want to use that email
        // address everywhere in the CRT app - e.g. for the Feedback tab or in the Draft tab"). ***
        // The Feedback tab shows the account's address, and the Drafts tab's Submit dialog shows it
        // and sends the submission with the account (ContactAddress has the rule). Raised by the tab
        // on every sign-in change (TabMaintainer.SignedInChanged).
        //
        // *** SIGNED IN IS SIGNED IN, WHATEVER THE CONFIGURATION TAB SAYS (owner request, 2026-10-02:
        // "As long as the maintainer is logged in, then use email from that, no matter what is
        // checked in "Configuration" tab. The maintainer will need to logoff to be forgotten"). ***
        // From 2026-10-01 the sign-in was withheld while the tab was turned off, and then also while
        // "Hide the Maintainer tab while no work is waiting" hid it (a code review: the Feedback
        // tab's "Sign out there" named a tab not on screen). That sent the owner's own submission
        // "without an account" while signed in - the queue had just emptied, so the tab was hidden.
        // Neither setting signs out, so neither takes the account away; signing out on the tab does
        // (turning the tab back on first, when it is off). The remembered sign-in is restored at
        // launch with the tab off too (StartMaintainerBadge).
        // ###########################################################################################
        private void ShareMaintainerSignIn()
        {
            ReviewSession? account = this.TabMaintainer.SignedIn;

            this.TabFeedback.UseMaintainerAccount(account);
            this.TabDrafts.UseMaintainerAccount(account);
        }

        // ###########################################################################################
        // THE TAB'S BADGE (owner request, 2026-09-30: "when a maintainer is logged in, and he/she
        // receives something new in queue (for him to process), then it should show as a badge in
        // the "Maintainer" tab ... It should check from server once every minute ... until it is
        // fully processed, including if it is awaiting in "BETA to PROD" queue").
        //
        // What it counts is the tab's (MaintainerModes.TabAttention); this only draws it, and says
        // whether it can be SEEN - the tab turned on, and CRT's window not minimised - which is what
        // keeps the minute check running while the tab is not on screen. A minimised window has no
        // row of tabs to show a badge in, so it asks nothing.
        //
        // *** STARTED ONCE CRT'S WINDOW IS UP, NEVER FROM THE CONSTRUCTOR *** - restoring the
        // remembered sign-in asks the server, and a slow one must not hold up the window
        // (TabMaintainer.Session.cs). StartAsync starts it; ticking the setting later starts it then.
        // ###########################################################################################
        private bool thisMaintainerBadgeMayStart;

        internal bool MaintainerTabBadgeCanBeSeen =>
            UserSettings.EnableMaintainerTab && this.WindowState != WindowState.Minimized;

        private void StartMaintainerBadge()
        {
            this.thisMaintainerBadgeMayStart = true;

            // The remembered sign-in, with the tab on or off - no request (ShareMaintainerSignIn).
            this.TabMaintainer.RestoreSignInQuietly();

            if (UserSettings.EnableMaintainerTab)
                _ = this.TabMaintainer.RestoreInBackgroundAsync();
        }

        // The count, or no badge at all for nothing waiting. No tooltip: the owner wants none on
        // these (2026-09-30, "no need to see it").
        private void ShowMaintainerTabBadge(int count)
        {
            string? text = MaintainerModes.AttentionBadge(count);

            this.MaintainerTabBadgeText.Text = text ?? string.Empty;
            this.MaintainerTabBadge.IsVisible = text is not null;

            // The badge IS the condition for "only while work is waiting", so every change to it
            // may show or hide the tab (owner request, 2026-10-01). Harmless when that setting is
            // off - TabIsShown then answers from EnableMaintainerTab alone.
            this.ApplyMaintainerTabVisibility();
        }

        // The badge as drawn, or null when none is - for tests.
        internal string? MaintainerTabBadgeForTests =>
            this.MaintainerTabBadge.IsVisible ? this.MaintainerTabBadgeText.Text : null;

        // ###########################################################################################
        // Called from OnMainTabControlSelectionChanged, which has already checked the event is the
        // tab control's own. The items ARE the TabItems, so RemovedItems says whether the tab was
        // just left - which also covers MoveSelectionOffHiddenTab's programmatic change.
        // ###########################################################################################
        private void ApplyMaintainerLayoutForSelection(SelectionChangedEventArgs e)
        {
            if (ReferenceEquals(this.MainTabControl?.SelectedItem, this.MaintainerTabItem))
                this.EnterMaintainerLayout();
            else if (e.RemovedItems.Contains(this.MaintainerTabItem))
                this.LeaveMaintainerLayout();
        }

        // Column definitions are changed IN PLACE: a new ColumnDefinitions collection leaves the
        // Grid measuring against the old one (CLAUDE.md, the Drafts tab's row heights).
        private void EnterMaintainerLayout()
        {
            if (this.IsMaintainerLayoutActive)
                return;

            this.thisWidthBeforeMaintainerLayout = this.RootGrid.ColumnDefinitions[0].Width;

            this.RootGrid.ColumnDefinitions[0].MinWidth = 0;
            this.RootGrid.ColumnDefinitions[0].Width = new GridLength(0);
            this.RootGrid.ColumnDefinitions[1].Width = new GridLength(0);

            this.LeftPanel.IsVisible = false;
            this.MainSplitter.IsVisible = false;
            this.WorklogBar.IsVisible = false;
        }

        private void LeaveMaintainerLayout()
        {
            if (this.thisWidthBeforeMaintainerLayout is not GridLength width)
                return;

            this.thisWidthBeforeMaintainerLayout = null;

            this.RootGrid.ColumnDefinitions[0].Width = width;
            this.RootGrid.ColumnDefinitions[0].MinWidth = Main.SidebarMinWidth;
            this.RootGrid.ColumnDefinitions[1].Width = new GridLength(Main.SidebarSplitterWidth);

            this.LeftPanel.IsVisible = true;
            this.MainSplitter.IsVisible = true;
            this.WorklogBar.IsVisible = UserSettings.EnableWorklog;
        }

        // The widths Main.axaml gives the sidebar column's minimum and the splitter column.
        private const double SidebarMinWidth = 80;
        private const double SidebarSplitterWidth = 4;

        // ###########################################################################################
        // Quitting CRT with unsaved changes in the Maintainer tab's table: the tab is brought forward
        // and asked, the way the Drafts tab's table is (OnWindowClosing). True when CRT may close.
        // ###########################################################################################
        private async Task<bool> ConfirmLeavingMaintainerTableAsync()
        {
            if (!this.TabMaintainer.HasUnsavedTableEdits)
                return true;

            if (this.MaintainerTabItem != null && this.MainTabControl != null)
                this.MainTabControl.SelectedItem = this.MaintainerTabItem;

            return await this.TabMaintainer.ConfirmLeavingTableAsync(this);
        }
    }
}
