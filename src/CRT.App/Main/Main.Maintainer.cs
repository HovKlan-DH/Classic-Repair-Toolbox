using System.Threading.Tasks;
using Avalonia.Controls;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // THE MAINTAINER TAB, as the window sees it (2026-09-29: the separate CRT Maintainer application
    // became a tab in CRT - owner decision, Assets/MaintainerTabMergePlan.md).
    //
    // WHAT THIS PART OWNS: whether the tab is shown (the Configuration tab's "Enable Maintainer tab",
    // UserSettings.EnableMaintainerTab), and the window's LAYOUT while it is selected. The tab
    // itself - sign-in, the four screens, the table - is TabMaintainer.
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
            this.MaintainerTabItem.IsVisible = isEnabled;

            if (isEnabled)
                return;

            this.MoveSelectionOffHiddenTab(this.MaintainerTabItem);
        }

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
