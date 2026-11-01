using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;

namespace CRT
{
    // ###########################################################################################
    // THE WIDTHS DRAGGED ON THE CONTRIBUTE TAB (owner request, 2026-10-29: "configurable (user
    // dragable) splitters ... one between the 'Board' selector and the content area, and one also
    // between the thumbnails of the schematic images and the colors etc. Should of course be saved -
    // probably per board").
    //
    //   - THE LEFT COLUMN, at the splitter beside it (WorkspaceSplitter), and THE PREVIEWS ON THE
    //     SCHEMATIC IMAGES' CARDS, at the handle beside them on any card (ComponentPane.PreviewWidth).
    //   - KEPT PER BOARD, a board never dragged on taking the width dragged last (ContributeWidths) -
    //     each board shown gets its own (ShowWidths, from ShowBoard).
    //   - THE SETTINGS' through the four seams below, handed in by Initialize; a tab a test builds has
    //     none - it keeps what is dragged for the run - unless the test hands its own.
    //   - THE LEFT COLUMN GONE WHILE THE TABLE IS OPEN (ShowWorkspaceColumns, 2026-10-29), and back at
    //     its width when it closes.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // How much room the component list keeps however low the window is: its heading, search box,
        // "Add component" and a few rows.
        internal const double ComponentListKeepsAtLeast = 230;

        // The board's entries scroll only when they would leave the list less than that.
        private void OnWorkspaceNavSizeChanged(object? sender, Avalonia.Controls.SizeChangedEventArgs e) =>
            this.BoardNavScroll.MaxHeight = System.Math.Max(120, e.NewSize.Height - TabDrafts.ComponentListKeepsAtLeast);

        internal Func<string, double>? NavWidthFor { get; set; }

        internal Action<string, double>? RememberNavWidth { get; set; }

        // What a category moved on "Board: Categories" does to the saved filter (board key, from, to).
        internal Action<string, string, string>? FollowCategoryMove { get; set; }

        internal Func<string, double>? PreviewWidthFor { get; set; }

        internal Action<string, double>? RememberPreviewWidth { get; set; }

        private void WireSplitters()
        {
            // The column never dragged away to nothing, nor over the workspace.
            this.WorkspaceGrid.ColumnDefinitions[0].MinWidth = ContributeWidths.NavRange.Min;
            this.WorkspaceGrid.ColumnDefinitions[0].MaxWidth = ContributeWidths.NavRange.Max;
            this.WorkspaceGrid.ColumnDefinitions[2].MinWidth = 360;

            // A GridSplitter marks its release handled - heard all the same, as the Workbooks tab's are.
            this.WorkspaceSplitter.AddHandler(
                InputElement.PointerReleasedEvent,
                (_, _) => this.NavWidthDragged(),
                RoutingStrategies.Bubble,
                handledEventsToo: true);

            this.ComponentPane.PreviewWidthChosen += (_, width) => this.PreviewWidthDragged(width);
        }

        // "<hardware>|<board>" - the settings' board key (Main.GetCurrentBoardKey's).
        private static string BoardKeyOf(HardwareBoardEntry? entry) =>
            entry is null ? string.Empty : $"{entry.HardwareName}|{entry.BoardName}";

        internal double NavWidthForTests => this.WorkspaceGrid.ColumnDefinitions[0].Width.Value;

        // ###########################################################################################
        // The board shown gets its widths - as last dragged on it, else as dragged last on any board.
        // While the table hides the left column, the width waits to be put back (ShowWorkspaceColumns).
        // ###########################################################################################
        private void ShowWidths()
        {
            string key = TabDrafts.BoardKeyOf(this.thisCurrentEntry);

            if (this.NavWidthFor is { } nav)
            {
                if (this.thisNavHiddenForTable)
                {
                    this.thisNavWidthBeforeTable = new GridLength(nav(key));
                }
                else
                {
                    this.WorkspaceGrid.ColumnDefinitions[0].Width = new GridLength(nav(key));
                }
            }

            if (this.PreviewWidthFor is { } preview)
            {
                this.ComponentPane.PreviewWidth = preview(key);
            }
        }

        // ###########################################################################################
        // THE TABLE TAKES THE WHOLE WORKSPACE (owner request, 2026-10-29: "When clicking Table view then
        // I see no reason to still see the Board data editable selections, as then the table view has
        // taken over"): the left column and its splitter go while the table is open, their columns to
        // nothing, and come back at the width they had. Run wherever the table opens or closes
        // (ShowWorkspace, ApplyTableMode). The columns' widths change - never the columns.
        // ###########################################################################################
        private bool thisNavHiddenForTable;

        // The splitter's column - 4px, as every splitter of CRT (the Schematics tab's; owner request,
        // 2026-10-30: "bulky and visible"), drawn in the global GridSplitter style.
        internal const double SplitterWidth = 4;

        private GridLength thisNavWidthBeforeTable = new(ContributeWidths.DefaultNavWidth);

        private void ShowWorkspaceColumns()
        {
            bool hide = this.IsTableOpen;

            if (hide == this.thisNavHiddenForTable)
            {
                return;
            }

            ColumnDefinition nav = this.WorkspaceGrid.ColumnDefinitions[0];
            ColumnDefinition gap = this.WorkspaceGrid.ColumnDefinitions[1];

            if (hide)
            {
                this.thisNavWidthBeforeTable = nav.Width;
                nav.MinWidth = 0;
                nav.Width = new GridLength(0);
                gap.Width = new GridLength(0);
                this.WorkspaceContent.Margin = default;
            }
            else
            {
                nav.MinWidth = ContributeWidths.NavRange.Min;
                nav.Width = this.thisNavWidthBeforeTable;
                gap.Width = new GridLength(TabDrafts.SplitterWidth);
                this.WorkspaceContent.Margin = new Avalonia.Thickness(8, 0, 0, 0);
            }

            this.thisNavHiddenForTable = hide;
            this.WorkspaceNav.IsVisible = !hide;
            this.WorkspaceSplitter.IsVisible = !hide;
        }

        internal bool IsNavShownForTests => this.WorkspaceNav.IsVisible && this.WorkspaceGrid.ColumnDefinitions[0].Width.Value > 0;

        // The splitter let go of: the column's width kept for the board shown.
        private void NavWidthDragged()
        {
            double width = this.WorkspaceGrid.ColumnDefinitions[0].ActualWidth;

            if (width > 0)
            {
                this.RememberNavWidth?.Invoke(TabDrafts.BoardKeyOf(this.thisCurrentEntry), width);
            }
        }

        // The column dragged to `width` and let go of - for a test.
        internal void DragNavWidthForTests(double width)
        {
            this.WorkspaceGrid.ColumnDefinitions[0].Width = new GridLength(width);
            this.RememberNavWidth?.Invoke(TabDrafts.BoardKeyOf(this.thisCurrentEntry), width);
        }

        // A card's previews' handle let go of: their width kept for the board the pane is on.
        private void PreviewWidthDragged(double width) =>
            this.RememberPreviewWidth?.Invoke(TabDrafts.BoardKeyOf(this.thisComponentEntry ?? this.thisCurrentEntry), width);
    }
}
