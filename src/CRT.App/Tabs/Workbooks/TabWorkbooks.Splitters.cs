using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;

namespace CRT
{
    // ###########################################################################################
    // THE TAB'S TWO SPLITTERS - the workbook list's and the entry list's widths, put back from the settings and saved as
    // a drag ends. In a part of its own since 2026-11-09 (code review: TabWorkbooks.axaml.cs had grown past 1,500 lines).
    //
    // Part of the TabWorkbooks partial class - see TabWorkbooks.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabWorkbooks
    {
        // ###########################################################################################
        // Subscribes both splitters' PointerReleased so a finished drag is saved.
        //
        // AddHandler with handledEventsToo: true, NOT a PointerReleased="..." attribute in the
        // markup: GridSplitter marks the event handled as it completes its own drag, so an ordinary
        // subscription never runs and the width is silently never saved. Main's own
        // OnMainSplitterPointerReleased and TabSchematics.ApplySchematicsSplitterRatio wire theirs
        // the same way, with the same comment - this tab had the markup form and neither of its two
        // widths was ever written.
        // ###########################################################################################
        private void WireSplitterPersistence()
        {
            this.OuterSplitter.AddHandler(
                InputElement.PointerReleasedEvent,
                this.OnOuterSplitterPointerReleased,
                RoutingStrategies.Bubble,
                handledEventsToo: true);

            this.BoardEntrySplitter.AddHandler(
                InputElement.PointerReleasedEvent,
                this.OnBoardEntrySplitterPointerReleased,
                RoutingStrategies.Bubble,
                handledEventsToo: true);
        }

        // Exposed to the test project so a splitter-persistence test can wire the handlers without
        // constructing a real Main (Initialize's normal caller) - the same idea as
        // ApplySplitterWidthsForTests.
        internal void WireSplitterPersistenceForTests() => this.WireSplitterPersistence();

        // ###########################################################################################
        // Restores both of this tab's splitters from UserSettings, so they do not flash back to
        // their design-time default width every time the app opens - the same pattern
        // TabSchematics.ApplySchematicsSplitterRatio and Main's own LeftPanelWidth follow. Unlike
        // the Schematics tab's splitter, these are plain pixel widths rather than a per-board ratio:
        // this tab's layout does not depend on which board is selected, so one app-wide value per
        // splitter is enough.
        // ###########################################################################################
        private void ApplySplitterWidths()
        {
            // Shut while the workbook list stands in the left panel (PlaceWorkbookList).
            if (!this.thisWorkbooksInLeftPanel)
            {
                this.OuterSplitGrid.ColumnDefinitions[0].Width = new GridLength(ClampPanelWidth(UserSettings.WorkbooksLeftPanelWidth));
            }

            this.BoardEntrySplitGrid.ColumnDefinitions[2].Width = new GridLength(ClampPanelWidth(UserSettings.WorkbooksEntryListWidth));
        }

        // ###########################################################################################
        // Keeps a restored panel width usable on the screen it is being restored ONTO.
        //
        // The saved value is a raw pixel width from whatever monitor the drag happened on, and this
        // runs once from Initialize, before the window has a meaningful size to compare against. A
        // width saved on a large monitor was applied verbatim on a small one: a 1100px panel on a
        // 1366px screen squeezes everything else to near-zero and puts the splitter off-screen, at
        // which point the only way back is hand-editing settings.json.
        //
        // Clamped to a fixed ceiling rather than to a fraction of the window because the window is
        // not laid out yet at this point. The ceiling is generous - it only catches values that are
        // unusable on any ordinary screen - and the floor keeps a panel dragged shut from restoring
        // as an invisible sliver with no splitter to grab.
        //
        // Matches the ColumnDefinition MinWidth of the two columns this clamp actually applies to -
        // OuterSplitGrid's left list column and BoardEntrySplitGrid's entry list column, both 200px
        // - on purpose: that MinWidth is the LIVE floor GridSplitter enforces while dragging, and a
        // lower value here would restore a width Avalonia immediately overrides back up to that
        // floor anyway, making this constant lie about what actually gets applied. OuterSplitGrid's
        // RIGHT column (the top-line/board/entry-list side) has its own higher 400px MinWidth
        // instead - see that column's own comment - but nothing here ever writes its width, since
        // it is the `*` column rather than a persisted one, so this constant does not need to know
        // about it.
        // ###########################################################################################
        private const double MinimumPanelWidth = 200.0;

        private const double MaximumPanelWidth = 900.0;

        private static double ClampPanelWidth(double width) =>
            double.IsNaN(width) || width < MinimumPanelWidth
                ? MinimumPanelWidth
                : Math.Min(width, MaximumPanelWidth);

        // Exposed to the test project so a splitter-persistence test can call this without
        // constructing a real Main (Initialize's normal caller) - the same idea as
        // SelectWorkbookForTests and SelectSchematicForTests.
        internal void ApplySplitterWidthsForTests() => this.ApplySplitterWidths();

        // ###########################################################################################
        // Saves the left workbook-list panel's width after the outer splitter is dragged. Deferred
        // via Post, matching Main's own OnLeftSplitterPointerReleased/TabSchematics's
        // OnSchematicsSplitterPointerReleased: the drag has not yet been applied to Bounds at the
        // moment PointerReleased fires, so reading it synchronously would save the width from before
        // this drag.
        // ###########################################################################################
        private void OnOuterSplitterPointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            Dispatcher.UIThread.Post(() => UserSettings.WorkbooksLeftPanelWidth = this.WorkbookListBorder.Bounds.Width);
        }

        // Same idea as OnOuterSplitterPointerReleased, for the board-pane/entry-list splitter.
        private void OnBoardEntrySplitterPointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            Dispatcher.UIThread.Post(() => UserSettings.WorkbooksEntryListWidth = this.SelectedSchematicEntriesBorder.Bounds.Width);
        }
    }
}
