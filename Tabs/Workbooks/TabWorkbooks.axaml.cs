using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE WORKBOOKS TAB - concept "C; Worklog tab" from the mockup, now fully functional.
    //
    // FILE MAP - this class is split across:
    //   TabWorkbooks.axaml.cs        (this file) construction, Main wiring, the splitters and their
    //                                persistence, the left-hand workbook list and its cards, and
    //                                the top-line above the board/entry split
    //   TabWorkbooks.BoardPreviews.cs the board pane and the entry list beside it - schematic
    //                                previews, entry badges, schematic selection, entry detail
    //                                cards, and the editor a badge opens
    //
    // What IS real:
    //   - the left panel, listing the workbooks WorklogManager holds for the selected board,
    //     rebuilt by RefreshWorkbooks;
    //   - clicking a card ACTIVATES that workbook app-wide (see SelectWorkbook): the choice is
    //     persisted to UserSettings.ActiveWorkbookIdByBoard and the worklog bar, "Show worklogs"
    //     and "Add worklog" all follow it. The tab does not switch away - the top-line and the
    //     board pane update in place;
    //   - the board pane (TabWorkbooks.BoardPreviews.cs): every schematic image with entries in the
    //     ACTIVE workbook, each entry drawn as the Schematics tab draws it - a dashed bounds
    //     rectangle plus an anchored badge where "show marked area" is on, a corner-parked badge
    //     where it is not;
    //   - clicking a badge opens the SAME WorklogEntryEditorWindow the Schematics tab opens,
    //     component-scope checklist included;
    //   - clicking elsewhere on a preview selects that schematic and switches the entry list on the
    //     right to it, one detail card per entry;
    //   - each of those detail cards carries its own "Delete worklog" button in its top-right
    //     corner (BuildDeleteWorklogButton/OnDeleteWorklogClick, in TabWorkbooks.BoardPreviews.cs),
    //     the per-worklog twin of "Delete workbook" below - confirmed via DeleteWorklogWindow,
    //     which cancels on Enter for the same reason that one does, then WorklogManager.DeleteEntry
    //     removes the entry's row AND its worklog_{id} attachment folder;
    //   - the top-line's second line shows the selected workbook's Note (the whole line collapsed
    //     when blank), and, right-aligned against both lines, "Edit workbook"/"Delete workbook"
    //     actions for it (OnEditWorkbookClick/OnDeleteWorkbookClick, below). Edit reopens
    //     CreateWorkbookWindow via InitializeForEdit - the SAME modal "Create new workbook" uses,
    //     pre-filled, its submit button relabelled "Update workbook" - so title/note editing has
    //     exactly one implementation rather than a second dialog to keep in sync. Delete confirms
    //     via DeleteWorkbookWindow (no minimize button, and Enter cancels rather than confirms -
    //     deleting is a click on the button, never a reflexive keypress), then
    //     WorklogManager.DeleteWorkbook removes the workbook's whole folder; the refresh that
    //     follows lands on the board's next workbook automatically, via the same
    //     ResolveActiveWorkbook stale-id fallback that already handles a workbook deleted by hand.
    //
    //   - the "Find a previous repair" box filters this whole tab as you type
    //     (OnFindRepairTextChanged -> RefreshWorkbooks): the workbook list, the board pane and the
    //     entry list all narrow to what matched, and the matched runs are highlighted wherever they
    //     are drawn. The query grammar and the fields it searches are Handlers/Data's own
    //     WorklogSearchQuery and WorklogSearchIndex - see their headers. RefreshWorkbooks re-reads
    //     the box itself rather than trusting a cached copy, so the filter survives every other
    //     refresh trigger too.
    //
    // Main.RefreshWorklogBar calls RefreshWorkbooks, so the list, the selected card, the top-line
    // and the board pane all follow a board change and any workbook edit through the one funnel the
    // rest of the worklog feature already refreshes through. The board pane has a second, narrower
    // entry point for when the component highlight cache changes - see
    // RefreshBoardPreviewsForCurrentSelection.
    //
    // Visibility is not decided here: Main.axaml.cs's ApplyWorklogBarVisibility shows and hides the
    // tab along with the worklog bar, both driven by "Enable Worklog" in Configuration.
    // ###########################################################################################
    public partial class TabWorkbooks : UserControl
    {
        // Set by Main during startup wiring, the same way TabSchematics gets its reference. The
        // board key lives on the main window (it is the hardware/board combo selection), so the
        // tab cannot resolve which board it is showing without this.
        public Main? MainWindow { get; set; }

        // The board key normally comes from the main window's hardware/board selection, which no
        // test constructs. This is the seam that lets the headless tests exercise the list itself
        // without standing up Main and its combo boxes - the same idea as UserSettings.LoadFrom
        // and DataManager.LoadFrom. Null in the running app, always.
        internal string? BoardKeyOverrideForTests { get; set; }

        // Same idea, for the board pane (TabWorkbooks.BoardPreviews.cs): it normally reads
        // Main.CurrentBoardData for the schematic image list, which again needs no test to stand up
        // Main itself. Null in the running app, always.
        internal BoardData? CurrentBoardDataOverrideForTests { get; set; }

        // What RefreshBoardPreviews actually reads - the override when a test set one, otherwise
        // the real main window's board data.
        private BoardData? CurrentBoardDataForPreviews =>
            this.CurrentBoardDataOverrideForTests ?? this.MainWindow?.CurrentBoardData;

        // Same idea again, for BuildWorklogEntryComponentScope (TabWorkbooks.BoardPreviews.cs): it
        // normally reads MainWindow.TabSchematicsControl.highlightRectsBySchematicAndLabel, which
        // again needs no test to stand up Main and a whole TabSchematics. Null in the running app,
        // always.
        //
        // The setter REJECTS a dictionary that is not OrdinalIgnoreCase-keyed. The real cache always
        // is - Main builds it that way at every one of its write sites, and TabSchematics declares
        // the field that way - so a seam that accepted a plain `new Dictionary<...>()` would let a
        // future test encode Ordinal lookup as the contract and pass, while the app, reading the
        // real cache, behaved the opposite way. A test seam that can certify behaviour opposite to
        // production's is worse than no seam.
        private Dictionary<string, Dictionary<string, List<Rect>>>? thisHighlightRectsOverrideForTests;

        internal Dictionary<string, Dictionary<string, List<Rect>>>? HighlightRectsBySchematicAndLabelOverrideForTests
        {
            get => this.thisHighlightRectsOverrideForTests;
            set
            {
                if (value != null && !ReferenceEquals(value.Comparer, StringComparer.OrdinalIgnoreCase))
                {
                    throw new ArgumentException(
                        "The highlight-rect cache is keyed OrdinalIgnoreCase in the running app; a test " +
                        "override with any other comparer would pin the opposite behaviour. Build it with " +
                        "new Dictionary<string, Dictionary<string, List<Rect>>>(StringComparer.OrdinalIgnoreCase).",
                        nameof(value));
                }

                this.thisHighlightRectsOverrideForTests = value;
            }
        }

        // What BuildWorklogEntryComponentScope actually reads - the override when a test set one,
        // otherwise the real main window's TabSchematics cache.
        private Dictionary<string, Dictionary<string, List<Rect>>>? HighlightRectsBySchematicAndLabelForPreviews =>
            this.thisHighlightRectsOverrideForTests ?? this.MainWindow?.TabSchematicsControl?.highlightRectsBySchematicAndLabel;

        // The workbook currently shown in the top-line, board and entry list. -1 means none -
        // either the board has no workbooks, or nothing has been clicked yet.
        //
        // Kept across a RefreshWorkbooks rebuild (see the end of that method): a board edit or a
        // theme change must not silently drop the user's selection back to "nothing chosen".
        private int thisSelectedWorkbookId = -1;

        // True when the list is showing workbook cards but none of them belongs to the currently
        // loaded board, so nothing can be shown on the right-hand side. Only reachable in AllBoards
        // scope; set by RefreshWorkbooks and read by RefreshBoardPreviews for its empty state.
        private bool thisHasWorkbooksOnOtherBoardsOnly;

        // ###########################################################################################
        // Whether the left-hand list shows every board's workbooks or only the currently loaded
        // board's - the Configuration tab's radio group below "Enable Worklog"
        // (UserSettings.WorkbooksScope). Read fresh on every RefreshWorkbooks rather than cached, so
        // flipping the setting takes effect on the very next refresh without this tab needing its own
        // change notification.
        // ###########################################################################################
        private static bool IsAllBoardsScope =>
            string.Equals(UserSettings.WorkbooksScope, "AllBoards", StringComparison.Ordinal);

        // The parsed "Find a previous repair" query. Parsed ONCE per keystroke here rather than per
        // record inside the filter loop - a board's worth of workbooks times their entries is a lot
        // of re-parsing of a string that has not changed.
        //
        // An empty query (the normal state) matches everything, so the unfiltered tab costs one
        // WorklogSearchQuery.IsEmpty check per record and nothing more.
        private WorklogSearchQuery thisSearchQuery = WorklogSearchQuery.Parse(null);

        // Set by RefreshWorkbooks for the workbooks that survived the filter, so the board pane and
        // the entry list can narrow themselves to the SAME matched entries rather than each
        // re-deciding what matched. Empty (not null) means "no filter is active" - see
        // MatchedEntryIdsForWorkbook.
        private readonly Dictionary<int, HashSet<int>> thisMatchedEntryIdsByWorkbookId = new();

        // ###########################################################################################
        // Coalesces keystrokes in the search box into one rebuild.
        //
        // Filtering costs a GetEntries per workbook - File.ReadAllText + Deserialize + a per-entry
        // normalise loop, uncached - plus a full board-pane rebuild, all synchronously on the UI
        // thread. Rebuilding per keystroke made typing one word on a board with a few dozen
        // workbooks hundreds of file reads, and the class header already calls reading entries twice
        // per pass "the single most expensive thing this tab did".
        //
        // 200ms: below the ~250ms at which a filter starts to feel detached from typing, and long
        // enough that an ordinary typing burst collapses into a single pass.
        // ###########################################################################################
        private DispatcherTimer? thisSearchDebounceTimer;

        private static readonly TimeSpan SearchDebounceInterval = TimeSpan.FromMilliseconds(200);

        // One pass's worth of entry reads, keyed by workbook id, so the filter and the board pane
        // that follows it do not each re-read the same entries.json. Cleared at the start of every
        // RefreshWorkbooks - it is a within-pass cache, NOT a cache of what is on disk, because an
        // entry save must be visible on the very next refresh.
        private readonly Dictionary<int, List<WorklogEntryRecord>> thisEntriesReadThisPass = new();

        public TabWorkbooks()
        {
            this.InitializeComponent();

            // The top-line's padlock needs the same ascent-overflow reservation the code-built cards
            // get, and markup cannot compute it - see the comment beside the control. The cards' own
            // glyphs are handled in BuildWorkbookCard.
            //
            // Null-guarded like every other named-control access in this class. This constructor runs
            // from Main's own InitializeComponent, so an unguarded dereference here does not degrade
            // the tab - it throws out of Main's constructor and the WHOLE MAIN WINDOW fails to
            // construct at startup. It also depended on the markup hardcoding a Text on that glyph,
            // a correctness-by-coincidence with no compile-time link.
            if (this.WorkbookHeaderStatusGlyph != null)
            {
                this.WorkbookHeaderStatusGlyph.Padding = Handlers.Geometry.FontAwesomeGlyphMetrics
                    .GetTopOverflowThicknessForText(this.WorkbookHeaderStatusGlyph.Text, this.WorkbookHeaderStatusGlyph.FontSize);
            }
        }

        // ###########################################################################################
        // Releases the full-resolution schematic bitmaps the board pane decoded - see
        // thisSchematicBitmapsByPath in TabWorkbooks.BoardPreviews.cs for why they are held for a
        // whole attachment rather than freed on each rebuild. Without this the last set outlives the
        // tab, exactly as WorklogEntryEditorWindow's own Closed handler documents for its thumbnails.
        //
        // DETACH IS NOT "THE TAB IS GOING AWAY". A TabControl detaches the previous tab's content
        // from the visual tree on every tab SWITCH, and this tab's Image controls keep their Source
        // pointing at these bitmaps while detached. Disposing without clearing the pane therefore
        // left every preview holding a dead Skia surface, and the next render pass over them threw
        // ObjectDisposedException on the RENDER thread - fatal in Avalonia, and reported as a crash
        // on switching away from Workbooks.
        //
        // So the pane is torn down FIRST and the bitmaps disposed second. Nothing then references a
        // disposed bitmap, and OnAttachedToVisualTree rebuilds the pane when the tab comes back -
        // the decode cost is the same one the first build already pays, and it only lands on a tab
        // the user has actually returned to.
        // ###########################################################################################
        protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnDetachedFromVisualTree(e);

            this.ClearBoardPreviewsBeforeDisposingBitmaps();
            this.DisposeSchematicBitmaps();
        }

        // ###########################################################################################
        // Rebuilds what OnDetachedFromVisualTree tore down, when the tab is selected again.
        //
        // RefreshWorkbooks is the whole-tab funnel (list, top-line and board pane), which is what a
        // returning tab needs: the board or the active workbook may have changed while it was away -
        // Main.RefreshWorklogBar's own calls reach this tab only while it is attached.
        //
        // COST, and why it is not conditional. This re-reads every workbook from disk and re-decodes
        // every schematic the active workbook has entries on, synchronously, on each return to the
        // tab. That is not incidental waste that a "did anything change?" guard could skip: detach
        // tore the pane down and disposed the bitmaps outright (it has to - see
        // OnDetachedFromVisualTree), so there is nothing left to reuse and the pane genuinely has to
        // be rebuilt from scratch whether or not anything changed.
        //
        // Making it cheaper means not disposing on detach, which is exactly the crash that pairing
        // was written to fix, or holding the workbooks in a cache that an out-of-app edit could make
        // stale. Neither is worth trading a correct tab for, so the cost is accepted deliberately
        // rather than overlooked.
        // ###########################################################################################
        protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
        {
            base.OnAttachedToVisualTree(e);

            // Guarded: this also fires during the very first attach, before Main has finished
            // wiring the tab up (Initialize supplies the board-key source), and RefreshWorkbooks
            // reads the board key. A null MainWindow with no test override means there is nothing
            // to rebuild yet, and the first real refresh arrives from Main a moment later.
            if (this.MainWindow != null || this.BoardKeyOverrideForTests != null)
            {
                this.RefreshWorkbooks();
            }
        }

        // ###########################################################################################
        // Hands the tab its main-window reference, matching TabSchematics/TabOverview/TabContribute.
        // The board key is the hardware/board combo selection, which lives on the main window, so
        // without this the tab cannot tell which board's workbooks to list.
        // ###########################################################################################
        public void Initialize(Main mainWindow)
        {
            this.MainWindow = mainWindow;
            this.thisActivateWorkbook = mainWindow.ActivateWorkbookAcrossBoards;
            this.ApplySplitterWidths();
            this.WireSplitterPersistence();
        }

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
        // Puts cursor focus in "Find a previous repair" - called by Main whenever this tab becomes
        // the selected one, so a user landing here can start typing straight away rather than having
        // to click the box first.
        //
        // This has to coexist with Main's own global "steal focus into ComponentSearchTextBox on any
        // click" handler, which normally treats a click anywhere as a reason to refocus that box.
        // Main excludes this tab from that handler by header the same way it already excludes
        // "Feedback"/"Configuration", so the two do not fight over the caret - see
        // OnMainPointerReleasedStealFocus's tab-header check. The user clicking anything ON this tab
        // (a workbook card, a pill, an entry) still moves focus normally; only the GLOBAL steal is
        // suppressed here, which is exactly what "clicking on something gives Filter components
        // priority again" needs once the user has left this tab.
        //
        // Posted at Background priority: called from the tab's SelectionChanged, which fires before
        // the tab's content has necessarily finished laying out on a first switch, and Focus() on a
        // control not yet part of a realized visual tree is a silent no-op.
        // ###########################################################################################
        public void FocusSearchBox()
        {
            Dispatcher.UIThread.Post(() => this.FindRepairTextBox?.Focus(), DispatcherPriority.Background);
        }

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
            this.OuterSplitGrid.ColumnDefinitions[0].Width = new GridLength(ClampPanelWidth(UserSettings.WorkbooksLeftPanelWidth));
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

        // ###########################################################################################
        // Rebuilds the left-hand workbook list for the currently selected board.
        //
        // Called from Main.RefreshWorklogBar rather than being wired to anything here: that method
        // is already the single place the app refreshes worklog state from (board changes, entry
        // saves, workbook creation and closure all funnel through it), so hanging this off it means
        // the list cannot go stale in a case the bar handles and this forgot.
        //
        // Rebuilding the whole panel each time, rather than diffing it, is deliberate - the list is
        // a handful of cards, and a rebuild cannot leave a stale card behind.
        //
        // Takes the board's workbooks when the caller has just read them (Main.RefreshWorklogBar has,
        // to resolve which one is active) and reads them itself otherwise. GetWorkbooksForBoard goes
        // through ReadAllWorkbooks, which enumerates EVERY workbook folder on disk and does a
        // File.Exists + ReadAllText + Deserialize per folder, uncached, with the board filter applied
        // only afterwards - and there is no delete feature, so that set only ever grows. Reading it
        // twice per refresh pass was the single most expensive thing this tab did.
        // ###########################################################################################
        // ###########################################################################################
        // Clears the "Find a previous repair" box, for a BOARD CHANGE specifically - called from
        // Main.OnBoardSelectionChanged before it refreshes.
        //
        // A board switch is a change of subject, not a refresh of the same view: carrying the query
        // over lands the user on a filtered (often empty) list for a board they just chose, reading
        // as "this board has nothing" while the reason sits in a text box at the top of the panel
        // they are not looking at. Main.OnHardwareSelectionChanged already clears
        // ComponentSearchTextBox for exactly this reason.
        //
        // Every OTHER refresh trigger - an entry save, a workbook create/delete - deliberately keeps
        // the query, since those are refreshes of the view the user is already looking at.
        //
        // Sets the field as well as the box because the box's TextChanged is debounced: without it,
        // the refresh that follows immediately would still parse the OLD text.
        // ###########################################################################################
        public void ClearSearchForBoardChange()
        {
            if (this.FindRepairTextBox == null)
                return;

            this.thisSearchDebounceTimer?.Stop();
            this.FindRepairTextBox.Text = string.Empty;
            this.thisSearchQuery = WorklogSearchQuery.Parse(null);
        }

        public void RefreshWorkbooks(List<WorkbookRecord>? boardWorkbooks = null)
        {
            if (this.WorkbookListPanel == null)
                return;

            // Re-read from the box rather than trusting the copy OnFindRepairTextChanged cached.
            // Every refresh path lands here - an entry save, a workbook create or delete - and each
            // of those must keep showing the filtered view rather than silently reverting to the
            // unfiltered one because it did not come through the (debounced) text handler. A BOARD
            // change is the one case that deliberately drops the query, via ClearSearchForBoardChange
            // above, before this runs.
            this.thisSearchQuery = WorklogSearchQuery.Parse(this.FindRepairTextBox?.Text);

            // Fresh per pass - this caches reads WITHIN one rebuild, never across them, so an entry
            // saved between refreshes is picked up by the next one. See the field's own comment.
            this.thisEntriesReadThisPass.Clear();

            string boardKey = this.BoardKeyOverrideForTests ?? this.MainWindow?.GetCurrentBoardKey() ?? string.Empty;

            // "Show all workbooks" (UserSettings.WorkbooksScope == "AllBoards") lists every workbook
            // on every board, not just this one - see the Configuration tab's radio group below
            // "Enable Worklog". boardWorkbooks, when supplied (Main.RefreshWorklogBar's normal call),
            // is already scope-aware - the full unfiltered set in AllBoards scope, this board's own
            // otherwise - so it is used as-is; the fallback here only matters for callers that read
            // straight off disk (e.g. the search debounce timer, OnAttachedToVisualTree).
            var workbooks = boardWorkbooks ?? (IsAllBoardsScope
                ? WorklogManager.GetAllWorkbooks()
                : WorklogManager.GetWorkbooksForBoard(boardKey));

            this.WorkbookListPanel.Children.Clear();

            // Which card is highlighted is decided by the SAME WorklogManager.ResolveActiveWorkbook
            // the worklog bar, "Show worklogs" and "Add worklog" use (via
            // Main.ResolveActiveWorkbookForBoard), reading the same saved
            // UserSettings.ActiveWorkbookIdByBoard - so this panel's highlighted card cannot
            // disagree with what the rest of the app is acting on for this board. There is no
            // in-memory tier: SelectWorkbook persists and lets the refresh that follows re-derive
            // the selection from that one saved id, whether the caller is Main or a test.
            //
            // NOTE that when nothing is saved this HIGHLIGHTS the newest workbook without saving
            // it. The highlight is therefore an assumption, not an activation - which is why
            // SelectWorkbook has no "already selected, nothing to do" guard; see its header.
            //
            // Resolved from the UNFILTERED (by search) set, deliberately: the search box narrows
            // what is SHOWN, it does not change which workbook the rest of the app is acting on.
            // Typing in it must never silently re-activate a different workbook, or "Add worklog"
            // would start writing into whichever workbook happened to survive the filter.
            //
            // Resolved from THIS BOARD's own workbooks specifically, even in AllBoards scope where
            // "workbooks" holds every board's - UserSettings.GetActiveWorkbookId(boardKey) is itself
            // board-scoped (nothing ever saves a cross-board id under it), and ResolveActiveWorkbook
            // falls back to "workbooks[0]" when nothing is saved, which in the unfiltered list would
            // be whichever board's workbook happens to be globally newest rather than this board's -
            // exactly the mismatch the board pane cannot render (see CurrentBoardDataForPreviews).
            // Same board-filtered input Main.RefreshWorklogBar's own activeWorkbook uses.
            var workbooksForThisBoard = IsAllBoardsScope
                ? workbooks.Where(w => string.Equals(w.BoardKey, boardKey, StringComparison.Ordinal)).ToList()
                : workbooks;
            var active = WorklogManager.ResolveActiveWorkbook(workbooksForThisBoard, UserSettings.GetActiveWorkbookId(boardKey));

            var shownWorkbooks = this.ApplySearchFilter(workbooks);

            // Which workbook this TAB shows in its top-line, board pane and entry list. Normally the
            // active one - but a search that filters the active workbook out has to move it, because
            // everything on the right-hand side belongs to whatever this names: leaving it on the
            // filtered-out workbook drew a top-line (with live Edit and DELETE buttons) for a
            // workbook that was not in the list, above a board pane that had gone blank because none
            // of its entries matched. Deleting from that state destroys a workbook the user cannot
            // see, and the blank pane reads as the workbook having lost its contents.
            //
            // This does NOT re-activate anything: ActiveWorkbookIdByBoard is untouched, so the bar,
            // "Show worklogs" and "Add worklog" keep acting on the real active workbook. It is a
            // display choice local to this tab, the same kind RefreshBoardPreviews already makes when
            // it picks which schematic to show.
            //
            // In "Show all workbooks" scope, shownWorkbooks can hold cards for OTHER boards too - but
            // the board pane can only ever render the CURRENTLY LOADED board's schematics (see
            // CurrentBoardDataForPreviews), so the fallback below is restricted to this board's own
            // workbooks regardless of scope. Picking a different board's card still works - it just
            // goes through SelectWorkbook -> ActivateWorkbookAcrossBoards, which switches the loaded
            // board first - this is only about what shows with no explicit click.
            var shownWorkbook = active != null && shownWorkbooks.Any(w => w.Id == active.Id)
                ? active
                : shownWorkbooks.Where(w => string.Equals(w.BoardKey, boardKey, StringComparison.Ordinal))
                    .FirstOrDefault();

            this.thisSelectedWorkbookId = shownWorkbook?.Id ?? -1;

            // Recorded for the board pane's empty state: in AllBoards scope the list can be showing
            // cards while nothing is selectable for THIS board, and a pane that just goes blank
            // there reads as the workbooks having lost their contents. See RefreshBoardPreviews.
            this.thisHasWorkbooksOnOtherBoardsOnly = shownWorkbook == null && shownWorkbooks.Count > 0;

            // "1 workbook" / "3 workbooks" - the count is the panel's heading, so it has to be
            // right for one as well as for none and many. Counts what is SHOWN, so it reads as the
            // result count while a search is active.
            this.WorkbookCountText.Text = WorklogEntryScope.FormatCount(shownWorkbooks.Count, "workbook", "workbooks");

            this.NoWorkbooksText.IsVisible = shownWorkbooks.Count == 0;

            // A search that matched nothing is a different situation from a board with no workbooks
            // at all, and saying "no workbooks" for it reads as data loss.
            this.NoWorkbooksText.Text = !this.thisSearchQuery.IsEmpty && workbooks.Count > 0
                ? "No workbooks or worklogs match your search."
                : NoWorkbooksDefaultText;

            foreach (var workbook in shownWorkbooks)
            {
                bool isSelected = workbook.Id == this.thisSelectedWorkbookId;
                this.WorkbookListPanel.Children.Add(this.BuildWorkbookCard(workbook, isSelected));
            }

            this.ApplyHeaderForWorkbook(shownWorkbook);
            this.RefreshBoardPreviews();
        }

        // The empty-list message the markup ships with, restored whenever a search is not the reason
        // the list is empty. Captured as a constant rather than read back off the control, which by
        // then may be showing the no-results message instead.
        private const string NoWorkbooksDefaultText =
            "No repairs recorded for this board yet.\n" +
            "Use \"Create new workbook\" above the tabs to start one.";

        // ###########################################################################################
        // Narrows a board's workbooks to those matching the current search, and records WHICH of each
        // surviving workbook's entries matched (thisMatchedEntryIdsByWorkbookId) so the board pane and
        // the entry list can narrow to the same set without re-running the query.
        //
        // A workbook matches when its OWN text matches, or when any of its entries does - searching
        // for a component you replaced should find the job it was replaced in, not nothing. When the
        // workbook itself matched but no individual entry did, every entry is treated as matched:
        // the user found the workbook they were after, and hiding all of its contents would make the
        // result look empty.
        //
        // Entries are read once per workbook here, which is the expensive part (GetEntries re-parses
        // the file per call) - so it is skipped entirely for an empty query, the normal case.
        // ###########################################################################################
        private List<WorkbookRecord> ApplySearchFilter(List<WorkbookRecord> workbooks)
        {
            this.thisMatchedEntryIdsByWorkbookId.Clear();

            if (this.thisSearchQuery.IsEmpty)
                return workbooks;

            var matched = new List<WorkbookRecord>();

            foreach (var workbook in workbooks)
            {
                var entries = this.GetEntriesForThisPass(workbook.Id);

                bool workbookTextMatches = this.thisSearchQuery.Matches(WorklogSearchIndex.ForWorkbook(workbook));

                var matchedEntryIds = entries
                    .Where(e => this.thisSearchQuery.Matches(WorklogSearchIndex.ForEntry(e)))
                    .Select(e => e.Id)
                    .ToHashSet();

                if (!workbookTextMatches && matchedEntryIds.Count == 0)
                    continue;

                // The workbook itself matched but none of its entries did - show all of them rather
                // than an empty workbook, see the header.
                if (workbookTextMatches && matchedEntryIds.Count == 0)
                    matchedEntryIds = entries.Select(e => e.Id).ToHashSet();

                this.thisMatchedEntryIdsByWorkbookId[workbook.Id] = matchedEntryIds;
                matched.Add(workbook);
            }

            return matched;
        }

        // ###########################################################################################
        // One workbook's entries, read at most once per refresh pass. GetEntries has no cache of its
        // own (File.ReadAllText + Deserialize + a per-entry normalise loop, every call), and within a
        // single pass the search filter and the board pane both want the same workbook's entries.
        //
        // Internal so TabWorkbooks.BoardPreviews.cs (the other half of this class) shares the same
        // pass cache rather than re-reading.
        // ###########################################################################################
        private List<WorklogEntryRecord> GetEntriesForThisPass(int workbookId)
        {
            if (this.thisEntriesReadThisPass.TryGetValue(workbookId, out var cached))
                return cached;

            var entries = WorklogManager.GetEntries(workbookId);
            this.thisEntriesReadThisPass[workbookId] = entries;
            return entries;
        }

        // ###########################################################################################
        // Which of a workbook's entries the current search matched, or null when no filter applies
        // and every entry should be shown. Null rather than "all the ids" so callers do not have to
        // build a set they will not use on the unfiltered path, which is the normal one.
        // ###########################################################################################
        private HashSet<int>? MatchedEntryIdsForWorkbook(int workbookId)
        {
            if (this.thisSearchQuery.IsEmpty)
                return null;

            return this.thisMatchedEntryIdsByWorkbookId.TryGetValue(workbookId, out var ids)
                ? ids
                : new HashSet<int>();
        }

        // ###########################################################################################
        // Rebuilds the tab as the user types, DEBOUNCED - see thisSearchDebounceTimer for why a
        // rebuild per keystroke is too expensive to run directly. Goes through RefreshWorkbooks
        // rather than filtering the existing cards in place, so the workbook list, the board pane and
        // the entry list all narrow together from one pass - the same funnel every other worklog
        // change already uses.
        //
        // Restarting the timer on each keystroke is what coalesces a typing burst: only the pause at
        // the end of it actually rebuilds. RefreshWorkbooks re-reads the box itself, so nothing about
        // the query has to be carried through the timer.
        // ###########################################################################################
        private void OnFindRepairTextChanged(object? sender, TextChangedEventArgs e)
        {
            if (this.thisSearchDebounceTimer == null)
            {
                this.thisSearchDebounceTimer = new DispatcherTimer { Interval = SearchDebounceInterval };
                this.thisSearchDebounceTimer.Tick += (_, _) =>
                {
                    this.thisSearchDebounceTimer.Stop();
                    this.RefreshWorkbooks();
                };
            }

            this.thisSearchDebounceTimer.Stop();
            this.thisSearchDebounceTimer.Start();
        }

        // ###########################################################################################
        // How SelectWorkbook activates a workbook app-wide. Set by Initialize to Main's own
        // ActivateWorkbook; left null only in headless tests, which supply their own via
        // ActivateWorkbookOverrideForTests.
        //
        // A settable delegate rather than a "if (MainWindow != null) ... else <do it locally>"
        // branch inside SelectWorkbook, and this is the point: with a branch, EVERY test ran the
        // no-MainWindow side and the shipped click path - persist, then rebuild from the saved id -
        // was pinned by nothing at all. One path now, whoever calls it.
        // ###########################################################################################
        private Action<string, int>? thisActivateWorkbook;

        // Lets a headless test stand in for Main.ActivateWorkbook - normally
        // "save the id, then RefreshWorkbooks", which is what the running app does via
        // Main.RefreshWorklogBar. Null in the running app, always; Initialize sets the real one.
        internal Action<string, int>? ActivateWorkbookOverrideForTests
        {
            get => this.thisActivateWorkbook;
            set => this.thisActivateWorkbook = value;
        }

        // ###########################################################################################
        // Selects a workbook by id, called when the user clicks a card. This "activates" the
        // workbook app-wide rather than only changing what this panel highlights: the worklog bar
        // above the tabs switches to showing it, "Show worklogs" on the Schematics tab draws ITS
        // entries, and "Add worklog" starts writing new entries into it - all in place of the
        // board's newest workbook, which is what every one of those otherwise defaults to.
        //
        // Deliberately does NOT switch tabs. Activating a workbook here is meant to be seen HERE -
        // the top-line and the board pane update in place for the newly active workbook - so the
        // user can browse several workbooks' schematics without leaving this tab. Only "Add
        // worklog" on the bar jumps to Schematics (see Main.OnWorklogAddEntryClick), because
        // drawing a new entry needs the actual schematic view to draw on.
        //
        // Activation goes through thisActivateWorkbook (Main.ActivateWorkbook in the running app),
        // which saves to UserSettings.ActiveWorkbookIdByBoard and calls RefreshWorklogBar - and
        // that calls back into RefreshWorkbooks above, which re-derives thisSelectedWorkbookId from
        // the same saved id. So there is exactly one path that decides what is selected and this
        // method does not duplicate it.
        //
        // NO "already selected, nothing to do" guard, deliberately. RefreshWorkbooks HIGHLIGHTS a
        // default card when the board has no saved activation, without saving anything - so the
        // card the user is looking at can be "the selected one" on screen while
        // ActiveWorkbookIdByBoard is empty. An early return on id equality made clicking that exact
        // card a no-op, leaving the highlight and the persisted state disagreeing: creating a newer
        // workbook afterwards then silently moved the user off it, and only clicking a different
        // card and back made the choice stick. An assumed highlight and a real activation must not
        // be indistinguishable to the persistence layer. Re-activating what is already active is
        // idempotent and cheap, so there is nothing to protect against here.
        // ###########################################################################################
        // Exposed to the test project so WorkbooksListTests can select a card without fighting
        // pointer-event routing against a UserControl that has no window - the same idea as
        // BoardKeyOverrideForTests. The running app always reaches this through a card's
        // PointerPressed handler in BuildWorkbookCard, never directly.
        internal void SelectWorkbookForTests(int workbookId, string? boardKeyOverride = null) =>
            this.SelectWorkbook(workbookId, boardKeyOverride);

        // ###########################################################################################
        // boardKey is the CLICKED CARD's own board key, not necessarily the currently selected
        // board's - in "Show all workbooks" scope a card can belong to a different board, and
        // thisActivateWorkbook (Main.ActivateWorkbookAcrossBoards) switches the app to it before
        // activating. In "current board" scope every shown card's board key already equals the
        // selected board's, so this is unchanged for that case.
        // ###########################################################################################
        private void SelectWorkbook(int workbookId, string? boardKey = null)
        {
            boardKey ??= this.BoardKeyOverrideForTests ?? this.MainWindow?.GetCurrentBoardKey() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(boardKey))
                return;

            this.thisActivateWorkbook?.Invoke(boardKey, workbookId);

            // AFTER the activation, not before: that call is what rebuilds the board pane and so
            // resolves which schematic this workbook actually shows (RefreshBoardPreviews sets
            // thisSelectedSchematicName), and propagating first would send the PREVIOUS workbook's
            // schematic over. Here rather than inside RefreshBoardPreviews because propagation must
            // only ever follow a user's own click - see
            // PropagateSelectedSchematicToSchematicsTab's header for what hooking the refresh path
            // broke.
            this.PropagateSelectedSchematicToSchematicsTab();
        }

        // ###########################################################################################
        // Updates the top-line's title and status pill for the given workbook, or clears it back to
        // the tab's construction-time placeholder when there is none (no board selected, or the
        // board has no workbooks).
        // ###########################################################################################
        // The workbook the top-line, and so the Edit/Delete buttons beside it, currently act on -
        // set alongside everything else ApplyHeaderForWorkbook draws, null when nothing is
        // selected (in which case the buttons are hidden, see below).
        private WorkbookRecord? thisHeaderWorkbook;

        private void ApplyHeaderForWorkbook(WorkbookRecord? workbook)
        {
            this.thisHeaderWorkbook = workbook;

            if (workbook == null)
            {
                this.WorkbookHeaderTitleText.Text = "No workbook selected";
                this.WorkbookHeaderStatusPill.IsVisible = false;
                this.WorkbookHeaderNoteText.IsVisible = false;
                this.WorkbookHeaderActionsPanel.IsVisible = false;
                this.WorkbookSummaryPanel.IsVisible = false;
                return;
            }

            string title = string.IsNullOrWhiteSpace(workbook.Title) ? "(untitled)" : workbook.Title;
            this.ApplyHighlightedText(this.WorkbookHeaderTitleText, $"#{workbook.Id} · {title}");

            // Border, glyph, label and the padlock's overshoot padding all applied by the ONE
            // shared informational styling - see WorklogInfoPillBuilder. This pill is declared in
            // the markup (it is long-lived and only its text changes), so it is restyled in place
            // rather than rebuilt.
            this.WorkbookHeaderStatusPill.IsVisible = true;
            WorklogInfoPillBuilder.ApplyStatePillVisual(
                this.WorkbookHeaderStatusPill,
                this.WorkbookHeaderStatusGlyph,
                this.WorkbookHeaderStatusText,
                workbook.Status);

            // Blank for most workbooks (Note is optional in the create/edit dialog), so the row is
            // collapsed rather than showing an empty muted TextBlock next to the pill.
            bool hasNote = !string.IsNullOrWhiteSpace(workbook.Note);
            this.WorkbookHeaderNoteText.IsVisible = hasNote;
            this.ApplyHighlightedText(this.WorkbookHeaderNoteText, hasNote ? workbook.Note : string.Empty, linkify: true);

            this.WorkbookHeaderActionsPanel.IsVisible = true;

            // The summary reads the workbook's entries through the within-pass cache, so a refresh
            // that has already read them for the search filter and the board pane does not read
            // the same files a third time - see GetEntriesForThisPass.
            this.ApplySummaryForWorkbook(workbook);
        }

        // ###########################################################################################
        // Opens the SAME modal "Create new workbook" uses, switched into edit mode via
        // InitializeForEdit so the user can change the description/note of the workbook the
        // top-line is currently showing.
        // ###########################################################################################
        private async void OnEditWorkbookClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisHeaderWorkbook == null)
                return;

            if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
                return;

            var dialog = new CreateWorkbookWindow();
            dialog.InitializeForEdit(this.thisHeaderWorkbook);

            var updated = await dialog.ShowDialog<WorkbookRecord?>(ownerWindow);
            if (updated == null)
                return;

            // The edited workbook is already the active one (Edit only ever acts on the workbook
            // the top-line is showing), so a bare refresh - not ActivateWorkbook - is enough to
            // pick up the new title/note everywhere it is drawn (this tab's list/top-line, and the
            // worklog bar via RefreshWorklogBar).
            this.MainWindow?.RefreshWorklogBar();
        }

        // ###########################################################################################
        // Deletes the workbook the top-line is currently showing, after the user confirms in
        // DeleteWorkbookWindow. The board's other workbooks are unaffected; RefreshWorklogBar's own
        // ResolveActiveWorkbook then picks the next one automatically - the saved
        // UserSettings.ActiveWorkbookIdByBoard entry (if it named the one just deleted) no longer
        // matches any workbook on disk, so it falls back to the board's newest remaining one,
        // exactly as it already does for a workbook deleted by hand outside the app.
        // ###########################################################################################
        private async void OnDeleteWorkbookClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisHeaderWorkbook == null)
                return;

            if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
                return;

            var workbook = this.thisHeaderWorkbook;

            var dialog = new DeleteWorkbookWindow();
            dialog.Initialize(workbook);

            bool? confirmed = await dialog.ShowDialog<bool?>(ownerWindow);
            if (confirmed != true)
                return;

            // An entry-drawing mode armed by "Add worklog" captured a workbook id when it started
            // (BeginWorklogEntryMode), and nothing about deleting that workbook cancels it. Without
            // this the cross cursor stays live on the Schematics tab after the workbook is gone, the
            // user draws an area, and AddEntry finds no folder - the work is discarded with nothing
            // but a log line. Main.ActivateWorkbook performs the same teardown, and for the same
            // reason: this is the other way "which workbook is being written to" can stop being
            // valid.
            this.MainWindow?.TabSchematicsControl?.CancelWorklogEntryMode();

            if (!WorklogManager.DeleteWorkbook(workbook.Id))
            {
                // The folder could not be removed (a photo held open in an external viewer, say).
                // Say so rather than returning silently: the user confirmed a destructive action and
                // an unchanged list with no message reads as "the click did not register", which
                // invites them to try again. Same treatment CreateWorkbookWindow gives a failed
                // create.
                await ShowWorkbookActionFailedAsync(
                    ownerWindow,
                    "Delete workbook",
                    $"Could not delete workbook #{workbook.Id} - see the log for details.\n\n" +
                    "It may be open in another program, for example a photo or file from the workbook.");
                return;
            }

            this.MainWindow?.RefreshWorklogBar();
        }

        // ###########################################################################################
        // A minimal "that did not work" modal for the destructive actions on this tab. The
        // create/edit dialog reports its own failures inline in its validation line, but a Delete
        // has no dialog left on screen by the time it fails - its confirmation has already closed -
        // so the message needs a window of its own.
        //
        // The title is a PARAMETER rather than the hardcoded "Delete workbook" it started as: both
        // deletes on this tab report through here now (a workbook, and a single worklog from an
        // entry card), and a "Delete workbook" title over a message about a worklog names the wrong
        // object at exactly the moment the user is trying to work out what failed.
        // ###########################################################################################
        private static async Task ShowWorkbookActionFailedAsync(Window ownerWindow, string title, string message)
        {
            var okButton = new Button
            {
                Content = "OK",
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 16, 0, 0)
            };

            var dialog = new Window
            {
                Title = title,
                Width = 420,
                SizeToContent = SizeToContent.Height,
                CanResize = false,
                CanMinimize = false,
                ShowInTaskbar = false,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background = ThemeResources.ResolveBrush("Bg", Brushes.White),
                Foreground = ThemeResources.ResolveBrush("Fg", Brushes.Black)
            };

            var body = new StackPanel { Margin = new Thickness(18), Spacing = 4 };
            body.Children.Add(new TextBlock { Text = message, TextWrapping = TextWrapping.Wrap });
            body.Children.Add(okButton);
            dialog.Content = body;

            okButton.Click += (_, _) => dialog.Close();
            dialog.AddHandler(
                KeyDownEvent,
                (_, args) =>
                {
                    if (args.Key == Key.Escape || args.Key == Key.Enter)
                    {
                        dialog.Close();
                        args.Handled = true;
                    }
                },
                RoutingStrategies.Tunnel);

            await dialog.ShowDialog(ownerWindow);
        }

        // ###########################################################################################
        // Builds one workbook card: the "#N" id with its status pill, the title, and the worklog
        // count with the start date. Clicking anywhere on the card selects that workbook - see
        // SelectWorkbook.
        //
        // Built in code rather than as an ItemsControl DataTemplate because the status pill's brush
        // has to be resolved through Application.Current with an explicit ThemeVariant (see
        // WorklogInfoPillBuilder), which a template binding cannot express - the same reason
        // Main builds the worklog bar's pill in code.
        //
        // An instance method, not static: it needs to reach SelectWorkbook on THIS tab from the
        // click handler it wires up.
        // ###########################################################################################
        // ###########################################################################################
        // A TextBlock whose search-matched runs are drawn with the Workbooks_SearchHit_* wash - the
        // "<highlight>This</highlight> is a <highlight>text</highlight>" behaviour the search was
        // asked for.
        //
        // With no search active (the normal case) this returns an ordinary single-Text TextBlock:
        // splitting into Inlines costs measurably more to lay out, and there is nothing to mark.
        //
        // The SPLIT itself is WorklogSearchQuery.SplitIntoSegments, not done here - every segment
        // boundary has to line up exactly or characters get dropped or doubled on screen, and that
        // maths is unit-tested on the Handlers side. This method only turns segments into runs.
        //
        // extraSetup runs on the finished block so callers can set the properties they need
        // (classes, alignment, weight) without this needing a parameter for each.
        // ###########################################################################################
        private TextBlock BuildHighlightedTextBlock(
            string text,
            double fontSize,
            TextWrapping wrapping = TextWrapping.NoWrap,
            Action<TextBlock>? extraSetup = null,
            bool linkify = false)
        {
            var block = new TextBlock
            {
                FontSize = fontSize,
                TextWrapping = wrapping
            };

            this.ApplyHighlightedText(block, text, linkify);

            extraSetup?.Invoke(block);
            return block;
        }

        // ###########################################################################################
        // Sets an EXISTING TextBlock's content with the search-matched runs marked - the same job as
        // BuildHighlightedTextBlock, for the blocks that come from the markup rather than from code
        // (the top-line's title and note).
        //
        // Clearing Inlines before every write matters: a block that was highlighted on the previous
        // pass keeps those runs otherwise, and setting Text alone would leave the old marked runs
        // rendering underneath the new value.
        //
        // linkify OPTS IN to rendering any web links in the text as clickable runs, and is off by
        // default. It belongs on the fields the user writes prose into - a workbook's Note, an
        // entry's Description - and NOT on titles or the "#N · Title" top line: a title is a
        // headline, and a URL typed into one is a label rather than something to navigate to.
        // Both markings compose in one pass (see TextLinkRenderer.ApplySegments), because a search
        // term routinely lands inside a URL and applying one split after the other would lose runs.
        // ###########################################################################################
        private void ApplyHighlightedText(TextBlock block, string text, bool linkify = false)
        {
            var segments = this.thisSearchQuery.IsEmpty
                ? null
                : this.thisSearchQuery.SplitIntoSegments(text);

            if (linkify)
            {
                TextLinkRenderer.ApplySegments(block, text, segments);
                return;
            }

            block.Inlines?.Clear();

            if (segments == null || !segments.Any(s => s.IsMatch))
            {
                block.Text = text;
                return;
            }

            var hitBackground = ThemeResources.ResolveBrush("Workbooks_SearchHit_Bg", Brushes.Yellow);
            var hitForeground = ThemeResources.ResolveBrush("Workbooks_SearchHit_Fg", Brushes.Black);

            // Text must be cleared before Inlines are added - a TextBlock carrying both renders the
            // Text and ignores the Inlines entirely, which showed up as highlighting that silently
            // did nothing on the markup-declared blocks.
            block.Text = null;

            foreach (var segment in segments)
            {
                var run = new Run(segment.Text);

                if (segment.IsMatch)
                {
                    run.Background = hitBackground;
                    run.Foreground = hitForeground;
                }

                block.Inlines!.Add(run);
            }
        }

        private Border BuildWorkbookCard(WorkbookRecord workbook, bool isSelected)
        {
            // --- Row 1: "#N" and the status pill -------------------------------------------------
            var idText = new TextBlock
            {
                Text = $"#{workbook.Id}",
                FontSize = 11,
                FontWeight = FontWeight.Bold,
                VerticalAlignment = VerticalAlignment.Center
            };

            // The one shared informational pill - see WorklogInfoPillBuilder. This was ~35 lines
            // building a 2px-outlined copy by hand, under a comment asserting it matched the
            // worklog bar and the entry pills; it did not, which is what "the pills are not
            // identical" was reported about.
            var statusPill = WorklogInfoPillBuilder.BuildStatePill(workbook.Status, fontSize: 10.0);
            statusPill.HorizontalAlignment = HorizontalAlignment.Right;

            var headerRow = new Grid();
            headerRow.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
            headerRow.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Star));
            Grid.SetColumn(idText, 0);
            Grid.SetColumn(statusPill, 1);
            headerRow.Children.Add(idText);
            headerRow.Children.Add(statusPill);

            // --- Row 2: the title ----------------------------------------------------------------
            // A workbook can be created without one, so fall back to something rather than leaving
            // the card's middle line blank and the card looking broken.
            var titleText = this.BuildHighlightedTextBlock(
                string.IsNullOrWhiteSpace(workbook.Title) ? "(untitled)" : workbook.Title,
                11,
                TextWrapping.Wrap);

            // --- Row 3: the worklog count and the start date -------------------------------------
            var metaText = new TextBlock
            {
                Text = BuildWorkbookMetaText(workbook),
                FontSize = 10,
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center
            };
            metaText.Classes.Add("WorkbooksMuted");

            var body = new StackPanel { Spacing = 3 };
            body.Children.Add(headerRow);
            body.Children.Add(titleText);
            body.Children.Add(metaText);

            // --- Row 4: which board this workbook belongs to, "Show all workbooks" scope only ----
            // Only shown when the list can hold more than one board's cards - in the normal
            // current-board scope every card already belongs to the board on screen, and naming it
            // on each card would just be noise.
            if (IsAllBoardsScope)
            {
                var boardLabelText = new TextBlock
                {
                    Text = this.MainWindow != null ? Main.FormatBoardKeyForDisplay(workbook.BoardKey) : workbook.BoardKey,
                    FontSize = 10,
                    TextWrapping = TextWrapping.Wrap,
                    VerticalAlignment = VerticalAlignment.Center
                };
                boardLabelText.Classes.Add("WorkbooksMuted");
                body.Children.Add(boardLabelText);
            }

            var card = new Border
            {
                // The record rides along on the card so SelectWorkbook can read it back off
                // WorkbookListPanel.Children without a second id->record lookup. The Hand cursor
                // comes from the WorkbooksJobCard style in the .axaml, not set here.
                Tag = workbook,
                Child = body
            };
            card.Classes.Add("WorkbooksJobCard");
            if (isSelected)
                card.Classes.Add("Selected");

            card.PointerPressed += (_, _) => this.SelectWorkbook(workbook.Id, workbook.BoardKey);

            return card;
        }

        // ###########################################################################################
        // The card's third line: "{x} worklogs · started 2026-August-26".
        //
        // The date format matches the worklog bar's (yyyy-MMMM-dd, invariant), so a workbook's start
        // date reads the same in both places. Invariant rather than the current culture on purpose:
        // the bar chose it so the month is always a name and never an ambiguous number, and the two
        // must not disagree.
        // ###########################################################################################
        private static string BuildWorkbookMetaText(WorkbookRecord workbook)
        {
            string startDate = workbook.StartDate.ToString("yyyy-MMMM-dd", CultureInfo.InvariantCulture);

            return $"{WorklogEntryScope.FormatCount(workbook.EntryCount, "worklog", "worklogs")} · started {startDate}";
        }

    }
}
