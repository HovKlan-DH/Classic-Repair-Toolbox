using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // The worklog bar: which workbook is active per board, the bar's own refresh, workbook
    // activation (from the Workbooks tab or the bar's own picker), the "Add worklog"/"Show
    // worklogs" buttons, and the Workbooks tab's own focus-on-select behaviour. See
    // Main.axaml.cs for the file map of the whole partial class.
    // ###########################################################################################
    public partial class Main
    {
        // ###########################################################################################
        // Resolves the ONE workbook the worklog bar shows, "Show worklogs" draws, and "Add worklog"
        // writes new entries into - the single notion of "the active workbook" every worklog-facing
        // control in Main and TabSchematics shares.
        //
        // Defaults to the board's newest workbook (open or closed - see RefreshWorklogBar's header
        // for why status is not a filter here). Selecting a workbook on the Workbooks tab
        // (TabWorkbooks.SelectWorkbook, via ActivateWorkbook below) overrides that default by saving
        // the choice in UserSettings.ActiveWorkbookIdByBoard, so "activating" an older or closed
        // workbook makes IT the one every other worklog surface acts on until something reactivates
        // the newest one again.
        //
        // The saved id is validated against this board's actual workbooks on every call rather than
        // trusted blindly: a workbook can be deleted from disk by hand, or the saved id can be stale
        // after switching boards, and an unvalidated id would otherwise make the bar quietly show
        // nothing (or another board's workbook, if ids ever collided) instead of falling back.
        //
        // The rule itself lives in WorklogManager.ResolveActiveWorkbook, not here: the Workbooks tab
        // needs the same answer for its highlighted card, and two copies of "saved id if valid, else
        // newest" is exactly how the card and the bar came to disagree. This wrapper only fetches
        // the two inputs.
        // ###########################################################################################
        // internal rather than private so the component popup's "attach capture to worklog" flow can
        // ask the same question this window does, rather than re-deriving "which workbook" a third
        // time - WorklogManager.ResolveActiveWorkbook's own header is explicit that it is the ONE
        // place that rule lives.
        internal static WorkbookRecord? ResolveActiveWorkbookForBoard(string boardKey) =>
            WorklogManager.ResolveActiveWorkbook(
                WorklogManager.GetWorkbooksForBoard(boardKey),
                UserSettings.GetActiveWorkbookId(boardKey));

        // ###########################################################################################
        // Activates a workbook for its board: persists it as the board's active workbook (so it
        // survives a tab switch, a board switch and back, and an app restart) and refreshes every
        // worklog surface that depends on "the active workbook" - the bar, and (via
        // TabSchematicsControl.SetShowWorklogEntriesList inside RefreshWorklogBar) the Schematics
        // tab's "Show worklogs" overlay for whichever schematic is on screen there.
        //
        // Called from TabWorkbooks.SelectWorkbook when the user clicks a card in the Workbooks tab -
        // see that method for why the caller ALSO switches to the Schematics tab, which this method
        // deliberately does not do itself (a board/data refresh must be able to call this without
        // stealing the user's current tab).
        // ###########################################################################################
        // ###########################################################################################
        // Sets the component highlight-rect cache and tells everything that reads it.
        //
        // The cache physically lives on TabSchematics (it is built as a side effect of that tab's
        // board load and most of its readers are there), but MAIN is what actually owns it: all four
        // writes are here - a board switch blanking it, the board load populating it, and both
        // region-filter paths rebuilding it - and TabSchematics never assigns it at all.
        //
        // Routing every write through this one method exists for the SECOND reader. The Workbooks
        // tab needs the same cache for a pill's "Mark components in scope" checklist, and the pane's
        // pills go on screen before the board load's fire-and-forget task has populated it: click one
        // in that window and the lookup missed, the checklist was silently skipped, and the two
        // modals were no longer identical - the exact bug this feature was written to fix, back as an
        // intermittent one. A region switch had the same gap from the other end, leaving the pane
        // stale against a cache that had just been rebuilt with a different region's rects.
        //
        // Refreshing the board pane from here rather than from each write site means a fifth write
        // added later cannot forget to.
        // ###########################################################################################
        private void SetComponentHighlightRects(Dictionary<string, Dictionary<string, List<Rect>>> highlightRects)
        {
            this.TabSchematicsControl.highlightRectsBySchematicAndLabel = highlightRects;
            this.TabWorkbooks?.RefreshBoardPreviewsForCurrentSelection();
        }

        public void ActivateWorkbook(string boardKey, int workbookId)
        {
            if (string.IsNullOrWhiteSpace(boardKey))
                return;

            // An entry-drawing mode started for the PREVIOUSLY active workbook captured that
            // workbook's id (BeginWorklogEntryMode), and nothing about switching tabs cancels it. So
            // without this, activating another workbook here left the cross cursor live on the
            // Schematics tab and the next drawn entry was written into the workbook the user had just
            // navigated away from, while every visible surface named the new one.
            // ApplyWorklogBarVisibility already performs the same teardown when the feature is
            // switched off; this is the other way "which workbook is being written to" can change.
            this.TabSchematicsControl.CancelWorklogEntryMode();

            UserSettings.SetActiveWorkbookId(boardKey, workbookId);
            this.RefreshWorklogBar();
        }

        // ###########################################################################################
        // Lets the worklog bar's own workbook picker activate ANY workbook directly - including one
        // for a board other than the one currently on screen, since the picker now lists every
        // workbook on every board (see RefreshWorklogBar). Guarded by _suppressWorklogJobBoxSave so
        // RefreshWorklogBar re-seeding the box's ItemsSource/SelectedItem does not loop back in here.
        // ###########################################################################################
        private void OnWorklogJobBoxSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this._suppressWorklogJobBoxSave)
                return;

            if (this.WorklogJobBox.SelectedItem is not WorkbookRecord selected)
                return;

            this.ActivateWorkbookAcrossBoards(selected.BoardKey, selected.Id);
        }

        // ###########################################################################################
        // Activates a workbook that may belong to a board other than the one currently on screen -
        // shared by the worklog bar's own picker (OnWorklogJobBoxSelectionChanged) and the Workbooks
        // tab's "Show all workbooks" scope (TabWorkbooks.SelectWorkbook), both of which can now list
        // workbooks for every board rather than just the current one.
        //
        // Same-board case: identical to clicking a card when the tab is scoped to the current board -
        // just calls ActivateWorkbook.
        //
        // Different-board case: the hardware/board selectors are switched FIRST. Persisting the
        // target workbook as its board's active one before that switch (rather than after, via
        // ActivateWorkbook) matters because switching BoardComboBox.SelectedItem starts the async
        // OnBoardSelectionChanged, and that method calls RefreshWorklogBar itself once the board
        // finishes loading (see its own comment on why that call moved after _currentBoardData is
        // assigned) - so by the time it does, ResolveActiveWorkbookForBoard must already see this
        // choice as the saved one, or the freshly loaded board would show its OWN newest workbook
        // instead of the one just picked. Calling ActivateWorkbook here too, after triggering the
        // switch, would be redundant at best and would run against the board being LEFT rather than
        // the one being entered.
        // ###########################################################################################
        public void ActivateWorkbookAcrossBoards(string boardKey, int workbookId)
        {
            if (string.IsNullOrWhiteSpace(boardKey))
                return;

            if (string.Equals(boardKey, this.GetCurrentBoardKey(), StringComparison.Ordinal))
            {
                this.ActivateWorkbook(boardKey, workbookId);
                return;
            }

            if (!TryResolveHardwareAndBoardForBoardKey(boardKey, out string hardwareName, out string boardName))
            {
                // The board this workbook belongs to is no longer in the synced data (its Excel entry
                // was removed by a later sync), so there is nothing to switch to. Refresh instead of
                // switching, so every surface goes back to naming the workbook that IS active rather
                // than one that cannot be reached.
                this.RefreshWorklogBar();
                return;
            }

            UserSettings.SetActiveWorkbookId(boardKey, workbookId);

            this.TabSchematicsControl.CancelWorklogEntryMode();

            // Switching HardwareComboBox first (if needed) synchronously repopulates BoardComboBox's
            // ItemsSource via OnHardwareSelectionChanged before BoardComboBox.SelectedItem is set -
            // setting the board first would pick an index into the OLD hardware's board list.
            //
            // _pendingBoardSelectionOverride makes that hardware switch land directly on the board we
            // are actually after. Without it OnHardwareSelectionChanged selects the new hardware's
            // SAVED last board, running a whole board load for a board nobody asked for before the
            // real one is set - see the field's own comment.
            var currentHardware = this.HardwareComboBox.SelectedItem as string;
            if (!string.Equals(currentHardware, hardwareName, StringComparison.OrdinalIgnoreCase))
            {
                this._pendingBoardSelectionOverride = boardName;
                this.HardwareComboBox.SelectedItem = hardwareName;

                // Cleared by OnHardwareSelectionChanged, which also selected boardName for us - so
                // the assignment below is a no-op in the normal case and only does real work if that
                // handler could not honour the override (the board vanished between the lookup above
                // and the switch).
            }

            this.BoardComboBox.SelectedItem = boardName;
        }

        // ###########################################################################################
        // Switches the main tab strip to Schematics, if it is not already there. A public wrapper
        // rather than exposing MainTabControl/SchematicsTabItem themselves to other tabs: those
        // fields are Avalonia's own x:Name-generated ones, and TabWorkbooks (the one other caller
        // that needs this, from SelectWorkbook) has no business reaching into Main's tab control
        // directly - the same reasoning OnWorklogAddEntryClick already followed for itself.
        // ###########################################################################################
        public void SwitchToSchematicsTab()
        {
            // Null-guarded like ApplyOscilloscopeTabVisibility and ApplyWorklogBarVisibility, its two
            // siblings that reach for the same control. This is public and callable from another tab
            // now, so it can no longer assume it only runs at a point where the tab control is up.
            if (this.MainTabControl == null || this.SchematicsTabItem == null)
                return;

            if (!ReferenceEquals(this.MainTabControl.SelectedItem, this.SchematicsTabItem))
                this.MainTabControl.SelectedItem = this.SchematicsTabItem;
        }

        // ###########################################################################################
        // Gives the Workbooks tab's own search box priority focus the moment it becomes the selected
        // tab - see TabWorkbooks.FocusSearchBox for how that coexists with the global "steal focus
        // into ComponentSearchTextBox" handler wired in the constructor (which now excludes this tab
        // by header). Every OTHER tab is left alone: this handler only ever hands focus TO the
        // Workbooks box, never takes it away when leaving - the global handler's own per-click logic
        // already covers every tab that wants ComponentSearchTextBox back, including this one once
        // the user clicks something on it.
        //
        // e.Source MUST be checked against MainTabControl itself. SelectionChanged is a bubbling
        // routed event, and several tabs hold their own ListBox/ComboBox controls (the Schematics
        // tab's thumbnail list among them) - without this guard, selecting a thumbnail or any other
        // nested list would bubble up through MainTabControl and re-fire this, stealing focus back
        // to the search box away from whatever the user had just clicked on the Workbooks tab.
        // ###########################################################################################
        private void OnMainTabControlSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (!ReferenceEquals(e.Source, this.MainTabControl))
            {
                return;
            }

            if (this.MainTabControl?.SelectedItem is TabItem { Header: "Workbooks" })
            {
                this.TabWorkbooks?.FocusSearchBox();
            }
        }

        // ###########################################################################################
        // Refreshes the worklog bar's content for the currently selected board: either the empty
        // "no jobs recorded" state, or that board's active workbook (see ResolveActiveWorkbookForBoard).
        //
        // The bar deliberately shows the active workbook whether it is Open or Closed. Resolving a
        // workbook's last outstanding entry closes it automatically, and showing only open workbooks
        // meant a finished workbook disappeared from the UI entirely - indistinguishable from having
        // been deleted. Status is presentation only: it picks the status dot's color and appears in
        // the label, and changes nothing about what the bar's buttons offer. "Add worklog" still works
        // on a closed workbook (adding a still-Open entry reopens it through the normal
        // RecomputeWorkbookStatus rule), as does the "Show worklogs" toggle.
        // ###########################################################################################
        public void RefreshWorklogBar()
        {
            if (this.WorklogNoJobText == null)
                return;

            string boardKey = this.GetCurrentBoardKey();

            // ONE disk scan for both lists. GetAllWorkbooks and GetWorkbooksForBoard each go through
            // ReadAllWorkbooks, which enumerates every workbook folder and does File.Exists +
            // ReadAllText + Deserialize per folder, uncached - so calling both would scan the whole
            // Workbook tree twice on every board change, entry save, workbook create/close/delete and
            // card click. The board-scoped list is a SUBSET of the full one, so it is filtered here
            // rather than re-read; the ordering (descending id, newest first) is GetAllWorkbooks' own
            // and is exactly what GetWorkbooksForBoard would have produced.
            var allWorkbooks = WorklogManager.GetAllWorkbooks();
            var workbooks = allWorkbooks
                .Where(w => string.Equals(w.BoardKey, boardKey, StringComparison.Ordinal))
                .ToList();

            var activeWorkbook = WorklogManager.ResolveActiveWorkbook(workbooks, UserSettings.GetActiveWorkbookId(boardKey));
            bool hasWorkbook = activeWorkbook != null;

            // The Workbooks tab's list is rebuilt from here rather than from its own wiring: this
            // method is already the single place worklog state is refreshed from - board changes,
            // entry saves, workbook creation and closure all reach it - so the tab cannot go stale
            // in a case the bar handles and the tab forgot about.
            //
            // Skipped entirely when the worklog feature is switched off. That rebuild decodes every
            // schematic image with an entry in the active workbook (full-resolution PNGs, hundreds of
            // MB of BGRA on a big board) and re-reads entries.json, and it ran on every board change
            // for a user who has the whole feature disabled and the tab hidden. Safe to skip because
            // ApplyWorklogBarVisibility rebuilds the tab on the enable side before revealing it, so
            // it cannot come back showing what was last rendered.
            //
            // Deliberately NOT also gated on "is the Workbooks tab currently selected". Clicking a
            // card in that tab goes ActivateWorkbook -> RefreshWorklogBar -> RefreshWorkbooks, so an
            // early-out on tab selection would make card clicks silently do nothing.
            if (UserSettings.EnableWorklog)
            {
                // The tab's own list can show every board's workbooks ("Show all workbooks", the
                // Configuration tab's radio group below "Enable Worklog") rather than just this
                // board's - allWorkbooks is already the unfiltered read from above, so passing it
                // through costs nothing extra; board is still used to resolve which workbook is
                // active and to filter the fallback when no explicit choice was made (see
                // TabWorkbooks.RefreshWorkbooks).
                bool showAllBoards = string.Equals(UserSettings.WorkbooksScope, "AllBoards", StringComparison.Ordinal);
                this.TabWorkbooks?.RefreshWorkbooks(showAllBoards ? allWorkbooks : workbooks);
            }

            this.WorklogNoJobText.IsVisible = !hasWorkbook;

            // The picker stays visible whenever ANY workbook exists anywhere, not just when THIS
            // board has one. It lists every workbook on every board and is the only way to reach
            // another board's workbook from the bar, so hiding it on a board with none of its own
            // hid it exactly where it is most useful - a board you have never worked on is precisely
            // where you want to jump to the job you were doing elsewhere. (The visibility rule was
            // written for the old read-only text box, which genuinely had nothing to show.)
            this.WorklogJobBox.IsVisible = allWorkbooks.Count > 0;

            // These three still follow the ACTIVE workbook: there is no status to show, no entries to
            // draw and nothing to add an entry to when this board has no workbook.
            this.WorklogJobStatusPanel.IsVisible = hasWorkbook;
            this.WorklogShowEntriesPanel.IsVisible = hasWorkbook;
            this.WorklogAddEntryButton.IsVisible = hasWorkbook;

            if (!hasWorkbook || activeWorkbook!.Id != this._worklogShowEntriesWorkbookId)
            {
                // A different (or no) workbook is now shown - re-seed the checkbox from the user's
                // saved preference and apply it to this workbook directly, rather than relying on
                // IsCheckedChanged (which will not fire below when the new value matches what the
                // checkbox already showed for the previous workbook).
                //
                // The write is suppressed because it is this code seeding the checkbox, not the
                // user clicking it: showByDefault is forced false whenever the board has no
                // workbook, and letting that reach the handler would persist it as the user's
                // preference. See _suppressWorklogShowEntriesSave.
                bool showByDefault = hasWorkbook && UserSettings.WorklogShowEntriesChecked;

                this._suppressWorklogShowEntriesSave = true;
                try
                {
                    this.WorklogShowEntriesCheckBox.IsChecked = showByDefault;
                }
                finally
                {
                    this._suppressWorklogShowEntriesSave = false;
                }

                int workbookId = showByDefault ? activeWorkbook!.Id : 0;
                this._worklogShowEntriesWorkbookId = workbookId;
                this.TabSchematicsControl.SetShowWorklogEntriesList(showByDefault, workbookId);
            }
            else
            {
                // Same workbook as before, but its ENTRIES may have just changed - an area redrawn,
                // "Show marked area" ticked, an entry added or deleted. The branch above only fires
                // when the shown WORKBOOK changes, so without this the schematic overlay and the
                // thumbnail pills kept drawing the entry as it was before the edit until something
                // else happened to rebuild them. Reported as the marker not updating.
                this.TabSchematicsControl.RefreshWorklogEntriesListForCurrentWorkbook();
            }

            // The picker lists EVERY workbook on every board, not just this board's own workbooks in
            // "workbooks" above - selecting one for a different board is how the bar can jump there
            // (see OnWorklogJobBoxSelectionChanged). "allWorkbooks" is the single disk read taken at
            // the top of this method; "workbooks" is its board-scoped subset.
            //
            // SelectedItem is looked up BY ID inside this same list rather than set to "activeWorkbook"
            // directly: WorkbookRecord has no Equals/GetHashCode override, so ComboBox's SelectedItem
            // only shows as selected when it is REFERENCE-equal to an item actually in ItemsSource -
            // and while activeWorkbook now comes from a filtered view of this very list (so reference
            // equality would in fact hold today), the by-id lookup keeps that from being a silent
            // dependency on how "workbooks" happens to be derived.
            //
            // Seeded whether or not a workbook is active: with no active workbook the box shows no
            // selection while still listing every OTHER board's workbooks, which is what makes it
            // usable as a navigator on a board that has none of its own. Suppressed the same way
            // OnWorklogShowEntriesCheckedChanged's seed is: this is RefreshWorklogBar re-populating
            // the list for the board on screen, not a user picking a workbook, and letting it reach
            // OnWorklogJobBoxSelectionChanged would call ActivateWorkbook -> RefreshWorklogBar again
            // for a selection nobody made.
            this._suppressWorklogJobBoxSave = true;
            try
            {
                this.WorklogJobBox.ItemsSource = allWorkbooks;
                this.WorklogJobBox.SelectedItem = hasWorkbook
                    ? allWorkbooks.FirstOrDefault(w => w.Id == activeWorkbook!.Id)
                    : null;
            }
            finally
            {
                this._suppressWorklogJobBoxSave = false;
            }

            if (activeWorkbook == null)
                return;

            // Border, padlock, label and the padlock's overshoot padding all applied by the ONE
            // shared informational styling - see Handlers/Theme/WorklogInfoPillBuilder.cs. This
            // pill is declared in Main.axaml (long-lived, only its text changes), so it is
            // restyled in place rather than rebuilt. It used to be styled here by hand at 2px,
            // which is exactly the drift that class now prevents.
            //
            // The status WORD lives in the pill, so WorklogJobStatusText below carries only the
            // counts and the start date - otherwise "Open" would read twice, once in each.
            Handlers.Theming.WorklogInfoPillBuilder.ApplyStatePillVisual(
                this.WorklogJobStatusPill,
                this.WorklogJobStatusDot,
                this.WorklogJobStatusPillText,
                activeWorkbook.Status);

            // A CLOSED workbook reports when it finished; an open one reports when it started -
            // the whole rule, word and date together, lives in the shared
            // WorklogEntryScope.FormatWorkbookDateLine, which the Workbooks tab's own cards use
            // too. This bar had the rule and the card had a hardcoded "started" against StartDate,
            // so the same workbook read "ended 2026-September-6" here and "started 2026-September-6"
            // on its card a row below. See that method for why "ended" rather than "closed", and
            // for the no-EndDate fallback.
            string dateLine = WorklogEntryScope.FormatWorkbookDateLine(
                activeWorkbook.Status,
                activeWorkbook.StartDate,
                activeWorkbook.EndDate);

            if (activeWorkbook.WorklogCount == 0)
            {
                this.WorklogJobStatusText.Text = $"No worklogs yet · {dateLine}";
                return;
            }

            string worklogWord = activeWorkbook.WorklogCount == 1 ? "worklog" : "worklogs";
            this.WorklogJobStatusText.Text =
                $"{activeWorkbook.WorklogCount} {worklogWord} · {dateLine}";
        }

        // ###########################################################################################
        // Opens the "Create new workbook" dialog for the currently selected board and ACTIVATES the
        // workbook it creates.
        //
        // ActivateWorkbook, not a bare RefreshWorklogBar: "which workbook is active" used to be
        // "the board's newest", which a just-created workbook always was, so a plain refresh was
        // enough. It is now ResolveActiveWorkbookForBoard, which prefers a previously-saved
        // ActiveWorkbookIdByBoard entry whenever that id still names a real workbook - so after the
        // user has ever clicked a card in the Workbooks tab, a refresh alone would leave the bar,
        // "Show worklogs" and "Add worklog" pointing at the OLD workbook and write the next drawn
        // entry into it rather than into the one just created and named.
        // ###########################################################################################
        private async void OnWorklogCreateWorkbookClick(object? sender, RoutedEventArgs e)
        {
            // "async void" - see OnBoardSelectionChanged. The prologue here is not exception-free
            // either: Initialize loads XAML, and ShowDialog can throw synchronously.
            try
            {
                string boardKey = this.GetCurrentBoardKey();
                if (string.IsNullOrWhiteSpace(boardKey))
                    return;

                var dialog = new CreateWorkbookWindow();
                dialog.Initialize(boardKey);

                var record = await dialog.ShowDialog<WorkbookRecord?>(this);
                if (record == null)
                    return;

                this.ActivateWorkbook(boardKey, record.Id);
            }
            catch (Exception ex)
            {
                Logger.Critical($"Creating a workbook failed: [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Switches to the Schematics tab and enters worklog entry-drawing mode for the ACTIVE
        // workbook - the one ResolveActiveWorkbookForBoard resolves, which the bar is already
        // showing. Not the Open-only lookup, so this keeps working on a closed workbook - the entry
        // it adds is Open, which reopens the workbook anyway.
        // ###########################################################################################
        private void OnWorklogAddEntryClick(object? sender, RoutedEventArgs e)
        {
            var activeWorkbook = ResolveActiveWorkbookForBoard(this.GetCurrentBoardKey());
            if (activeWorkbook == null)
                return;

            this.SwitchToSchematicsTab();

            if (!this.TabSchematicsControl.BeginWorklogEntryMode(activeWorkbook.Id))
                return;

            this.WorklogAddEntryButton.IsEnabled = false;
            this.WorklogCancelEntryButton.IsVisible = true;
        }

        // ###########################################################################################
        // Cancels the in-progress worklog entry-drawing mode. TabSchematics calls back into
        // ResetWorklogEntryModeButtons() once it has actually torn the mode down, so the buttons
        // stay in sync whether cancellation came from here, Escape, or the entry editor closing.
        // ###########################################################################################
        private void OnWorklogCancelEntryClick(object? sender, RoutedEventArgs e)
        {
            this.TabSchematicsControl.CancelWorklogEntryMode();
        }

        // ###########################################################################################
        // Toggles the "Show worklogs" list view for the workbook the bar is showing, and saves the
        // checked state as the user's default for next time. TabSchematics scopes what it actually
        // draws to the schematic currently on screen - see TabSchematics.Worklog.cs's
        // RefreshWorklogEntriesListOverlay for why. Uses the same latest-workbook lookup the bar
        // does, so a closed workbook's entries stay viewable.
        //
        // Only a real user toggle is saved: RefreshWorklogBar seeds the checkbox programmatically
        // and suppresses the save, because the value it seeds is derived from the board on screen
        // rather than from the user's intent. It applies the overlay itself, so returning early
        // here loses nothing.
        // ###########################################################################################
        private void OnWorklogShowEntriesCheckedChanged(object? sender, RoutedEventArgs e)
        {
            if (this._suppressWorklogShowEntriesSave)
                return;

            bool isChecked = this.WorklogShowEntriesCheckBox.IsChecked == true;
            UserSettings.WorklogShowEntriesChecked = isChecked;

            var activeWorkbook = ResolveActiveWorkbookForBoard(this.GetCurrentBoardKey());
            int workbookId = isChecked && activeWorkbook != null ? activeWorkbook.Id : 0;

            this._worklogShowEntriesWorkbookId = workbookId;
            this.TabSchematicsControl.SetShowWorklogEntriesList(isChecked && activeWorkbook != null, workbookId);
        }

        // ###########################################################################################
        // Makes the "Show worklogs" text act like part of the checkbox it labels, since a bare
        // TextBlock does not react to clicks on its own. Toggling the checkbox itself fires
        // IsCheckedChanged, which does the actual work and persists the preference.
        // ###########################################################################################
        private void OnWorklogShowEntriesLabelPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            this.WorklogShowEntriesCheckBox.IsChecked = !(this.WorklogShowEntriesCheckBox.IsChecked == true);
        }

        // ###########################################################################################
        // Restores the worklog bar's "Add worklog" / "Cancel" buttons to their idle state.
        // Called by TabSchematics whenever worklog entry-drawing mode ends, regardless of trigger.
        // ###########################################################################################
        public void ResetWorklogEntryModeButtons()
        {
            this.WorklogAddEntryButton.IsEnabled = true;
            this.WorklogCancelEntryButton.IsVisible = false;
        }
    }
}
