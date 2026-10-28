using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE WORKSPACE (CRT 4.0 step 2, 2026-10-25 - the plan: "One tab instead of Contribute and
    // Drafts ... A My drafts list in the header switches between boards ... The table view is one view
    // of the same draft").
    //
    // Owns: the board the tab is about - the one shown in CRT's drop-down lists (ShowBoard, from
    // Main); its header - the draft's own row, or what the board is while it has no draft; the left
    // column - the board's parts with their counts (ContributeBoardParts), its components with a search
    // and the draft's "Changed" mark, "Add component" and "Table view"; and "My drafts", which shows
    // the board picked, everywhere in CRT.
    //
    // *** A FIRST CHANGE MAKES THE DRAFT. *** Every action on the left works on the board's draft,
    // and on a board with none it makes one first (Main.EnsureDraftAsync - the published file
    // copied, as every editor's first save does), so nobody has to "start a draft" before working.
    //
    // *** NOT YET THE PLAN'S WHOLE SCREEN. *** A component opens in the component pane, "Add
    // component" opens its form for a new one, and the board's pictures, files and links, credits and
    // important signals open there too (TabDrafts.Component.cs), the pictures added, named, matched
    // to their KiCad views, ordered and removed on their cards since step 3 - but KiCad data still
    // opens its window.
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        private HardwareBoardEntry? thisCurrentEntry;
        private BoardData? thisCurrentBoard;
        private string thisCurrentRegion = string.Empty;
        private IReadOnlySet<string>? thisDraftedLabels;

        // "My drafts" falls back to its placeholder after every pick - set from code, not a pick.
        private bool thisResettingMyDrafts;

        // For tests, in place of Main: showing a board picked in "My drafts", and making a board's
        // first draft.
        internal Action<HardwareBoardEntry>? SelectBoardOverrideForTests { get; set; }

        internal Func<HardwareBoardEntry, Task<bool>>? EnsureDraftOverrideForTests { get; set; }

        // The board the tab is about - none before a board is shown.
        internal HardwareBoardEntry? CurrentEntry => this.thisCurrentEntry;

        // ###########################################################################################
        // The board shown in CRT's drop-down lists, as it reads now (its draft, when it has one), with
        // the region shown and the components its draft changed - from Main, after every board load.
        // A table open on ANOTHER board's draft is closed (what it holds saved first): the table is a
        // view of the board the tab is about.
        // ###########################################################################################
        internal void ShowBoard(HardwareBoardEntry? entry, BoardData? board, string region, IReadOnlySet<string>? draftedLabels)
        {
            // Why no draft was made is about the board it was pressed on - never left standing under
            // another one (code review, 2026-10-09).
            if (entry is null || this.thisCurrentEntry is null || !TabDrafts.SameBoard(entry, this.thisCurrentEntry))
            {
                this.ShowDraftProblem(null);
            }

            this.thisCurrentEntry = entry;
            this.thisCurrentBoard = entry is null ? null : board;
            this.thisCurrentRegion = region ?? string.Empty;
            this.thisDraftedLabels = draftedLabels;

            // Main's marks are up to date with the board it has just read.
            this.thisChangedHere.Clear();

            // The start panel's lists are worked out once below - unless the start panel coming back
            // as what was open closes has done it already (code review, 2026-10-26: twice per board).
            int readinessAsked = this.thisReadinessAsk;

            if (this.thisTableEntry is { } tableEntry && (entry is null || !TabDrafts.SameBoard(tableEntry, entry)))
            {
                this.CloseTableForAnotherBoard(tableEntry);
            }

            // The component pane too - what it holds handed to its draft and saved as it goes.
            if (this.thisComponentEntry is { } paneEntry && (entry is null || !TabDrafts.SameBoard(paneEntry, entry)))
            {
                this.CloseComponent();
            }

            this.ShowWorkspace();

            // How ready a new board is, or what an existing one lacks (TabDrafts.Readiness.cs).
            if (this.thisReadinessAsk == readinessAsked)
            {
                this.ShowReadiness();
            }
        }

        // ###########################################################################################
        // *** ANOTHER BOARD IS SHOWN ALREADY, SO THIS CLOSE CANNOT BE CALLED OFF (code review, 2026-10-26).
        // *** It asked when the save failed, and Cancel left the table open on the previous board's
        // draft under the new board's header - typing went into a draft not shown. What the table
        // holds goes into its draft (CollectHeldChanges) and a save of it starts; one that cannot be
        // made waits in the draft held in memory - its header says so when its board is shown again,
        // and quitting asks (OtherUnsavedDrafts).
        // ###########################################################################################
        private void CloseTableForAnotherBoard(HardwareBoardEntry tableEntry)
        {
            DraftSession? session = DraftManager.Sessions.Find(DraftManager.DraftsRoot, tableEntry.ExcelDataFile);
            session?.CollectHeldChanges(DraftSaveWording.ReadAgainStep);

            this.CloseTableWithoutAsking();

            if (session is { HasUnsavedChanges: true })
            {
                _ = TabDrafts.SaveLeftDraftAsync(session);
            }

            this.LetGoOfDraftMadeOnlyToLook(tableEntry);
        }

        // ###########################################################################################
        // *** A DRAFT MADE ONLY TO OPEN SOMETHING ON THE BOARD GOES AGAIN when that closes with nothing
        // in it (code review, 2026-10-26). *** Opening a component - "Edit this component" is on
        // everybody's right-click menu - a board part or the table makes the board's draft first
        // (OnCurrentDraftAsync); looked at and closed, it stayed - the tab's title red, an empty draft in
        // "My drafts". Only a draft such an opening made this run, with nothing waiting to be saved, no
        // change against the published board and nothing else open on it - never a new board's.
        // ###########################################################################################
        private string? thisDraftMadeToLook;

        private void LetGoOfDraftMadeOnlyToLook(HardwareBoardEntry entry)
        {
            if (!string.Equals(this.thisDraftMadeToLook, entry.ExcelDataFile, StringComparison.OrdinalIgnoreCase) ||
                this.IsTableOpenFor(entry) || this.IsComponentOpenFor(entry))
            {
                return;
            }

            this.thisDraftMadeToLook = null;

            if (DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile) is { HasUnsavedChanges: true })
            {
                return;
            }

            DraftStatus? status = DraftStatusReader.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, entry.ExcelDataFile, entry.IsPublished);

            if (status is null || status.IsNewBoard ||
                DraftStatusReader.CountChangesCached(status) > 0 ||
                DraftStatusReader.CountOtherChangesCached(status, DraftManager.DraftsRoot).Any ||
                TabDrafts.HoldsAnotherOrder(entry, status))
            {
                return;
            }

            Logger.Info($"The draft of [{entry.ExcelDataFile}] was made only to open something on it, and holds nothing - let go of again");
            this.DiscardConfirmed(entry);
        }

        // ###########################################################################################
        // Whether the draft holds its rows in another order than the published board - pictures moved
        // up or down, a row dragged in the table: no change the counts above see, but the
        // contributor's, kept (BoardRowOrder; code review, 2026-10-28). The draft as held in memory
        // when there is a session, else its file.
        // ###########################################################################################
        private static bool HoldsAnotherOrder(HardwareBoardEntry entry, DraftStatus status)
        {
            BoardData? draft = DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile)?.Content.Board
                ?? (File.Exists(status.WorkbookPath) ? BoardDataReader.ReadWorkbookUncached(status.WorkbookPath) : null);
            BoardData? published = File.Exists(status.PublishedWorkbookPath) ? BoardDataReader.ReadWorkbookUncached(status.PublishedWorkbookPath) : null;

            return draft is not null && published is not null && BoardRowOrder.DiffersFrom(published, draft);
        }

        // A draft no editor shows any more, saved off the UI thread - a failure is the session's, said
        // where it is shown next.
        private static async Task SaveLeftDraftAsync(DraftSession session)
        {
            DraftWorkbookEditOutcome outcome = await session.SaveAsync();

            if (outcome != DraftWorkbookEditOutcome.Saved)
            {
                Logger.Warning($"The draft of [{session.ExcelDataFile}] could not be saved as its table closed ({outcome}) - its change waits in memory");
            }
        }

        private static bool SameBoard(HardwareBoardEntry first, HardwareBoardEntry second) =>
            string.Equals(first.ExcelDataFile, second.ExcelDataFile, StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // Everything the workspace shows, read again: the header, the left column, "My drafts".
        // Cheap - no file is read - so it runs after every RefreshDrafts too.
        // ###########################################################################################
        private void ShowWorkspace()
        {
            this.ShowHeader();
            this.ShowBoardParts();
            this.ShowComponents();
            this.ShowMyDrafts();

            bool hasBoard = this.thisCurrentBoard is not null;
            this.WorkspaceNav.IsEnabled = hasBoard;
            this.ComponentPane.IsVisible = !this.IsTableOpen && this.IsComponentOpen;
            this.ShowStartPanel(!this.IsTableOpen && !this.IsComponentOpen);
            this.TableViewButton.Content = this.thisCurrentEntry is { } entry && this.IsTableOpenFor(entry) ? "Close table view" : "Table view";

            string revision = this.thisCurrentBoard?.RevisionDate?.Trim() ?? string.Empty;
            this.PublishedOnText.Text = revision.Length == 0 ? string.Empty : $"The board's data was last published on {revision}.";
            this.PublishedOnText.IsVisible = revision.Length > 0 && this.thisCurrentEntry?.IsDraftOnly != true;
        }

        // ###########################################################################################
        // The header: the draft's row of the board the table is open on, else of the board shown -
        // or, with no draft, what the board is and that the first change makes one.
        // ###########################################################################################
        private void ShowHeader()
        {
            HardwareBoardEntry? headerEntry = this.thisTableEntry ?? this.thisCurrentEntry;

            DraftListItem? row = headerEntry is null
                ? null
                : this.Drafts.FirstOrDefault(item => string.Equals(item.ExcelDataFile, headerEntry.ExcelDataFile, StringComparison.OrdinalIgnoreCase));

            // The same row object as the list's, so a row kept by KeepUnchangedRows stays the very
            // control being pressed.
            if (!ReferenceEquals(this.CurrentDraftRow.Content, row))
            {
                this.CurrentDraftRow.Content = row;
            }

            this.CurrentDraftRow.IsVisible = row is not null;
            this.NoDraftPanel.IsVisible = row is null;

            this.NoDraftBoardText.Text = headerEntry?.ToString() ?? "No board shown";
            this.NoDraftText.Text = ContributeBoardParts.NoDraftLine(hasBoard: headerEntry is not null);
        }

        // The board's parts on the left, each with how many it holds.
        private void ShowBoardParts()
        {
            ContributeBoardParts parts = ContributeBoardParts.Of(this.BoardForList());

            this.PicturesCountText.Text = ContributeBoardParts.CountText(parts.Pictures);
            this.SignalsCountText.Text = ContributeBoardParts.CountText(parts.ImportantSignals);
            this.FilesCountText.Text = ContributeBoardParts.CountText(parts.FilesAndLinks);
            this.CreditsCountText.Text = ContributeBoardParts.CountText(parts.Credits);
        }

        // ###########################################################################################
        // The components of the board shown in the region shown - the list on CRT's left with its
        // "drafted" mark, by the one rule that builds it (ComponentListBuilder) - narrowed by the
        // search box's words. With the component pane or the table open, the draft as they left it
        // (BoardForList), so a name typed in the pane is in the list at once.
        // ###########################################################################################
        private void ShowComponents()
        {
            BoardData? board = this.BoardForList();

            List<ComponentListItem> items = board is null
                ? []
                : ComponentListBuilder.BuildComponentItems(
                    board,
                    this.thisCurrentRegion,
                    categoryFilter: null,
                    this.ComponentSearchBox.Text ?? string.Empty,
                    this.DraftedLabelsForList());

            this.ComponentsItemsControl.ItemsSource = items;

            this.NoComponentsText.Text = ContributeBoardParts.NoComponentsLine(
                hasBoard: board is not null,
                searching: !string.IsNullOrWhiteSpace(this.ComponentSearchBox.Text));
            this.NoComponentsText.IsVisible = items.Count == 0;
        }

        internal IReadOnlyList<ComponentListItem> ComponentsShownForTests =>
            this.ComponentsItemsControl.ItemsSource as IReadOnlyList<ComponentListItem> ?? [];

        private void OnComponentSearchChanged(object? sender, TextChangedEventArgs e) => this.ShowComponents();

        // ---- "My drafts" --------------------------------------------------------------------------

        private void ShowMyDrafts()
        {
            this.MyDraftsComboBox.PlaceholderText = ContributeBoardParts.MyDraftsLine(this.Drafts.Count);
            this.MyDraftsComboBox.IsEnabled = this.Drafts.Count > 0;
        }

        private void OnMyDraftsSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisResettingMyDrafts || this.MyDraftsComboBox.SelectedItem is not DraftListItem picked)
            {
                return;
            }

            // Back to its placeholder: the list is a way to go somewhere, not a second board name.
            this.thisResettingMyDrafts = true;
            this.MyDraftsComboBox.SelectedItem = null;
            this.thisResettingMyDrafts = false;

            this.ShowDraftOf(picked.Entry);
        }

        // ###########################################################################################
        // Shows the board of a draft picked in "My drafts" - in CRT's drop-down lists, so the whole of
        // CRT is on it; the board's load then shows it here (ShowBoard).
        // ###########################################################################################
        internal void ShowDraftOf(HardwareBoardEntry entry)
        {
            if (this.SelectBoardOverrideForTests is { } select)
            {
                select(entry);
                return;
            }

            this.thisMainWindow?.SelectBoard(entry);
        }

        // ---- the left column's actions ------------------------------------------------------------

        // ###########################################################################################
        // `action` on the draft of the board shown - made first when it has none (Main.EnsureDraftAsync:
        // the published file copied, the board read again). Nothing when no draft could be made; Main
        // says why on the tab's own line.
        // ###########################################################################################
        internal async Task OnCurrentDraftAsync(Func<HardwareBoardEntry, Task> action)
        {
            if (this.thisCurrentEntry is not { } entry || string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                return;
            }

            if (!DraftBoardSource.HasDraft(DraftManager.DraftsRoot, entry.ExcelDataFile))
            {
                bool made = this.EnsureDraftOverrideForTests is { } ensure
                    ? await ensure(entry)
                    : this.thisMainWindow is not null && await this.thisMainWindow.EnsureDraftAsync(entry);

                if (!made)
                {
                    return;
                }

                // Let go of again should what is opened close with nothing in it (LetGoOfDraftMadeOnlyToLook).
                this.thisDraftMadeToLook = entry.ExcelDataFile;
                this.RefreshDrafts();
            }

            await action(entry);
        }

        // The table view on one of its sheets - a board part the plan's screens are not built for yet.
        internal Task ShowTableOnSheetAsync(string sheetName) =>
            this.OnCurrentDraftAsync(async entry =>
            {
                await this.ShowTableAsync(entry);
                this.TableEditor.SelectSheetNamed(sheetName);
            });

        // "Table view", or "Close table view" while it is open on this board.
        internal Task ToggleTableViewAsync() =>
            this.thisCurrentEntry is { } entry && this.IsTableOpenFor(entry)
                ? this.CloseTableAsync()
                : this.OnCurrentDraftAsync(this.ShowTableAsync);

        // The board's pictures in the pane - added, named, matched to their KiCad views, ordered, removed
        // and coloured there (slice 4b; step 3 moved the "Schematic images" window in).
        private void OnPicturesClick(object? sender, RoutedEventArgs e) => this.ManagePictures();

        private void ManagePictures() =>
            this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.Pictures), "Pictures");

        private void OnKiCadClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(() => this.OnCurrentDraftAsync(entry => this.ManageFilesAsync(entry)), "KiCad data");

        // The board's own lists, in the pane on the right (slice 4a - TabDrafts.Component.cs).
        private void OnImportantSignalsClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.ImportantSignals), "Important signals");

        private void OnFilesAndLinksClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.FilesAndLinks), "Files and links");

        private void OnCreditsClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.Credits), "Credits");

        private void OnTableViewClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(this.ToggleTableViewAsync, "Table view");

        // A component, in the component pane on the right (TabDrafts.Component.cs).
        private void OnComponentClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ComponentListItem item } && item.BoardLabel.Length > 0)
            {
                this.RunGuarded(() => this.OpenComponentAsync(item.BoardLabel), "Open component");
            }
        }

        // A new component, named in the component pane's form (TabDrafts.Component.cs).
        private void OnAddComponentClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(this.AddComponentAsync, "Add component");

        // ###########################################################################################
        // A click's work, started and watched: nothing may leave a click handler that awaits (App lets
        // such an exception end CRT - AsyncVoidHandlersGuardTests), so it is logged instead.
        // ###########################################################################################
        private async void RunGuarded(Func<Task> work, string what)
        {
            try
            {
                await work();
            }
            catch (Exception ex)
            {
                Logger.Critical($"The Contribute tab's [{what}] failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // Why "Table view", a board part or "Add a new board" made no draft - or nothing, to clear
        // it. Said about the board it was pressed on: a new board shown clears it.
        // ###########################################################################################
        internal void ShowDraftProblem(string? problem)
        {
            this.DraftProblemText.Text = problem ?? string.Empty;
            this.DraftProblemText.IsVisible = !string.IsNullOrWhiteSpace(problem);
        }

        internal string? DraftProblemForTests => this.DraftProblemText.IsVisible ? this.DraftProblemText.Text : null;
    }
}
