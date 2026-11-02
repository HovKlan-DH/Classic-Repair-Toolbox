using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    // Main); its header - the draft's own row, or what the board is while it has no draft - with "Table
    // view" at its end; the left column - the board's parts with their counts (ContributeBoardParts),
    // and below them its components with a search, the draft's "Draft" mark and "Add component" -
    // hidden while the table is open; and "My drafts", which shows the board picked, everywhere in CRT.
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

            bool anotherBoard = entry is null || this.thisCurrentEntry is null || !TabDrafts.SameBoard(entry, this.thisCurrentEntry);

            // "Board name" is a view of the board shown - another one's page is no use.
            if (anotherBoard)
            {
                this.thisBoardNamePicked = false;
            }

            this.thisCurrentEntry = entry;
            this.thisCurrentBoard = entry is null ? null : board;
            this.thisCurrentRegion = region ?? string.Empty;
            this.thisDraftedLabels = draftedLabels;

            // What is open on the right shows the region CRT shows - either switch picks it (2026-11-01).
            if (!anotherBoard)
            {
                this.ComponentPane.UseRegion(this.thisCurrentRegion);
            }

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

            // And "KiCad data" (TabDrafts.KiCad.cs) - its changes still being written end by themselves.
            if (this.thisKiCadEntry is { } kiCadEntry && (entry is null || !TabDrafts.SameBoard(kiCadEntry, entry)))
            {
                this.CloseKiCadNow();
            }

            // The widths dragged for this board (TabDrafts.Splitters.cs) - not on a board read again,
            // and only once the previous board's pane is closed (code review, 2026-10-30: set before, the
            // new board's preview width read the old board's pictures again just as they went).
            if (anotherBoard)
            {
                this.ShowWidths();
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
        //
        // *** AND UNTIL ITS FIRST CHANGE IT IS NOT SHOWN AS A DRAFT AT ALL (owner request, 2026-11-01: "It
        // should not show it as draft until I have done my first edit, whatever that might be"). *** No
        // row, no "Discard draft", no red tab title: RefreshDrafts leaves it out (ShowsAsDraft) until the
        // draft held in memory takes a change (thisDraftChangedSinceLook, by its Changed event - and by
        // its files, for a write the session does not see).
        // ###########################################################################################
        private string? thisDraftMadeToLook;

        private bool thisDraftChangedSinceLook;

        private DraftSession? thisDraftWatchedForChange;

        internal bool IsDraftOnlyLookedAtForTests(string identity) => this.IsUntouchedLook(identity);

        private bool IsUntouchedLook(string identity) =>
            !this.thisDraftChangedSinceLook && string.Equals(this.thisDraftMadeToLook, identity, StringComparison.OrdinalIgnoreCase);

        private void ForgetDraftMadeToLook()
        {
            this.thisDraftMadeToLook = null;
            this.thisDraftChangedSinceLook = false;
            this.StopWatchingDraftForChange();
        }

        private void StopWatchingDraftForChange()
        {
            if (this.thisDraftWatchedForChange is { } watched)
            {
                watched.Changed -= this.OnLookedAtDraftChanged;
                this.thisDraftWatchedForChange = null;
            }
        }

        private async Task WatchDraftForFirstChangeAsync(string identity)
        {
            string draftsRoot = DraftManager.DraftsRoot;
            DraftSession? session = await Task.Run(() => DraftManager.Sessions.Get(draftsRoot, identity));

            if (session is not null && this.IsUntouchedLook(identity))
            {
                this.StopWatchingDraftForChange();
                this.thisDraftWatchedForChange = session;
                session.Changed += this.OnLookedAtDraftChanged;
            }
        }

        // Its first change: shown as the draft it now is - posted, as a direct write raises this off the UI
        // thread, and an edit raises it while its own change is still being taken in.
        private void OnLookedAtDraftChanged(object? sender, DraftSessionChange change) =>
            Dispatcher.UIThread.Post(
                () =>
                {
                    if (!ReferenceEquals(sender, this.thisDraftWatchedForChange))
                    {
                        return;
                    }

                    this.StopWatchingDraftForChange();
                    this.thisDraftChangedSinceLook = true;

                    if (this.thisMainWindow is { } main)
                    {
                        main.ApplyContributeTabState();
                    }
                    else
                    {
                        this.RefreshDrafts();
                    }
                },
                DispatcherPriority.Background);

        // ###########################################################################################
        // Whether this draft's row is shown: unless it was made only to look and has not changed. A write
        // the session does not see (a KiCad import's files) is found on disk as the list is read again.
        // ###########################################################################################
        private bool ShowsAsDraft(HardwareBoardEntry entry, DraftStatus status)
        {
            if (!this.IsUntouchedLook(entry.ExcelDataFile))
            {
                return true;
            }

            bool changed =
                DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile) is { HasUnsavedChanges: true } or { CanUndo: true } ||
                DraftStatusReader.CountChangesCached(status) > 0 ||
                DraftStatusReader.CountOtherChangesCached(status, DraftManager.DraftsRoot).Any;

            if (changed)
            {
                this.StopWatchingDraftForChange();
                this.thisDraftChangedSinceLook = true;
            }

            return changed;
        }

        // ###########################################################################################
        // Quitting with a draft only looked at still open on the right: gone with the run, or the next
        // run would show it as a draft of the contributor's. Only by its files - CRT is closing, and shows
        // nothing again.
        // ###########################################################################################
        internal void LetGoOfUntouchedDraftOnQuit()
        {
            if (this.thisCurrentEntry is { } entry && this.IsUntouchedLook(entry.ExcelDataFile) && TabDrafts.HoldsNothing(entry))
            {
                Logger.Info($"The draft of [{entry.ExcelDataFile}] was only looked at - let go of as CRT closes");
                this.ForgetDraftMadeToLook();
                _ = this.DiscardOverrideForTests?.Invoke(entry.ExcelDataFile) ?? DraftManager.DiscardDraft(entry.ExcelDataFile);
            }
        }

        // ###########################################################################################
        // The same just before an update's restart - which can still fail, and CRT then carries on (code
        // review, 2026-11-01). So not only by its files: what is open on the draft closes with it
        // (DiscardConfirmed), or the next change would find no folder to save into.
        // ###########################################################################################
        internal void LetGoOfUntouchedDraftBeforeRestart()
        {
            if (this.thisCurrentEntry is { } entry && this.IsUntouchedLook(entry.ExcelDataFile) && TabDrafts.HoldsNothing(entry))
            {
                Logger.Info($"The draft of [{entry.ExcelDataFile}] was only looked at - let go of before the restart");
                this.DiscardConfirmed(entry);
            }
        }

        // ###########################################################################################
        // The board something is OPENING on right now - the pane, the table or "KiCad data", each of
        // which first closes what was open on the right (2026-10-30). That close is a hand-over, not
        // the draft left alone: a draft made only to look was let go of on the way, and what was being
        // opened found no draft. While a board is being opened on, its draft is not let go of; the
        // marker stays, so it is when the newly opened thing closes with nothing in it.
        // ###########################################################################################
        private HardwareBoardEntry? thisOpeningOn;

        private OpeningScope OpeningOn(HardwareBoardEntry entry)
        {
            var scope = new OpeningScope(this, this.thisOpeningOn);
            this.thisOpeningOn = entry;
            return scope;
        }

        private sealed class OpeningScope(TabDrafts tab, HardwareBoardEntry? before) : IDisposable
        {
            public void Dispose() => tab.thisOpeningOn = before;
        }

        private void LetGoOfDraftMadeOnlyToLook(HardwareBoardEntry entry)
        {
            if (!string.Equals(this.thisDraftMadeToLook, entry.ExcelDataFile, StringComparison.OrdinalIgnoreCase) ||
                this.IsTableOpenFor(entry) || this.IsComponentOpenFor(entry) || this.IsKiCadOpenFor(entry) ||
                (this.thisOpeningOn is { } opening && TabDrafts.SameBoard(opening, entry)))
            {
                return;
            }

            this.ForgetDraftMadeToLook();

            if (!TabDrafts.HoldsNothing(entry))
            {
                return;
            }

            Logger.Info($"The draft of [{entry.ExcelDataFile}] was made only to open something on it, and holds nothing - let go of again");
            this.DiscardConfirmed(entry);
        }

        // Whether the board's draft holds nothing: nothing waiting to be saved, no change against the
        // published board - rows, files, calibrations, their order and its regions alike (DraftOtherChanges;
        // nothing waiting, its file is what it holds) - and never a new board's.
        private static bool HoldsNothing(HardwareBoardEntry entry)
        {
            if (DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile) is { HasUnsavedChanges: true })
            {
                return false;
            }

            DraftStatus? status = DraftStatusReader.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, entry.ExcelDataFile, entry.IsPublished);

            return status is not null && !status.IsNewBoard &&
                   DraftStatusReader.CountChangesCached(status) == 0 &&
                   !DraftStatusReader.CountOtherChangesCached(status, DraftManager.DraftsRoot).Any;
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

        // Board status shows while nothing else is open on the right - ONE thing there at a time.
        private bool NothingElseOpenOnTheRight =>
            !this.IsTableOpen && !this.IsComponentOpen && !this.IsKiCadOpen && !this.thisBoardNamePicked;

        // ###########################################################################################
        // Everything the workspace shows, read again: the header, the left column, "My drafts".
        // Cheap - no file is read - so it runs after every RefreshDrafts too.
        // ###########################################################################################
        private void ShowWorkspace()
        {
            this.ShowHeader();
            this.ShowBoardParts();
            this.ShowComponents();
            this.ShowNavSelection();
            this.ShowMyDrafts();

            bool hasBoard = this.thisCurrentBoard is not null;
            this.WorkspaceNav.IsEnabled = hasBoard;
            this.TableViewButton.IsEnabled = hasBoard;
            this.ComponentPane.IsVisible = !this.IsTableOpen && this.IsComponentOpen;
            this.KiCadPanelHost.IsVisible = !this.IsTableOpen && !this.IsComponentOpen && this.IsKiCadOpen;
            this.BoardNamePanel.IsVisible = this.thisBoardNamePicked && !this.IsTableOpen && !this.IsComponentOpen && !this.IsKiCadOpen;
            this.ShowStartPanel(this.NothingElseOpenOnTheRight);
            this.ShowWorkspaceColumns();
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
            this.HeaderPanel.Classes.Set("hasDraft", row is not null);

            // The toolbar's buttons act on the row shown - the same buttons, only their context moves,
            // so a press survives the list being read again (KeepUnchangedRows' reason, here for free).
            if (!ReferenceEquals(this.HeaderToolbar.DataContext, row))
            {
                this.HeaderToolbar.DataContext = row;
            }

            this.HistoryButtons.IsVisible = row is not null;
            this.DraftActions.IsVisible = row is not null;

            // The draft is local - said while there is one (owner request, 2026-10-30).
            this.LocalDraftNotice.IsVisible = row is not null;
            this.LocalDraftNoticeText.Text = ContributeBoardParts.LocalDraftNotice(isNewBoard: headerEntry?.IsDraftOnly == true);

            this.NoDraftText.Text = ContributeBoardParts.NoDraftLine(hasBoard: headerEntry is not null);
        }

        internal bool IsLocalDraftNoticeShownForTests => this.LocalDraftNotice.IsVisible;

        internal string LocalDraftNoticeForTests => this.LocalDraftNoticeText.Text ?? string.Empty;

        // The board's parts on the left, each with how many it holds - the components counted by
        // ShowComponents, as the list shows them.
        private void ShowBoardParts()
        {
            ContributeBoardParts parts = ContributeBoardParts.Of(this.BoardForList());

            this.PicturesCountText.Text = ContributeBoardParts.CountText(parts.Pictures);
            this.SignalsCountText.Text = ContributeBoardParts.CountText(parts.ImportantSignals);
            this.FilesCountText.Text = ContributeBoardParts.CountText(parts.Files);
            this.LinksCountText.Text = ContributeBoardParts.CountText(parts.Links);
            this.CreditsCountText.Text = ContributeBoardParts.CountText(parts.Credits);
            this.CategoriesCountText.Text = ContributeBoardParts.CountText(parts.Categories);
            this.RegionsCountText.Text = ContributeBoardParts.RegionsText(this.BoardForList());

            this.ShowKiCadCount();
        }

        // ###########################################################################################
        // HOW MANY KiCad FILES THE BOARD HAS (owner request, 2026-10-30: "show the amount of files
        // imported/used (if any)") - the names its board file lists under "KiCad data/", the draft's
        // when it has one (its session's list, and a file copied in and not listed yet), else the
        // published board's. Read off the UI thread - a board file is parsed for its list - and shown
        // when it arrives, only while the board it was read for is still the one shown; the last count
        // stays meanwhile for that board, and "none" is not said before it is known.
        // ###########################################################################################
        private string? thisKiCadCountBoard;

        private int thisKiCadCountAsk;

        // What the count was last asked from, for a board with a draft held in memory: the same again
        // is not counted again (code review, 2026-10-30: every refresh of the tab - each edit in the
        // pane - walked the draft's KiCad folder for it). Forgotten when "KiCad data" catches up
        // (CatchUpAfterBoardFiles): a file it copied in and could not list is in the folder, not the list.
        private (string Identity, string Published, List<string> Listed)? thisKiCadCountedFrom;

        internal void ForgetKiCadCount() => this.thisKiCadCountedFrom = null;

        private void ShowKiCadCount()
        {
            HardwareBoardEntry? entry = this.thisCurrentEntry;

            if (entry is null || string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                this.thisKiCadCountBoard = null;
                this.KiCadCountText.Text = string.Empty;
                return;
            }

            if (!string.Equals(this.thisKiCadCountBoard, entry.ExcelDataFile, StringComparison.OrdinalIgnoreCase))
            {
                this.KiCadCountText.Text = string.Empty;
            }

            string identity = entry.ExcelDataFile;
            string draftsRoot = DraftManager.DraftsRoot;
            DraftSession? session = DraftManager.Sessions.Find(draftsRoot, identity);
            List<string>? listed = session?.ListedFiles.Keys.ToList();
            string published = this.PublishedBoardFileOf(entry);

            // Without a draft in memory nothing is edited, and the files on disk decide - counted each time.
            if (listed is not null &&
                this.thisKiCadCountedFrom is { } last &&
                string.Equals(last.Identity, identity, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(last.Published, published, StringComparison.Ordinal) &&
                last.Listed.SequenceEqual(listed, StringComparer.Ordinal))
            {
                return;
            }

            this.thisKiCadCountedFrom = null;

            int ask = ++this.thisKiCadCountAsk;
            this.KiCadCountsForTests++;

            this.RunGuarded(async () =>
            {
                int count = await Task.Run(() => TabDrafts.KiCadFileNamesOf(draftsRoot, identity, listed, published).Count);

                if (ask == this.thisKiCadCountAsk)
                {
                    this.thisKiCadCountBoard = identity;
                    this.KiCadCountText.Text = ContributeBoardParts.CountText(count);

                    // Kept only once it is shown - a read that failed is asked again next time.
                    this.thisKiCadCountedFrom = listed is null ? null : (identity, published, listed);
                }
            }, "KiCad files counted");
        }

        // The KiCad files of a board: the draft's list - handed over from its session, else read from
        // its board file - with the files copied in and not listed yet; with no draft, the published
        // board's list.
        private static List<string> KiCadFileNamesOf(string draftsRoot, string identity, IReadOnlyList<string>? listed, string publishedBoardFile)
        {
            string folder = DraftFolderLayout.GetBoardFolder(draftsRoot, identity);
            string draftBoardFile = Path.Combine(folder, DataTreeLayout.BoardFileName);

            IEnumerable<string>? names = listed
                ?? (File.Exists(draftBoardFile) ? TabDrafts.KiCadNamesListedIn(draftBoardFile) : null);

            if (names is not null)
            {
                return ContributeBoardParts.KiCadFileNames(names.Concat(DraftFiles.ArrivedKiCadNames(folder)));
            }

            return File.Exists(publishedBoardFile)
                ? ContributeBoardParts.KiCadFileNames(TabDrafts.KiCadNamesListedIn(publishedBoardFile))
                : [];
        }

        // A board file's KiCad names, kept while the file's length and write time stay as read - the
        // list is asked for after every refresh of the drafts, and a board file is parsed for it.
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, (string Stamp, List<string> Names)> KiCadNamesRead =
            new(StringComparer.OrdinalIgnoreCase);

        private static List<string> KiCadNamesListedIn(string boardFile)
        {
            var info = new FileInfo(boardFile);
            string stamp = $"{info.Length}|{info.LastWriteTimeUtc.Ticks}";

            if (TabDrafts.KiCadNamesRead.TryGetValue(boardFile, out var known) && known.Stamp == stamp &&
                StampSettling.IsSettled(info.LastWriteTimeUtc))
            {
                return known.Names;
            }

            List<string> names = (DraftFiles.ListedIn(boardFile)?.Keys ?? [])
                .Where(SubmissionKiCadFiles.IsInKiCadFolder)
                .ToList();

            TabDrafts.KiCadNamesRead[boardFile] = (stamp, names);
            return names;
        }

        internal string KiCadCountForTests => this.KiCadCountText.Text ?? string.Empty;

        // How many times the KiCad files were counted - for a test.
        internal int KiCadCountsForTests { get; private set; }

        // The components the list shows - ONE list for the tab's lifetime, changed in place.
        private readonly ObservableCollection<ComponentListItem> thisComponentItems = new();

        // True while the list's selection is set from code, to follow the pane - no pick then.
        private bool thisFollowingComponentSelection;

        // ###########################################################################################
        // The components of the board shown in the region shown - the list on CRT's left with its
        // "drafted" mark, by the one rule that builds it (ComponentListBuilder) - narrowed by the
        // search box's words. With the component pane or the table open, the draft as they left it
        // (BoardForList), so a name typed in the pane is in the list at once.
        //
        // *** CHANGED IN PLACE (InPlaceList.Sync), never a new list: *** this runs after every save
        // of the drafts list, and a list replaced under the pointer scrolled back to its top and let
        // go of its selection.
        // ###########################################################################################
        private void ShowComponents()
        {
            BoardData? board = this.BoardForList();
            string search = this.ComponentSearchBox.Text ?? string.Empty;
            IReadOnlySet<string>? drafted = this.DraftedLabelsForList();

            // Built once, then narrowed by the search's own rule - the count is of them all.
            List<ComponentListItem> all = board is null
                ? []
                : ComponentListBuilder.BuildComponentItems(board, this.thisCurrentRegion, categoryFilter: null, string.Empty, drafted);

            List<ComponentListItem> items = string.IsNullOrWhiteSpace(search)
                ? all
                : all.Where(item => ComponentListBuilder.MatchesSearch(item.DisplayText, search)).ToList();

            this.thisFollowingComponentSelection = true;

            try
            {
                InPlaceList.Sync(
                    this.thisComponentItems,
                    items,
                    (shown, wanted) => string.Equals(shown.BoardLabel, wanted.BoardLabel, StringComparison.Ordinal) &&
                                       string.Equals(shown.DisplayText, wanted.DisplayText, StringComparison.Ordinal) &&
                                       string.Equals(shown.SelectionKey, wanted.SelectionKey, StringComparison.Ordinal) &&
                                       shown.IsDrafted == wanted.IsDrafted,
                    wanted => wanted,
                    keep: this.ComponentsListBox.SelectedItem as ComponentListItem);
            }
            finally
            {
                this.thisFollowingComponentSelection = false;
            }

            this.ComponentsCountText.Text = ContributeBoardParts.CountText(all.Count);

            this.NoComponentsText.Text = ContributeBoardParts.NoComponentsLine(
                hasBoard: board is not null,
                searching: !string.IsNullOrWhiteSpace(search));
            this.NoComponentsText.IsVisible = items.Count == 0;
        }

        internal IReadOnlyList<ComponentListItem> ComponentsShownForTests => this.thisComponentItems;

        private void OnComponentSearchChanged(object? sender, TextChangedEventArgs e)
        {
            this.ShowComponents();
            this.ShowNavSelection();
        }

        // ---- which entry of the left column is open ------------------------------------------------

        // The entry of the left column shown selected - the part open in the pane; none while a
        // component is, which the list below them shows.
        internal string? SelectedNavEntryForTests =>
            new[] { this.NameNavButton, this.StatusNavButton, this.RegionsNavButton, this.PicturesNavButton, this.CategoriesNavButton, this.KiCadNavButton, this.SignalsNavButton, this.FilesNavButton, this.LinksNavButton, this.CreditsNavButton }
                .FirstOrDefault(button => button.Classes.Contains("selected"))?.Name;

        // ###########################################################################################
        // THE LEFT COLUMN FOLLOWS WHAT IS OPEN (owner request, 2026-10-29: "Selecting a component will
        // highlight it as normal, so it is possible to see in list which component is selected"). The
        // part open in the pane is selected; a component open selects itself in the list below them.
        // Worked out from the pane every time, never kept beside it, so a part that refused to close (a
        // schematic image's name it cannot take) or a table that stayed open leaves the column right.
        // ###########################################################################################
        private void ShowNavSelection()
        {
            PaneBoardPart? part = this.IsComponentOpen ? this.ComponentPane.BoardPart : null;
            bool componentOpen = this.IsComponentOpen && part is null;
            bool kiCadShown = this.IsKiCadOpen && !this.IsComponentOpen && !this.IsTableOpen;

            // Anything else open on the right lets "Name" go.
            if (this.IsTableOpen || this.IsComponentOpen || this.IsKiCadOpen)
            {
                this.thisBoardNamePicked = false;
            }

            TabDrafts.ShowSelected(this.NameNavButton, this.thisBoardNamePicked);

            // "Overview" while it is what the right side shows - nothing else open (2026-10-30).
            TabDrafts.ShowSelected(this.StatusNavButton, !this.IsTableOpen && !this.IsComponentOpen && !this.IsKiCadOpen && !this.thisBoardNamePicked);
            TabDrafts.ShowSelected(this.RegionsNavButton, part == PaneBoardPart.Regions);
            TabDrafts.ShowSelected(this.PicturesNavButton, part == PaneBoardPart.Pictures);
            TabDrafts.ShowSelected(this.CategoriesNavButton, part == PaneBoardPart.Categories);
            TabDrafts.ShowSelected(this.KiCadNavButton, kiCadShown);
            TabDrafts.ShowSelected(this.SignalsNavButton, part == PaneBoardPart.ImportantSignals);
            TabDrafts.ShowSelected(this.FilesNavButton, part == PaneBoardPart.Files);
            TabDrafts.ShowSelected(this.LinksNavButton, part == PaneBoardPart.Links);
            TabDrafts.ShowSelected(this.CreditsNavButton, part == PaneBoardPart.Credits);

            string label = componentOpen ? this.ComponentPane.Label : string.Empty;
            ComponentListItem? selected = label.Length == 0
                ? null
                : this.thisComponentItems.FirstOrDefault(item => string.Equals(item.BoardLabel, label, StringComparison.OrdinalIgnoreCase));

            // A selection the keys moved, waiting to open, is left where the keys put it - the pane
            // follows it once they pause (OpenKeyPickedComponent).
            if (ReferenceEquals(this.ComponentsListBox.SelectedItem, selected) || this.thisComponentKeyTimer?.IsEnabled == true)
            {
                return;
            }

            this.thisFollowingComponentSelection = true;

            try
            {
                this.ComponentsListBox.SelectedItem = selected;
            }
            finally
            {
                this.thisFollowingComponentSelection = false;
            }

            if (selected is not null)
            {
                this.ComponentsListBox.ScrollIntoView(selected);
            }
        }

        private static void ShowSelected(Button button, bool selected) => button.Classes.Set("selected", selected);

        // ###########################################################################################
        // The "Open components" of "Overview": the first component on the right, when none is - the
        // list below the board's entries is always there, so the keyboard goes to its search too.
        // ###########################################################################################
        internal void PickComponents()
        {
            if (!(this.IsComponentOpen && this.ComponentPane.BoardPart is null) &&
                this.thisComponentItems.FirstOrDefault() is { } first)
            {
                this.OpenPickedComponent(first);
            }

            this.ComponentSearchBox.Focus();
        }

        // ###########################################################################################
        // A component picked in the list - opened in the pane. A selection set from code to follow the
        // pane, or the one already open, opens nothing; the column follows what the pane did either way.
        //
        // *** A KEY MOVING THE SELECTION OPENS THE COMPONENT ONLY ONCE THE KEYS PAUSE *** (code review,
        // 2026-10-30): Down five times opened - and saved - each component on the way, every one under
        // the BusyOverlay, which took the next key press. A key's selection waits ComponentKeyPause;
        // Enter opens it at once, and a click always does.
        // ###########################################################################################
        internal static readonly TimeSpan ComponentKeyPause = TimeSpan.FromMilliseconds(400);

        private DispatcherTimer? thisComponentKeyTimer;

        // True while a key the list moves its selection with is being handled.
        private bool thisComponentKeyMoving;

        internal bool IsComponentKeyPickWaitingForTests => this.thisComponentKeyTimer?.IsEnabled == true;

        private void WireComponentsList()
        {
            this.thisComponentKeyTimer = new DispatcherTimer { Interval = TabDrafts.ComponentKeyPause };
            this.thisComponentKeyTimer.Tick += (_, _) => this.OpenKeyPickedComponent();

            this.ComponentsListBox.AddHandler(InputElement.KeyDownEvent, this.OnComponentsListKeyDown, RoutingStrategies.Tunnel);
            this.ComponentsListBox.AddHandler(InputElement.KeyDownEvent, (_, _) => this.thisComponentKeyMoving = false, RoutingStrategies.Bubble, handledEventsToo: true);
        }

        private void OnComponentsListKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && this.thisComponentKeyTimer?.IsEnabled == true)
            {
                this.OpenKeyPickedComponent();
                e.Handled = true;
                return;
            }

            this.thisComponentKeyMoving = e.Key is Key.Up or Key.Down or Key.PageUp or Key.PageDown or Key.Home or Key.End;
        }

        // The keys have paused (or Enter): the component they stopped on opened - while the list still
        // has the keyboard. Gone elsewhere meanwhile, the list shows what is open again.
        internal void OpenKeyPickedComponent()
        {
            this.thisComponentKeyTimer?.Stop();

            if (this.ComponentsListBox.IsKeyboardFocusWithin && this.ComponentsListBox.SelectedItem is ComponentListItem item)
            {
                this.OpenPickedComponent(item);
                return;
            }

            this.ShowNavSelection();
        }

        private void OnComponentsListSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisFollowingComponentSelection ||
                this.ComponentsListBox.SelectedItem is not ComponentListItem item ||
                item.BoardLabel.Length == 0)
            {
                return;
            }

            if (this.thisComponentKeyMoving && this.thisComponentKeyTimer is { } timer)
            {
                timer.Stop();
                timer.Start();
                return;
            }

            this.thisComponentKeyTimer?.Stop();
            this.OpenPickedComponent(item);
        }

        private void OpenPickedComponent(ComponentListItem item)
        {
            if (item.BoardLabel.Length == 0)
            {
                return;
            }

            if (this.IsComponentOpen && this.ComponentPane.BoardPart is null &&
                string.Equals(this.ComponentPane.Label, item.BoardLabel, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            this.RunGuarded(async () =>
            {
                await this.OpenComponentAsync(item.BoardLabel);
                this.ShowNavSelection();
            }, "Open component");
        }

        // A component picked in the list, as a click on it picks it - for a test.
        internal void PickComponentForTests(string label) =>
            this.ComponentsListBox.SelectedItem = this.thisComponentItems.First(item => string.Equals(item.BoardLabel, label, StringComparison.OrdinalIgnoreCase));

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
                // Let go of again should what is opened close with nothing in it (LetGoOfDraftMadeOnlyToLook),
                // and shown as a draft only from its first change - said before it is made, as making it
                // reads the list again.
                this.ForgetDraftMadeToLook();
                this.thisDraftMadeToLook = entry.ExcelDataFile;

                bool made = this.EnsureDraftOverrideForTests is { } ensure
                    ? await ensure(entry)
                    : this.thisMainWindow is not null && await this.thisMainWindow.EnsureDraftAsync(entry);

                if (!made)
                {
                    this.ForgetDraftMadeToLook();
                    return;
                }

                await this.WatchDraftForFirstChangeAsync(entry.ExcelDataFile);
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

        // ###########################################################################################
        // "BOARD STATUS" (owner request, 2026-10-30: the start page "as a new menu item in Board contents
        // ... and then remove those Close buttons on all those pages"): what is open on the right closes -
        // what it holds saved as it goes; a schematic image's name it cannot take keeps the Pictures part
        // open, the reason under its heading, as the pane's Close did - the components' list closes, and
        // the page is worked out again, so it is up to date whenever it is asked for. "Overview" in the
        // column and "Board: Overview" over the page since 2026-11-01 (owner request).
        // ###########################################################################################
        private void OnBoardStatusClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(this.ShowBoardStatusAsync, "Board status");

        internal async Task ShowBoardStatusAsync()
        {
            // Asked BEFORE anything closes (code review, 2026-10-30): a close shows the page, which
            // works it out as it comes back, and asked after it, the page was worked out a second time.
            bool shownAlready = this.StartPanel.IsVisible;

            if (this.IsComponentOpen)
            {
                if (this.ComponentPane.KeepsRefusedPictureName())
                {
                    return;
                }

                this.CloseComponent();
            }

            await this.CloseKiCadAsync();

            if (this.IsTableOpen && !await this.CloseTableAsync())
            {
                return;
            }

            this.thisBoardNamePicked = false;
            this.ShowWorkspace();

            // Coming back works it out by itself (ShowStartPanel); shown already, it is worked out again.
            if (shownAlready)
            {
                this.ShowReadiness();
            }
        }

        // The board's schematic images in the pane - added, named, matched to their KiCad views,
        // ordered, deleted and colored there (slice 4b; step 3 moved the "Schematic images" window in).
        private void OnPicturesClick(object? sender, RoutedEventArgs e) => this.ManagePictures();

        private void ManagePictures() =>
            this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.Pictures), "Schematic images");

        // "KiCad data" on the right, in place of a component or a part (TabDrafts.KiCad.cs) - a window
        // until 2026-10-30.
        private void OnKiCadClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(this.OpenKiCadAsync, "KiCad data");

        // The board's own lists, in the pane on the right (slice 4a - TabDrafts.Component.cs). Files
        // and Links are an entry each since 2026-10-29.
        private void OnImportantSignalsClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.ImportantSignals), "Important signals");

        private void OnRegionsClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.Regions), "Regions");

        private void OnCategoriesClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.Categories), "Categories");

        private void OnFilesClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.Files), "Files");

        private void OnLinksClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.Links), "Links");

        private void OnCreditsClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.Credits), "Credits");

        private void OnTableViewClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(this.ToggleTableViewAsync, "Table view");

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
