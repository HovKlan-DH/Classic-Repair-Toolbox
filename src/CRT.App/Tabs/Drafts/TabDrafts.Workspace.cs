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
    // Main); its header - the draft's own row, or what the board is while it has no draft - with "UI" and
    // "Table" at its end; the left column's parts with their counts (ContributeBoardParts - its entries
    // are TabDrafts.Entries.cs) and "Add component"; and "My drafts", which shows the board picked,
    // everywhere in CRT.
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

            // "Board name" is a view of the board shown - another one's page is no use - and the component
            // picked was one of its.
            if (anotherBoard)
            {
                this.thisBoardNamePicked = false;
                this.thisNoComponentPage = null;
                this.thisPickedComponent = string.Empty;
            }

            this.thisCurrentEntry = entry;
            this.thisCurrentBoard = entry is null ? null : board;
            this.thisCurrentRegion = region ?? string.Empty;
            this.thisDraftedLabels = draftedLabels;

            // A published board with no draft says its own errors and warnings - checked once this tab is shown
            // (TabDrafts.Problems.cs).
            this.ReadPublishedProblemsOnceShown();

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
        // in it (code review, 2026-10-26). *** Opening a component - "Edit component" is on everybody's
        // popup - a board part or the table makes the board's draft first
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
        private HardwareBoardEntry? thisDraftMadeToLook;

        private bool thisDraftChangedSinceLook;

        private DraftSession? thisDraftWatchedForChange;

        internal bool IsDraftOnlyLookedAtForTests(string identity) => this.IsUntouchedLook(identity);

        private bool IsUntouchedLook(string identity) =>
            !this.thisDraftChangedSinceLook && string.Equals(this.thisDraftMadeToLook?.ExcelDataFile, identity, StringComparison.OrdinalIgnoreCase);

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
        // nothing again. Whichever board it is: another may be shown by now (code review, 2026-11-02).
        // ###########################################################################################
        internal void LetGoOfUntouchedDraftOnQuit()
        {
            foreach (HardwareBoardEntry entry in this.LooksHoldingNothing())
            {
                Logger.Info($"The draft of [{entry.ExcelDataFile}] was only looked at - let go of as CRT closes");
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
            foreach (HardwareBoardEntry entry in this.LooksHoldingNothing())
            {
                Logger.Info($"The draft of [{entry.ExcelDataFile}] was only looked at - let go of before the restart");
                this.DiscardConfirmed(entry);
            }
        }

        // ###########################################################################################
        // *** A STEP TO REDO IS WORK THAT CAN STILL BE ASKED FOR (2026-11-02). *** A draft made only to
        // look, changed, undone and closed holds nothing - and was deleted, its Redo with it, so a change
        // undone a moment ago was gone for good. It stays while the Redo can bring the change back
        // (LetGoOfDraftMadeOnlyToLook). That ends with the run, so as CRT closes or restarts it goes -
        // left, the next run would show a draft with nothing in it. Brought back, it holds something, and
        // is a draft like any other.
        // ###########################################################################################
        private readonly List<HardwareBoardEntry> thisLooksKeptForRedo = [];

        // The drafts made only to look that hold nothing, marked or kept for their Redo, each once - told
        // now, as letting go of one forgets the marker and this list stays as it is for the caller.
        private List<HardwareBoardEntry> LooksHoldingNothing()
        {
            List<HardwareBoardEntry> looks = [.. this.thisLooksKeptForRedo];

            if (this.thisDraftMadeToLook is { } marked && !looks.Exists(look => TabDrafts.SameBoard(look, marked)))
            {
                looks.Add(marked);
            }

            this.thisLooksKeptForRedo.Clear();
            this.ForgetDraftMadeToLook();

            return looks.Where(TabDrafts.HoldsNothing).ToList();
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
            if (this.thisDraftMadeToLook is not { } look || !TabDrafts.SameBoard(look, entry) ||
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

            if (DraftManager.CanRedo(entry.ExcelDataFile))
            {
                Logger.Info($"The draft of [{entry.ExcelDataFile}] holds nothing but a step to redo - kept while CRT runs");

                if (!this.thisLooksKeptForRedo.Exists(look => TabDrafts.SameBoard(look, entry)))
                {
                    this.thisLooksKeptForRedo.Add(entry);
                }

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
            !this.IsTableOpen && !this.IsComponentOpen && !this.IsKiCadOpen && !this.thisBoardNamePicked && this.thisNoComponentPage is null;

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
            this.ViewSwitch.IsEnabled = hasBoard;
            this.ComponentPane.IsVisible = !this.IsTableOpen && this.IsComponentOpen;
            this.KiCadPanelHost.IsVisible = !this.IsTableOpen && !this.IsComponentOpen && this.IsKiCadOpen;
            this.BoardNamePanel.IsVisible = this.thisBoardNamePicked && !this.IsTableOpen && !this.IsComponentOpen && !this.IsKiCadOpen;
            this.ShowNoComponentPage();
            this.ShowStartPanel(this.NothingElseOpenOnTheRight);

            string revision = this.thisCurrentBoard?.RevisionDate?.Trim() ?? string.Empty;
            this.PublishedOnText.Text = revision.Length == 0 ? string.Empty : $"The board's data was last published on {revision}.";
            this.PublishedOnText.IsVisible = revision.Length > 0 && this.thisCurrentEntry?.IsDraftOnly != true;
        }

        // ###########################################################################################
        // The header: the draft's row of the board the table is open on, else of the board shown - and
        // the tab's top line, naming that board.
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

            this.NoBoardText.Text = ContributeBoardParts.NoBoardLine;
            this.NoBoardText.IsVisible = headerEntry is null;

            Handlers.Theming.BoldRuns.ShowMarked(this.TopLineText, ContributeBoardParts.TopLine(headerEntry?.HardwareName, headerEntry?.BoardName));
        }

        internal string TopLineForTests => Handlers.Theming.BoldRuns.TextOf(this.TopLineText);

        // A component's page with no component picked: where to pick one.
        private void ShowNoComponentPage()
        {
            bool shown = this.thisNoComponentPage is not null && !this.IsTableOpen && !this.IsComponentOpen && !this.IsKiCadOpen && !this.thisBoardNamePicked;

            this.NoComponentPanel.IsVisible = shown;

            if (shown && this.thisNoComponentPage is { } page)
            {
                this.NoComponentTitle.Text = ContributeBoardParts.NoComponentTitle(page);
                this.NoComponentText.Text = this.IsEveryTakenOver
                    ? ContributeBoardParts.EveryComponentSelected(page)
                    : ContributeBoardParts.NoComponentPicked(page);
            }
        }

        internal bool IsLocalDraftNoticeShownForTests => this.LocalDraftNotice.IsVisible;

        internal string LocalDraftNoticeForTests => this.LocalDraftNoticeText.Text ?? string.Empty;

        // The board's parts on the left, each with how many it holds.
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
        // stays meanwhile for that board, and "0" is not said before it is known.
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

        // ###########################################################################################
        // THE LISTS IT PICKS COMPONENTS IN STAND ON THE FAR LEFT (owner request, 2026-11-03; Main has them, in
        // place of the other tabs' while this one is shown - Main.ContributeLists.cs). What they show is held
        // here, as only this tab knows the draft as it is being edited: ONE collection each for the tab's
        // lifetime, changed in place.
        // ###########################################################################################
        private readonly ObservableCollection<ComponentListItem> thisComponentItems = new();

        private readonly ObservableCollection<string> thisCategoryItems = new();

        // The list on the far left - null for a tab a test builds until it hands it one.
        private ListBox? thisComponentList;

        // What narrows the list: the categories unticked, and the words typed over it (Main keeps them per
        // board - Main.ContributeLists.cs).
        private HashSet<string> thisCategoriesUnticked = new(StringComparer.OrdinalIgnoreCase);

        private string thisComponentSearch = string.Empty;

        // True while the list's selection is set from code, to follow the pane - no pick then.
        private bool thisFollowingComponentSelection;

        internal ObservableCollection<ComponentListItem> ComponentItems => this.thisComponentItems;

        internal ObservableCollection<string> CategoryItems => this.thisCategoryItems;

        // The lists were read again - whoever shows them follows what else hangs on the board.
        internal event EventHandler? ComponentListsShown;

        // ###########################################################################################
        // The list components are picked in: a pick in it opens the component, and a key moving its selection
        // waits for the keys to pause (OnComponentsListSelectionChanged).
        // ###########################################################################################
        internal void UseComponentList(ListBox? list)
        {
            this.thisComponentKeyTimer?.Stop();

            if (this.thisComponentList is { } before)
            {
                before.SelectionChanged -= this.OnComponentsListSelectionChanged;
                before.Tapped -= this.OnComponentsListTapped;
                before.RemoveHandler(InputElement.KeyDownEvent, this.OnComponentsListKeyDown);
                before.RemoveHandler(InputElement.KeyDownEvent, this.OnComponentsListKeyHandled);
            }

            this.thisComponentList = list;

            if (list is not null)
            {
                list.SelectionChanged += this.OnComponentsListSelectionChanged;
                list.Tapped += this.OnComponentsListTapped;
                list.AddHandler(InputElement.KeyDownEvent, this.OnComponentsListKeyDown, RoutingStrategies.Tunnel);
                list.AddHandler(InputElement.KeyDownEvent, this.OnComponentsListKeyHandled, RoutingStrategies.Bubble, handledEventsToo: true);

                this.ShowComponents();
                this.ShowNavSelection();
            }
        }

        // The categories unticked and the words typed over the list, as they now read. The third thing that
        // narrows it, "With issues", is the tab's own (TabDrafts.Problems.cs).
        internal void UseComponentFilter(IReadOnlyCollection<string> unticked, string search)
        {
            this.KeepComponentFilter(unticked, search);

            this.ShowComponents();
            this.ShowNavSelection();
        }

        // The same for a board about to be shown, which reads the list itself (ShowBoard).
        internal void KeepComponentFilter(IReadOnlyCollection<string> unticked, string search)
        {
            this.thisCategoriesUnticked = new HashSet<string>(unticked, StringComparer.OrdinalIgnoreCase);
            this.thisComponentSearch = search ?? string.Empty;
        }

        // ###########################################################################################
        // The categories and the components of the board shown in the region shown - the list on CRT's
        // left with its "drafted" mark, by the one rule that builds it (ComponentListBuilder) - narrowed by
        // the categories unticked and the words typed. With the component pane or the table open, the draft
        // as they left it (BoardForList), so a name typed in the pane is in the list at once.
        //
        // *** CHANGED IN PLACE (InPlaceList.Sync), never a new list: *** this runs after every save
        // of the drafts list, and a list replaced under the pointer scrolled back to its top and let
        // go of its selection.
        //
        // *** NEVER INSIDE ITSELF: *** a category picked that goes from the list is a selection changed while
        // the list is still being changed, so whoever shows the lists asks IsReadingComponentLists - that is
        // nobody's choice - and the list's own selection is thisFollowingComponentSelection, the box's
        // thisFollowingIssuesBox.
        // ###########################################################################################
        private bool thisShowingComponents;

        internal bool IsReadingComponentLists => this.thisShowingComponents;

        private void ShowComponents()
        {
            this.thisShowingComponents = true;

            try
            {
                this.ReadComponentLists();
            }
            finally
            {
                this.thisShowingComponents = false;
            }

            this.ComponentListsShown?.Invoke(this, EventArgs.Empty);
        }

        private void ReadComponentLists()
        {
            BoardData? board = this.BoardForList();

            // The list can be narrowed by what the checks found (TabDrafts.Problems.cs).
            this.PlaceProblems();
            this.LetGoOfIssuesTickWithNoIssues();

            List<ComponentListItem> items = board is null
                ? []
                : this.NarrowedToIssues(ComponentListBuilder.BuildComponentItems(board, this.thisCurrentRegion, this.CategoriesTicked(board), this.thisComponentSearch, this.DraftedLabelsForList()));

            // ###########################################################################################
            // *** THE LIST COMES BACK TO THE SELECTION ONCE, NOT ROW BY ROW (owner report, 2026-11-04: "some
            // kind of slow scroll behaviour, where it populates the list and scrolls to the same component
            // afterwards"). *** Every row put in or taken out above the selection moves its index, and the
            // list by itself asks to be scrolled to each index in turn, posted, a layout pass apiece: 392
            // scrolls for 399 rows, seen as the list scrolling down slowly. Off while it changes, then once -
            // only when the selection moved at all, or a save below it would pull the list back to it.
            // ###########################################################################################
            ListBox? list = this.thisComponentList;
            int selectedWas = list?.SelectedIndex ?? -1;
            bool scrollsItself = list?.AutoScrollToSelectedItem == true;

            if (list is not null)
            {
                list.AutoScrollToSelectedItem = false;
            }

            this.thisFollowingComponentSelection = true;

            try
            {
                InPlaceList.Sync(
                    this.thisCategoryItems,
                    board is null ? [] : ComponentListBuilder.BuildDistinctCategories(board),
                    (shown, wanted) => string.Equals(shown, wanted, StringComparison.Ordinal),
                    wanted => wanted);

                InPlaceList.Sync(
                    this.thisComponentItems,
                    items,
                    (shown, wanted) => string.Equals(shown.BoardLabel, wanted.BoardLabel, StringComparison.Ordinal) &&
                                       string.Equals(shown.DisplayText, wanted.DisplayText, StringComparison.Ordinal) &&
                                       string.Equals(shown.SelectionKey, wanted.SelectionKey, StringComparison.Ordinal) &&
                                       shown.IsDrafted == wanted.IsDrafted,
                    wanted => wanted,
                    keep: this.thisComponentList?.SelectedItem as ComponentListItem);
            }
            finally
            {
                this.thisFollowingComponentSelection = false;

                if (list is not null)
                {
                    list.AutoScrollToSelectedItem = scrollsItself;
                }
            }

            if (list is { SelectedIndex: >= 0 } && list.SelectedIndex != selectedWas)
            {
                list.ScrollIntoView(list.SelectedIndex);
            }

            // A row just made says what the checks found on it too (TabDrafts.Problems.cs).
            this.ShowProblemsInComponentList();
            this.ShowIssuesBox();
        }

        // The categories that let their components through: null for all of them. A component with no
        // category has no tick that could hide it, and is what has to be found to give it one.
        private HashSet<string>? CategoriesTicked(BoardData board)
        {
            if (this.thisCategoriesUnticked.Count == 0)
            {
                return null;
            }

            var ticked = new HashSet<string>(ComponentListBuilder.BuildDistinctCategories(board), StringComparer.OrdinalIgnoreCase) { string.Empty };

            ticked.ExceptWith(this.thisCategoriesUnticked);

            return ticked;
        }

        internal IReadOnlyList<ComponentListItem> ComponentsShownForTests => this.thisComponentItems;

        // ---- which entry of the left column is open ------------------------------------------------

        // ###########################################################################################
        // THE LEFT COLUMN FOLLOWS WHAT IS OPEN (owner request, 2026-10-29: "Selecting a component will
        // highlight it as normal, so it is possible to see in list which component is selected"). The entry
        // open is selected (EntryOpen), so a part that refused to close (a schematic image's name it
        // cannot take) or a table that stayed open leaves the column right; the component picked stays
        // selected in the list whatever is open beside it (owner request, 2026-11-03).
        // ###########################################################################################
        private void ShowNavSelection()
        {
            // Anything else open on the right lets "Name" go.
            if (this.IsTableOpen || this.IsComponentOpen || this.IsKiCadOpen)
            {
                this.thisBoardNamePicked = false;
            }

            this.FollowPickedComponent();

            ContributeEntry? current = this.EntryOpen();

            foreach ((Button button, ContributeEntry entry) in this.NavEntries())
            {
                TabDrafts.ShowSelected(button, entry == current);
            }

            this.ShowComponentPagesInNav();
            this.ShowTableEntriesEnabled();
            this.ShowViewSwitch();
            this.ShowProblems();

            if (this.thisComponentList is not { } list)
            {
                return;
            }

            // The table's rows follow the components picked, whatever is open (BoardTableEditor.Components.cs).
            this.TableEditor.ShowOnlyComponents(this.IsTableOpen ? this.TableLabels : []);

            if (this.thisPickedSeveral.Count > 1)
            {
                this.ShowSeveralSelected(list);
                return;
            }

            string label = this.thisPickedComponent;
            ComponentListItem? selected = label.Length == 0
                ? null
                : this.thisComponentItems.FirstOrDefault(item => string.Equals(item.BoardLabel, label, StringComparison.OrdinalIgnoreCase));

            // A selection the keys moved, waiting to open, is left where the keys put it - the pane
            // follows it once they pause (OpenKeyPickedComponent).
            if ((ReferenceEquals(list.SelectedItem, selected) && list.SelectedItems?.Count <= 1) || this.thisComponentKeyTimer?.IsEnabled == true)
            {
                return;
            }

            this.thisFollowingComponentSelection = true;

            try
            {
                list.SelectedItem = selected;
            }
            finally
            {
                this.thisFollowingComponentSelection = false;
            }

            if (selected is not null)
            {
                list.ScrollIntoView(selected);
            }
        }

        private static void ShowSelected(Button button, bool selected) => button.Classes.Set("selected", selected);

        // ###########################################################################################
        // The "Open components" of "Overview": the component picked, else the first - and the keyboard to
        // the filter over the list on the far left.
        // ###########################################################################################
        internal void PickComponents()
        {
            ComponentListItem? item =
                this.thisComponentItems.FirstOrDefault(listed => string.Equals(listed.BoardLabel, this.thisPickedComponent, StringComparison.OrdinalIgnoreCase)) ??
                this.thisComponentItems.FirstOrDefault();

            if (item is not null)
            {
                this.OpenPickedComponent(item);
            }

            this.thisMainWindow?.FocusComponentFilter();
        }

        // ###########################################################################################
        // A component picked in the list - opened in the pane on the page it stands on, or with "Table"
        // picked, gone to on its sheet. A selection set from code to follow the pane, or the one already
        // open, opens nothing; the column follows what the pane did either way.
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
        }

        private void OnComponentsListKeyHandled(object? sender, KeyEventArgs e) => this.thisComponentKeyMoving = false;

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

            if (this.thisComponentList is { IsKeyboardFocusWithin: true } && this.SelectedInList() is { Count: > 1 } several)
            {
                this.OpenPickedComponents(several);
                return;
            }

            if (this.thisComponentList is { IsKeyboardFocusWithin: true, SelectedItem: ComponentListItem item })
            {
                this.OpenPickedComponent(item);
                return;
            }

            this.ShowNavSelection();
        }

        private void OnComponentsListSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisFollowingComponentSelection)
            {
                return;
            }


            // Several selected - Ctrl or Shift with a click, Shift with the keys - edited as one (owner request, 2026-11-08).
            if (this.SelectedInList() is { Count: > 1 } several)
            {
                if (this.thisComponentKeyMoving && this.thisComponentKeyTimer is { } waiting)
                {
                    waiting.Stop();
                    waiting.Start();
                    return;
                }

                this.thisComponentKeyTimer?.Stop();
                this.OpenPickedComponents(several);
                return;
            }

            if (this.thisComponentList?.SelectedItem is not ComponentListItem item ||
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

            // The list stands outside the tab, so "CRT has to be updated" over the tab does not cover it.
            if (this.IsUpdateRequiredShown)
            {
                this.ShowNavSelection();
                return;
            }

            string label = item.BoardLabel;

            this.thisPickedComponent = label;
            this.thisPickedSeveral = [];

            // One click is a selection changed and then a tap, each coming here - in the table too (code
            // review, 2026-11-04: the table was read and built twice). Only while that open, or what the pane shows, is the
            // latest pick (code review, 2026-11-09): a Shift+click opening several in between, a click on the one still
            // opening was dropped, and the several - ending last - took the pick back.
            bool sameOpenUnderWay = string.Equals(this.thisOpeningComponent, label, StringComparison.OrdinalIgnoreCase) &&
                                    this.thisOpeningGeneration == this.thisPickGeneration;
            bool shownAlready = !this.thisTableMode && this.IsPaneOnComponent(label) && this.thisPicksOpening == 0;

            if (sameOpenUnderWay || shownAlready)
            {
                return;
            }

            this.thisOpeningComponent = label;

            int generation = ++this.thisPickGeneration;
            this.thisOpeningGeneration = generation;
            this.thisPicksOpening++;

            this.RunGuarded(async () =>
            {
                try
                {
                    if (this.HoldOpeningForTests is { } hold)
                    {
                        await hold;
                    }

                    await this.OpenPickedInTheViewPickedAsync(label, () => generation == this.thisPickGeneration);
                }
                finally
                {
                    // A newer open of its own leaves its mark.
                    if (this.thisOpeningGeneration == generation)
                    {
                        this.thisOpeningComponent = null;
                    }

                    this.thisPicksOpening--;
                }

                this.ShowNavSelection();
            }, "Open component");
        }

        // The pick the open of thisOpeningComponent was started for.
        private int thisOpeningGeneration;

        // ###########################################################################################
        // *** A PICK OVERTAKEN BY A NEWER ONE OPENS NOTHING (2026-11-08: a Shift+click failed now and then). *** Opening goes
        // through the BusyOverlay and reads the draft off the UI thread, so a click and a Shift+click right after it each open
        // - and the first, ending last, put the pane on its one component and the pick back to it. Each pick takes the next
        // number; an open shows only while its number is the latest, and while any is under way the tab does not follow what
        // the pane shows (FollowPickedComponent).
        // ###########################################################################################
        private int thisPickGeneration;

        private int thisPicksOpening;

        // For a test: what an open from the list waits for first, so a newer pick can be made while it is under way.
        internal Task? HoldOpeningForTests { get; set; }

        private Task OpenPickedInTheViewPickedAsync(string label, Func<bool> stillWanted)
        {
            if (!stillWanted())
            {
                return Task.CompletedTask;
            }

            if (!this.thisTableMode)
            {
                return this.OpenComponentAsync(label, stillWanted);
            }

            ContributeEntry entry = this.EntryOpen() is { } shown && ContributeEntries.PageOf(shown) is not null
                ? shown
                : ContributeEntries.OfPage(this.ComponentPane.Page);

            return this.ShowEntryInTableAsync(entry, takeFocus: false);
        }

        // Several picked in the list, as Ctrl+click picks them - for a test.
        internal void PickComponentsForTests(params string[] labels)
        {
            if (this.thisComponentList is not { SelectedItems: { } selected } list)
            {
                return;
            }

            this.thisFollowingComponentSelection = true;

            try
            {
                selected.Clear();

                foreach (ComponentListItem item in this.thisComponentItems.Where(item => labels.Contains(item.BoardLabel, StringComparer.OrdinalIgnoreCase)))
                {
                    selected.Add(item);
                }
            }
            finally
            {
                this.thisFollowingComponentSelection = false;
            }

            this.OnComponentsListSelectionChanged(list, new SelectionChangedEventArgs(Avalonia.Controls.Primitives.SelectingItemsControl.SelectionChangedEvent, Array.Empty<object>(), Array.Empty<object>()));
        }

        // A component picked in the list, as a click on it picks it - for a test, whose tab has no list unless
        // it hands it one.
        internal void PickComponentForTests(string label)
        {
            ComponentListItem item = this.thisComponentItems.First(item => string.Equals(item.BoardLabel, label, StringComparison.OrdinalIgnoreCase));

            // Already selected, a click changes no selection - and is what opens it again.
            if (this.thisComponentList is { } list && !ReferenceEquals(list.SelectedItem, item))
            {
                list.SelectedItem = item;
            }
            else
            {
                this.OpenPickedComponent(item);
            }
        }

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
        //
        // *** A DRAFT THE ACTION OPENED NOTHING ON GOES AGAIN (code review, 2026-11-02). *** "Use" with
        // nothing left to take, a KiCad open called off: nothing is left to close, so nothing let go of
        // it. Hidden until the next look forgot it, it then stood in "My drafts" with no changes - so that
        // look lets go of an earlier one too, should something still have left it behind.
        // ###########################################################################################
        internal async Task OnCurrentDraftAsync(Func<HardwareBoardEntry, Task> action)
        {
            if (this.thisCurrentEntry is not { } entry || string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                return;
            }

            if (!DraftBoardSource.HasDraft(DraftManager.DraftsRoot, entry.ExcelDataFile))
            {
                if (this.thisDraftMadeToLook is { } earlier)
                {
                    this.LetGoOfDraftMadeOnlyToLook(earlier);
                }

                // Let go of again should what is opened close with nothing in it (LetGoOfDraftMadeOnlyToLook),
                // and shown as a draft only from its first change - said before it is made, as making it
                // reads the list again.
                this.ForgetDraftMadeToLook();
                this.thisDraftMadeToLook = entry;

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

            // Asks nothing of a draft that was there before, or while what the action opened is open.
            this.LetGoOfDraftMadeOnlyToLook(entry);
        }

        // A sheet's table - "Show in the table" under the problems found on it - never a sheet in two, as
        // "Components" has no such button (BoardStatusPage.WithProblemsOf). The left column follows
        // (TabDrafts.Entries.cs).
        internal Task ShowTableOnSheetAsync(string sheetName) =>
            this.OnCurrentDraftAsync(async entry =>
            {
                await this.ShowTableAsync(entry);

                if (this.IsTableOpenFor(entry))
                {
                    this.TableEditor.ShowSheetPart(sheetName, BoardTableSheetPart.Whole);
                }
            });

        // ###########################################################################################
        // "BOARD STATUS" (owner request, 2026-10-30: the start page "as a new menu item in Board contents
        // ... and then remove those Close buttons on all those pages"): what is open on the right closes -
        // what it holds saved as it goes; a schematic image's name it cannot take keeps the Pictures part
        // open, the reason under its heading, as the pane's Close did - the components' list closes, and
        // the page is worked out again, so it is up to date whenever it is asked for. "Overview" in the
        // column since 2026-11-01, the first entry and the page's headline too since 2026-11-03 (owner
        // requests).
        // ###########################################################################################
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
            this.thisNoComponentPage = null;
            this.ShowWorkspace();

            // Coming back works it out by itself (ShowStartPanel); shown already, it is worked out again.
            if (shownAlready)
            {
                this.ShowReadiness();
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
        // Why "Table", a board part or "Add a new board" made no draft - or nothing, to clear
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
