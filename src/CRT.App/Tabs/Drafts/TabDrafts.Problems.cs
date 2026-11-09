using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;

namespace CRT
{
    // ###########################################################################################
    // THE DRAFT'S ERRORS AND WARNINGS, WHERE THEY ARE (owner request, 2026-11-04: "move the error and warning
    // badges to where the problem is ... and then those errors and warnings should be clearly visible in
    // there"; two counts in the draft's header until then).
    //
    // Owns: a red and an amber badge on each entry of the left column and on each component in the list on
    // the far left, the box narrowing that list to those with any, and on the right, over the page open, a
    // panel saying what they are. "Overview" has them all; with a component selected its pages have that
    // one's alone. Where a problem goes and which are shown is ContributeProblems'.
    //
    // *** READ WHEN THE DRAFTS ARE READ (RefreshDrafts, the one read the header's counts came from), so
    // it follows a save - a second after a change, as those did - and an edit made in Excel. *** What
    // the page itself marks at once stays its own: a box's "Required before submitting", a cell's corner.
    //
    // *** NOT OVER A TABLE: *** it marks every problem on its cell and has its own "Errors" and "Warnings"
    // pills, and every row of the panel would take a row of the table's. The badges stay.
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // What the checks found on each draft, as the list last read it.
        private Dictionary<string, BoardProblemsRead> thisProblemsRead = new(StringComparer.OrdinalIgnoreCase);

        // Where they go, for the board shown - worked out only when the read is another.
        private BoardProblemsRead? thisProblemsPlacedFrom;
        private List<EntryProblem> thisProblemsPlaced = [];
        private Dictionary<string, BoardProblemCounts> thisProblemsOfComponents = new(StringComparer.OrdinalIgnoreCase);

        private readonly Dictionary<ContributeEntry, (StackPanel Panel, Border Errors, TextBlock ErrorsText, Border Warnings, TextBlock WarningsText)> thisNavBadges = [];

        private readonly ObservableCollection<EntryProblemRow> thisEntryProblemRows = [];

        internal IReadOnlyList<EntryProblemRow> EntryProblemRowsForTests => this.thisEntryProblemRows;

        internal bool IsEntryProblemsPanelShownForTests => this.EntryProblemsPanel.IsVisible;

        // What an entry's badges say - "" for one not shown - for a test.
        internal (string Errors, string Warnings) NavBadgesForTests(ContributeEntry entry)
        {
            var badges = this.thisNavBadges[entry];

            return (badges.Errors.IsVisible ? badges.ErrorsText.Text ?? string.Empty : string.Empty,
                    badges.Warnings.IsVisible ? badges.WarningsText.Text ?? string.Empty : string.Empty);
        }

        private void WireProblems()
        {
            this.EntryProblemsList.ItemsSource = this.thisEntryProblemRows;

            foreach ((Button button, ContributeEntry entry) in this.NavEntries())
            {
                this.thisNavBadges[entry] = TabDrafts.AddBadgesTo(button);
            }
        }

        // ###########################################################################################
        // The badges into an entry: between its name and its count, so the counts stay under each other.
        // An entry with a name alone gets a row for them.
        // ###########################################################################################
        private static (StackPanel, Border, TextBlock, Border, TextBlock) AddBadgesTo(Button button)
        {
            (Border errors, TextBlock errorsText) = TabDrafts.Badge("error");
            (Border warnings, TextBlock warningsText) = TabDrafts.Badge("warning");

            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 3,
                Margin = new Thickness(6, 0),
                VerticalAlignment = VerticalAlignment.Center,
                IsVisible = false,
                Children = { errors, warnings },
            };

            if (button.Content is not Grid row)
            {
                var name = (Control)button.Content!;

                button.Content = null;
                row = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), Children = { name } };
                button.Content = row;
            }

            row.ColumnDefinitions.Insert(1, new ColumnDefinition(GridLength.Auto));

            foreach (Control child in row.Children.Where(child => Grid.GetColumn(child) >= 1))
            {
                Grid.SetColumn(child, Grid.GetColumn(child) + 1);
            }

            Grid.SetColumn(panel, 1);
            row.Children.Add(panel);

            return (panel, errors, errorsText, warnings, warningsText);
        }

        private static (Border, TextBlock) Badge(string kind)
        {
            var text = new TextBlock();
            var badge = new Border { Child = text, IsVisible = false };

            badge.Classes.Add("NavBadge");
            badge.Classes.Add(kind);

            return (badge, text);
        }

        // ###########################################################################################
        // *** A PUBLISHED BOARD WITH NO DRAFT SAYS ITS OWN ERRORS AND WARNINGS (owner decision, 2026-11-07: the launch check
        // of every board, written to the log alone, "needs to run for everyone and its findings should show just like any
        // other error or warning in the "Contribute" tab"). *** The same checks a draft gets, on the published board file,
        // read off the UI thread (DraftStatusReader.PublishedProblemsCached) - in the badges, the component list's and
        // "Overview"'s panels as a draft's are. A draft made then has its own, as ever.
        //
        // *** ONLY ONCE THIS TAB IS SHOWN (code review, 2026-11-08). *** Nothing but this tab shows them, and every board
        // shown on another tab - switching back and forth on "Schematics" - checked its whole published board, every row and
        // every file it cites, with only one board's read kept: each switch checked the board again. Shown, the tab checks
        // the board on screen (OnAttachedToVisualTree), and the board status's panels wait for that very check rather than
        // run a second one beside it (PublishedProblemsOf).
        // ###########################################################################################
        private (string Identity, BoardProblemsRead Read)? thisPublishedProblems;

        private int thisPublishedProblemsAsk;

        // The board on screen's published board not checked, the tab being off screen as it was shown.
        private bool thisPublishedProblemsStale;

        private void ReadPublishedProblemsOnceShown()
        {
            if (this.thisCurrentEntry is not { } entry)
            {
                this.thisPublishedProblemsStale = false;
                return;
            }

            if (!this.IsOnScreen())
            {
                this.thisPublishedProblemsStale = true;
                return;
            }

            this.thisPublishedProblemsStale = false;
            _ = this.ReadPublishedProblemsAsync(entry);
        }

        // ###########################################################################################
        // *** A DATA SYNC OVER: THE BOARD SHOWN IS CHECKED AGAIN (code review, 2026-11-08). *** The board files sync before
        // CRT's window opens and the files they name after it, so a published board checked meanwhile said "is not in the
        // downloaded data" of files still on their way - and kept saying it until the board or the tab was shown again.
        // Checked again, a read finds what has arrived since (DraftStatusReader keeps a read only while every file it
        // missed is still missing), and the board status's panels with it. From Main, after the launch's sync and a check
        // asked for; off screen, both wait for the tab to be shown.
        // ###########################################################################################
        internal void CheckPublishedBoardAgain()
        {
            // A check still running began before the files arrived: never handed back for these (PublishedProblemsOf).
            this.thisPublishedCheckGeneration++;
            this.ReadPublishedProblemsOnceShown();
            this.ShowReadiness();
        }

        // The tab shown: the check the board on screen waited for.
        private void ReadPublishedProblemsIfStale()
        {
            if (this.thisPublishedProblemsStale)
            {
                this.ReadPublishedProblemsOnceShown();
            }
        }

        private async Task ReadPublishedProblemsAsync(HardwareBoardEntry entry)
        {
            int ask = ++this.thisPublishedProblemsAsk;
            string identity = entry.ExcelDataFile;

            if (this.thisPublishedProblems is { } held && !string.Equals(held.Identity, identity, StringComparison.OrdinalIgnoreCase))
            {
                this.thisPublishedProblems = null;
            }

            if (entry.IsDraftOnly || string.IsNullOrWhiteSpace(identity))
            {
                return;
            }

            try
            {
                // A board with a draft has the draft's own - asked off the UI thread, as the check is.
                string draftsRoot = DraftManager.DraftsRoot;

                if (await Task.Run(() => DraftBoardSource.HasDraft(draftsRoot, identity)))
                {
                    return;
                }

                BoardProblemsRead read = await this.PublishedProblemsOf(entry);

                if (ask != this.thisPublishedProblemsAsk || this.thisCurrentEntry is not { } shown || !TabDrafts.SameBoard(shown, entry))
                {
                    return;
                }

                this.thisPublishedProblems = (identity, read);
                this.ShowProblems();
            }
            catch (Exception ex)
            {
                Logger.Warning($"The published board [{identity}] could not be checked for errors and warnings - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // The published board's check, off the UI thread - the one still running for the board when there is one, so the
        // badges and the board status's panels share it. A check done is not kept here: DraftStatusReader keeps its read
        // while the board file is as it was, so asking again after a data update checks the file as it is.
        // ###########################################################################################
        private (string Identity, int Generation, Task<BoardProblemsRead> Check)? thisPublishedCheck;

        // Moved by a data sync's end (CheckPublishedBoardAgain): a check started before it is never handed back after it
        // (code review, 2026-11-08) - it began before the files the sync brought down.
        private int thisPublishedCheckGeneration;

        private Task<BoardProblemsRead> PublishedProblemsOf(HardwareBoardEntry entry)
        {
            string identity = entry.ExcelDataFile;

            if (this.thisPublishedCheck is { } running &&
                !running.Check.IsCompleted &&
                running.Generation == this.thisPublishedCheckGeneration &&
                string.Equals(running.Identity, identity, StringComparison.OrdinalIgnoreCase))
            {
                return running.Check;
            }

            string published = this.PublishedBoardFileOf(entry);
            string dataRoot = DataManager.DataRoot;
            string draftFolder = DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, identity);

            this.PublishedChecksForTests++;
            Task<BoardProblemsRead> check = Task.Run(() => DraftStatusReader.PublishedProblemsCached(published, dataRoot, draftFolder, identity));
            this.thisPublishedCheck = (identity, this.thisPublishedCheckGeneration, check);

            return check;
        }

        // How many published boards' checks were started - for the tests.
        internal int PublishedChecksForTests { get; private set; }

        // The published board's read, for the board shown while it has no draft's.
        private BoardProblemsRead? PublishedProblemsOfBoardShown() =>
            this.thisCurrentEntry is { } entry &&
            this.thisPublishedProblems is { } held &&
            string.Equals(held.Identity, entry.ExcelDataFile, StringComparison.OrdinalIgnoreCase)
                ? held.Read
                : null;

        // ###########################################################################################
        // Where the read goes, for the board shown - only worked out for a read not placed yet.
        // ###########################################################################################
        private void PlaceProblems()
        {
            BoardProblemsRead? read = this.thisCurrentEntry is { } entry && this.thisProblemsRead.TryGetValue(entry.ExcelDataFile, out BoardProblemsRead? found)
                ? found
                : this.PublishedProblemsOfBoardShown();

            if (!ReferenceEquals(read, this.thisProblemsPlacedFrom))
            {
                this.thisProblemsPlacedFrom = read;
                this.thisProblemsPlaced = read is null ? [] : ContributeProblems.Placed(read);
                this.thisProblemsOfComponents = ContributeProblems.CountsByComponent(this.thisProblemsPlaced);
            }
        }

        // ###########################################################################################
        // The badges and the page's panel, as the checks last read and what is open now - from whatever
        // moves either (ShowNavSelection).
        // ###########################################################################################
        private void ShowProblems()
        {
            this.PlaceProblems();

            Dictionary<ContributeEntry, BoardProblemCounts> shown = ContributeProblems.CountsShown(this.thisProblemsPlaced, this.PickedLabels);

            foreach ((ContributeEntry entry, var badges) in this.thisNavBadges)
            {
                BoardProblemCounts counts = shown.GetValueOrDefault(entry);

                badges.Errors.IsVisible = counts.Errors > 0;
                badges.Warnings.IsVisible = counts.Warnings > 0;
                badges.Panel.IsVisible = counts.Errors > 0 || counts.Warnings > 0;
                badges.ErrorsText.Text = counts.Errors.ToString(System.Globalization.CultureInfo.InvariantCulture);
                badges.WarningsText.Text = counts.Warnings.ToString(System.Globalization.CultureInfo.InvariantCulture);
                ToolTip.SetTip(badges.Errors, counts.Errors > 0 ? ContributeProblems.ErrorsTip(counts.Errors) : null);
                ToolTip.SetTip(badges.Warnings, counts.Warnings > 0 ? ContributeProblems.WarningsTip(counts.Warnings) : null);
            }

            this.ShowProblemsOfEntryOpen();
            this.ShowProblemsInComponentList();
        }

        // ###########################################################################################
        // What the checks found on each component, on its row of the list on the far left (owner request,
        // 2026-11-04: "this should also show on each component in the component list in the left-side"). On
        // the rows as they are: a row made again lets go of the selection, and the list scrolls back to it.
        // ###########################################################################################
        private void ShowProblemsInComponentList()
        {
            foreach (ComponentListItem item in this.thisComponentItems)
            {
                item.Problems = this.thisProblemsOfComponents.GetValueOrDefault(item.BoardLabel);
            }
        }

        // ###########################################################################################
        // "WITH ISSUES", THE BOX OVER THE LIST ON THE FAR LEFT (owner requests, 2026-11-04: "it should only
        // show components with either a warning or an error"; and "I only want to see this checkbox when it at
        // least has one component with an issue"). Main hands it over as the tab is shown; on every other tab it
        // is not there. Its row follows its IsVisible (Main.axaml), and a click on the row is Main's.
        //
        //   - EACH BOARD ITS OWN, FOR THE RUN, as all that is done to the lists here is.
        //   - WHAT IS BEING FIXED STAYS LISTED until another is picked, as the table keeps the row its cursor
        //     is on: picked for its error, its row going a second after the fix - the selection with it - would
        //     pull the list out from under the one working on it. Picked with nothing wrong when the box is
        //     ticked, it goes: the pane stays on it, as for a category unticked.
        //   - THE LAST ISSUE GONE TAKES THE TICK WITH IT: left ticked underneath, the box would come back ticked
        //     with the next warning - any new component has one - and narrow the list by itself.
        //   - THE LIST COUNTS WHAT THE ROWS' BADGES COUNT, every region's, so the box is there though the one
        //     with an issue is of the other region.
        // ###########################################################################################
        private CheckBox? thisIssuesBox;

        private readonly HashSet<string> thisBoardsOnlyWithIssues = new(StringComparer.OrdinalIgnoreCase);

        // The component picked, while it is listed for an issue it had or has - "" for none.
        private string thisPickedListedForIssues = string.Empty;

        // True while the box is set from code - no choice of anybody's.
        private bool thisFollowingIssuesBox;

        private bool OnlyWithIssues => this.thisCurrentEntry is { } entry && this.thisBoardsOnlyWithIssues.Contains(entry.ExcelDataFile);

        internal void UseIssuesBox(CheckBox? box)
        {
            if (this.thisIssuesBox is { } handedOver)
            {
                handedOver.IsCheckedChanged -= this.OnIssuesBoxChanged;
            }

            this.thisIssuesBox = box;

            if (box is not null)
            {
                this.ShowIssuesBox();
                box.IsCheckedChanged += this.OnIssuesBoxChanged;
            }
        }

        // "Clear" on the list: nothing narrowed. The list is read again by whoever asks.
        internal void LetGoOfOnlyWithIssues()
        {
            if (this.thisCurrentEntry is { } entry)
            {
                this.thisBoardsOnlyWithIssues.Remove(entry.ExcelDataFile);
            }
        }

        private void OnIssuesBoxChanged(object? sender, Avalonia.Interactivity.RoutedEventArgs e)
        {
            if (this.thisFollowingIssuesBox || this.thisCurrentEntry is not { } entry)
            {
                return;
            }

            if (this.thisIssuesBox?.IsChecked == true)
            {
                this.thisBoardsOnlyWithIssues.Add(entry.ExcelDataFile);
            }
            else
            {
                this.thisBoardsOnlyWithIssues.Remove(entry.ExcelDataFile);
            }

            this.ShowComponents();
            this.ShowNavSelection();
        }

        private void ShowIssuesBox()
        {
            if (this.thisIssuesBox is not { } box)
            {
                return;
            }

            this.thisFollowingIssuesBox = true;

            try
            {
                box.IsVisible = this.thisProblemsOfComponents.Count > 0;
                box.IsChecked = this.OnlyWithIssues;
            }
            finally
            {
                this.thisFollowingIssuesBox = false;
            }
        }

        private bool HasIssues(string label) => this.thisProblemsOfComponents.ContainsKey(label);

        // Before the list is worked out: with no issue left nothing is narrowed, now or by the next one.
        private void LetGoOfIssuesTickWithNoIssues()
        {
            if (this.thisProblemsOfComponents.Count == 0)
            {
                this.LetGoOfOnlyWithIssues();
            }
        }

        private List<ComponentListItem> NarrowedToIssues(List<ComponentListItem> items)
        {
            if (!this.OnlyWithIssues)
            {
                this.thisPickedListedForIssues = string.Empty;

                return items;
            }

            string picked = this.thisPickedComponent;

            if (this.HasIssues(picked))
            {
                this.thisPickedListedForIssues = picked;
            }
            else if (!string.Equals(this.thisPickedListedForIssues, picked, StringComparison.OrdinalIgnoreCase))
            {
                this.thisPickedListedForIssues = string.Empty;
            }

            return items
                .Where(item => this.HasIssues(item.BoardLabel) || string.Equals(item.BoardLabel, this.thisPickedListedForIssues, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        // ###########################################################################################
        // The panel of the page open, under its headline (owner request, 2026-11-04: "should come below the
        // headline everywhere") - not over a table, which marks them on its cells.
        //
        //   - "OVERVIEW" HAS NONE (owner request, 2026-11-04: "there is no need to show the first panel ... The
        //     user can scroll down and see exactly where the issues are"): its own panels say each on the part
        //     it is about, and its badge how many the draft has.
        //   - WITH A COMPONENT SELECTED, A COMPONENT'S PAGE LISTS THAT ONE'S AND NOTHING ELSE (owner request,
        //     2026-11-04: "I do not want to see any warnings or errors for other components"), so with nothing
        //     wrong there is no panel. With none selected, every component's, with the way to each.
        // ###########################################################################################
        private void ShowProblemsOfEntryOpen()
        {
            ContributeEntry? open = this.IsTableOpen ? null : this.EntryOpen();

            List<EntryProblem> here = open is null or ContributeEntry.Overview
                ? []
                : ContributeProblems.OnEntry(this.thisProblemsPlaced, open.Value, this.PickedLabels);

            InPlaceList.Sync(
                this.thisEntryProblemRows,
                here.Take(ContributeProblems.ListedAtMost).Select(this.RowOfProblem).ToList(),
                (shown, wanted) => shown.ShowsTheSameAs(wanted),
                wanted => wanted);

            this.EntryProblemsPanel.IsVisible = here.Count > 0;

            if (here.Count == 0 || open is not { } entry)
            {
                return;
            }

            this.PutEntryProblemsPanelIn(this.ProblemsHostOfPageOpen());

            BoardProblemCounts counts = BoardProblemCounts.Of(here.Select(item => item.Problem));
            BoardStatusState state = counts.Errors > 0 ? BoardStatusState.Error : BoardStatusState.ToDo;

            TabDrafts.UseStateClasses(this.EntryProblemsPanel, state);
            TabDrafts.UseStateClasses(this.EntryProblemsIcon, state);
            this.EntryProblemsIcon.Text = TabDrafts.StatusIconOf(state);

            this.EntryProblemsTitle.Text = ContributeProblems.Heading(counts);
            this.EntryProblemsHelp.Text = ContributeProblems.Help(counts, ofComponents: this.thisPickedComponent.Length == 0 && ContributeEntries.PageOf(entry) is not null);

            int notListed = here.Count - this.thisEntryProblemRows.Count;

            this.EntryProblemsMore.Text = notListed > 0 ? ContributeProblems.More(notListed) : string.Empty;
            this.EntryProblemsMore.IsVisible = notListed > 0;
        }

        // Under the headline of the page open. No problem has its place on "KiCad: Files" or "Board: Name"
        // (ContributeProblems.EntryOf), so what is left is the pane - a component or a part of the board - and
        // a component's page with none picked. In the pane under the whole block, the headline and a
        // component's buttons, so those stay where they are as a component with issues is followed by one
        // without.
        private Decorator ProblemsHostOfPageOpen() => this.IsComponentOpen ? this.ComponentPane.ProblemsHost : this.NoComponentProblemsHost;

        internal Decorator? EntryProblemsHostForTests => this.EntryProblemsPanel.Parent as Decorator;

        private void PutEntryProblemsPanelIn(Decorator host)
        {
            if (ReferenceEquals(this.EntryProblemsPanel.Parent, host))
            {
                return;
            }

            if (this.EntryProblemsPanel.Parent is Decorator old)
            {
                old.Child = null;
            }

            host.Child = this.EntryProblemsPanel;
        }

        private EntryProblemRow RowOfProblem(EntryProblem item) =>
            new(
                item,
                ContributeProblems.LevelWord(item.Problem.Level),
                canOpen: item.ComponentLabel.Length > 0 && (!this.IsPaneOnComponent(item.ComponentLabel) || !this.ShowsRegionOf(item)),
                new ActionCommand(() => this.RunGuarded(() => this.OpenComponentOfProblemAsync(item), "Open the component of a problem")));

        private bool ShowsRegionOf(EntryProblem item)
        {
            string region = ContributeProblems.RegionToOpen(item, BoardRegions.Of(this.BoardForList()));

            return region.Length == 0 || string.Equals(region, this.thisCurrentRegion, StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // The component a problem is about, on the page it is on - as a pick in the list would, whichever
        // view is picked (ShowEntryAsync) - in the region its row is of, or the page does not list that row
        // (code review, 2026-11-04). The region is all of CRT's, as the buttons under the list pick it.
        // ###########################################################################################
        internal System.Threading.Tasks.Task OpenComponentOfProblemAsync(EntryProblem item)
        {
            if (!this.ShowsRegionOf(item))
            {
                string region = ContributeProblems.RegionToOpen(item, BoardRegions.Of(this.BoardForList()));

                if (this.thisMainWindow is { } main)
                {
                    main.PickRegion(region);
                }
                else
                {
                    this.UseRegion(region);
                }
            }

            this.thisPickedComponent = item.ComponentLabel;
            this.thisPickedSeveral = [];

            return this.ShowEntryAsync(item.Entry);
        }
    }

    // One line of the panel over a page: what is wrong, and the way to the component it is about.
    public sealed class EntryProblemRow
    {
        internal EntryProblemRow(EntryProblem source, string level, bool canOpen, ICommand openCommand)
        {
            this.Source = source;
            this.IsError = source.Problem.Level == BoardProblemLevel.Error;
            this.IsWarning = !this.IsError;
            this.Level = level;
            this.Message = source.Text;
            this.Component = source.ComponentLabel;
            this.CanOpen = canOpen;
            this.OpenLabel = source.ComponentLabel.Length > 0 ? ContributeProblems.OpenLabel(source.ComponentLabel) : string.Empty;
            this.OpenCommand = openCommand;
        }

        // What it is about, as its OpenCommand holds it.
        internal EntryProblem Source { get; }

        public bool IsError { get; }

        public bool IsWarning { get; }

        public string Level { get; }

        public string Message { get; }

        public string Component { get; }

        public bool CanOpen { get; }

        public string OpenLabel { get; }

        public ICommand OpenCommand { get; }

        // Kept in place when it draws the same and opens the same - a row made again under a press throws the
        // click away, and one saying the same of another page went on opening the page left behind (code
        // review, 2026-11-04). By value, so a read of the checks finding the same keeps its rows.
        internal bool ShowsTheSameAs(EntryProblemRow other) =>
            this.CanOpen == other.CanOpen &&
            this.Source == other.Source;
    }
}
