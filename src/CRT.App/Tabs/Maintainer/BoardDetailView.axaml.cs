using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace CRT
{
    // ###########################################################################################
    // The right-hand side of the "Boards" screen (owner request, 2026-09-27): one board, as the
    // server's BoardOverviewFlow gives it. The list it sits beside is TabMaintainer.Boards.cs.
    //
    // *** IN SIX VIEWS *** (five since 2026-10-03 - owner request: "all the same functionalities, as
    // the 'Contributor Submissions' has" - and History since 2026-10-04) - BoardSections says what
    // each is and why:
    //
    //   BoardDetailView.axaml.cs         - the header, reading one board, and the Contributor and
    //                                 Statistics views
    //   BoardDetailView.Sections.cs      - which view is shown, and reading the one that needs the server
    //   BoardDetailView.Table.cs         - Board data: BETA's board, and publishing a change straight to
    //                                 BETA (the reason asked in PublishBoardChangeWindow)
    //   BoardDetailView.Files.cs         - Files: every file the board uses in BETA
    //   BoardDetailView.Maintainers.cs   - Maintainer: who maintains it, a list (changing that is the
    //                                 administrator's, Account > Maintainers - MaintainerPoolView)
    //   BoardDetailView.History.cs       - History: a card per submission, with what it changed
    //   BoardDetailView.Stable.cs        - the BETA / Stable switch on Board data and Files, and the
    //                                 stable source's read-only table and file tree (2026-10-04)
    //   BoardDetailView.Compare.cs       - "Compare sources": either table coloured against the other
    //                                 source (2026-10-09)
    //
    // (The stage line under the name - BoardDetailView.Stages.cs - was removed on 2026-10-09, owner
    // request; it is in git history.)
    //
    // *** THE WORDS ARE BoardsDisplay's AND BoardSections', AND THE FACTS THE SERVER's. *** Nothing
    // here decides anything.
    //
    // A board not in CRT's drop-down lists yet gets BoardPlacementView in its Maintainer view
    // (2026-09-27), from the listing the Boards screen reads beside its list (UseListing).
    // ###########################################################################################
    public partial class BoardDetailView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        // A counter that drops an answer the selection has moved past.
        private int thisRequest;

        public BoardDetailView()
        {
            this.InitializeComponent();
            this.WireTable();
            this.WireCompare();
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
            this.GetControl<BoardPlacementView>("PlacementView").Initialize(client, session);
        }

        // BETA's drop-down lists and the boards not in them - read with the Boards list.
        private BoardListingAnswer? thisListing;

        // Told after a placement was saved, so the screen can read the lists again.
        public Func<Task>? AfterPlacementSaved
        {
            get => this.GetControl<BoardPlacementView>("PlacementView").AfterSaved;
            set => this.GetControl<BoardPlacementView>("PlacementView").AfterSaved = value;
        }

        // ###########################################################################################
        // The drop-down lists as last read. Null (not read, or signed out) shows no placement and no
        // listed line - never "not listed" on the strength of a request that failed.
        // ###########################################################################################
        public void UseListing(BoardListingAnswer? listing)
        {
            this.thisListing = listing;
            this.ShowListing();
        }

        internal BoardPlacementView PlacementForTests => this.GetControl<BoardPlacementView>("PlacementView");

        private void ShowListing()
        {
            string? boardId = this.ShownBoard?.BoardId;

            if (this.FindControl<TextBlock>("BoardListingText") is TextBlock listed)
            {
                string? line = BoardPlacementDisplay.ListedLine(this.thisListing, boardId);
                listed.Text = line ?? string.Empty;
                listed.IsVisible = line is not null;
            }

            // Not in the lists: said above the views, since the placement is in one of them.
            if (this.FindControl<TextBlock>("BoardPlacementText") is TextBlock placement)
            {
                string? line = BoardSections.PlacementLine(this.thisListing, boardId);
                placement.Text = line ?? string.Empty;
                placement.IsVisible = line is not null;
            }

            this.GetControl<BoardPlacementView>("PlacementView").Show(this.thisListing, boardId);
        }

        // The board on screen, or null.
        public BoardOverviewEntry? ShownBoard { get; private set; }

        // ###########################################################################################
        // Shows one board: what the list already knows at once, the rest when it arrives - its detail,
        // then the view on show when that view is read from the server (Board data, Files). Null
        // empties the panel. An answer for a board no longer chosen is dropped.
        //
        // `keepShown` is the queue's own check re-reading the board already on screen: what is
        // shown stays until the new answer replaces it, rather than blanking every minute - and the
        // table and the file list are read again only when BETA moved under them (2026-10-04), the
        // table never while it holds a change not published yet.
        //
        // *** ANOTHER BOARD EMPTIES THE TABLE AND THE FILES. *** The caller has asked about unsaved
        // table changes first (TabMaintainer.OnBoardsSelectionChanged); the same board shown again
        // keeps both.
        // ###########################################################################################
        public async Task ShowBoardAsync(BoardOverviewEntry? board, bool keepShown = false)
        {
            int request = ++this.thisRequest;

            bool sameBoard = board is not null && string.Equals(
                this.ShownBoard?.BoardId, board.BoardId, StringComparison.Ordinal);

            bool quiet = keepShown && sameBoard;

            if (!sameBoard)
                this.ForgetBoardContent();

            if (!quiet)
            {
                this.ShowSummary(board);
                this.ShowSections(null);
                this.ShowMessage(null, isError: false);
            }

            if (board is null || this.thisClient is null || this.thisSession is null)
                return;

            if (!quiet)
                this.ShowMessage("Reading the board's contributors and submissions...", isError: false);

            ReviewApiResult<BoardDetailAnswer> detail = await this.thisClient.GetBoardDetailAsync(this.thisSession, board.BoardId);

            if (request != this.thisRequest)
                return;

            if (!detail.IsOk)
            {
                // A quiet re-read that failed leaves what is shown; the queue's check says why.
                if (!quiet)
                    this.ShowMessage(detail.Message, isError: true);

                return;
            }

            this.ShowMessage(null, isError: false);
            this.ShowDetail(detail.Value!);

            // The table or the files read before BETA moved are read again (owner report,
            // 2026-10-04) - see CatchUpWithBetaAsync. Nothing is held yet for another board.
            await this.CatchUpWithBetaAsync(detail.Value!.Board);

            if (quiet)
                return;

            if (request == this.thisRequest)
                await this.LoadShownSectionAsync();
        }

        // Signed out: empty, with nothing of the previous account's on screen.
        public void Clear()
        {
            this.thisRequest++;
            this.thisListing = null;
            this.ForgetBoardContent();
            this.ShowSummary(null);
            this.ShowSections(null);
            this.ShowMessage(null, isError: false);
        }

        // Draws an answer without a server - for tests.
        internal void ShowDetailForTests(BoardDetailAnswer detail)
        {
            this.thisRequest++;
            this.ShowDetail(detail);
        }

        private void ShowDetail(BoardDetailAnswer detail)
        {
            // Who maintains it, for whether a table read earlier is out of date (CatchUpWithBetaAsync).
            this.thisShownMaintainers = detail.Maintainers.Select(maintainer => maintainer.AccountId).ToList();

            // The detail's own facts - newer than the list's, if anything moved between the two.
            this.ShowSummary(detail.Board);
            this.ShowSections(detail);

            // Why BETA's table cannot be changed, above whichever view is open (code review,
            // 2026-10-09) - not known from an older server, which leaves it to the table.
            if (detail.MayEdit is bool mayEdit)
                this.ShowReadOnlyNotice(BoardSections.ReadOnlyReason(mayEdit, detail.MayNotEditReason));
        }

        // ###########################################################################################
        // The name, and - only when it is so - where its data is, who maintains it and that it is
        // closed (owner request, 2026-10-03: "only show where there is something odd/off"); or, with
        // nothing chosen, what to do. The views show once a board is chosen.
        // ###########################################################################################
        private void ShowSummary(BoardOverviewEntry? board)
        {
            this.ShownBoard = board;

            if (this.FindControl<TextBlock>("BoardTitleText") is TextBlock title)
            {
                title.Text = board is null ? "Select a board" : BoardsDisplay.Name(board);
                title.FontWeight = board is null ? FontWeight.Normal : FontWeight.SemiBold;
                title.Opacity = board is null ? 0.7 : 1;
            }

            if (this.FindControl<TextBlock>("BoardStateText") is TextBlock state)
            {
                // What is still to be done in bold, as in the list (owner request, 2026-09-27).
                TabMaintainer.ShowParts(state, board is null ? [] : BoardsDisplay.ListLineParts(board));
                state.IsVisible = board is not null;
            }

            this.SetShown("BoardSectionBar", board is not null);
            this.SetShown("SectionsPanel", board is not null);

            // Which half of Board data and Files the board can show (BoardDetailView.Stable.cs).
            this.ApplyTree();

            this.ShowListing();
        }

        // ###########################################################################################
        // What the detail answers, each in its view. Each says so when it is empty - "Nobody
        // maintains this board" is the very thing somebody opens this screen to find, and a blank
        // space would hide it.
        // ###########################################################################################
        private void ShowSections(BoardDetailAnswer? detail)
        {
            if (this.FindControl<StackPanel>("ContributorsSection") is not StackPanel contributors)
                return;

            // The maintainers are BoardDetailView.Maintainers.cs', the history BoardDetailView.History.cs'.
            this.ShowMaintainers(detail);
            this.ShowHistory(detail);

            contributors.Children.Clear();

            this.ShowViews(detail);

            if (detail is null)
                return;

            // ---- Contributor ------------------------------------------------------------------
            contributors.Children.Add(BoardDetailView.Heading(BoardsDisplay.ContributorsHeading(detail.Contributors.Count), detail.Contributors.Count == 0));

            // No addresses for a board this account does not maintain (2026-10-05) - said, once.
            if (detail.AddressesHidden && detail.Contributors.Count > 0)
                contributors.Children.Add(BoardDetailView.Note(BoardsDisplay.AddressesHiddenLine));

            // Each by name, in bold (owner request, 2026-10-09).
            foreach (BoardContributorEntry contributor in detail.Contributors)
            {
                contributors.Children.Add(BoardDetailView.TwoLines(
                    BoardsDisplay.ContributorRuns(contributor, detail.AddressesHidden),
                    [PersonRuns.Plain(BoardsDisplay.ContributorRecord(contributor))],
                    note: null));
            }
        }

        // ###########################################################################################
        // THE STATISTICS VIEW (owner request, 2026-10-03; graphs "we can take this later"): how often
        // CRT users look at the board (2026-09-27) - the counts, the countries most views come from,
        // the BETA-source views counted apart, and what a view is. No counts from the server (an
        // older one, or counts it could not read): said so - never "no views", which would be a
        // claim about the board.
        // ###########################################################################################
        private void ShowViews(BoardDetailAnswer? detail)
        {
            if (this.FindControl<StackPanel>("ViewsSection") is not StackPanel section)
                return;

            section.Children.Clear();

            if (detail is null)
                return;

            if (detail.Views is not BoardViewStatistics views)
            {
                section.Children.Add(BoardDetailView.Heading(BoardsDisplay.NoViewCountsLine, isEmpty: true));
                return;
            }

            bool none = BoardsDisplay.HasNoViews(views);
            section.Children.Add(BoardDetailView.Heading(BoardsDisplay.ViewsHeading(views), none));

            if (none)
                return;

            section.Children.Add(BoardDetailView.Counts(BoardsDisplay.ViewCountRuns(views), faint: false));

            if (BoardsDisplay.ViewCountriesLine(views) is string countries)
                section.Children.Add(BoardDetailView.Line(countries));

            if (BoardsDisplay.ViewBetaRuns(views) is IReadOnlyList<ReviewNoteRun> beta)
                section.Children.Add(BoardDetailView.Counts(beta, faint: true));

            // The views per day, as a graph (owner request, 2026-10-09) - from a server that sends them.
            if (views.Daily is IReadOnlyList<BoardViewDay> daily)
                section.Children.Add(this.ViewsPerDay(daily));

            section.Children.Add(new TextBlock
            {
                Text = BoardsDisplay.ViewExplanation,
                FontSize = 11,
                Opacity = 0.7,
                TextWrapping = TextWrapping.Wrap
            });
        }

        // ###########################################################################################
        // "Views per day", the range - 30, 90 or 365 days, the first to start with and the one chosen
        // kept for the next board, as the view itself is - and the graph (BoardViewsChart). The days
        // are UTC days, as the server counts them.
        // ###########################################################################################
        private int thisViewsRange = ViewsChartGeometry.Ranges[0];

        // The day the graph ends on - today, in UTC; a test sets its own.
        internal DateOnly? TodayForTests { get; set; }

        private StackPanel ViewsPerDay(IReadOnlyList<BoardViewDay> daily)
        {
            var panel = new StackPanel { Spacing = 6, Margin = new Avalonia.Thickness(0, 10, 0, 0) };
            var chart = new BoardViewsChart { Name = "ViewsChart" };
            var range = new StackPanel { Orientation = Avalonia.Layout.Orientation.Horizontal };
            DateOnly today = this.TodayForTests ?? DateOnly.FromDateTime(DateTime.UtcNow);

            void Show(int days)
            {
                this.thisViewsRange = days;
                chart.Days = ViewsChartGeometry.Days(daily, today, days);

                foreach (Button button in range.Children.OfType<Button>())
                    button.Classes.Set("Selected", Equals(button.Tag, days));
            }

            for (int index = 0; index < ViewsChartGeometry.Ranges.Count; index++)
            {
                int days = ViewsChartGeometry.Ranges[index];

                var button = new Button
                {
                    Name = $"ViewsRange{days}Button",
                    Content = $"{days} days",
                    Tag = days,
                    Classes = { "Segment" }
                };

                if (index == 0)
                    button.Classes.Add("First");

                if (index == ViewsChartGeometry.Ranges.Count - 1)
                    button.Classes.Add("Last");

                button.Click += (_, _) => Show(days);
                range.Children.Add(button);
            }

            var heading = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            heading.Children.Add(new TextBlock
            {
                Text = BoardsDisplay.ViewsPerDayHeading,
                FontSize = 14,
                FontWeight = FontWeight.SemiBold,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
            });
            Grid.SetColumn(range, 1);
            heading.Children.Add(range);

            panel.Children.Add(heading);
            panel.Children.Add(chart);
            panel.Children.Add(new TextBlock
            {
                Text = BoardsDisplay.ViewsPerDayExplanation,
                FontSize = 11,
                Opacity = 0.7,
                TextWrapping = TextWrapping.Wrap
            });

            Show(ViewsChartGeometry.Ranges.Contains(this.thisViewsRange) ? this.thisViewsRange : ViewsChartGeometry.Ranges[0]);

            return panel;
        }

        private static TextBlock Counts(IReadOnlyList<ReviewNoteRun> runs, bool faint)
        {
            var block = new TextBlock { TextWrapping = TextWrapping.Wrap };

            if (faint)
            {
                block.FontSize = 11;
                block.Opacity = 0.7;
            }

            TabMaintainer.ShowCounts(block, runs);
            return block;
        }

        // A section's heading; faint when it is the whole section ("Nobody ..."), since then it is
        // an answer rather than a title over one.
        private static TextBlock Heading(string text, bool isEmpty) =>
            new()
            {
                Text = text,
                FontSize = 14,
                FontWeight = isEmpty ? FontWeight.Normal : FontWeight.SemiBold,
                Opacity = isEmpty ? 0.7 : 1,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Avalonia.Thickness(0, 0, 0, 2)
            };

        private static TextBlock Line(string text) =>
            new() { Text = text, TextWrapping = TextWrapping.Wrap };

        // A line naming somebody - the person in bold (PersonRuns).
        private static TextBlock Line(IReadOnlyList<ReviewNoteRun> runs)
        {
            var block = new TextBlock { TextWrapping = TextWrapping.Wrap };
            TabMaintainer.ShowCounts(block, runs);
            return block;
        }

        // A grey line under a heading that says how to read the list below it.
        private static TextBlock Note(string text) =>
            new()
            {
                Text = text,
                FontSize = 11,
                Opacity = 0.7,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Avalonia.Thickness(0, 0, 0, 4)
            };

        // One entry: its main line, a grey line under it, and - when there is one - what the
        // contributor was told, in italics. The people in the two lines in bold (PersonRuns).
        private static StackPanel TwoLines(IReadOnlyList<ReviewNoteRun> main, IReadOnlyList<ReviewNoteRun> footer, string? note)
        {
            var panel = new StackPanel { Spacing = 1 };

            panel.Children.Add(BoardDetailView.Line(main));

            if (footer.Any(run => run.Text.Length > 0))
            {
                var grey = new TextBlock { FontSize = 11, Opacity = 0.7, TextWrapping = TextWrapping.Wrap };
                TabMaintainer.ShowCounts(grey, footer);
                panel.Children.Add(grey);
            }

            if (note is not null)
                panel.Children.Add(new TextBlock { Text = note, FontSize = 11, FontStyle = FontStyle.Italic, Opacity = 0.85, TextWrapping = TextWrapping.Wrap });

            return panel;
        }

        private void SetShown(string name, bool shown)
        {
            if (this.FindControl<Control>(name) is Control control)
                control.IsVisible = shown;
        }

        // Every text the panel shows, top to bottom, in every view - for tests.
        internal IReadOnlyList<string> TextsForTests() =>
            new[] { "BoardTitleText", "BoardStateText", "BoardListingText", "BoardPlacementText", "MessageText" }
                .Select(name => this.FindControl<TextBlock>(name))
                .Where(block => block is { IsVisible: true })
                .Select(block => TabMaintainer.TextOf(block!))
                .Concat(new[] { "MaintainersSection", "ViewsSection", "ContributorsSection", "HistoryItemsSection" }
                    .SelectMany(name => BoardDetailView.TextsIn(this.FindControl<StackPanel>(name))))
                .ToList();

        // The texts of one view's panel, top to bottom - for tests.
        internal IReadOnlyList<string> SectionTextsForTests(string panelName) =>
            BoardDetailView.TextsIn(this.FindControl<StackPanel>(panelName)).ToList();

        private static IEnumerable<string> TextsIn(Panel? panel)
        {
            if (panel is null)
                yield break;

            foreach (Control child in panel.Children)
            {
                foreach (string text in BoardDetailView.TextsOf(child))
                    yield return text;
            }
        }

        // A control's texts - through a card's border too (the History view's).
        private static IEnumerable<string> TextsOf(Control? control)
        {
            switch (control)
            {
                case TextBlock block:
                    yield return TabMaintainer.TextOf(block);
                    break;

                case Panel inner:
                    foreach (string text in BoardDetailView.TextsIn(inner))
                        yield return text;
                    break;

                case Decorator decorator:
                    foreach (string text in BoardDetailView.TextsOf(decorator.Child))
                        yield return text;
                    break;
            }
        }

        internal void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);
    }
}
