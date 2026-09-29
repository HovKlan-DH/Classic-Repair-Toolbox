using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // The right-hand side of the "Systems" screen (owner request, 2026-09-27): one system's
    // maintainers, contributors and recent submissions, as the server's SystemOverviewFlow gives
    // them. The list it sits beside is TabMaintainer.Systems.cs.
    //
    // *** THE WORDS ARE SystemsDisplay's, AND THE FACTS THE SERVER's. *** This file fetches one
    // answer and lays it out in three sections; nothing here decides anything.
    //
    // A system not in CRT's drop-down lists yet gets SystemPlacementView above those sections
    // (2026-09-27), from the listing the Systems screen reads beside its list (UseListing).
    //
    // An ADMINISTRATOR sets the maintainers here too (2026-09-27) - SystemView.Maintainers.cs.
    // ###########################################################################################
    public partial class SystemView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        // A counter that drops an answer the selection has moved past.
        private int thisRequest;

        public SystemView()
        {
            this.InitializeComponent();
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
            this.GetControl<SystemPlacementView>("PlacementView").Initialize(client, session);
        }

        // BETA's drop-down lists and the systems not in them - read with the Systems list.
        private SystemListingAnswer? thisListing;

        // Told after a placement was saved, so the screen can read the lists again.
        public Func<Task>? AfterPlacementSaved
        {
            get => this.GetControl<SystemPlacementView>("PlacementView").AfterSaved;
            set => this.GetControl<SystemPlacementView>("PlacementView").AfterSaved = value;
        }

        // ###########################################################################################
        // The drop-down lists as last read. Null (not read, or signed out) shows no placement and no
        // listed line - never "not listed" on the strength of a request that failed.
        // ###########################################################################################
        public void UseListing(SystemListingAnswer? listing)
        {
            this.thisListing = listing;
            this.ShowListing();
        }

        internal SystemPlacementView PlacementForTests => this.GetControl<SystemPlacementView>("PlacementView");

        private void ShowListing()
        {
            string? systemId = this.ShownSystem?.SystemId;

            if (this.FindControl<TextBlock>("SystemListingText") is TextBlock listed)
            {
                string? line = SystemPlacementDisplay.ListedLine(this.thisListing, systemId);
                listed.Text = line ?? string.Empty;
                listed.IsVisible = line is not null;
            }

            this.GetControl<SystemPlacementView>("PlacementView").Show(this.thisListing, systemId);
        }

        // The system on screen, or null.
        public SystemOverviewEntry? ShownSystem { get; private set; }

        // ###########################################################################################
        // Shows one system: what the list already knows at once, the rest when it arrives. Null
        // empties the panel. An answer for a system no longer chosen is dropped.
        //
        // `keepShown` is the queue's own check re-reading the system already on screen: what is
        // shown stays until the new answer replaces it, rather than blanking every minute.
        // ###########################################################################################
        public async Task ShowSystemAsync(SystemOverviewEntry? system, bool keepShown = false)
        {
            int request = ++this.thisRequest;

            bool quiet = keepShown && system is not null && string.Equals(
                this.ShownSystem?.SystemId, system.SystemId, System.StringComparison.Ordinal);

            if (!quiet)
            {
                this.ShowSummary(system);
                this.ShowSections(null);
                this.ShowMessage(null, isError: false);
            }

            if (system is null || this.thisClient is null || this.thisSession is null)
                return;

            if (!quiet)
                this.ShowMessage("Reading the system's contributors and submissions...", isError: false);

            ReviewApiResult<SystemDetailAnswer> detail = await this.thisClient.GetSystemDetailAsync(this.thisSession, system.SystemId);

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

            // An administrator's list of accounts to add from - read with a system freshly opened,
            // so an account made by accepting an invitation meanwhile is there to choose.
            if (!quiet)
                await this.ReadAccountsAsync();
        }

        // Signed out: empty, with nothing of the previous account's on screen.
        public void Clear()
        {
            this.thisRequest++;
            this.thisListing = null;
            this.thisAccounts.Clear();
            this.ShowPoolMessage(null, isError: false);
            this.ShowSummary(null);
            this.ShowSections(null);
            this.ShowMessage(null, isError: false);
        }

        // Draws an answer without a server - for tests.
        internal void ShowDetailForTests(SystemDetailAnswer detail)
        {
            this.thisRequest++;
            this.ShowDetail(detail);
        }

        private void ShowDetail(SystemDetailAnswer detail)
        {
            // The detail's own facts - newer than the list's, if anything moved between the two.
            this.ShowSummary(detail.System);
            this.ShowSections(detail);
        }

        // The name, where its data is, and its revisions - or, with nothing chosen, what to do.
        private void ShowSummary(SystemOverviewEntry? system)
        {
            this.ShownSystem = system;

            if (this.FindControl<TextBlock>("SystemTitleText") is TextBlock title)
            {
                title.Text = system is null ? "Select a system" : SystemsDisplay.Name(system);
                title.FontWeight = system is null ? FontWeight.Normal : FontWeight.SemiBold;
                title.Opacity = system is null ? 0.7 : 1;
            }

            if (this.FindControl<TextBlock>("SystemStateText") is TextBlock state)
            {
                // What is still to be done in bold, as in the list (owner request, 2026-09-27).
                TabMaintainer.ShowParts(state, system is null ? [] : SystemsDisplay.ListLineParts(system));
                state.IsVisible = system is not null;
            }

            if (this.FindControl<TextBlock>("SystemRevisionsText") is TextBlock revisions)
            {
                string? line = system is null ? null : SystemsDisplay.Revisions(system);
                revisions.Text = line ?? string.Empty;
                revisions.IsVisible = line is not null;
            }

            this.ShowListing();
        }

        // ###########################################################################################
        // The three sections. Each says so when it is empty - "Nobody maintains this system" is the
        // very thing somebody opens this screen to find, and a blank space would hide it.
        // ###########################################################################################
        private void ShowSections(SystemDetailAnswer? detail)
        {
            var contributors = this.FindControl<StackPanel>("ContributorsSection");
            var submissions = this.FindControl<StackPanel>("SubmissionsSection");

            if (contributors is null || submissions is null)
                return;

            // The maintainers - with the administrator's controls - are SystemView.Maintainers.cs'.
            this.thisDetail = detail;
            this.ShowMaintainers(detail);

            contributors.Children.Clear();
            submissions.Children.Clear();

            this.ShowViews(detail?.Views);

            if (detail is null)
                return;

            contributors.Children.Add(SystemView.Heading(SystemsDisplay.ContributorsHeading(detail.Contributors.Count), detail.Contributors.Count == 0));

            foreach (SystemContributorEntry contributor in detail.Contributors)
            {
                contributors.Children.Add(SystemView.TwoLines(
                    SystemsDisplay.ContributorName(contributor),
                    SystemsDisplay.ContributorRecord(contributor),
                    note: null));
            }

            // What has happened to it, newest first and date first (owner request, 2026-09-27) - its
            // submissions' sending and decisions among everything else. An older server sends no
            // history, and then the submissions are listed on their own as before.
            if (detail.History is IReadOnlyList<SystemHistoryEntry> history)
            {
                submissions.Children.Add(SystemView.Heading(SystemsDisplay.HistoryHeading(history.Count), history.Count == 0));

                foreach (SystemHistoryEntry entry in history)
                {
                    submissions.Children.Add(SystemView.TwoLines(
                        SystemsDisplay.HistoryLine(entry),
                        SystemsDisplay.HistoryFooter(entry),
                        SystemsDisplay.HistoryNote(entry)));
                }

                return;
            }

            submissions.Children.Add(SystemView.Heading(SystemsDisplay.SubmissionsHeading(detail.Submissions.Count), detail.Submissions.Count == 0));

            foreach (SystemSubmissionEntry submission in detail.Submissions)
            {
                StackPanel lines = SystemView.TwoLines(
                    SystemsDisplay.SubmissionTitle(submission),
                    SystemsDisplay.SubmissionFooter(submission),
                    SystemsDisplay.SubmissionComment(submission));

                // Its contributor discarded their own draft since sending it (owner request,
                // 2026-09-28) - in the failure colour, as on the queue row.
                if (submission.DraftDiscardedUtc is DateTimeOffset discarded)
                {
                    lines.Children.Add(new TextBlock
                    {
                        Text = DraftDiscardWording.Mark(discarded),
                        FontSize = 11,
                        FontWeight = FontWeight.SemiBold,
                        Foreground = Brushes.IndianRed,
                        TextWrapping = TextWrapping.Wrap
                    });
                }

                submissions.Children.Add(lines);
            }
        }

        // ###########################################################################################
        // How often CRT users look at the board (owner request, 2026-09-27): the counts, the countries
        // most views come from, the BETA-source views counted apart, and what a view is. No counts
        // from the server (an older one, or counts it could not read): no section at all - never
        // "no views", which would be a claim about the board.
        // ###########################################################################################
        private void ShowViews(BoardViewStatistics? views)
        {
            if (this.FindControl<StackPanel>("ViewsSection") is not StackPanel section)
                return;

            section.Children.Clear();
            section.IsVisible = views is not null;

            if (views is null)
                return;

            bool none = SystemsDisplay.HasNoViews(views);
            section.Children.Add(SystemView.Heading(SystemsDisplay.ViewsHeading(views), none));

            if (none)
                return;

            section.Children.Add(SystemView.Counts(SystemsDisplay.ViewCountRuns(views), faint: false));

            if (SystemsDisplay.ViewCountriesLine(views) is string countries)
                section.Children.Add(SystemView.Line(countries));

            if (SystemsDisplay.ViewBetaRuns(views) is IReadOnlyList<ReviewNoteRun> beta)
                section.Children.Add(SystemView.Counts(beta, faint: true));

            section.Children.Add(new TextBlock
            {
                Text = SystemsDisplay.ViewExplanation,
                FontSize = 11,
                Opacity = 0.7,
                TextWrapping = TextWrapping.Wrap
            });
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

        // One entry: its main line, a grey line under it, and - when there is one - what the
        // contributor was told, in italics.
        private static StackPanel TwoLines(string main, string footer, string? note)
        {
            var panel = new StackPanel { Spacing = 1 };

            panel.Children.Add(new TextBlock { Text = main, TextWrapping = TextWrapping.Wrap });

            if (footer.Length > 0)
                panel.Children.Add(new TextBlock { Text = footer, FontSize = 11, Opacity = 0.7, TextWrapping = TextWrapping.Wrap });

            if (note is not null)
                panel.Children.Add(new TextBlock { Text = note, FontSize = 11, FontStyle = FontStyle.Italic, Opacity = 0.85, TextWrapping = TextWrapping.Wrap });

            return panel;
        }

        // Every text the panel shows, top to bottom - for tests.
        internal IReadOnlyList<string> TextsForTests() =>
            new[] { "SystemTitleText", "SystemStateText", "SystemRevisionsText", "SystemListingText", "MessageText" }
                .Select(name => this.FindControl<TextBlock>(name))
                .Where(block => block is { IsVisible: true })
                .Select(block => TabMaintainer.TextOf(block!))
                .Concat(new[] { "MaintainersSection", "ViewsSection", "ContributorsSection", "SubmissionsSection" }
                    .SelectMany(name => SystemView.TextsIn(this.FindControl<StackPanel>(name))))
                .ToList();

        private static IEnumerable<string> TextsIn(Panel? panel)
        {
            if (panel is null)
                yield break;

            foreach (Control child in panel.Children)
            {
                if (child is TextBlock block)
                    yield return TabMaintainer.TextOf(block);
                else if (child is Panel inner)
                {
                    foreach (string text in SystemView.TextsIn(inner))
                        yield return text;
                }
            }
        }

        internal void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);
    }
}
