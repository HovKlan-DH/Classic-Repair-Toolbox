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
    // The right-hand side of the "Systems" screen (owner request, 2026-09-27): one system, as the
    // server's SystemOverviewFlow gives it. The list it sits beside is TabMaintainer.Systems.cs.
    //
    // *** IN SIX VIEWS *** (five since 2026-10-03 - owner request: "all the same functionalities, as
    // the 'Contributor Submissions' has" - and History since 2026-10-04) - SystemSections says what
    // each is and why:
    //
    //   SystemView.axaml.cs         - the header, reading one system, and the Contributor and
    //                                 Statistics views
    //   SystemView.Sections.cs      - which view is shown, and reading the one that needs the server
    //   SystemView.Table.cs         - Board data: BETA's board, and publishing a change straight to
    //                                 BETA (the reason asked in PublishSystemChangeWindow)
    //   SystemView.Files.cs         - Files: every file the system uses in BETA
    //   SystemView.Maintainers.cs   - Maintainer: who maintains it, a list (changing that is the
    //                                 administrator's, Account > Maintainers - MaintainerPoolView)
    //   SystemView.History.cs       - History: a card per submission, with what it changed
    //   SystemView.Stages.cs        - the stage line under the name: its newest submission, BETA
    //                                 and the stable source (2026-10-04)
    //   SystemView.Stable.cs        - the BETA / Stable switch on Board data and Files, and the
    //                                 stable source's read-only table and file tree (2026-10-04)
    //
    // *** THE WORDS ARE SystemsDisplay's AND SystemSections', AND THE FACTS THE SERVER's. *** Nothing
    // here decides anything.
    //
    // A system not in CRT's drop-down lists yet gets SystemPlacementView in its Maintainer view
    // (2026-09-27), from the listing the Systems screen reads beside its list (UseListing).
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
            this.WireTable();
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

            // Not in the lists: said above the views, since the placement is in one of them.
            if (this.FindControl<TextBlock>("SystemPlacementText") is TextBlock placement)
            {
                string? line = SystemSections.PlacementLine(this.thisListing, systemId);
                placement.Text = line ?? string.Empty;
                placement.IsVisible = line is not null;
            }

            this.GetControl<SystemPlacementView>("PlacementView").Show(this.thisListing, systemId);
        }

        // The system on screen, or null.
        public SystemOverviewEntry? ShownSystem { get; private set; }

        // ###########################################################################################
        // Shows one system: what the list already knows at once, the rest when it arrives - its detail,
        // then the view on show when that view is read from the server (Board data, Files). Null
        // empties the panel. An answer for a system no longer chosen is dropped.
        //
        // `keepShown` is the queue's own check re-reading the system already on screen: what is
        // shown stays until the new answer replaces it, rather than blanking every minute - and the
        // table and the file list are read again only when BETA moved under them (2026-10-04), the
        // table never while it holds a change not published yet.
        //
        // *** ANOTHER SYSTEM EMPTIES THE TABLE AND THE FILES. *** The caller has asked about unsaved
        // table changes first (TabMaintainer.OnSystemsSelectionChanged); the same system shown again
        // keeps both.
        // ###########################################################################################
        public async Task ShowSystemAsync(SystemOverviewEntry? system, bool keepShown = false)
        {
            int request = ++this.thisRequest;

            bool sameSystem = system is not null && string.Equals(
                this.ShownSystem?.SystemId, system.SystemId, StringComparison.Ordinal);

            bool quiet = keepShown && sameSystem;

            if (!sameSystem)
                this.ForgetSystemContent();

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

            // The table or the files read before BETA moved are read again (owner report,
            // 2026-10-04) - see CatchUpWithBetaAsync. Nothing is held yet for another system.
            await this.CatchUpWithBetaAsync(detail.Value!.System);

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
            this.ForgetSystemContent();
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
            // Who maintains it, for whether a table read earlier is out of date (CatchUpWithBetaAsync).
            this.thisShownMaintainers = detail.Maintainers.Select(maintainer => maintainer.AccountId).ToList();

            // Its submissions, for the stage line's "Submitted" (SystemView.Stages.cs).
            this.UseStageSubmissions(detail);

            // The detail's own facts - newer than the list's, if anything moved between the two.
            this.ShowSummary(detail.System);
            this.ShowSections(detail);
        }

        // ###########################################################################################
        // The name, and - only when it is so - where its data is, who maintains it and that it is
        // closed (owner request, 2026-10-03: "only show where there is something odd/off"), and its
        // revisions; or, with nothing chosen, what to do. The views show once a system is chosen.
        // ###########################################################################################
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

            // Where it is: its submission, BETA, the stable source (SystemView.Stages.cs) - which
            // replaced the revisions line, whose revisions it carries.
            this.ShowStages(system);

            this.SetShown("SystemSectionBar", system is not null);
            this.SetShown("SectionsPanel", system is not null);

            // Which half of Board data and Files the system can show (SystemView.Stable.cs).
            this.ApplyTree();

            this.ShowListing();
        }

        // ###########################################################################################
        // What the detail answers, each in its view. Each says so when it is empty - "Nobody
        // maintains this system" is the very thing somebody opens this screen to find, and a blank
        // space would hide it.
        // ###########################################################################################
        private void ShowSections(SystemDetailAnswer? detail)
        {
            if (this.FindControl<StackPanel>("ContributorsSection") is not StackPanel contributors)
                return;

            // The maintainers are SystemView.Maintainers.cs', the history SystemView.History.cs'.
            this.ShowMaintainers(detail);
            this.ShowHistory(detail);

            contributors.Children.Clear();

            this.ShowViews(detail);

            if (detail is null)
                return;

            // ---- Contributor ------------------------------------------------------------------
            contributors.Children.Add(SystemView.Heading(SystemsDisplay.ContributorsHeading(detail.Contributors.Count), detail.Contributors.Count == 0));

            // No addresses for a system this account does not maintain (2026-10-05) - said, once.
            if (detail.AddressesHidden && detail.Contributors.Count > 0)
                contributors.Children.Add(SystemView.Note(SystemsDisplay.AddressesHiddenLine));

            foreach (SystemContributorEntry contributor in detail.Contributors)
            {
                contributors.Children.Add(SystemView.TwoLines(
                    SystemsDisplay.ContributorName(contributor, detail.AddressesHidden),
                    SystemsDisplay.ContributorRecord(contributor),
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
        private void ShowViews(SystemDetailAnswer? detail)
        {
            if (this.FindControl<StackPanel>("ViewsSection") is not StackPanel section)
                return;

            section.Children.Clear();

            if (detail is null)
                return;

            if (detail.Views is not BoardViewStatistics views)
            {
                section.Children.Add(SystemView.Heading(SystemsDisplay.NoViewCountsLine, isEmpty: true));
                return;
            }

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

        private void SetShown(string name, bool shown)
        {
            if (this.FindControl<Control>(name) is Control control)
                control.IsVisible = shown;
        }

        // Every text the panel shows, top to bottom, in every view - for tests.
        internal IReadOnlyList<string> TextsForTests() =>
            new[] { "SystemTitleText", "SystemStateText", "SystemListingText", "SystemPlacementText", "MessageText" }
                .Select(name => this.FindControl<TextBlock>(name))
                .Where(block => block is { IsVisible: true })
                .Select(block => TabMaintainer.TextOf(block!))
                .Concat(new[] { "MaintainersSection", "ViewsSection", "ContributorsSection", "HistoryItemsSection" }
                    .SelectMany(name => SystemView.TextsIn(this.FindControl<StackPanel>(name))))
                .ToList();

        // The texts of one view's panel, top to bottom - for tests.
        internal IReadOnlyList<string> SectionTextsForTests(string panelName) =>
            SystemView.TextsIn(this.FindControl<StackPanel>(panelName)).ToList();

        private static IEnumerable<string> TextsIn(Panel? panel)
        {
            if (panel is null)
                yield break;

            foreach (Control child in panel.Children)
            {
                foreach (string text in SystemView.TextsOf(child))
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
                    foreach (string text in SystemView.TextsIn(inner))
                        yield return text;
                    break;

                case Decorator decorator:
                    foreach (string text in SystemView.TextsOf(decorator.Child))
                        yield return text;
                    break;
            }
        }

        internal void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);
    }
}
