using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using Handlers.Theming;

namespace CRT
{
    // ###########################################################################################
    // "SERVER VERSION" ON THE MAINTAINER TAB'S "ACCOUNT" SCREEN (owner requests, 2026-11-24: "make the "Account"
    // page a little prettier to look at, so it is more easy to overview", of this entry, and 'have a "Changelog"
    // like there is in "About" ... so I can check what has changed in server also').
    //
    // It only draws: TabMaintainer.Account.cs asks the server and hands over what came, ServerVersionDisplay words
    // it. The panels are made again at every answer - two or three cards, nothing to keep. The changelog is the
    // server's own (ServerStatusAnswer.Changelog), drawn by the "About" tab's means - GitHubMarkdown reads it,
    // ChangelogVersionHeadings words each version's heading "Version 12.2.0, ...", MarkdownRenderer draws each
    // block as it scrolls into view - and is built again only when its text changed, so reading the server again
    // neither rebuilds it nor moves the page.
    // ###########################################################################################
    public partial class ServerVersionView : UserControl
    {
        // The Font Awesome icon of each state - a shape of its own for each, so nothing rests on the color alone
        // (the "Contribute" tab's "Board status" icons).
        internal static string IconOf(ServerPanelState state) => state switch
        {
            ServerPanelState.Good => "",
            ServerPanelState.Waiting => "",
            ServerPanelState.Problem => "",
            _ => "",
        };

        // The changelog whose blocks are listed now - null when none is.
        private string? thisChangelogShown;

        private MarkdownRenderer? thisRenderer;

        public ServerVersionView()
        {
            this.InitializeComponent();
        }

        // Whether any panel is shown - none before the first answer, or once cleared.
        public bool HasPanels => this.PanelsHost.Children.Count > 0;

        // ###########################################################################################
        // The panels, in the order given. None clears them.
        // ###########################################################################################
        public void ShowPanels(IReadOnlyList<ServerPanel>? panels)
        {
            this.PanelsHost.Children.Clear();

            foreach (ServerPanel panel in panels ?? [])
                this.PanelsHost.Children.Add(ServerVersionView.CardOf(panel));
        }

        // ###########################################################################################
        // The changelog under the panels. `answered`: the server said how it stands - then its changelog is
        // listed, or that it sends none (an older server). Not answered, there is nothing to say about it.
        // ###########################################################################################
        public void ShowChangelog(string? markdown, bool answered)
        {
            string? changelog = answered && !string.IsNullOrWhiteSpace(markdown) ? markdown : null;

            this.ChangelogTitle.IsVisible = answered;
            this.NoChangelogText.IsVisible = answered && changelog is null;

            if (changelog == this.thisChangelogShown)
                return;

            this.LetGoOfChangelogBlocks();
            this.thisChangelogShown = changelog;

            if (changelog is null)
                return;

            if (this.thisRenderer is null)
            {
                MarkdownRenderer renderer = new(this.FontSize, url => ExternalTargetLauncher.OpenOrSay(url));
                this.thisRenderer = renderer;

                // A block that leaves the view gives its place to the next one, and in between the list hands the
                // template no block at all - that builds nothing (the "About" tab's lesson).
                this.PageItems.ItemTemplate = new FuncDataTemplate<MarkdownBlockItem?>((item, _) => item == null ? null : renderer.Build(item));
            }

            foreach (MarkdownBlockItem item in MarkdownRenderer.Arrange(ChangelogVersionHeadings.Of(GitHubMarkdown.Parse(changelog))))
                this.PageItems.Items.Add(item);
        }

        // Everything gone - signed out, another account.
        public void Clear()
        {
            this.ShowPanels(null);
            this.ShowChangelog(null, answered: false);
        }

        // The list's first item is the page; the changelog's blocks follow it. Taken from the end.
        private void LetGoOfChangelogBlocks()
        {
            for (int index = this.PageItems.Items.Count - 1; index >= 1; index--)
                this.PageItems.Items.RemoveAt(index);
        }

        // ###########################################################################################
        // One card: the icon left, then the title, the line saying how it stands, and the facts in two columns
        // under them - the label quiet, the value's pieces bold where ServerVersionDisplay marks them.
        // ###########################################################################################
        private static Border CardOf(ServerPanel panel)
        {
            string state = panel.State.ToString();

            var icon = new TextBlock { Text = ServerVersionView.IconOf(panel.State) };
            icon.Classes.Add("ServerPanelIcon");
            icon.Classes.Add(state);

            var title = new TextBlock { Text = panel.Title };
            title.Classes.Add("ServerPanelTitle");

            var summary = new TextBlock { Text = panel.Summary };
            summary.Classes.Add("ServerPanelSummary");
            summary.Classes.Add(state);

            var words = new StackPanel { Spacing = 2 };
            words.Children.Add(title);
            words.Children.Add(summary);

            if (panel.Rows.Count > 0)
                words.Children.Add(ServerVersionView.FactsOf(panel.Rows));

            Grid.SetColumn(words, 1);

            var layout = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*") };
            layout.Children.Add(icon);
            layout.Children.Add(words);

            var card = new Border { Child = layout };
            card.Classes.Add("ServerPanel");
            card.Classes.Add(state);

            return card;
        }

        private static Grid FactsOf(IReadOnlyList<ServerPanelRow> rows)
        {
            var facts = new Grid
            {
                ColumnDefinitions = new ColumnDefinitions("Auto,*"),
                RowSpacing = 4,
                Margin = new Thickness(0, 8, 0, 0),
            };

            for (int index = 0; index < rows.Count; index++)
            {
                facts.RowDefinitions.Add(new RowDefinition(GridLength.Auto));

                var label = new TextBlock { Text = rows[index].Label };
                label.Classes.Add("ServerFactLabel");
                Grid.SetRow(label, index);

                var value = new TextBlock();
                value.Classes.Add("ServerFactValue");
                BoldRuns.Show(value, rows[index].Value.Select(run => (run.Text, run.IsCount)).ToList());
                Grid.SetRow(value, index);
                Grid.SetColumn(value, 1);

                facts.Children.Add(label);
                facts.Children.Add(value);
            }

            return facts;
        }

        // -----------------------------------------------------------------------------------
        // For tests
        // -----------------------------------------------------------------------------------

        // Every card as lines - "Server: This CRT and the server speak the same API", "Version: 4.6.0", ... - read off
        // the controls shown, so a test sees what is drawn.
        internal IReadOnlyList<string> PanelLinesForTests =>
            this.PanelsHost.Children.OfType<Border>().SelectMany(ServerVersionView.LinesOf).ToList();

        // The changelog's blocks listed, and the page's own lines about it.
        internal int ChangelogBlocksForTests => this.PageItems.Items.Count - 1;

        internal object? FirstChangelogBlockForTests => this.PageItems.Items.Count > 1 ? this.PageItems.Items[1] : null;

        internal bool ChangelogTitleShownForTests => this.ChangelogTitle.IsVisible;

        internal bool NoChangelogShownForTests => this.NoChangelogText.IsVisible;

        // The classes of each card - its state among them - in order.
        internal IReadOnlyList<string> PanelStatesForTests =>
            this.PanelsHost.Children.OfType<Border>()
                .Select(card => card.Classes.FirstOrDefault(name => name != "ServerPanel") ?? string.Empty)
                .ToList();

        private static IEnumerable<string> LinesOf(Border card)
        {
            if (card.Child is not Grid { Children: [_, StackPanel words] })
                yield break;

            TextBlock[] heading = words.Children.OfType<TextBlock>().ToArray();
            yield return BoldRuns.TextOf(heading[0]) + ": " + BoldRuns.TextOf(heading[1]);

            if (words.Children.OfType<Grid>().FirstOrDefault() is not Grid facts)
                yield break;

            TextBlock[] cells = facts.Children.OfType<TextBlock>().ToArray();

            for (int index = 0; index + 1 < cells.Length; index += 2)
                yield return BoldRuns.TextOf(cells[index]) + ": " + BoldRuns.TextOf(cells[index + 1]);
        }
    }
}
