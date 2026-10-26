using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "REVIEW CHANGES" (CRT 4.0 step 1's last slice, 2026-10-25 - the plan: "Review changes shows a
    // readable summary, and the Submit dialog shows the same ... Undo changes to IC1 - one component,
    // back to its published values").
    //
    // The draft's summary (DraftChangeSummary), one group per component and one for the board's own
    // parts. Each component offers "Undo changes to [IC1]": the component back to its published
    // values as ONE step of the draft's history (the host's `undoComponent`, DraftEdits), then the
    // summary read again - so the group goes, and the row's Undo can bring it back.
    // ###########################################################################################
    public partial class ReviewChangesWindow : Window
    {
        private Func<string, Task<(DraftChangeSummary Summary, string? Failure)>>? thisUndoComponent;

        public ReviewChangesWindow()
        {
            this.InitializeComponent();

            this.AddHandler(KeyDownEvent, this.OnWindowKeyDown, RoutingStrategies.Tunnel);
        }

        // `undoComponent` puts one component back and answers the summary afterwards, with why it
        // could not when it could not. Null offers no undo.
        public void Initialize(
            string boardDisplayName,
            DraftChangeSummary summary,
            Func<string, Task<(DraftChangeSummary Summary, string? Failure)>>? undoComponent = null)
        {
            ArgumentNullException.ThrowIfNull(summary);

            this.BoardNameText.Text = boardDisplayName;
            this.thisUndoComponent = undoComponent;

            this.ShowSummary(summary);
        }

        // ###########################################################################################
        // One bordered group per component and for the board - its title, its "Undo changes to"
        // button (a component's, and only with somewhere to send it), and a line per change.
        // ###########################################################################################
        private void ShowSummary(DraftChangeSummary summary)
        {
            this.GroupsPanel.Children.Clear();
            this.NothingText.IsVisible = summary.IsEmpty;

            foreach (DraftChangeGroup group in summary.Groups)
            {
                var heading = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };

                heading.Children.Add(new TextBlock
                {
                    Text = group.Title,
                    FontWeight = FontWeight.SemiBold,
                    TextWrapping = TextWrapping.Wrap,
                    VerticalAlignment = VerticalAlignment.Center,
                });

                if (group.IsComponent && this.thisUndoComponent is not null)
                {
                    // Red - it throws the component's changes away (CLAUDE.md, button colours).
                    var undo = new Button
                    {
                        Content = ReviewChangesWindow.UndoLabel(group.Label),
                        Tag = group.Label,
                        Margin = new Avalonia.Thickness(12, 0, 0, 0),
                        VerticalAlignment = VerticalAlignment.Center,
                    };

                    undo.Classes.Add("Danger");
                    ToolTip.SetTip(undo, $"Put [{group.Label}] back as it is published - every row of it. Undo on the draft's row takes this back.");
                    undo.Click += this.OnUndoComponentClick;

                    Grid.SetColumn(undo, 1);
                    heading.Children.Add(undo);
                }

                var lines = new StackPanel { Spacing = 2, Margin = new Avalonia.Thickness(12, 4, 0, 0) };

                foreach (string line in group.Lines)
                {
                    lines.Children.Add(new TextBlock { Text = "- " + line, TextWrapping = TextWrapping.Wrap });
                }

                var content = new StackPanel { Spacing = 0 };
                content.Children.Add(heading);
                content.Children.Add(lines);

                this.GroupsPanel.Children.Add(new Border
                {
                    Padding = new Avalonia.Thickness(10, 8),
                    BorderThickness = new Avalonia.Thickness(1),
                    BorderBrush = ThemeResources.ResolveBrush("Form_Border"),
                    Background = ThemeResources.ResolveBrush("Form_Bg"),
                    Child = content,
                });
            }
        }

        // "Undo changes to [IC1]" - what the button throws away, said on it (the plan's wording).
        internal static string UndoLabel(string label) => $"Undo changes to [{label}]";

        private async void OnUndoComponentClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button { Tag: string label })
                {
                    await this.UndoComponentAsync(label);
                }
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Undoing the changes to a component failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // The component put back, under this window's "please wait", and the summary shown again - or
        // why it could not be done, with the summary as it now is.
        // ###########################################################################################
        internal async Task UndoComponentAsync(string label)
        {
            if (this.thisUndoComponent is not { } undo)
            {
                return;
            }

            (DraftChangeSummary summary, string? failure) =
                await BusyOverlay.RunLocalAsync(this, CrtWaitWording.UndoingComponentChanges, () => undo(label));

            this.StatusText.Text = failure is null ? string.Empty : $"[{label}] could not be put back. {failure}";
            this.StatusText.IsVisible = failure is not null;

            this.ShowSummary(summary);
        }

        // ---- for tests ----------------------------------------------------------------------------

        internal IReadOnlyList<string> GroupTitlesForTests =>
            this.GroupsPanel.Children
                .OfType<Border>()
                .Select(border => ((TextBlock)((Grid)((StackPanel)border.Child!).Children[0]).Children[0]).Text ?? string.Empty)
                .ToList();

        internal IReadOnlyList<Button> UndoButtonsForTests =>
            this.GroupsPanel.Children
                .OfType<Border>()
                .SelectMany(border => ((Grid)((StackPanel)border.Child!).Children[0]).Children.OfType<Button>())
                .ToList();

        internal string StatusTextForTests => this.StatusText.IsVisible ? this.StatusText.Text ?? string.Empty : string.Empty;

        internal bool ShowsNothingForTests => this.NothingText.IsVisible;

        private void OnWindowKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
                e.Handled = true;
            }
        }

        private void OnCloseClick(object? sender, RoutedEventArgs e) => this.Close();
    }
}
