using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE HISTORY VIEW (owner request, 2026-10-04: "it should show the full history of what has
    // happened with this board, in an 'easy to overview' way"). SystemHistoryDisplay decides and words
    // everything; this part only draws it:
    //
    //   - a heading counting what is there;
    //   - per month, newest first, a faint month heading;
    //   - each item in a row of two columns - its DAY on the left, so the left edge reads as a
    //     timeline, and on the right either a submission's CARD (a thin border around everything
    //     that happened to it and what it changed) or another event's two lines.
    //
    // Drawn with the system's detail, like the Contributor and Statistics views - nothing to read
    // from the server on choosing it.
    // ###########################################################################################
    public partial class SystemView
    {
        private void ShowHistory(SystemDetailAnswer? detail)
        {
            if (this.FindControl<StackPanel>("HistoryItemsSection") is not StackPanel section)
                return;

            section.Children.Clear();

            if (detail is null)
                return;

            IReadOnlyList<HistoryMonth> months = SystemHistoryDisplay.Build(detail);

            int cards = months.Sum(month => month.Items.OfType<HistorySubmissionCard>().Count());
            int events = months.Sum(month => month.Items.OfType<HistoryEventLine>().Count());

            section.Children.Add(SystemView.Heading(SystemHistoryDisplay.Heading(cards, events), cards + events == 0));

            foreach (HistoryMonth month in months)
            {
                section.Children.Add(new TextBlock
                {
                    Text = month.Heading,
                    FontSize = 13,
                    FontWeight = FontWeight.SemiBold,
                    Opacity = 0.75,
                    Classes = { "HistoryMonth" },
                    Margin = new Thickness(0, 10, 0, 0)
                });

                foreach (HistoryItem item in month.Items)
                    section.Children.Add(SystemView.HistoryRow(item));
            }
        }

        // The day on the left, the item on the right.
        private static Grid HistoryRow(HistoryItem item)
        {
            var row = new Grid { ColumnDefinitions = new ColumnDefinitions("52,*") };

            var day = new TextBlock
            {
                Text = SystemHistoryDisplay.DayOf(item.AtUtc),
                FontSize = 12,
                Opacity = 0.7,
                Margin = new Thickness(0, item is HistorySubmissionCard ? 7 : 1, 8, 0)
            };

            Control content = item switch
            {
                HistorySubmissionCard card => SystemView.HistoryCard(card),
                HistoryEventLine line => SystemView.TwoLines(line.What, line.Footer, line.Note),
                _ => new TextBlock()
            };

            Grid.SetColumn(content, 1);
            row.Children.Add(day);
            row.Children.Add(content);

            return row;
        }

        // ###########################################################################################
        // One submission's card: what the contributor wrote (bold), where it stands now, what
        // happened to it (grey, oldest first), what the contributor was told (italic), the red mark of
        // a discarded draft, and what it changed - each sheet a heading with its kinds of change set
        // in under it, the counts in bold.
        // ###########################################################################################
        private static Border HistoryCard(HistorySubmissionCard card)
        {
            var panel = new StackPanel { Spacing = 2 };

            panel.Children.Add(new TextBlock { Text = card.Title, FontWeight = FontWeight.SemiBold, TextWrapping = TextWrapping.Wrap });
            panel.Children.Add(new TextBlock { Text = card.State, FontSize = 12, TextWrapping = TextWrapping.Wrap });

            foreach (string step in card.Steps)
                panel.Children.Add(new TextBlock { Text = step, FontSize = 11, Opacity = 0.7, TextWrapping = TextWrapping.Wrap });

            if (card.Told is string told)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = told,
                    FontSize = 11,
                    FontStyle = FontStyle.Italic,
                    Opacity = 0.85,
                    TextWrapping = TextWrapping.Wrap
                });
            }

            // Its contributor discarded the draft since sending it (2026-09-28) - in the failure
            // colour, as on the queue row.
            if (card.DraftDiscarded is string discarded)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = discarded,
                    FontSize = 11,
                    FontWeight = FontWeight.SemiBold,
                    Foreground = Brushes.IndianRed,
                    TextWrapping = TextWrapping.Wrap
                });
            }

            if (card.Changes.Count > 0)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = SystemHistoryDisplay.ChangesHeading,
                    FontSize = 12,
                    Margin = new Thickness(0, 6, 0, 0),
                    TextWrapping = TextWrapping.Wrap
                });

                foreach (HistoryChangeLine line in card.Changes)
                {
                    var block = new TextBlock
                    {
                        FontSize = 12,
                        TextWrapping = TextWrapping.Wrap,
                        FontWeight = line.IsHeading ? FontWeight.SemiBold : FontWeight.Normal,
                        Margin = new Thickness(line.IsHeading ? 10 : 24, line.IsHeading ? 2 : 0, 0, 0)
                    };

                    TabMaintainer.ShowCounts(block, line.Runs);
                    panel.Children.Add(block);
                }
            }
            else if (card.NoSummary is string none)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = none,
                    FontSize = 11,
                    FontStyle = FontStyle.Italic,
                    Opacity = 0.7,
                    Margin = new Thickness(0, 4, 0, 0),
                    TextWrapping = TextWrapping.Wrap
                });
            }

            var border = new Border
            {
                Child = panel,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(10, 6),
                Classes = { "HistoryCard" }
            };

            border.Bind(Border.BorderBrushProperty, border.GetResourceObservable("Table_BorderRowLine"));

            return border;
        }
    }
}
