using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE QUEUE LIST, GROUPED BY BOARD (owner request, 2026-09-26: "let's try with your
    // 'recommendation is to group the queue by board'"):
    //
    //     Commodore / C64 / 250407             [New board]
    //        Corrected the pinout pictures for U8 and added U10.
    //        Waiting 10 hours - replaces a shared file
    //        Added the missing CIA pictures.
    //        Waiting 2 days - with the other approver             (dimmed)
    //
    // Each submission was six lines until then - labelled Manufacturer / Hardware / Board and two
    // badges on every row - which the project owner found "quite hard to overview". The board is
    // said once, in a HEADING, and each submission is its comment and one grey line. Only what is
    // unusual is marked: "New board" on a heading, and a submission that does not wait for this
    // account is dimmed and says it is with the other approver. The words are ReviewQueueDisplay's.
    //
    // *** A HEADING IS A DISABLED ListBoxItem. *** Disabled, it can be neither clicked nor reached
    // with the arrow keys, so the selection is always a submission - the WorklogAttachCaptureWindow
    // group headers' way. Which is also why the list is read by ITEM, never by index: a heading
    // sits before every group, so a position in the list is not a position in the queue.
    //
    // *** THE OPENED SUBMISSION FOLLOWS WHAT IT SAYS ITSELF. *** The queue judges "awaits you" with
    // the shared-files flag stored at create; the detail judges it against the tree as it is now,
    // and is what the Approve button follows. So when the detail arrives its answers are put on the
    // list (UpdateQueueEntry) - it cannot show an entry as waiting for you beside an Approve that is
    // off.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        // The list's parts by submission id and by board - rebuilt with the list.
        private readonly Dictionary<long, QueueEntry> thisQueueEntries = [];
        private readonly Dictionary<string, QueueHeading> thisQueueHeadings = new(StringComparer.Ordinal);

        // The list's items: each board's heading, then its submissions.
        private List<ListBoxItem> BuildQueueItems(IReadOnlyList<ReviewQueueRow> rows, DateTimeOffset now)
        {
            this.thisQueueEntries.Clear();
            this.thisQueueHeadings.Clear();

            var items = new List<ListBoxItem>();

            foreach (ReviewQueueGroup group in ReviewQueueDisplay.Group(rows))
            {
                var heading = new QueueHeading(group, isFirst: items.Count == 0);
                this.thisQueueHeadings[group.BoardId] = heading;
                items.Add(heading.Item);

                foreach (ReviewQueueRow row in group.Rows)
                {
                    var entry = new QueueEntry(row, now);
                    this.thisQueueEntries[row.Id] = entry;
                    items.Add(entry.Item);
                }
            }

            return items;
        }

        // The submission selected in the list, or null.
        private ReviewQueueRow? SelectedQueueRow =>
            (this.FindControl<ListBox>("QueueList")?.SelectedItem as ListBoxItem)?.Tag as ReviewQueueRow;

        // Selects a submission's entry - or nothing, for null or a submission not in the list.
        private void SelectQueueRow(long? id)
        {
            if (this.FindControl<ListBox>("QueueList") is not ListBox list)
                return;

            list.SelectedItem = id is long value && this.thisQueueEntries.TryGetValue(value, out QueueEntry? entry)
                ? entry.Item
                : null;
        }

        // Puts the opened submission's own answers on the list - see the header.
        private void UpdateQueueEntry(ReviewQueueRow row, bool? isNewBoard, bool? awaitsYou)
        {
            if (this.thisQueueHeadings.TryGetValue(row.BoardId ?? string.Empty, out QueueHeading? heading))
                heading.ShowBadge(isNewBoard);

            if (this.thisQueueEntries.TryGetValue(row.Id, out QueueEntry? entry))
                entry.Show(entry.Row with { AwaitsYou = awaitsYou }, DateTimeOffset.UtcNow);

            // The Review button's badge counts the rows as they now stand.
            this.UpdateModeBadges();
        }

        // Every text the list shows, top to bottom - for tests.
        internal IReadOnlyList<string> QueueTextsForTests() =>
            this.FindControl<ListBox>("QueueList")?.ItemsSource is IEnumerable<ListBoxItem> items
                ? items.SelectMany(item => item.Content is Control content ? TabMaintainer.VisibleTexts(content) : []).ToList()
                : [];

        internal ReviewQueueRow? SelectedQueueRowForTests => this.SelectedQueueRow;

        internal bool QueueEntryIsDimmedForTests(long id) =>
            this.thisQueueEntries.TryGetValue(id, out QueueEntry? entry) && entry.IsDimmed;

        // The texts shown, in order - hidden parts (a badge not given) left out.
        private static IEnumerable<string> VisibleTexts(Control control)
        {
            if (!control.IsVisible)
                yield break;

            if (control is TextBlock block)
            {
                yield return TabMaintainer.TextOf(block);
                yield break;
            }

            IEnumerable<Control> children = control switch
            {
                Panel panel => panel.Children,
                Decorator { Child: Control child } => [child],
                _ => []
            };

            foreach (Control child in children)
            {
                foreach (string text in TabMaintainer.VisibleTexts(child))
                    yield return text;
            }
        }

        // ###########################################################################################
        // One board's heading: its name and, for a board with nothing published, "New board".
        // A thin line above it, except above the first, keeps the boards apart.
        // ###########################################################################################
        private sealed class QueueHeading
        {
            private readonly Border thisBadge;

            public QueueHeading(ReviewQueueGroup group, bool isFirst)
            {
                var line = new WrapPanel { Orientation = Orientation.Horizontal, ItemSpacing = 8, LineSpacing = 4 };

                line.Children.Add(new TextBlock
                {
                    Text = ReviewQueueDisplay.BoardHeading(group.BoardId),
                    FontSize = 13,
                    FontWeight = FontWeight.SemiBold,
                    VerticalAlignment = VerticalAlignment.Center
                });

                this.thisBadge = new Border
                {
                    Classes = { "Badge" },
                    Child = new TextBlock { Text = ReviewQueueDisplay.NewBoardBadge }
                };
                line.Children.Add(this.thisBadge);

                var frame = new Border { Child = line };

                if (!isFirst)
                    frame.Classes.Add("QueueDivider");

                this.Item = new ListBoxItem
                {
                    IsEnabled = false,
                    Classes = { "QueueHeading" },
                    Content = frame
                };

                this.ShowBadge(group.IsNewBoard);
            }

            public ListBoxItem Item { get; }

            public void ShowBadge(bool? isNewBoard) =>
                this.thisBadge.IsVisible = ReviewQueueDisplay.BoardBadge(isNewBoard) is not null;
        }

        // ###########################################################################################
        // One submission: what the contributor wrote, and one grey line - the wait and the
        // two-approval notes. Dimmed when it does not wait for this account.
        // ###########################################################################################
        private sealed class QueueEntry
        {
            private readonly StackPanel thisContent;
            private readonly TextBlock thisFooter;

            // "Contributor discarded their draft on ..." (owner request, 2026-09-28) - in the failure
            // colour, and never dimmed with the rest: it is the one thing on the row to act on first.
            private readonly TextBlock thisDiscarded;

            public QueueEntry(ReviewQueueRow row, DateTimeOffset now)
            {
                this.thisContent = new StackPanel { Spacing = 1 };

                this.thisContent.Children.Add(new TextBlock
                {
                    Text = ReviewQueueDisplay.Comment(row),
                    FontSize = 13,
                    TextWrapping = TextWrapping.Wrap,

                    // A long description is cut short in the list; it is the submission's own words,
                    // and the rest of it is not what a maintainer picks the next one by.
                    MaxLines = 3,
                    TextTrimming = TextTrimming.CharacterEllipsis
                });

                this.thisFooter = new TextBlock
                {
                    FontSize = 11,
                    Opacity = 0.7,
                    TextWrapping = TextWrapping.Wrap
                };
                this.thisContent.Children.Add(this.thisFooter);

                this.thisDiscarded = new TextBlock
                {
                    FontSize = 11,
                    FontWeight = FontWeight.SemiBold,
                    Foreground = Brushes.IndianRed,
                    TextWrapping = TextWrapping.Wrap,
                    IsVisible = false
                };
                this.thisContent.Children.Add(this.thisDiscarded);

                this.Row = row;
                this.Item = new ListBoxItem
                {
                    Tag = row,
                    Classes = { "QueueEntry" },
                    Content = this.thisContent
                };

                this.Show(row, now);
            }

            public ListBoxItem Item { get; }

            public ReviewQueueRow Row { get; private set; }

            public bool IsDimmed => this.Row.AwaitsYou == false;

            public void Show(ReviewQueueRow row, DateTimeOffset now)
            {
                this.Row = row;

                string footer = ReviewQueueDisplay.Footer(row, now);
                this.thisFooter.Text = footer;
                this.thisFooter.IsVisible = footer.Length > 0;

                this.thisDiscarded.Text = row.DraftDiscardedUtc is DateTimeOffset discarded ? DraftDiscardWording.Mark(discarded) : string.Empty;
                this.thisDiscarded.IsVisible = row.DraftDiscardedUtc is not null;

                this.thisContent.Opacity = this.IsDimmed ? 0.55 : 1;
            }
        }
    }
}
