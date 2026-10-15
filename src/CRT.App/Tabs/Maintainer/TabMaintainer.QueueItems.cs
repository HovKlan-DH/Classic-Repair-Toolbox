using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
    //
    // *** THE LISTS ARE CHANGED IN PLACE, NEVER BUILT AGAIN (code review, 2026-10-13). *** Every list
    // on this tab - the queue, the BETA list, the Boards list - was replaced wholesale on each minute
    // check (`ItemsSource = new list`). A click whose press landed on a row just replaced was lost
    // (the Drafts tab's "I need to click it twice" of 2026-10-10), and putting the selection back
    // scrolled the list to it. Now each list keeps one collection of items: a row still there keeps
    // its item - the same object, its content brought up to date - a new one is put in where it
    // belongs, a gone one is taken out (InPlaceList.Sync), and the selected item is never made again.
    // ShowItemsInPlace below does it for the BETA and Boards lists, ShowQueueItems for the queue,
    // whose board headings are items too.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        // The list's parts by submission id and by board - kept in step with the list's items.
        private readonly Dictionary<long, QueueEntry> thisQueueEntries = [];
        private readonly Dictionary<string, QueueHeading> thisQueueHeadings = new(StringComparer.Ordinal);

        // The queue list's items: each board's heading, then its submissions - changed in place.
        private readonly ObservableCollection<ListBoxItem> thisQueueItems = [];

        // ###########################################################################################
        // Puts `rows` on the queue list, grouped by board, IN PLACE - see the header. An entry is
        // kept by submission id and a heading by board; each kept one shows its row as it is now
        // (the wait in its footer moves on with `now`), and the first heading carries no divider.
        // ###########################################################################################
        private void ShowQueueItems(ListBox list, IReadOnlyList<ReviewQueueRow> rows, DateTimeOffset now)
        {
            if (!ReferenceEquals(list.ItemsSource, this.thisQueueItems))
                list.ItemsSource = this.thisQueueItems;

            // What each item on the list is now, so one still wanted is kept rather than made again.
            var parts = new Dictionary<ListBoxItem, object>();

            foreach (QueueEntry entry in this.thisQueueEntries.Values)
                parts[entry.Item] = entry;

            foreach (QueueHeading heading in this.thisQueueHeadings.Values)
                parts[heading.Item] = heading;

            List<QueueLine> wanted = [];

            foreach (ReviewQueueGroup group in ReviewQueueDisplay.Group(rows))
            {
                wanted.Add(new QueueLine(group, null));

                foreach (ReviewQueueRow row in group.Rows)
                    wanted.Add(new QueueLine(group, row));
            }

            InPlaceList.Sync(
                this.thisQueueItems,
                wanted,
                (item, line) => parts.TryGetValue(item, out object? part) && line.IsShownBy(part),
                line =>
                {
                    object part = line.Row is { } row ? new QueueEntry(row, now) : new QueueHeading(line.Group, isFirst: false);
                    ListBoxItem item = part is QueueEntry made ? made.Item : ((QueueHeading)part).Item;
                    parts[item] = part;
                    return item;
                },
                keep: list.SelectedItem as ListBoxItem);

            this.thisQueueEntries.Clear();
            this.thisQueueHeadings.Clear();

            for (int index = 0; index < wanted.Count; index++)
            {
                QueueLine line = wanted[index];

                switch (parts[this.thisQueueItems[index]])
                {
                    case QueueHeading heading:
                        heading.Show(line.Group, isFirst: index == 0);
                        this.thisQueueHeadings[line.Group.BoardId] = heading;
                        break;

                    case QueueEntry entry when line.Row is { } row:
                        entry.Show(row, now);
                        entry.Item.Tag = row;
                        this.thisQueueEntries[row.Id] = entry;
                        break;
                }
            }
        }

        // One line of the queue list: a board's heading (no Row), or one of its submissions.
        private sealed record QueueLine(ReviewQueueGroup Group, ReviewQueueRow? Row)
        {
            public bool IsShownBy(object part) => part switch
            {
                QueueHeading heading => this.Row is null && string.Equals(heading.BoardId, this.Group.BoardId, StringComparison.Ordinal),
                QueueEntry entry => this.Row is { } row && entry.Row.Id == row.Id,
                _ => false
            };
        }

        // ###########################################################################################
        // The BETA and Boards lists, in place (see the header): one item per row, kept by `key`; a
        // kept item whose row reads differently now (`showsTheSame` false) gets its content made
        // again - the ITEM stays, so a click on it and the scroll position stay too. The selected
        // item is never made again (InPlaceList.Sync's `keep`).
        // ###########################################################################################
        private static void ShowItemsInPlace<T>(
            ListBox list,
            ObservableCollection<ListBoxItem> items,
            IReadOnlyList<T> rows,
            Func<T, string> key,
            Func<T, Control> content,
            Func<ListBoxItem, T, bool> showsTheSame)
            where T : class
        {
            if (!ReferenceEquals(list.ItemsSource, items))
                list.ItemsSource = items;

            InPlaceList.Sync(
                items,
                rows,
                (item, row) => item.Tag is T held && string.Equals(key(held), key(row), StringComparison.Ordinal),
                row => new ListBoxItem { Tag = row, Classes = { "ListEntry" }, Content = content(row) },
                keep: list.SelectedItem as ListBoxItem);

            for (int index = 0; index < rows.Count; index++)
            {
                ListBoxItem item = items[index];

                if (!showsTheSame(item, rows[index]))
                    item.Content = content(rows[index]);

                item.Tag = rows[index];
            }
        }

        // The submission selected in the list, or null.
        private ReviewQueueRow? SelectedQueueRow =>
            (this.FindControl<ListBox>("QueueList")?.SelectedItem as ListBoxItem)?.Tag as ReviewQueueRow;

        // ###########################################################################################
        // Selects a submission's entry - or nothing, for null or a submission not in the list. Only
        // when that changes anything: setting the selection again scrolls the list to it.
        // ###########################################################################################
        private void SelectQueueRow(long? id)
        {
            if (this.FindControl<ListBox>("QueueList") is not ListBox list)
                return;

            ListBoxItem? wanted = id is long value && this.thisQueueEntries.TryGetValue(value, out QueueEntry? entry)
                ? entry.Item
                : null;

            if (!ReferenceEquals(list.SelectedItem, wanted))
                list.SelectedItem = wanted;
        }

        // The queue list's items as they are - for tests, which hold an unchanged row to its object.
        internal IReadOnlyList<ListBoxItem> QueueItemsForTests => this.thisQueueItems;

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
            private readonly Border thisFrame;

            public QueueHeading(ReviewQueueGroup group, bool isFirst)
            {
                this.BoardId = group.BoardId;

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

                this.thisFrame = new Border { Child = line };

                this.Item = new ListBoxItem
                {
                    IsEnabled = false,
                    Classes = { "QueueHeading" },
                    Content = this.thisFrame
                };

                this.Show(group, isFirst);
            }

            public ListBoxItem Item { get; }

            public string BoardId { get; }

            // The heading as its group now reads: the divider above every heading but the first,
            // and "New board" when nothing of it is published.
            public void Show(ReviewQueueGroup group, bool isFirst)
            {
                this.thisFrame.Classes.Set("QueueDivider", !isFirst);
                this.ShowBadge(group.IsNewBoard);
            }

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
            private readonly TextBlock thisComment;
            private readonly TextBlock thisFooter;

            // "Contributor discarded their draft on ..." (owner request, 2026-09-28) - in the failure
            // colour, and never dimmed with the rest: it is the one thing on the row to act on first.
            private readonly TextBlock thisDiscarded;

            public QueueEntry(ReviewQueueRow row, DateTimeOffset now)
            {
                this.thisContent = new StackPanel { Spacing = 1 };

                this.thisComment = new TextBlock
                {
                    FontSize = 13,
                    TextWrapping = TextWrapping.Wrap,

                    // A long description is cut short in the list; it is the submission's own words,
                    // and the rest of it is not what a maintainer picks the next one by.
                    MaxLines = 3,
                    TextTrimming = TextTrimming.CharacterEllipsis
                };
                this.thisContent.Children.Add(this.thisComment);

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

                // Its words too: an entry is kept across a refresh (ShowQueueItems), and an amended
                // submission can say something new.
                this.thisComment.Text = ReviewQueueDisplay.Comment(row);

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
