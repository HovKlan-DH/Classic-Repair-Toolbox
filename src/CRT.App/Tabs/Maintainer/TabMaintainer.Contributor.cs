using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE CONTRIBUTOR VIEW (owner request, 2026-09-30: "all information about the contributor, to
    // get an honest opinion if this person can be trusted ... the trustworthiness history of the
    // contributor"): who sent the submission, whether the address was checked, their counts, and
    // every other submission they sent - what it was, where, how it went and what they were told.
    //
    // The facts arrive with the submission detail (the server's ContributorHistory); the words are
    // ReviewContributorHistory's. Nothing is asked of the server here, so switching to the view is
    // instant. A turned-down submission is drawn in the failure colour, so a record with
    // rejections in it is seen at a glance.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        // Null empties the view - nothing chosen, or the detail not read.
        private void ShowContributorHistory(ReviewContributorFacts? facts)
        {
            if (this.FindControl<StackPanel>("ContributorSection") is not StackPanel section)
                return;

            section.Children.Clear();

            if (facts is null)
                return;

            var who = new StackPanel { Spacing = 2 };

            who.Children.Add(new TextBlock
            {
                Text = ReviewContributorHistory.Heading(facts),
                FontSize = 16,
                FontWeight = FontWeight.SemiBold,
                TextWrapping = TextWrapping.Wrap
            });

            if (ReviewContributorHistory.AccountLine(facts) is string account)
            {
                who.Children.Add(new TextBlock
                {
                    Text = account,
                    Opacity = 0.75,
                    TextWrapping = TextWrapping.Wrap,

                    // Unverified is the one to notice.
                    FontWeight = facts.SignedIn == false ? FontWeight.SemiBold : FontWeight.Normal
                });
            }

            var counts = new TextBlock { TextWrapping = TextWrapping.Wrap, Margin = new Avalonia.Thickness(0, 4, 0, 0) };
            TabMaintainer.ShowCounts(counts, ReviewContributorHistory.Counts(facts));
            who.Children.Add(counts);

            section.Children.Add(who);

            IReadOnlyList<ContributorSubmissionEntry> listed = facts.Submissions ?? [];

            var list = new StackPanel { Spacing = 8 };

            // None at all when there is nothing to list - the counts line has said so already.
            if (ReviewContributorHistory.SubmissionsHeading(facts) is string heading)
            {
                list.Children.Add(new TextBlock
                {
                    Text = heading,
                    FontSize = 14,
                    FontWeight = listed.Count == 0 ? FontWeight.Normal : FontWeight.SemiBold,
                    Opacity = listed.Count == 0 ? 0.7 : 1,
                    TextWrapping = TextWrapping.Wrap
                });
            }

            foreach (ContributorSubmissionEntry submission in listed)
                list.Children.Add(TabMaintainer.ContributorSubmissionRow(submission));

            section.Children.Add(list);
        }

        // What they wrote, then "#12 - board - state - sent ...", then what they were told.
        private static StackPanel ContributorSubmissionRow(ContributorSubmissionEntry submission)
        {
            var row = new StackPanel { Spacing = 1 };

            row.Children.Add(new TextBlock { Text = ReviewContributorHistory.Title(submission), TextWrapping = TextWrapping.Wrap });

            var footer = new TextBlock
            {
                Text = ReviewContributorHistory.Footer(submission),
                FontSize = 11,
                TextWrapping = TextWrapping.Wrap
            };

            if (ReviewContributorHistory.IsTurnedDown(submission))
            {
                footer.Foreground = Brushes.IndianRed;
                footer.FontWeight = FontWeight.SemiBold;
            }
            else
            {
                footer.Opacity = 0.7;
            }

            row.Children.Add(footer);

            if (ReviewContributorHistory.Comment(submission) is string comment)
            {
                row.Children.Add(new TextBlock
                {
                    Text = comment,
                    FontSize = 11,
                    FontStyle = FontStyle.Italic,
                    Opacity = 0.85,
                    TextWrapping = TextWrapping.Wrap
                });
            }

            return row;
        }

        // Every text the view shows, top to bottom - for tests.
        internal IReadOnlyList<string> ContributorViewTextsForTests()
        {
            var texts = new List<string>();

            void Walk(Panel panel)
            {
                foreach (Control child in panel.Children)
                {
                    if (child is TextBlock block)
                        texts.Add(TabMaintainer.TextOf(block));
                    else if (child is Panel inner)
                        Walk(inner);
                }
            }

            if (this.FindControl<StackPanel>("ContributorSection") is StackPanel section)
                Walk(section);

            return texts;
        }
    }
}
