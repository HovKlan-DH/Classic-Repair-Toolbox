using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.LogicalTree;
using CRT;
using Handlers.Theming;
using Handlers.Tours;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// *** WHAT A TOUR CALLS A BUTTON, A PANEL OR A TAB, THE WINDOW STILL CALLS IT SO (owner question, 2026-11-12: "how can we
// do this, so we make sure that we catch that any UI changes are being caught and adapted for the introduction"). ***
// A tour names things in bold ("<b>Mark all</b> selects every component"), and a button renamed left the step saying a
// word no control says any more - nothing failed. Every bold name of every tour's steps and of its first card's foot must
// be the text of a control in CRT's window, or of the two windows a tour speaks of: the worklog's editor ("Note", "Open")
// and "Errors and warnings" ("Copy all"). A key and an example typed in the filter box are no control's, and listed here.
//
// MainToursTests fails on a control a step FRAMES that is gone; .claude/hooks/remind-tours.sh names the tours whose tab
// changed while the tours did not - a new feature or a changed one no test can see.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TourLabelsTests
{
    // Keys, and examples of what to type - no control's text, on purpose.
    private static readonly HashSet<string> NoControls = new(StringComparer.Ordinal)
    {
        "Shift", "Ctrl", "Cmd", "F11",
        "U1", "CIA", "6526",
    };

    [Fact]
    public void Every_bold_name_in_a_tour_is_a_control_of_the_window()
    {
        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var editor = new WorklogEntryEditorWindow();
            var problems = new ProblemsWindow();

            try
            {
                HashSet<string> said = [.. new Control[] { main, editor, problems }.SelectMany(TextsOf)];

                List<string> missing = TourCatalogue.All
                    .SelectMany(tour => tour.Steps.SelectMany(step => step.Body.Parts.Select(part => (tour.TabName, step.Headline, part.Text))))
                    .Append(("(every first card)", "foot", TourWording.TakeAgainFoot))
                    .SelectMany(where => BoldRuns.Marked(where.Item3).Where(run => run.IsBold).Select(run => (where.Item1, where.Item2, Name: run.Text)))
                    .Where(bold => !NoControls.Contains(bold.Name) && !said.Contains(bold.Name))
                    .Select(bold => $"{bold.Item1} / {bold.Item2}: [{bold.Name}]")
                    .Distinct()
                    .ToList();

                Assert.True(missing.Count == 0, "A tour names what no control of CRT's window says - a button renamed? Change the tour with it:\n" + string.Join("\n", missing));
            }
            finally
            {
                editor.Close();
                problems.Close();
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    // The detector reads what a person reads: a button's words, a label, a tab's header, a placeholder.
    [Fact]
    public void The_texts_read_are_a_buttons_a_labels_a_headers_and_a_placeholders()
    {
        UiTest.Run(() =>
        {
            var panel = new StackPanel
            {
                Children =
                {
                    new Button { Content = "Mark all" },
                    new TextBlock { Text = "Labels visible" },
                    new ComboBox { PlaceholderText = "My drafts" },
                },
            };
            var bold = new TextBlock();
            BoldRuns.ShowMarked(bold, "<b>Show</b> worklogs");
            panel.Children.Add(bold);

            var tabs = new TabControl { Items = { new TabItem { Header = "Workbooks" } } };
            var root = new StackPanel { Children = { panel, tabs } };

            HashSet<string> said = [.. TextsOf(root)];

            Assert.Contains("Mark all", said);
            Assert.Contains("Labels visible", said);
            Assert.Contains("My drafts", said);
            Assert.Contains("Show worklogs", said);
            Assert.Contains("Workbooks", said);
            Assert.DoesNotContain("Mark", said);
        });
    }

    // Every text a person reads in `root` and under it, trimmed.
    private static IEnumerable<string> TextsOf(Control root) =>
        new[] { (ILogical)root }.Concat(root.GetLogicalDescendants())
            .SelectMany(TextOf)
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .Select(text => text!.Trim());

    private static IEnumerable<string?> TextOf(ILogical item)
    {
        switch (item)
        {
            case TextBlock block:
                yield return BoldRuns.TextOf(block);
                break;

            case ComboBox box:
                yield return box.PlaceholderText;
                break;

            case TextBox box:
                yield return box.PlaceholderText;
                break;

            case HeaderedContentControl headered:
                yield return headered.Header as string;
                yield return headered.Content as string;
                break;

            case HeaderedSelectingItemsControl selecting:
                yield return selecting.Header as string;
                break;

            case ContentControl content:
                yield return content.Content as string;
                break;
        }
    }
}
