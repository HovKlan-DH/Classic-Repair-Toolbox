using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The modal that files a just-captured oscilloscope image into a worklog entry.
//
// Its whole design claim is that ONE dialog does the job - it names the workbook, ranks the entries
// and takes the comment - so the flow from the capture banner is button, dialog, done. These tests
// pin the parts of that claim which are UI rather than ranking (WorklogAttachTargetsTests covers
// the ordering itself): that the preselected row is the ranked-first one, that "Create new worklog"
// is always reachable, that the button says which of the two things the click does, and that the
// two-band ordering is stated by non-selectable GROUP HEADERS inside the dropdown.
//
// The headers matter because the ordering was reported as illogical twice. The list has two bands -
// component matches, then everything else - and with nothing naming them a reader sees "#2, #1, #3"
// and concludes it is simply unsorted. A header must never be selectable, or the dialog can be
// submitted with a heading as its target.
//
// Built but never shown: ShowDialog blocks, and everything asserted here is set by Initialize.
[Collection("HeadlessUi")]
public sealed class WorklogAttachCaptureWindowTests
{
    private static WorklogEntryRecord Entry(int id, string title, params string[] componentLabels) =>
        new() { Id = id, Title = title, State = "Open", ComponentLabels = componentLabels.ToList() };

    private static WorklogAttachCaptureWindow BuildWindow(
        IReadOnlyList<WorklogEntryRecord> entries,
        string componentLabel)
    {
        var window = new WorklogAttachCaptureWindow();
        window.Initialize(
            "capture.png",
            new WorkbookRecord { Id = 3, Title = "Dave's C64" },
            entries,
            componentLabel);
        return window;
    }

    private static ComboBox EntryCombo(WorklogAttachCaptureWindow window) =>
        window.GetControl<ComboBox>("EntryComboBox");

    // The dialog preselects the ranked-first row, which is what makes the common case a single
    // Attach click. Asserted with a component match that is NOT first by id, so a list in plain id
    // order - which is what everything below the matched band uses - would fail this.
    [Fact]
    public void The_preselected_worklog_is_the_one_scoping_the_measured_component()
    {
        UiTest.Run(() =>
        {
            // The match is #2, deliberately NOT the lowest id: with the matched entry first by id
            // anyway, a dialog that did no ranking at all would pass this test.
            var window = BuildWindow(
                new[] { Entry(1, "Keyboard fault", "U1"), Entry(2, "Video fault", "U8") },
                "U8");

            // By VALUE, not index: a group header occupies index 0 whenever the list is grouped.
            Assert.Contains("#2", EntryCombo(window).SelectedItem!.ToString());
        });
    }

    // Probing before anything has been written down is how diagnosis starts, so the "create one
    // now" route is offered whether or not the workbook already has entries - and it is always the
    // LAST row, so it never displaces a real entry from the preselected slot.
    [Fact]
    public void A_new_worklog_can_always_be_created_and_never_takes_the_preselected_slot()
    {
        UiTest.Run(() =>
        {
            var withEntries = BuildWindow(new[] { Entry(1, "Video fault", "U8") }, "U8");
            var items = EntryCombo(withEntries).ItemsSource!.Cast<object>().ToList();

            Assert.Contains("Create new worklog", items[^1].ToString());

            // With no entries at all the create row is the only one, and preselecting it is right.
            var empty = BuildWindow(Array.Empty<WorklogEntryRecord>(), "U8");
            var emptyItems = EntryCombo(empty).ItemsSource!.Cast<object>().ToList();

            Assert.Single(emptyItems);
            Assert.Contains("Create new worklog", EntryCombo(empty).SelectedItem!.ToString());
        });
    }

    // The button has to say which of two different things the click does: attaching there and then,
    // or opening the full editor on a draft. A button still reading "Attach" for the create row
    // would promise the dialog was the end of it.
    [Fact]
    public void The_button_says_whether_it_attaches_or_creates()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow(new[] { Entry(1, "Video fault", "U8") }, "U8");
            var button = window.GetControl<Button>("AttachButton");

            Assert.Equal("Attach to existing worklog", button.Content);

            // Selecting the create row is what the user does when none of the listed worklogs fit.
            // Found by value rather than by index, since group headers shift every position.
            var combo = EntryCombo(window);
            combo.SelectedItem = combo.ItemsSource!.Cast<object>()
                .Last(item => item is not ComboBoxItem);

            Assert.Equal("Create worklog", button.Content);
        });
    }

    // The workbook is named because this dialog opens from the component popup, which can be sitting
    // over a schematic while the user has been looking at the oscilloscope - the worklog bar that
    // normally makes the active workbook obvious is not in view.
    [Fact]
    public void The_target_workbook_is_named()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow(new[] { Entry(1, "Video fault", "U8") }, "U8");

            Assert.Equal("#3 - Dave's C64", window.GetControl<TextBlock>("WorkbookText").Text);
        });
    }

    // The two bands are NAMED in the list itself, which is the whole fix for "the order is not
    // logical": the grouping is stated where it is being applied, not in a caption under a box that
    // is closed at the time.
    [Fact]
    public void The_two_bands_are_named_by_headers_inside_the_dropdown()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow(
                new[] { Entry(2, "Video fault", "U8"), Entry(1, "Keyboard fault", "U1") },
                "U8");

            var labels = EntryCombo(window).ItemsSource!.Cast<object>()
                .Select(item => item is ComboBoxItem header ? VisibleTextOf(header) : item.ToString())
                .ToList();

            // The matched header names the component, so it explains itself without the caption.
            Assert.Equal("Worklogs with [U8] in scope", labels[0]);
            Assert.Equal("#2 - Video fault", labels[1]);
            Assert.Equal("All other worklogs", labels[2]);
            Assert.Equal("#1 - Keyboard fault", labels[3]);
        });
    }

    // The matched heading picks the component out in bold inside brackets, so the thing the
    // grouping is keyed on is visible at a glance rather than buried in a sentence. A TextBlock
    // cannot mix weights within one Text, so the heading is built from Runs - which means its Text
    // is null and only its Inlines carry the words, the same shape the Workbooks summary strip's
    // bold numbers take.
    [Fact]
    public void The_matched_heading_picks_the_component_out_in_bold()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow(
                new[] { Entry(2, "Video fault", "U8"), Entry(1, "Keyboard fault", "U1") },
                "U8");

            var header = EntryCombo(window).ItemsSource!.Cast<object>().OfType<ComboBoxItem>().First();
            var block = Assert.IsType<TextBlock>(header.Content);

            // The whole sentence still reads correctly once the runs are joined back together.
            Assert.Equal("Worklogs with [U8] in scope", VisibleTextOf(header));

            // ONLY the component is bold - a heading in which everything is bold emphasises nothing.
            var boldRuns = block.Inlines!.OfType<Run>().Where(run => run.FontWeight == FontWeight.Bold).ToList();
            var bold = Assert.Single(boldRuns);
            Assert.Equal("U8", bold.Text);
        });
    }

    // The plain "All other worklogs" heading has nothing to emphasise, so it stays a bare string
    // rather than a TextBlock built to hold no bold run at all.
    [Fact]
    public void The_other_heading_stays_a_plain_string()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow(
                new[] { Entry(2, "Video fault", "U8"), Entry(1, "Keyboard fault", "U1") },
                "U8");

            var others = EntryCombo(window).ItemsSource!.Cast<object>().OfType<ComboBoxItem>().Last();

            Assert.Equal("All other worklogs", Assert.IsType<string>(others.Content));
        });
    }

    // A header's words can live in Content directly (a plain string) or in a TextBlock's Inlines
    // (the bolded one), and a block carrying Inlines has Text == null - so a reader looking only at
    // Text sees that heading as blank.
    private static string VisibleTextOf(ComboBoxItem header) => header.Content switch
    {
        string text => text,
        TextBlock block when block.Inlines != null && block.Inlines.Count > 0 =>
            string.Concat(block.Inlines.OfType<Run>().Select(run => run.Text)),
        TextBlock block => block.Text ?? string.Empty,
        _ => string.Empty
    };

    // A header must be unselectable, or the dialog can be submitted with a heading as its target.
    // Asserted on IsEnabled: a disabled ComboBoxItem is skipped by both mouse and keyboard, which a
    // SelectionChanged guard would not achieve without bouncing the selection first.
    [Fact]
    public void Group_headers_can_never_be_selected()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow(
                new[] { Entry(2, "Video fault", "U8"), Entry(1, "Keyboard fault", "U1") },
                "U8");

            var headers = EntryCombo(window).ItemsSource!.Cast<object>().OfType<ComboBoxItem>().ToList();

            Assert.Equal(2, headers.Count);
            Assert.All(headers, header => Assert.False(header.IsEnabled));

            // And the preselected row is a real worklog, not the leading header at index 0.
            Assert.Contains("#2", EntryCombo(window).SelectedItem!.ToString());
        });
    }

    // A header must not be mistakable for a selectable row - it was, at 0.75 opacity, because a
    // ComboBox's own rows carry no styling to contrast against. It is also aligned to the OUTER
    // edge, with its members indented under it, so the grouping reads as a hierarchy.
    //
    // The indent assertion is against the FLUENT THEME'S default item padding (11px), not against
    // the header's own 6px: an un-indented row already sits further right than the header, so the
    // obvious comparison passes with the container hook removed entirely. Verified by removing it.
    [Fact]
    public void Group_headers_are_faint_and_sit_outside_the_rows_they_head()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow(
                new[] { Entry(2, "Video fault", "U8"), Entry(1, "Keyboard fault", "U1") },
                "U8");

            var header = EntryCombo(window).ItemsSource!.Cast<object>().OfType<ComboBoxItem>().First();

            Assert.True(header.Opacity < 0.5);

            // The indent has to clear the Fluent theme's OWN 11px item padding, not merely the
            // header's 6px - a row left at the theme default already sits further right than the
            // header does, so comparing the two proves nothing about the hook having run.
            // Called once: it shows and closes the window, so a second call reads a closed one.
            double indentLeft = EntryIndentLeft(window);

            Assert.True(indentLeft > 11, "worklog rows are not indented under their heading");
            Assert.True(indentLeft > header.Padding.Left);
        });
    }

    // Reads the left padding the container hook applies to a realised worklog row. Attaching the
    // window to a real Window and showing it is what makes Avalonia realise those containers at
    // all - an ItemsSource alone generates none.
    private static double EntryIndentLeft(WorklogAttachCaptureWindow window)
    {
        window.Show();
        try
        {
            Dispatcher.UIThread.RunJobs();

            var combo = EntryCombo(window);
            combo.IsDropDownOpen = true;
            Dispatcher.UIThread.RunJobs();

            // Matched on the EntryChoice specifically, not on "not a string": the matched heading
            // is a TextBlock now that it carries a bolded component, so a not-a-string test picks
            // the HEADER up as if it were a worklog row and reads its outer padding instead.
            var row = combo.GetRealizedContainers()
                .OfType<ComboBoxItem>()
                .First(container => container.Content is not null
                    && container.Content.GetType().Name == "EntryChoice");

            return row.Padding.Left;
        }
        finally
        {
            window.Close();
        }
    }

    // With nothing matching there is only one band, so a lone "All other worklogs" heading would
    // name a distinction that is not being drawn - and imply a matched group exists somewhere.
    [Fact]
    public void No_headers_are_shown_when_nothing_matches_the_component()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow(new[] { Entry(1, "Video fault", "U1") }, "U8");

            Assert.Empty(EntryCombo(window).ItemsSource!.Cast<object>().OfType<ComboBoxItem>());
        });
    }
}
