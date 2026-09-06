using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The Workbooks tab's "Find a previous repair" box: what it filters, and the highlighting it puts
// on the runs that matched.
//
// The query LANGUAGE itself is pinned down in WorklogSearchQueryTests (pure, no UI) and which
// fields are searched in WorklogSearchIndexTests - these tests are only about the tab actually
// applying them: cards disappearing, entry lists narrowing, and matched text rendering as marked
// Runs rather than plain Text.
//
// COLLECTION NOTE: "HeadlessUi" for the same reason WorkbooksListTests is - it constructs a
// control and so needs the shared dispatcher thread, while also driving WorklogManager's and
// UserSettings' statics. See that file's own header for why that is safe here.
[Collection("HeadlessUi")]
public sealed class WorkbooksSearchTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private readonly string thisBoardKey = "Commodore 64|250469 " + Guid.NewGuid().ToString("N");

    public void Dispose()
    {
        this.LoadWorklog();
        this.thisWorkspace.Dispose();
    }

    private string LoadWorklog()
    {
        string root = this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N"));
        WorklogManager.LoadFrom(root);
        return root;
    }

    private static TabWorkbooks BuildTab(string boardKey)
    {
        var tab = new TabWorkbooks { BoardKeyOverrideForTests = boardKey };
        tab.ActivateWorkbookOverrideForTests = (key, workbookId) =>
        {
            UserSettings.SetActiveWorkbookId(key, workbookId);
            tab.RefreshWorkbooks();
        };
        tab.RefreshWorkbooks();
        return tab;
    }

    // Types into the real search box, then refreshes.
    //
    // The explicit RefreshWorkbooks is NOT the test taking a shortcut past the handler: these tabs
    // are never attached to a visual tree (no test constructs Main), and a detached TextBox does
    // not raise TextChanged, so the tab's own OnFindRepairTextChanged cannot fire here. Setting the
    // real control's Text and then refreshing exercises the same path the running app takes, since
    // RefreshWorkbooks deliberately re-reads the box rather than trusting a cached copy - which is
    // exactly what makes the filter survive every OTHER refresh trigger too (a board switch, an
    // entry save), not just a keystroke.
    private static void Search(TabWorkbooks tab, string query)
    {
        tab.GetControl<TextBox>("FindRepairTextBox").Text = query;
        tab.RefreshWorkbooks();
    }

    private static StackPanel ListPanel(TabWorkbooks tab) =>
        tab.GetControl<StackPanel>("WorkbookListPanel");

    private static StackPanel EntriesPanel(TabWorkbooks tab) =>
        tab.GetControl<StackPanel>("SelectedSchematicEntriesPanel");

    private static string CountText(TabWorkbooks tab) =>
        tab.GetControl<TextBlock>("WorkbookCountText").Text ?? string.Empty;

    // A TextBlock's visible text whether it was set as plain Text or built from highlighted Runs -
    // once a search marks part of a block, Text is null and the content lives in Inlines, so a
    // reader that only looked at Text would see every highlighted card as blank.
    private static string VisibleText(TextBlock block)
    {
        if (block.Text != null)
            return block.Text;

        return block.Inlines == null
            ? string.Empty
            : string.Concat(block.Inlines.OfType<Run>().Select(r => r.Text));
    }

    // The top-line's title, read the Inlines-aware way - a search matching the title highlights it,
    // which leaves TextBlock.Text null.
    private static string HeaderText(TabWorkbooks tab) =>
        VisibleText(tab.GetControl<TextBlock>("WorkbookHeaderTitleText"));

    private static List<string> TextsIn(Control root) =>
        root.GetSelfAndVisualDescendants()
            .OfType<TextBlock>()
            .Select(VisibleText)
            .ToList();

    // The runs a search marked, across a whole subtree - what the user sees highlighted.
    private static List<string> HighlightedRuns(Control root) =>
        root.GetSelfAndVisualDescendants()
            .OfType<TextBlock>()
            .Where(b => b.Inlines != null)
            .SelectMany(b => b.Inlines!.OfType<Run>())
            .Where(r => r.Background != null)
            .Select(r => r.Text ?? string.Empty)
            .ToList();

    // -------------------------------------------------------------- Filtering the workbook list

    [Fact]
    public void An_empty_search_box_shows_every_workbook()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen", "");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);

            // An empty box is not a filter - clearing the field must bring everything back rather
            // than leaving the tab blank.
            Assert.Equal(2, ListPanel(tab).Children.Count);

            Search(tab, "black");
            Assert.Single(ListPanel(tab).Children);

            Search(tab, "");
            Assert.Equal(2, ListPanel(tab).Children.Count);
        });
    }

    [Fact]
    public void A_workbook_is_found_by_its_own_title()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen", "");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            Search(tab, "pla");

            var card = Assert.Single(ListPanel(tab).Children);
            Assert.Contains("Dead PLA", TextsIn((Control)card));

            // The heading counts what is SHOWN, so it reads as the result count.
            Assert.Equal("1 workbook", CountText(tab));
        });
    }

    [Fact]
    public void A_workbook_is_found_by_its_note()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen", "collected on tuesday");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            Search(tab, "tuesday");

            var card = Assert.Single(ListPanel(tab).Children);
            Assert.Contains("Black screen", TextsIn((Control)card));
        });
    }

    // The point of searching a worklog: find the JOB by something recorded inside it, not just by
    // the title somebody gave the workbook months ago.
    [Fact]
    public void A_workbook_is_found_through_text_in_one_of_its_entries()
    {
        this.LoadWorklog();
        var withEntry = WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen", "");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA", "");

        WorklogManager.AddEntry(
            withEntry!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "Replaced U18", "the 6510 ran hot", "Issue", "Open", Array.Empty<string>());

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            Search(tab, "6510");

            var card = Assert.Single(ListPanel(tab).Children);
            Assert.Contains("Black screen", TextsIn((Control)card));
        });
    }

    [Fact]
    public void A_search_matching_nothing_empties_the_list_and_says_so()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            Search(tab, "nothingmatchesthis");

            Assert.Empty(ListPanel(tab).Children);

            // "No results" must not read as "no workbooks recorded" - that looks like data loss.
            var empty = tab.GetControl<TextBlock>("NoWorkbooksText");
            Assert.True(empty.IsVisible);
            Assert.Contains("match your search", empty.Text ?? string.Empty);
        });
    }

    [Fact]
    public void Every_term_must_match_because_space_means_and()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen on a C64", "");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black keyboard", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);

            Search(tab, "black");
            Assert.Equal(2, ListPanel(tab).Children.Count);

            Search(tab, "black screen");
            Assert.Single(ListPanel(tab).Children);
        });
    }

    [Fact]
    public void A_minus_term_excludes_matching_workbooks()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen on a C64", "");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black keyboard", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            Search(tab, "black -keyboard");

            var card = Assert.Single(ListPanel(tab).Children);
            Assert.Contains("Black screen on a C64", TextsIn((Control)card));
        });
    }

    [Fact]
    public void A_quoted_phrase_matches_the_whole_run_including_its_spaces()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen", "");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Screen is black", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);

            // Unquoted, both match - the two words appear in each, in some order.
            Search(tab, "black screen");
            Assert.Equal(2, ListPanel(tab).Children.Count);

            // Quoted, only the one carrying that exact contiguous run does.
            Search(tab, "\"black screen\"");
            var card = Assert.Single(ListPanel(tab).Children);
            Assert.Contains("Black screen", TextsIn((Control)card));
        });
    }

    [Fact]
    public void Searching_is_not_case_sensitive()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black SCREEN", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            Search(tab, "screen");

            Assert.Single(ListPanel(tab).Children);
        });
    }

    // -------------------------------------------------------------- Filtering the entry list

    [Fact]
    public void The_entry_list_narrows_to_the_entries_that_matched()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen", "");

        WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "Replaced U18", "the 6510 ran hot", "Issue", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(
            workbook.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "Recapped the PSU", "", "Issue", "Open", Array.Empty<string>());

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            tab.CurrentBoardDataOverrideForTests = this.BoardDataWithSchematic("Sheet 1");
            tab.RefreshBoardPreviewsForCurrentSelection();

            // Both entries are on the schematic to begin with.
            Assert.Equal(2, EntriesPanel(tab).Children.Count);

            Search(tab, "6510");

            var card = Assert.Single(EntriesPanel(tab).Children);
            Assert.Contains("Replaced U18", TextsIn((Control)card));
        });
    }

    // A workbook found by its OWN text keeps all of its entries visible - the user found the job
    // they were looking for, and blanking its contents would make the result look empty.
    [Fact]
    public void A_workbook_matched_by_its_own_title_still_shows_all_its_entries()
    {
        this.LoadWorklog();
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Jensens C64", "");

        WorklogManager.AddEntry(
            workbook!.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "Replaced U18", "", "Issue", "Open", Array.Empty<string>());
        WorklogManager.AddEntry(
            workbook.Id, "Sheet 1", new Avalonia.Rect(0, 0, 10, 10),
            "Recapped the PSU", "", "Issue", "Open", Array.Empty<string>());

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            tab.CurrentBoardDataOverrideForTests = this.BoardDataWithSchematic("Sheet 1");
            tab.RefreshBoardPreviewsForCurrentSelection();

            Search(tab, "jensens");

            Assert.Single(ListPanel(tab).Children);
            Assert.Equal(2, EntriesPanel(tab).Children.Count);
        });
    }

    // -------------------------------------------------------------- The filtered-out active workbook

    // The top-line, its Edit/Delete buttons and the whole right-hand side belong to whatever the tab
    // shows. Leaving that on a workbook the search has hidden meant "Delete workbook" destroyed a
    // workbook that was not in the list - so the shown workbook moves to one that survived, WITHOUT
    // re-activating anything.
    [Fact]
    public void Filtering_out_the_active_workbook_moves_the_top_line_to_a_shown_one()
    {
        this.LoadWorklog();
        var active = WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen", "");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);

            // Make the FIRST workbook the active one, then search for something only the other has.
            tab.SelectWorkbookForTests(active!.Id);
            Assert.Contains("Black screen", HeaderText(tab));

            Search(tab, "PLA");

            var card = (Control)Assert.Single(ListPanel(tab).Children);
            Assert.Contains("Dead PLA", TextsIn(card));

            // The top-line follows the list rather than naming the hidden workbook.
            Assert.Contains("Dead PLA", HeaderText(tab));
            Assert.DoesNotContain("Black screen", HeaderText(tab));
        });
    }

    // Moving the SHOWN workbook must not move the ACTIVE one: the worklog bar, "Show worklogs" and
    // "Add worklog" all act on the saved active id, and typing in a search box must never redirect
    // where the next drawn entry is written.
    [Fact]
    public void Filtering_does_not_change_which_workbook_is_activated_app_wide()
    {
        this.LoadWorklog();
        var active = WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen", "");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            tab.SelectWorkbookForTests(active!.Id);

            Search(tab, "PLA");

            Assert.Equal(active.Id, UserSettings.GetActiveWorkbookId(this.thisBoardKey));
        });
    }

    // The filter is cleared by the user and by nothing else - the button at the end of the box.
    // It empties the box AND the filter, and rebuilds the list unfiltered on its own (it used to be
    // ClearSearchForBoardChange, whose callers refreshed separately).
    [Fact]
    public void Clearing_the_search_empties_the_box_and_the_filter()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen", "");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);

            Search(tab, "PLA");
            Assert.Single(ListPanel(tab).Children);

            tab.ClearSearch();

            Assert.Equal(string.Empty, tab.GetControl<TextBox>("FindRepairTextBox").Text);
            Assert.Equal(2, ListPanel(tab).Children.Count);
        });
    }

    // THE REPORTED BUG. A board change must NOT clear the query.
    //
    // In "Show all workbooks" scope a search routinely matches a workbook on another board, and
    // clicking that result is how the user follows the search to it - which switches the board. The
    // old behaviour cleared on that switch, so the filter vanished exactly when it had found what
    // was asked for and had to be retyped.
    //
    // A board change reaches this tab as a plain RefreshWorkbooks (Main.RefreshWorklogBar), which is
    // what this drives - the clearing call that used to sit in front of it is gone.
    [Fact]
    public void A_board_change_keeps_the_search_and_its_filtering()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen", "");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);

            Search(tab, "PLA");
            Assert.Single(ListPanel(tab).Children);

            // What a board change does to this tab.
            tab.RefreshWorkbooks();

            Assert.Equal("PLA", tab.GetControl<TextBox>("FindRepairTextBox").Text);
            Assert.Single(ListPanel(tab).Children);
        });
    }

    // The filter also survives the other refresh triggers - an entry save, a workbook create or
    // delete - all of which land in RefreshWorkbooks the same way. Pinned alongside the board-change
    // case so "nothing clears it implicitly" is stated as a rule rather than one example.
    [Fact]
    public void Creating_a_workbook_keeps_the_search_applied()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Black screen", "");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);

            Search(tab, "PLA");
            Assert.Single(ListPanel(tab).Children);

            WorklogManager.CreateWorkbook(this.thisBoardKey, "Another job entirely", "");
            tab.RefreshWorkbooks();

            // Still filtered: the new workbook does not match, so it is not shown.
            Assert.Equal("PLA", tab.GetControl<TextBox>("FindRepairTextBox").Text);
            Assert.Single(ListPanel(tab).Children);
        });
    }

    // The clear button is the ONLY way to drop the query now, so it has to be visible whenever there
    // is a query - and absent when there is not, where it would be a control that does nothing.
    [Fact]
    public void The_clear_button_appears_only_while_there_is_something_to_clear()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            var clearButton = tab.GetControl<Button>("ClearSearchButton");

            Assert.False(clearButton.IsVisible);

            Search(tab, "PLA");
            Assert.True(clearButton.IsVisible);

            tab.ClearSearch();
            Assert.False(clearButton.IsVisible);
        });
    }

    // The box reserves room on its right while the button is over it, so typed text cannot run
    // underneath the button - and gives that room back when the button goes away.
    [Fact]
    public void The_search_box_reserves_room_for_the_clear_button_only_while_it_is_shown()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            var box = tab.GetControl<TextBox>("FindRepairTextBox");

            double unfiltered = box.Padding.Right;

            Search(tab, "PLA");
            double filtered = box.Padding.Right;

            Assert.True(
                filtered > unfiltered,
                $"the box reserved no extra room for the button ({filtered} vs {unfiltered})");

            tab.ClearSearch();

            Assert.Equal(unfiltered, box.Padding.Right, precision: 3);
        });
    }

    // -------------------------------------------------------------- Highlighting

    [Fact]
    public void The_matched_run_is_highlighted_and_the_rest_is_not()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "This is a text", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);

            // The request's own example: searching "this text" marks both words, nothing else.
            Search(tab, "this text");

            var card = (Control)Assert.Single(ListPanel(tab).Children);
            var highlighted = HighlightedRuns(card);

            Assert.Contains("This", highlighted);
            Assert.Contains("text", highlighted);
            Assert.DoesNotContain(" is a ", highlighted);
        });
    }

    [Fact]
    public void Highlighting_preserves_the_original_casing_and_the_full_text()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead CPU here", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            Search(tab, "cpu");

            var card = (Control)Assert.Single(ListPanel(tab).Children);

            // The mark is drawn over the text as written, not over a lowercased copy...
            Assert.Contains("CPU", HighlightedRuns(card));

            // ...and splitting it into runs must not drop or duplicate a character.
            Assert.Contains("Dead CPU here", TextsIn(card));
        });
    }

    [Fact]
    public void Nothing_is_highlighted_when_no_search_is_active()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead CPU here", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);

            var card = (Control)Assert.Single(ListPanel(tab).Children);
            Assert.Empty(HighlightedRuns(card));

            // With no query the block carries plain Text rather than Runs - cheaper to lay out,
            // and there is nothing to mark.
            Assert.Contains("Dead CPU here", TextsIn(card));
        });
    }

    // Clearing the box has to remove the marks as well as restore the rows - a stale highlight left
    // behind would claim a match that is no longer being searched for.
    [Fact]
    public void Clearing_the_search_removes_the_highlighting()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead CPU here", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);

            Search(tab, "cpu");
            Assert.NotEmpty(HighlightedRuns((Control)ListPanel(tab).Children[0]));

            Search(tab, "");
            var card = (Control)Assert.Single(ListPanel(tab).Children);
            Assert.Empty(HighlightedRuns(card));
            Assert.Contains("Dead CPU here", TextsIn(card));
        });
    }

    [Fact]
    public void An_excluded_term_is_never_highlighted()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA here", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);
            Search(tab, "dead -cpu");

            var card = (Control)Assert.Single(ListPanel(tab).Children);
            var highlighted = HighlightedRuns(card);

            Assert.Contains("Dead", highlighted);
            Assert.DoesNotContain("cpu", highlighted, StringComparer.OrdinalIgnoreCase);
        });
    }

    // A minimal but genuinely valid 1x1 PNG - the board pane does a plain File-based
    // `new Bitmap(path)`, so the fixture only has to be a file libpng can decode. Same bytes and
    // same reasoning as WorkbooksBoardPreviewTests' own copy.
    private static readonly byte[] OnePixelPng =
    {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
        0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41,
        0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
        0x42, 0x60, 0x82
    };

    // The tab's own minimal BoardData - just enough for the board pane to have a schematic to group
    // entries under, without standing up a real board Excel file.
    private BoardData BoardDataWithSchematic(string schematicName)
    {
        string imagePath = this.thisWorkspace.Path_(schematicName + ".png");
        File.WriteAllBytes(imagePath, OnePixelPng);

        return new BoardData
        {
            Schematics = new List<BoardSchematicEntry>
            {
                new() { SchematicName = schematicName, SchematicImageFile = imagePath },
            },
        };
    }

    // ###########################################################################################
    // CLEARING THE SEARCH MUST NOT LEAVE A SECOND REBUILD ARMED BEHIND IT.
    //
    // ClearSearch assigns the box's Text, and that assignment raises TextChanged - whose handler
    // RESTARTS the debounce timer. Stopping the timer BEFORE the assignment therefore simply re-arms
    // it, and the immediate rebuild ClearSearch performs was followed ~200ms later by a second,
    // identical full pass: every workbook's worklogs re-read from disk, the whole board pane
    // re-laid-out, and a visible flicker on a workbook with several previews. It also re-entered
    // RefreshBoardPreviews, which the tab's own comments warn can re-enter while a badge's editor is
    // open.
    //
    // The fix is ordering - stop AFTER the assignment. Fails against the version that stopped first.
    // ###########################################################################################
    //
    // DRIVEN THROUGH RaiseSearchTextChangedForTests rather than by assigning the box's Text: a
    // headless TextBox does not raise TextChanged even with the tab attached to a shown window
    // (which is why every other test in this file sets Text and then calls RefreshWorkbooks by
    // hand). So the handler is invoked directly, which is exactly what the real assignment does -
    // and it is invoked AGAIN after ClearSearch's own assignment, standing in for the raise the
    // running app performs there. That second call is the whole scenario: it is the one the old
    // ordering left un-stopped.
    [Fact]
    public void Clearing_the_search_leaves_no_debounced_rebuild_pending()
    {
        this.LoadWorklog();
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead PLA", "");

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey);

            // Typing arms the timer - which is what makes the assertions below meaningful rather
            // than merely observing a timer that was never started.
            tab.GetControl<TextBox>("FindRepairTextBox").Text = "PLA";
            tab.RaiseSearchTextChangedForTests();
            Assert.True(tab.IsSearchRebuildPendingForTests);

            // ClearSearch as the app runs it: it empties the box, which raises TextChanged (re-arming
            // the timer), and only then stops it.
            tab.SimulateSearchTextChangedForTests = tab.RaiseSearchTextChangedForTests;
            tab.ClearSearch();

            Assert.False(tab.IsSearchRebuildPendingForTests);
        });
    }
}
