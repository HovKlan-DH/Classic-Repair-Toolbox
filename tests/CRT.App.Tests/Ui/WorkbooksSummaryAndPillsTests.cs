using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using CRT;
using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// Three things the Workbooks tab gained together, all reported or asked for as one complaint about
// the tab not reading as one consistent surface:
//
//  1. THE INFORMATIONAL PILLS. Status pills and category chips that are NOT selectable had drifted
//     into four different looks across four surfaces - some 2px-outlined in the status colour, some
//     1px-outlined in grey. They now all come from WorklogInfoPillBuilder: 1px, in the thing's own
//     colour. These tests pin the border WIDTH and the border COLOUR, because those are exactly the
//     two axes that had drifted.
//
//  2. CLICKABLE ENTRY CARDS. A card in the right-hand list opens the same full editor its pill on
//     the board pane opens. Pinned by the shared code path, since the click itself opens a modal a
//     headless test cannot dismiss - see the test's own comment.
//
//  3. THE SUMMARY STRIP. Its headline, its collapsed-by-default state and its persistence.
//
// COLLECTION NOTE: "HeadlessUi", like every other UI test - constructing a control needs the shared
// dispatcher thread. This class ALSO writes UserSettings (the summary's expanded flag), but a class
// can only join one collection and "HeadlessUi" is the mandatory one; every test here points
// UserSettings at its own temp file first, which is what WorkbooksListTests does for the same
// reason. Every assertion runs inside UiTest.Run - reading a control property from the test thread
// throws even for a plain GetControl.
[Collection("HeadlessUi")]
public sealed class WorkbooksSummaryAndPillsTests : IDisposable
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
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));

        // LoadFrom against a path that does not exist RETURNS EARLY without resetting the static
        // _data - it only logs "using defaults" - so a value another test in this collection wrote
        // is still there. The summary's expanded flag is one this class depends on starting false,
        // so it is reset explicitly rather than assumed.
        UserSettings.WorkbooksSummaryExpanded = false;

        // The currency is the other one, for the same reason: the headline prints it beside the
        // cost, so a code another test left behind would change the string these tests assert on.
        UserSettings.WorklogCurrencyCode = WorklogCurrency.DefaultCode;

        return root;
    }

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

    private string WriteSchematicImage(string fileName)
    {
        string path = this.thisWorkspace.Path_(fileName);
        File.WriteAllBytes(path, OnePixelPng);
        return path;
    }

    private static TabWorkbooks BuildTab(string boardKey, BoardData boardData, int selectedWorkbookId)
    {
        var tab = new TabWorkbooks
        {
            BoardKeyOverrideForTests = boardKey,
            CurrentBoardDataOverrideForTests = boardData
        };

        tab.ActivateWorkbookOverrideForTests = (key, workbookId) =>
        {
            UserSettings.SetActiveWorkbookId(key, workbookId);
            tab.RefreshWorkbooks();
        };

        tab.RefreshWorkbooks();
        tab.SelectWorkbookForTests(selectedWorkbookId);
        return tab;
    }

    private static StackPanel EntriesPanel(TabWorkbooks tab) =>
        tab.GetControl<StackPanel>("SelectedSchematicEntriesPanel");

    // Every Border anywhere under a control, so a pill can be found without depending on the exact
    // nesting of the row it sits in.
    private static IEnumerable<Border> AllBorders(Control root)
    {
        if (root is Border border)
            yield return border;

        if (root is Panel panel)
        {
            foreach (var child in panel.Children.OfType<Control>())
                foreach (var found in AllBorders(child))
                    yield return found;
        }
        else if (root is Border { Child: Control inner })
        {
            foreach (var found in AllBorders(inner))
                yield return found;
        }
        else if (root is ContentControl { Content: Control content })
        {
            foreach (var found in AllBorders(content))
                yield return found;
        }
    }

    private static string AllText(Control root)
    {
        if (root is TextBlock block)
            return block.Text ?? string.Join("", block.Inlines?.Select(i => (i as Avalonia.Controls.Documents.Run)?.Text ?? "") ?? Array.Empty<string>());

        if (root is Panel panel)
            return string.Join(" ", panel.Children.OfType<Control>().Select(AllText));

        if (root is Border { Child: Control inner })
            return AllText(inner);

        if (root is ContentControl { Content: Control content })
            return AllText(content);

        return string.Empty;
    }

    // ###########################################################################################
    // The pill carrying the given word - the INNERMOST border whose text contains it.
    //
    // Innermost matters: an entry card is itself a Border, and its text contains every word inside
    // it, so a plain "first border containing the word" search returns the CARD and then asserts on
    // the card's own 1px grey outline instead of the pill's coloured one. That found a real pill
    // colour of #ffe6e6e6 on the first run of this test - the card, not the pill.
    // ###########################################################################################
    private static Border PillContaining(Control root, string word) =>
        AllBorders(root)
            .Where(b => b.Child is Control child && AllText(child).Contains(word, StringComparison.Ordinal))
            .OrderBy(b => AllText(b).Length)
            .First();

    private static IReadOnlyList<Avalonia.Controls.Documents.Run> Runs(TextBlock block) =>
        block.Inlines?.OfType<Avalonia.Controls.Documents.Run>().ToList()
        ?? new List<Avalonia.Controls.Documents.Run>();

    private static string InlineText(TextBlock block) =>
        string.Join("", Runs(block).Select(r => r.Text));

    private static bool IsNumber(string text) =>
        double.TryParse(text, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out _);

    private static Color ColorOf(IBrush? brush) =>
        brush is ISolidColorBrush solid ? solid.Color : Colors.Transparent;

    private int CreateWorkbookWithEntry(string schematicName, string category, string state, out int workbookId)
    {
        var workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "Repair job", "Found at the tip");
        Assert.NotNull(workbook);
        workbookId = workbook!.Id;

        var entry = WorklogManager.AddEntry(
            workbookId, schematicName, new Rect(0, 0, 10, 10),
            "Bad cap", "Leaked electrolytic.", category, state, Array.Empty<string>());

        Assert.NotNull(entry);
        return entry!.Id;
    }

    // Logs one Work done line against an entry, through the real persistence path - the totals
    // strip reads what is on disk, so an in-memory record would not reach it.
    private static void LogWorkDone(int workbookId, int entryId, double hours, double cost)
    {
        var entry = WorklogManager.GetEntries(workbookId).Single(e => e.Id == entryId);

        entry.WorkDoneItems.Add(new WorklogWorkDoneRecord
        {
            Id = entry.WorkDoneItems.Count + 1,
            Text = "Recap",
            Date = DateTime.Now,
            HoursSpent = hours,
            Cost = cost
        });

        Assert.True(WorklogManager.UpdateEntry(workbookId, entry));
    }

    private static BoardData BuildBoardData(string schematicName, string imagePath) =>
        new()
        {
            Schematics = new List<BoardSchematicEntry>
            {
                new() { SchematicName = schematicName, SchematicImageFile = imagePath }
            }
        };

    // ###########################################################################################
    // 1. THE INFORMATIONAL PILLS
    // ###########################################################################################

    // ###########################################################################################
    // The reported complaint, in one assertion: the status pill on an entry card and the one on the
    // tab's top-line must be the SAME pill. They were 1px grey and 2px coloured respectively, which
    // is precisely "the pill is not identical for the status".
    //
    // Asserts the width AND the colour, because matching on only one of them is what let these two
    // drift while each carried a comment claiming they matched.
    // ###########################################################################################
    [Fact]
    public void The_entry_card_and_top_line_status_pills_have_the_same_border_width_and_colour()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Issue", "Closed", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);

            var cardPill = PillContaining((Control)EntriesPanel(tab).Children[0], "Closed");
            var headerPill = tab.GetControl<Border>("WorkbookHeaderStatusPill");

            Assert.Equal(new Thickness(1), cardPill.BorderThickness);
            Assert.Equal(new Thickness(1), headerPill.BorderThickness);
            Assert.Equal(ColorOf(headerPill.BorderBrush), ColorOf(cardPill.BorderBrush));
        });
    }

    // The border is the STATE's own colour, not a neutral grey - that coloured border is what makes
    // these read as information at a glance, and it is what point 2 in the report was pointing at.
    // Open and Closed must therefore differ.
    [Fact]
    public void A_status_pills_border_is_the_state_colour_so_open_and_closed_differ()
    {
        UiTest.Run(() =>
        {
            // Resolved INSIDE UiTest.Run: these keys live in App.axaml's ThemeDictionaries and are
            // read through Application.Current, which exists only on the UI thread. Resolving them
            // on the test thread returns the fallback for both, so the assertion below would have
            // compared two fallbacks and passed while proving nothing.
            var open = WorklogInfoPillBuilder.ResolveStateColor("Open");
            var closed = WorklogInfoPillBuilder.ResolveStateColor("Closed");

            var openPill = WorklogInfoPillBuilder.BuildStatePill("Open");
            var closedPill = WorklogInfoPillBuilder.BuildStatePill("Closed");

            Assert.Equal(open, ColorOf(openPill.BorderBrush));
            Assert.Equal(closed, ColorOf(closedPill.BorderBrush));
            Assert.NotEqual(ColorOf(openPill.BorderBrush), ColorOf(closedPill.BorderBrush));
        });
    }

    // ###########################################################################################
    // The category chip gets the same treatment, asked for in the same breath as the pill: it had
    // shipped grey-outlined beside a colour-outlined status pill, which made the two look like
    // different kinds of control sitting next to each other.
    // ###########################################################################################
    [Fact]
    public void A_category_chips_border_is_its_own_category_colour_at_one_pixel()
    {
        UiTest.Run(() =>
        {
            foreach (string category in new[] { "Note", "Cosmetic", "Issue" })
            {
                var chip = WorklogInfoPillBuilder.BuildCategoryChip(category);

                Assert.Equal(new Thickness(1), chip.BorderThickness);
                Assert.Equal(WorklogInfoPillBuilder.ResolveCategoryColor(category), ColorOf(chip.BorderBrush));
            }
        });
    }

    // An entry's card shows its real category and state, not a fixed pair - the chip and pill are
    // built per entry.
    [Fact]
    public void An_entry_card_shows_its_own_category_and_state()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Cosmetic", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);
            string text = AllText((Control)EntriesPanel(tab).Children[0]);

            Assert.Contains("Cosmetic", text);
            Assert.Contains("Open", text);
        });
    }

    // ###########################################################################################
    // 2. CLICKABLE ENTRY CARDS
    // ###########################################################################################

    // ###########################################################################################
    // A card must LOOK clickable, and the Hand cursor is the only thing on it that says so - there
    // is no button and no hover chrome. This is the whole affordance, so it is worth pinning.
    //
    // The click itself is deliberately not driven here: it opens a modal via ShowDialog, which a
    // headless test cannot dismiss, and the same restriction already applies to the board pane's
    // pills. What makes the two provably identical is that both go through the ONE OpenEntryEditor -
    // see TabWorkbooks.BoardPreviews.cs, where the pill handler is now a two-line wrapper around it.
    // ###########################################################################################
    [Fact]
    public void An_entry_card_carries_a_hand_cursor_marking_it_clickable()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Issue", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);
            var card = (Border)EntriesPanel(tab).Children[0];

            Assert.Equal(new Cursor(StandardCursorType.Hand).ToString(), card.Cursor?.ToString());
        });
    }

    // ###########################################################################################
    // 3. THE SUMMARY STRIP
    // ###########################################################################################

    // ###########################################################################################
    // The headline is the always-visible half, and it must show the workbook's REAL totals rather
    // than a placeholder - the mockup-era literals on this tab were a reported problem once already.
    //
    // Read through Inlines, not Text: the numbers are bold Runs and the words are not, so the
    // block's Text is null. A reader that only looked at Text would see this as blank - the same
    // trap the search highlighting on this tab already documents.
    // ###########################################################################################
    [Fact]
    public void The_summary_headline_shows_the_workbooks_real_totals_and_counts_worklogs()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Issue", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);

            Assert.True(tab.GetControl<StackPanel>("WorkbookSummaryPanel").IsVisible);

            var headline = tab.GetControl<TextBlock>("WorkbookSummaryHeadlineText");

            Assert.Null(headline.Text);

            // NEITHER a time NOR a cost item: this workbook has no Work done lines at all, and
            // both of those figures are omitted when there is nothing to report rather than shown
            // as "0 minutes . 0 USD" - a line of real figures should not spend two columns saying
            // that two figures are absent. Asked for directly.
            //
            // "open" DOES stay at zero (it is 1 here) - see BuildHeadlineStats: it counts worklogs
            // that exist rather than reporting a missing total, and "0 open" tells a reader the job
            // is finished.
            Assert.Equal("1 worklog · 1 open", InlineText(headline));
        });
    }

    // ###########################################################################################
    // The other half of the above: once a workbook HAS time logged, the headline says it in hours
    // and minutes rather than in the decimal hours it is stored as - 1.75 h is "1 hour and 45
    // minutes". Decimal hours survives in exactly one place in the app, the NumericUpDown in the
    // full editor that types the value.
    //
    // Driven through WorklogManager.UpdateEntry and the tab's own refresh rather than by calling
    // WorkbookSummary directly, so it pins the WIRING - a strip still printing "1.75 h" would pass
    // every WorklogDurationFormatter unit test.
    // ###########################################################################################
    [Fact]
    public void The_summary_headline_says_the_time_in_hours_and_minutes()
    {
        this.LoadWorklog();
        int entryId = this.CreateWorkbookWithEntry("Sheet 1", "Issue", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        LogWorkDone(workbookId, entryId, hours: 1.25, cost: 0);
        LogWorkDone(workbookId, entryId, hours: 0.5, cost: 0);

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);

            string headline = InlineText(tab.GetControl<TextBlock>("WorkbookSummaryHeadlineText"));

            Assert.Contains("1 hour and 45 minutes", headline);

            // The two halves of the duration read as ONE figure: no " . " separator between them,
            // which is what Stat.JoinedToPrevious exists for. Without it the strip would say
            // "1 hour and . 45 minutes".
            Assert.DoesNotContain("1 hour and ·", headline);
        });
    }

    // ###########################################################################################
    // The Configuration tab's currency reaches this strip - the whole point of the setting is that
    // a repairer in Denmark sees DKK rather than a bare number a customer has to guess at.
    //
    // Driven through UserSettings and the tab's own refresh rather than by calling WorkbookSummary
    // directly, so it pins the WIRING: the tab reads the setting and passes it to BuildHeadlineStats.
    // A tab that kept printing the bare number would still pass every WorkbookSummary unit test.
    // ###########################################################################################
    [Fact]
    public void The_summary_headline_prints_the_configured_currency()
    {
        this.LoadWorklog();
        int entryId = this.CreateWorkbookWithEntry("Sheet 1", "Issue", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        // A REAL cost, because a zero one is now omitted from the headline entirely - with no cost
        // logged there would be no currency code on the line to assert against, and this test would
        // pass while proving nothing about the setting reaching the strip.
        LogWorkDone(workbookId, entryId, hours: 0, cost: 120);

        // Restored afterwards: UserSettings is static and LoadWorklog's own reset only runs at the
        // START of a test, so a code left set here would reach whatever ran next in this collection.
        try
        {
            UserSettings.WorklogCurrencyCode = "DKK";

            UiTest.Run(() =>
            {
                var tab = BuildTab(this.thisBoardKey, boardData, workbookId);

                var headline = tab.GetControl<TextBlock>("WorkbookSummaryHeadlineText");

                Assert.Equal("1 worklog · 120 DKK · 1 open", InlineText(headline));
            });
        }
        finally
        {
            UserSettings.WorklogCurrencyCode = WorklogCurrency.DefaultCode;
        }
    }

    // ###########################################################################################
    // The NUMBERS are bold and the words are not - asked for explicitly. Asserting on the runs
    // rather than on the finished string is the only way to see the difference at all.
    // ###########################################################################################
    [Fact]
    public void Only_the_numbers_in_the_summary_are_bold()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Issue", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);
            var runs = Runs(tab.GetControl<TextBlock>("WorkbookSummaryHeadlineText"));

            var bold = runs.Where(r => r.FontWeight == FontWeight.Bold).Select(r => r.Text ?? "").ToList();
            var plain = runs.Where(r => r.FontWeight != FontWeight.Bold).Select(r => r.Text ?? "").ToList();

            // Every bold run is a number, and no plain run is.
            Assert.NotEmpty(bold);
            Assert.All(bold, t => Assert.True(IsNumber(t), $"bold run [{t}] is not a number"));
            Assert.All(plain, t => Assert.False(IsNumber(t), $"plain run [{t}] is a number"));
        });
    }

    // Collapsed by default: the headline is worth one line on every visit, the four-line breakdown
    // is not - see UserSettings.WorkbooksSummaryExpanded.
    [Fact]
    public void The_summary_breakdown_starts_collapsed()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Issue", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);

            Assert.False(tab.GetControl<StackPanel>("WorkbookSummaryDetailPanel").IsVisible);
        });
    }

    [Fact]
    public void Toggling_the_summary_shows_the_breakdown_and_toggling_again_hides_it()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Issue", "Closed", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);
            var detail = tab.GetControl<StackPanel>("WorkbookSummaryDetailPanel");

            tab.ToggleSummaryForTests();
            Assert.True(detail.IsVisible);

            tab.ToggleSummaryForTests();
            Assert.False(detail.IsVisible);
        });
    }

    // ###########################################################################################
    // The expanded choice PERSISTS, and is re-applied on the next refresh. Without this the strip
    // folded itself shut on every board change and every entry save - the tab rebuilds this whole
    // header on each - which is exactly the behaviour that makes a collapsible panel annoying
    // enough to be worse than no panel.
    // ###########################################################################################
    [Fact]
    public void An_expanded_summary_stays_expanded_across_a_refresh()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Issue", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);

            tab.ToggleSummaryForTests();
            Assert.True(UserSettings.WorkbooksSummaryExpanded);

            tab.RefreshWorkbooks();

            Assert.True(tab.GetControl<StackPanel>("WorkbookSummaryDetailPanel").IsVisible);
        });
    }

    // The components line is hidden outright for a workbook that scopes none - a permanent "0
    // components in scope" row trains the eye to skip the whole block.
    [Fact]
    public void The_components_line_is_hidden_when_the_workbook_scopes_no_components()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Note", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);
            tab.ToggleSummaryForTests();

            Assert.False(tab.GetControl<TextBlock>("WorkbookSummaryComponentsText").IsVisible);
        });
    }

    // With no workbook selected there is nothing to summarise, so the strip goes away entirely
    // rather than showing zeroes under "No workbook selected" - the same rule the Edit/Delete/Export
    // buttons follow.
    [Fact]
    public void The_summary_strip_is_hidden_when_no_workbook_is_selected()
    {
        this.LoadWorklog();

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, new BoardData(), 0);

            Assert.False(tab.GetControl<StackPanel>("WorkbookSummaryPanel").IsVisible);
            Assert.False(tab.GetControl<StackPanel>("WorkbookHeaderActionsPanel").IsVisible);
        });
    }

    // ###########################################################################################
    // The category and state counts are drawn as the SAME non-selectable pills the rest of the app
    // uses, each carrying its count - asked for directly, replacing two plain lines of text.
    //
    // Every category and state is present including the zeroes: a "0 Issue" pill says this workbook
    // records no faults, and a row that dropped the empty ones would change width as the workbook
    // was worked on.
    // ###########################################################################################
    [Fact]
    public void The_category_and_state_counts_are_drawn_as_counted_pills()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Note", "Closed", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);
            tab.ToggleSummaryForTests();

            var categories = tab.GetControl<WrapPanel>("WorkbookSummaryCategoryPanel").Children
                .OfType<Border>().Select(AllText).ToList();
            var states = tab.GetControl<WrapPanel>("WorkbookSummaryStatePanel").Children
                .OfType<Border>().Select(AllText).ToList();

            Assert.Equal(3, categories.Count);
            Assert.Equal(2, states.Count);

            // The count LEADS the pill - "1 {icon} Note" - which is the order the sentence reads in.
            Assert.StartsWith("1", categories[0]);
            Assert.Contains("Note", categories[0]);
            Assert.StartsWith("0", categories[1]);
            Assert.Contains("Cosmetic", categories[1]);
            Assert.Contains("Closed", states[1]);
        });
    }

    // A counted pill is still the ordinary informational pill: 1px, in its own colour. It would be
    // no use matching the rest of the app everywhere except the one place showing several at once.
    [Fact]
    public void A_counted_pill_keeps_the_ordinary_informational_outline()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Issue", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);
            tab.ToggleSummaryForTests();

            var issuePill = tab.GetControl<WrapPanel>("WorkbookSummaryCategoryPanel").Children
                .OfType<Border>().Single(b => AllText(b).Contains("Issue", StringComparison.Ordinal));

            Assert.Equal(new Thickness(1), issuePill.BorderThickness);
            Assert.Equal(WorklogInfoPillBuilder.ResolveCategoryColor("Issue"), ColorOf(issuePill.BorderBrush));
        });
    }

    // ###########################################################################################
    // A ZERO-count pill is muted to a neutral grey rather than its category/state colour - asked
    // for directly, after the summary strip's five always-present pills (including the zeroes,
    // see BuildCountPills' own header) blended together with nothing to draw the eye to which
    // categories/states actually have entries.
    //
    // Pinned against Workbooks_ZeroCount_Fg - its OWN dedicated token, not one of the tab's two
    // body-text greys: Muted_Fg, then Faint_Fg, were each tried here in turn and both reported as
    // "still looks black"/"needs to be lighter" once bolded onto a small pill, since both are tuned
    // to stay readable as PROSE rather than to read as washed-out/disabled - see MutedPillBrush's
    // own comment. Muted pills also drop the count's bold weight (see BuildCountLabel).
    //
    // Only the COUNTED path is muted at zero - a workbook with one Note entry and zero Cosmetic/
    // Issue ones still has a real "Open" or "Closed" state pill elsewhere on this very tab (the
    // top-line, a workbook card, an entry card) that must never greyed out just because it shares
    // a resolver with this one.
    // ###########################################################################################
    [Fact]
    public void A_zero_count_pill_is_muted_to_a_neutral_grey_instead_of_its_own_colour()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Note", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);
            tab.ToggleSummaryForTests();

            var mutedFg = ColorOf(ThemeResources.ResolveBrush("Workbooks_ZeroCount_Fg"));

            // Cosmetic and Issue are both zero - the workbook has one Note entry only.
            var cosmeticChip = tab.GetControl<WrapPanel>("WorkbookSummaryCategoryPanel").Children
                .OfType<Border>().Single(b => AllText(b).Contains("Cosmetic", StringComparison.Ordinal));
            var issueChip = tab.GetControl<WrapPanel>("WorkbookSummaryCategoryPanel").Children
                .OfType<Border>().Single(b => AllText(b).Contains("Issue", StringComparison.Ordinal));

            Assert.Equal(mutedFg, ColorOf(cosmeticChip.BorderBrush));
            Assert.Equal(mutedFg, ColorOf(issueChip.BorderBrush));

            // The muted "0" itself must not be bold - a bold zero in this lighter grey still read
            // as solid/dark at pill size, which is what made Muted_Fg look black in the first place.
            var cosmeticCount = TextBlocksIn(cosmeticChip).First(b => b.Text == "0");
            Assert.Equal(FontWeight.Normal, cosmeticCount.FontWeight);

            // Note itself has a real count (1) and must keep its own category colour.
            var noteChip = tab.GetControl<WrapPanel>("WorkbookSummaryCategoryPanel").Children
                .OfType<Border>().Single(b => AllText(b).Contains("Note", StringComparison.Ordinal));
            Assert.Equal(WorklogInfoPillBuilder.ResolveCategoryColor("Note"), ColorOf(noteChip.BorderBrush));

            // Closed is zero (the one entry is Open) and must be muted the same way.
            var closedPill = tab.GetControl<WrapPanel>("WorkbookSummaryStatePanel").Children
                .OfType<Border>().Single(b => AllText(b).Contains("Closed", StringComparison.Ordinal));
            Assert.Equal(mutedFg, ColorOf(closedPill.BorderBrush));

            // The workbook's own top-line status pill reports a REAL state (Open, uncounted) and
            // must not be muted just because BuildStatePill also serves the summary's zero counts.
            var topLineStatusPill = tab.GetControl<Border>("WorkbookHeaderStatusPill");
            Assert.Equal(WorklogInfoPillBuilder.ResolveStateColor("Open"), ColorOf(topLineStatusPill.BorderBrush));
        });
    }

    // ###########################################################################################
    // A COUNTED pill carries NO icon; an uncounted one still does.
    //
    // Asked for after the counted pills shipped with their icons: a padlock or a category glyph
    // sitting between a number and its label reads as a third piece of information rather than as
    // decoration - "2 [lock] Open" invites the question of what the lock is counting.
    //
    // The uncounted half of the assertion matters just as much: on an entry card the glyph is the
    // only thing separating Open from Closed at a glance, so dropping it everywhere would have
    // traded one problem for a worse one.
    //
    // Counted by the number of TextBlocks in the pill: an icon is its own block, so a counted pill
    // is (count, label) and an uncounted one is (glyph, label) - both two, which is why this looks
    // at the FontFamily rather than the count. The icon block is the only one in a Font Awesome
    // family.
    // ###########################################################################################
    [Fact]
    public void A_counted_pill_drops_its_icon_while_an_uncounted_one_keeps_it()
    {
        UiTest.Run(() =>
        {
            var fontAwesome = ThemeResources.ResolveFontAwesomeSolid();

            foreach (string state in new[] { "Open", "Closed" })
            {
                Assert.DoesNotContain(TextBlocksIn(WorklogInfoPillBuilder.BuildStatePill(state, count: 2)),
                    b => Equals(b.FontFamily, fontAwesome));

                Assert.Contains(TextBlocksIn(WorklogInfoPillBuilder.BuildStatePill(state)),
                    b => Equals(b.FontFamily, fontAwesome));
            }

            foreach (string category in new[] { "Note", "Cosmetic", "Issue" })
            {
                var counted = TextBlocksIn(WorklogInfoPillBuilder.BuildCategoryChip(category, count: 2));

                // The count and the label, and nothing else - no glyph block of any family.
                Assert.Equal(new[] { "2", category }, counted.Select(b => b.Text));
            }
        });
    }

    private static IReadOnlyList<TextBlock> TextBlocksIn(Border pill) =>
        pill.Child is Panel panel ? panel.Children.OfType<TextBlock>().ToList() : new List<TextBlock>();

    // ###########################################################################################
    // WorklogInfoPillBuilder's own muting rule, exercised directly rather than through the summary
    // strip: count: 0 replaces the state/category colour with Workbooks_ZeroCount_Fg on both the
    // border and every text/label run, AND drops the count's bold weight, while a non-zero count
    // (or no count at all, an ordinary uncounted pill) keeps the real colour and its usual weight.
    //
    // Its own dedicated token, and non-bold rather than bold: Muted_Fg and then Faint_Fg were each
    // tried here in turn and both reported as still too dark at this pill's small bold size - see
    // MutedPillBrush and BuildCountLabel's own comments.
    // ###########################################################################################
    [Fact]
    public void BuildStatePill_and_BuildCategoryChip_mute_only_a_zero_count()
    {
        UiTest.Run(() =>
        {
            var mutedFg = ColorOf(ThemeResources.ResolveBrush("Workbooks_ZeroCount_Fg"));

            var zeroStatePill = WorklogInfoPillBuilder.BuildStatePill("Open", count: 0);
            Assert.Equal(mutedFg, ColorOf(zeroStatePill.BorderBrush));
            Assert.All(TextBlocksIn(zeroStatePill), b => Assert.Equal(mutedFg, ColorOf(b.Foreground)));
            Assert.Equal(FontWeight.Normal, TextBlocksIn(zeroStatePill).First(b => b.Text == "0").FontWeight);

            var countedStatePill = WorklogInfoPillBuilder.BuildStatePill("Open", count: 3);
            Assert.Equal(WorklogInfoPillBuilder.ResolveStateColor("Open"), ColorOf(countedStatePill.BorderBrush));
            Assert.Equal(FontWeight.Bold, TextBlocksIn(countedStatePill).First(b => b.Text == "3").FontWeight);

            var uncountedStatePill = WorklogInfoPillBuilder.BuildStatePill("Open");
            Assert.Equal(WorklogInfoPillBuilder.ResolveStateColor("Open"), ColorOf(uncountedStatePill.BorderBrush));

            var zeroChip = WorklogInfoPillBuilder.BuildCategoryChip("Issue", count: 0);
            Assert.Equal(mutedFg, ColorOf(zeroChip.BorderBrush));
            Assert.All(TextBlocksIn(zeroChip), b => Assert.Equal(mutedFg, ColorOf(b.Foreground)));
            Assert.Equal(FontWeight.Normal, TextBlocksIn(zeroChip).First(b => b.Text == "0").FontWeight);

            var countedChip = WorklogInfoPillBuilder.BuildCategoryChip("Issue", count: 3);
            Assert.Equal(WorklogInfoPillBuilder.ResolveCategoryColor("Issue"), ColorOf(countedChip.BorderBrush));
            Assert.Equal(FontWeight.Bold, TextBlocksIn(countedChip).First(b => b.Text == "3").FontWeight);

            var uncountedChip = WorklogInfoPillBuilder.BuildCategoryChip("Issue");
            Assert.Equal(WorklogInfoPillBuilder.ResolveCategoryColor("Issue"), ColorOf(uncountedChip.BorderBrush));
        });
    }

    // ###########################################################################################
    // 4. EXPORT
    // ###########################################################################################

    // ###########################################################################################
    // BOTH export formats have their own visible button. The ZIP was originally reachable only as a
    // second file type inside the save dialog, which made it invisible from the tab - reported as
    // "I do not see the Export to ZIP anywhere". A format the user cannot see does not exist.
    //
    // Also asserts they carry no icon: the fa-regular file-pdf glyph rendered as a blank box in the
    // shipped font subset, so both are plain text labels.
    // ###########################################################################################
    [Fact]
    public void Both_export_formats_have_their_own_visible_button_with_no_icon()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Note", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);

            Assert.True(tab.GetControl<StackPanel>("WorkbookHeaderActionsPanel").IsVisible);
            Assert.Equal("Export to PDF", tab.GetControl<Button>("ExportWorkbookButton").Content);
            Assert.Equal("Export to ZIP", tab.GetControl<Button>("ExportWorkbookZipButton").Content);
        });
    }

    // ###########################################################################################
    // The document an export would write, built through the tab's own path - so the board data the
    // tab holds really does reach the exported document's schematic sections.
    //
    // The WRITING is not covered: it needs QuestPDF to produce bytes, and asserting on those tests
    // the library rather than this app. What is worth pinning is that the tab hands the model the
    // right board and the right entries, which is the part that can silently regress.
    // ###########################################################################################
    [Fact]
    public void The_export_document_is_built_from_the_tabs_own_board_data_and_entries()
    {
        this.LoadWorklog();
        int entryId = this.CreateWorkbookWithEntry("Sheet 1", "Issue", "Closed", out int workbookId);
        string imagePath = this.WriteSchematicImage("sheet1.png");
        var boardData = BuildBoardData("Sheet 1", imagePath);

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);
            var workbook = WorklogManager.GetWorkbooksForBoard(this.thisBoardKey).Single(w => w.Id == workbookId);

            var document = tab.BuildExportDocumentForTests(workbook);

            Assert.Equal("Repair job", document.Title);
            Assert.Equal(1, document.Totals.WorklogCount);

            var section = Assert.Single(document.Sections);
            Assert.Equal("Sheet 1", section.SchematicName);

            // The board data's schematic image reached the document - the piece that would break if
            // the tab stopped resolving paths against the data root.
            Assert.Equal(imagePath, section.SchematicImagePath);
            Assert.Equal(entryId, Assert.Single(section.Entries).Record.Id);
        });
    }

    // The exported file name identifies the workbook without carrying its title - see
    // WorkbookExportModel.BuildFileBaseName. Checked through the tab so the board key it puts in
    // the name is the one the tab is actually showing.
    [Fact]
    public void The_suggested_export_file_name_names_the_workbook_and_its_board()
    {
        this.LoadWorklog();
        this.CreateWorkbookWithEntry("Sheet 1", "Note", "Open", out int workbookId);
        var boardData = BuildBoardData("Sheet 1", this.WriteSchematicImage("sheet1.png"));

        UiTest.Run(() =>
        {
            var tab = BuildTab(this.thisBoardKey, boardData, workbookId);
            var workbook = WorklogManager.GetWorkbooksForBoard(this.thisBoardKey).Single(w => w.Id == workbookId);

            string name = WorkbookExportModel.BuildFileBaseName(tab.BuildExportDocumentForTests(workbook));

            Assert.StartsWith($"Workbook_{workbookId}_", name);
            Assert.DoesNotContain("Repair job", name);
        });
    }
}
