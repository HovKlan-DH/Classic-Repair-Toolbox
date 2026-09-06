using Avalonia;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// The pure parts of a worklog entry that BOTH tabs rendering one depend on: which components its
// marked area covers, what its Work done rows total, and the singular/plural count wording.
//
// These were private members of two different UserControls, one copy each, and the component-scope
// computation decides whether the "Mark components in scope" checklist appears in the editor modal.
// The Schematics tab and the Workbooks tab are required to open the SAME modal, and a fix applied
// to one copy and missed on the other is exactly the divergence already reported once - so what is
// pinned here is the single shared implementation both now call.
//
// No collection attribute: nothing here touches a static singleton, a control, or the filesystem.
public class WorklogEntryScopeTests
{
    private static BoardData BuildBoardData(params string[] boardLabels)
    {
        var data = new BoardData();
        foreach (string label in boardLabels)
        {
            data.Components.Add(new ComponentEntry
            {
                BoardLabel = label,
                FriendlyName = label + " name",
            });
        }

        return data;
    }

    private static WorklogEntryRecord EntryCovering(string schematicName, Rect area) => new()
    {
        Id = 1,
        SchematicName = schematicName,
        AreaX = area.X,
        AreaY = area.Y,
        AreaWidth = area.Width,
        AreaHeight = area.Height,
    };

    private static Dictionary<string, Dictionary<string, List<Rect>>> RectCache(
        string schematicName, params (string Label, Rect Rect)[] rects)
    {
        var byLabel = new Dictionary<string, List<Rect>>(StringComparer.OrdinalIgnoreCase);
        foreach (var (label, rect) in rects)
        {
            byLabel[label] = new List<Rect> { rect };
        }

        // OrdinalIgnoreCase, matching how Main builds the real cache at all four of its write sites
        // and how TabSchematics declares the field.
        return new Dictionary<string, Dictionary<string, List<Rect>>>(StringComparer.OrdinalIgnoreCase)
        {
            [schematicName] = byLabel,
        };
    }

    [Fact]
    public void Components_whose_highlight_the_area_touches_are_in_scope()
    {
        var boardData = BuildBoardData("U1", "U2");
        var cache = RectCache("Sheet 1",
            ("U1", new Rect(0, 0, 10, 10)),
            ("U2", new Rect(100, 100, 10, 10)));

        var scope = WorklogEntryScope.BuildComponentsInScope(
            boardData, cache, EntryCovering("Sheet 1", new Rect(5, 5, 20, 20)));

        Assert.NotNull(scope);
        Assert.Equal(new[] { "U1" }, scope!.Select(c => c.BoardLabel));
    }

    // NULL and EMPTY are NOT interchangeable to the caller, and this is the reason the distinction is
    // pinned here rather than left to the two call sites: empty means "this area genuinely covers no
    // component", null means "the scope is unknown", and only null leaves the entry's saved
    // ComponentLabels alone. Returning empty for an unknown scope would wipe the user's selection the
    // first time they saved.
    [Fact]
    public void An_area_touching_nothing_is_empty_scope_not_unknown_scope()
    {
        var boardData = BuildBoardData("U1");
        var cache = RectCache("Sheet 1", ("U1", new Rect(100, 100, 10, 10)));

        var scope = WorklogEntryScope.BuildComponentsInScope(
            boardData, cache, EntryCovering("Sheet 1", new Rect(0, 0, 5, 5)));

        Assert.NotNull(scope);
        Assert.Empty(scope!);
    }

    [Fact]
    public void No_board_data_means_the_scope_is_unknown()
    {
        var cache = RectCache("Sheet 1", ("U1", new Rect(0, 0, 10, 10)));

        Assert.Null(WorklogEntryScope.BuildComponentsInScope(
            null, cache, EntryCovering("Sheet 1", new Rect(0, 0, 5, 5))));
    }

    [Fact]
    public void No_highlight_cache_at_all_means_the_scope_is_unknown()
    {
        Assert.Null(WorklogEntryScope.BuildComponentsInScope(
            BuildBoardData("U1"), null, EntryCovering("Sheet 1", new Rect(0, 0, 5, 5))));
    }

    // A cache that exists but has no key for THIS entry's schematic. Reachable on a board load, where
    // the cache is populated by a fire-and-forget task after the previews are already on screen, and
    // on a region switch, where HighlightRectBuilder legitimately drops every highlight belonging to
    // the other region. Unknown, not empty - saving must not discard the entry's component labels
    // just because the rects for its schematic are not loaded.
    [Fact]
    public void A_cache_with_no_entry_for_that_schematic_means_the_scope_is_unknown()
    {
        var cache = RectCache("Sheet 2", ("U1", new Rect(0, 0, 10, 10)));

        Assert.Null(WorklogEntryScope.BuildComponentsInScope(
            BuildBoardData("U1"), cache, EntryCovering("Sheet 1", new Rect(0, 0, 5, 5))));
    }

    [Fact]
    public void An_entry_with_no_schematic_name_has_an_unknown_scope()
    {
        var cache = RectCache("Sheet 1", ("U1", new Rect(0, 0, 10, 10)));

        Assert.Null(WorklogEntryScope.BuildComponentsInScope(
            BuildBoardData("U1"), cache, EntryCovering("   ", new Rect(0, 0, 5, 5))));
    }

    // ------------------------------------------------------------------- work done totals

    [Fact]
    public void Work_done_totals_sum_hours_and_cost_across_every_row()
    {
        var entry = new WorklogEntryRecord
        {
            WorkDoneItems =
            {
                new WorklogWorkDoneRecord { HoursSpent = 1.5, Cost = 20.0 },
                new WorklogWorkDoneRecord { HoursSpent = 2.25, Cost = 5.5 },
            },
        };

        var (hours, cost) = WorklogEntryScope.GetWorkDoneTotals(entry);

        Assert.Equal(3.75, hours);
        Assert.Equal(25.5, cost);
    }

    // Zero rather than a blank or a null: both callers render these into a label unconditionally.
    [Fact]
    public void An_entry_with_no_work_done_rows_totals_zero()
    {
        Assert.Equal((0.0, 0.0), WorklogEntryScope.GetWorkDoneTotals(new WorklogEntryRecord()));
        Assert.Equal((0.0, 0.0), WorklogEntryScope.GetWorkDoneTotals(null));
    }

    // ------------------------------------------------------------------------ pluralisation

    // The plural rule has to hold at zero as well as at one and many - "1 entries" is the classic
    // way this goes wrong, and there were several open-coded copies of the ternary before this.
    [Theory]
    [InlineData(0, "0 entries")]
    [InlineData(1, "1 entry")]
    [InlineData(2, "2 entries")]
    [InlineData(11, "11 entries")]
    public void A_count_takes_the_singular_only_at_one(int count, string expected)
    {
        Assert.Equal(expected, WorklogEntryScope.FormatCount(count, "entry", "entries"));
    }

    // ---------------------------------------------------------------------------------------------
    // FormatWorkbookDate - the date shown in the worklog bar above the tabs and on the workbook
    // cards. One formatter for both, so those two cannot disagree; they used to hold a format string
    // each.
    // ---------------------------------------------------------------------------------------------

    // THE REPORTED POINT: no leading zero on the day. The 6th is "2026-September-6", not
    // "2026-September-06".
    [Fact]
    public void A_single_digit_day_carries_no_leading_zero()
    {
        Assert.Equal("2026-September-6", WorklogEntryScope.FormatWorkbookDate(new DateTime(2026, 9, 6)));
        Assert.Equal("2026-January-1", WorklogEntryScope.FormatWorkbookDate(new DateTime(2026, 1, 1)));
    }

    // The other half: suppressing the padding must not truncate a two-digit day. Worth its own case
    // because "d" and "dd" only differ for days 1-9, so a test written on a single-digit day alone
    // passes for both formats.
    [Fact]
    public void A_two_digit_day_is_shown_in_full()
    {
        Assert.Equal("2026-September-26", WorklogEntryScope.FormatWorkbookDate(new DateTime(2026, 9, 26)));
        Assert.Equal("2026-December-31", WorklogEntryScope.FormatWorkbookDate(new DateTime(2026, 12, 31)));
    }

    // The month is a NAME and the format is invariant, so the date cannot be read as 06-09 or 09-06
    // depending on the machine's locale - which is the whole reason this format was chosen over a
    // numeric one.
    [Fact]
    public void The_month_is_always_an_english_name_regardless_of_the_current_culture()
    {
        var previous = System.Globalization.CultureInfo.CurrentCulture;

        try
        {
            System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("da-DK");

            Assert.Equal("2026-September-6", WorklogEntryScope.FormatWorkbookDate(new DateTime(2026, 9, 6)));
        }
        finally
        {
            System.Globalization.CultureInfo.CurrentCulture = previous;
        }
    }

    // ---------------------------------------------------------------------------------------------
    // FormatWorkbookDateLine - the WORD and the DATE together. Both the worklog bar above the tabs
    // and the workbook cards in the Workbooks tab print this line for the same workbook, feet apart
    // on the same screen, and they disagreed: the bar chose "ended"/EndDate for a closed workbook
    // while the card wrote "started" against StartDate unconditionally. These tests are why the
    // word and the date are chosen by ONE function - pairing them at the call site is what let the
    // two drift.
    // ---------------------------------------------------------------------------------------------

    // An OPEN workbook reports when it started, even when a stale EndDate is sitting on the record.
    // That is not hypothetical: RecomputeWorkbookStatus deliberately leaves EndDate standing when a
    // workbook REOPENS, so an open workbook carrying a finish date is the normal state of any job
    // that was closed and then had a worklog reopened in it.
    [Fact]
    public void An_open_workbook_reports_when_it_started_even_carrying_a_stale_end_date()
    {
        Assert.Equal(
            "started 2026-September-6",
            WorklogEntryScope.FormatWorkbookDateLine("Open", new DateTime(2026, 9, 6), null));

        Assert.Equal(
            "started 2026-September-6",
            WorklogEntryScope.FormatWorkbookDateLine("Open", new DateTime(2026, 9, 6), new DateTime(2026, 10, 1)));
    }

    // THE REPORTED POINT: a closed workbook reports when it ENDED - the word AND the date move
    // together. A version that changed only the word would produce "ended 2026-September-6" against
    // the start date, which is a wrong fact rather than a wrong label, so both halves are asserted.
    [Fact]
    public void A_closed_workbook_reports_the_end_word_and_the_end_date_together()
    {
        Assert.Equal(
            "ended 2026-October-1",
            WorklogEntryScope.FormatWorkbookDateLine("Closed", new DateTime(2026, 9, 6), new DateTime(2026, 10, 1)));
    }

    // A closed workbook with NO EndDate falls back to BOTH halves of the started form - every
    // workbook closed before that field existed reads this way, and there is deliberately no
    // migration. Showing the start it really has beats inventing a finish date.
    [Fact]
    public void A_closed_workbook_with_no_end_date_falls_back_to_the_whole_started_line()
    {
        Assert.Equal(
            "started 2026-September-6",
            WorklogEntryScope.FormatWorkbookDateLine("Closed", new DateTime(2026, 9, 6), null));
    }

    // The status is matched the way IsWorkbookStatusOpen matches it - trimmed and case-insensitive,
    // since a status read back off disk (or hand-edited) can carry other casing. An unrecognised
    // status reads as open, matching how every status pill in the app already falls back.
    [Theory]
    [InlineData("Closed")]
    [InlineData("closed")]
    [InlineData("  CLOSED  ")]
    public void A_closed_status_is_recognised_whatever_its_casing_or_padding(string status)
    {
        Assert.Equal(
            "ended 2026-October-1",
            WorklogEntryScope.FormatWorkbookDateLine(status, new DateTime(2026, 9, 6), new DateTime(2026, 10, 1)));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Something else")]
    public void An_unrecognised_status_reads_as_open_and_reports_the_start(string? status)
    {
        Assert.Equal(
            "started 2026-September-6",
            WorklogEntryScope.FormatWorkbookDateLine(status, new DateTime(2026, 9, 6), new DateTime(2026, 10, 1)));
    }

    // The line uses the SAME date formatter as FormatWorkbookDate above, so the no-leading-zero rule
    // it pins holds here too - a second formatter inside this method would be exactly the drift both
    // of these exist to prevent.
    [Fact]
    public void The_line_carries_the_shared_date_format()
    {
        var date = new DateTime(2026, 12, 31);

        Assert.Equal(
            $"started {WorklogEntryScope.FormatWorkbookDate(date)}",
            WorklogEntryScope.FormatWorkbookDateLine("Open", date, null));
    }

}
