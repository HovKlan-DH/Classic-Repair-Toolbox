using System.Globalization;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// What one workbook adds up to across its entries - the numbers the Workbooks tab's summary strip
// shows and the exported PDF prints as its opening section.
//
// Both surfaces read these from the ONE computation, which is the point of pinning it here: a
// document handed to a customer must not be able to report different totals from the screen the
// repairer was looking at when they produced it.
//
// No collection attribute: nothing here touches a static singleton, a control, or the filesystem.
public class WorkbookSummaryTests
{
    private static WorklogEntryRecord Entry(
        int id,
        string category = "Note",
        string state = "Open",
        double hours = 0,
        double cost = 0,
        int comments = 0,
        int links = 0,
        int photos = 0,
        int files = 0,
        string[]? components = null,
        string[]? completed = null)
    {
        var entry = new WorklogEntryRecord
        {
            Id = id,
            Category = category,
            State = state
        };

        if (hours > 0 || cost > 0)
        {
            entry.WorkDoneItems.Add(new WorklogWorkDoneRecord { Id = 1, HoursSpent = hours, Cost = cost });
        }

        for (int i = 0; i < comments; i++)
            entry.Comments.Add(new WorklogCommentRecord { Id = i + 1, Text = $"comment {i}" });

        for (int i = 0; i < links; i++)
            entry.Links.Add(new WorklogLinkRecord { Id = i + 1, Url = $"https://example.com/{i}" });

        for (int i = 0; i < photos; i++)
            entry.Photos.Add(new WorklogAttachmentRecord { Id = i + 1, FileName = $"photo{i}.png" });

        for (int i = 0; i < files; i++)
            entry.Files.Add(new WorklogAttachmentRecord { Id = i + 1, FileName = $"file{i}.pdf" });

        if (components != null)
            entry.ComponentLabels.AddRange(components);

        if (completed != null)
            entry.CompletedComponentLabels.AddRange(completed);

        return entry;
    }

    // A workbook with no entries is an ordinary state - one just created - not a failure, and both
    // callers render whatever comes back. Returning null instead would make every caller guard.
    [Fact]
    public void A_workbook_with_no_entries_summarizes_to_all_zeroes_rather_than_null()
    {
        var totals = WorkbookSummary.Summarize(null);

        Assert.Equal(0, totals.WorklogCount);
        Assert.Equal(0.0, totals.TotalHours);
        Assert.Equal(0.0, totals.TotalCost);
        Assert.Equal(0, totals.OpenWorklogCount);
        Assert.Equal(0, totals.ClosedWorklogCount);
    }

    [Fact]
    public void Hours_and_cost_are_summed_across_every_entrys_work_done_rows()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            Entry(1, hours: 2.5, cost: 100),
            Entry(2, hours: 1.25, cost: 30.5),
            Entry(3)
        });

        Assert.Equal(3, totals.WorklogCount);
        Assert.Equal(3.75, totals.TotalHours);
        Assert.Equal(130.5, totals.TotalCost);
    }

    [Fact]
    public void Comments_links_photos_and_files_are_counted_across_the_whole_workbook()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            Entry(1, comments: 2, links: 1, photos: 3, files: 1),
            Entry(2, comments: 1, photos: 1)
        });

        Assert.Equal(3, totals.CommentCount);
        Assert.Equal(1, totals.LinkCount);
        Assert.Equal(4, totals.PhotoCount);
        Assert.Equal(1, totals.FileCount);
    }

    // Every category and state key is present even at zero - see WorkbookSummary.Categories. A row
    // that silently omitted the empty ones would make the reader work out which were missing.
    [Fact]
    public void Every_category_and_state_is_reported_even_when_no_entry_uses_it()
    {
        var totals = WorkbookSummary.Summarize(new[] { Entry(1, category: "Issue", state: "Closed") });

        Assert.Equal(0, totals.EntriesByCategory["Note"]);
        Assert.Equal(0, totals.EntriesByCategory["Cosmetic"]);
        Assert.Equal(1, totals.EntriesByCategory["Issue"]);
        Assert.Equal(0, totals.EntriesByState["Open"]);
        Assert.Equal(1, totals.EntriesByState["Closed"]);
    }

    // An entry carrying a category or state from a future build (or a hand-edited entries.json)
    // must not be miscounted as one of the known values - that would quietly misreport the
    // workbook. It still counts toward WorklogCount, so nothing goes missing.
    [Fact]
    public void An_unrecognised_category_is_counted_in_no_bucket_but_still_counts_as_an_entry()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            Entry(1, category: "Note"),
            Entry(2, category: "Catastrophic", state: "Deferred")
        });

        Assert.Equal(2, totals.WorklogCount);
        Assert.Equal(1, totals.EntriesByCategory["Note"]);
        Assert.Equal(0, totals.EntriesByCategory["Cosmetic"]);
        Assert.Equal(0, totals.EntriesByCategory["Issue"]);

        // Only the recognised entry is in a state bucket, so the two do not add up to WorklogCount.
        Assert.Equal(1, totals.OpenWorklogCount);
        Assert.Equal(0, totals.ClosedWorklogCount);
    }

    // ###########################################################################################
    // The same component reached from two entries counts ONCE. Summing them instead would report
    // more components than the board physically has, which reads as a data error to anyone who
    // knows the board - and "checked U1, then replaced U1" is an entirely normal pair of entries.
    // ###########################################################################################
    [Fact]
    public void A_component_scoped_by_two_entries_counts_once()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            Entry(1, components: new[] { "U1", "U2" }),
            Entry(2, components: new[] { "U1", "C5" })
        });

        Assert.Equal(3, totals.ComponentCount);
    }

    // Component labels are compared case-insensitively everywhere else in this app, so "u1" and
    // "U1" are the same chip here too.
    [Fact]
    public void Component_labels_are_matched_case_insensitively()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            Entry(1, components: new[] { "U1" }),
            Entry(2, components: new[] { "u1" })
        });

        Assert.Equal(1, totals.ComponentCount);
    }

    [Fact]
    public void Completed_components_are_counted_separately_from_those_in_scope()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            Entry(1, components: new[] { "U1", "U2", "C5" }, completed: new[] { "U1" })
        });

        Assert.Equal(3, totals.ComponentCount);
        Assert.Equal(1, totals.CompletedComponentCount);
    }

    // Blank labels are skipped rather than counted as a component called "": an empty string in
    // ComponentLabels comes from hand-edited data and is not a component.
    [Fact]
    public void A_blank_component_label_is_not_counted()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            Entry(1, components: new[] { "U1", "", "   " })
        });

        Assert.Equal(1, totals.ComponentCount);
    }

    // ###########################################################################################
    // The headline is what the collapsed strip shows and what the PDF prints first, so its exact
    // shape matters. Hours and cost use the SAME "0.##" InvariantCulture formatting as the entry
    // detail card and the editor's own summary line - a Danish user's comma decimal separator in
    // one place and a point in another would look like two different numbers.
    //
    // "WORKLOGS", not "entries": the app calls these worklogs everywhere the user can see one, and
    // "entry" was internal vocabulary leaking out through this one line. Reported directly.
    // ###########################################################################################
    [Fact]
    public void The_headline_counts_worklogs_not_entries()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            Entry(1, hours: 2.5, cost: 100, state: "Open"),
            Entry(2, hours: 1, cost: 30, state: "Closed"),
            Entry(3, state: "Open")
        });

        // The time reads as HOURS AND MINUTES, not as the decimal hours it is stored as - 3.5
        // is "3 hours and 30 minutes". See WorklogDurationFormatter; decimal hours survives in
        // exactly one place, the NumericUpDown that types it.
        //
        // Asserted on the PARTS, which is what both renderers consume - the summary strip and the
        // PDF each bold the numbers and leave the words plain, so neither ever builds this as one
        // string. Joining them here is the test saying what the line reads as end to end without
        // requiring a shipped method that only the test would call.
        Assert.Equal("3 worklogs · 3 hours and 30 minutes · 130 DKK · 2 open",
            JoinHeadline(WorkbookSummary.BuildHeadlineStats(totals, "DKK")));
    }

    // ###########################################################################################
    // The headline as one string, joined the way both renderers lay the parts out: " · " between
    // stats, and NOTHING before a part flagged JoinedToPrevious (a duration's second half), so the
    // separator never lands inside "2 hours and 30 minutes".
    //
    // Lives in the test rather than in WorkbookSummary because nothing that ships wants a finished
    // headline string - see that class's own note where FormatHeadline used to be.
    // ###########################################################################################
    private static string JoinHeadline(IReadOnlyList<WorkbookSummary.Stat> stats)
    {
        var text = new System.Text.StringBuilder();

        for (int i = 0; i < stats.Count; i++)
        {
            if (i > 0 && !stats[i].JoinedToPrevious)
                text.Append(" · ");

            text.Append(stats[i].Prefix).Append(stats[i].Number).Append(stats[i].Suffix);
        }

        return text.ToString();
    }

    // A one-worklog workbook must not read "1 worklogs" - the kind of bug that shipped once
    // elsewhere in this app.
    [Fact]
    public void A_single_worklog_headline_reads_worklog_not_worklogs()
    {
        var totals = WorkbookSummary.Summarize(new[] { Entry(1) });

        Assert.StartsWith("1 worklog ·", JoinHeadline(WorkbookSummary.BuildHeadlineStats(totals, "DKK")));
    }

    // ###########################################################################################
    // The headline splits into parts so the UI can render the NUMBERS bold and the words around
    // them plain - which one string cannot express. The split lives here, beside the code that
    // produced the numbers, rather than being re-derived in the UI by hunting for digits: "0.5 h"
    // and any number inside a workbook title would both have to be guessed about.
    // ###########################################################################################
    [Fact]
    public void The_headline_parts_separate_each_number_from_its_words()
    {
        var totals = WorkbookSummary.Summarize(new[] { Entry(1, hours: 2.5, cost: 100) });

        var stats = WorkbookSummary.BuildHeadlineStats(totals, "DKK");

        // A DURATION carrying both halves contributes TWO bold numbers - "2 hours and 30
        // minutes" - which is why the headline is built as a list rather than a fixed four-part
        // array.
        Assert.Equal(new[] { "1", "2", "30", "100", "1" }, stats.Select(s => s.Number));

        // The currency code rides in the cost's SUFFIX, not in its Number - the strip bolds every
        // Number and leaves the suffixes plain, so a code in the number would come out bold while
        // the words beside it did not.
        Assert.Equal(new[] { " worklog", " hours and ", " minutes", " DKK", " open" },
            stats.Select(s => s.Suffix));

        // Only the duration's SECOND half is joined to the one before it, so a renderer's " . "
        // separator never lands inside "2 hours and 30 minutes" and splits one figure into two.
        Assert.Equal(new[] { false, false, true, false, false }, stats.Select(s => s.JoinedToPrevious));

        // Those same parts, laid out the way both renderers lay them out, read as one sentence -
        // which is what pins the Prefix/Number/Suffix split above to something a reader recognises.
        Assert.Equal("1 worklog · 2 hours and 30 minutes · 100 DKK · 1 open", JoinHeadline(stats));
    }

    // ###########################################################################################
    // The category and state counts come back as (label, count) pairs, which the tab renders as
    // its own non-selectable pills. Every value is present INCLUDING the zeroes - a "0 Issue" pill
    // says this workbook records no faults, and dropping the empty ones would both hide that and
    // make the row change width as a workbook is worked on.
    // ###########################################################################################
    [Fact]
    public void The_category_and_state_counts_come_back_as_pairs_including_the_zeroes()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            Entry(1, category: "Note", state: "Open"),
            Entry(2, category: "Note", state: "Closed")
        });

        Assert.Equal(
            new[] { ("Note", 2), ("Cosmetic", 0), ("Issue", 0) },
            WorkbookSummary.BuildCategoryCounts(totals));

        Assert.Equal(
            new[] { ("Open", 1), ("Closed", 1) },
            WorkbookSummary.BuildStateCounts(totals));
    }

    // The flat breakdown strings are the PDF's - it has no pills to put a count in and wants one
    // plain sentence per line. Count first, matching the pills the tab draws from the same pairs.
    [Fact]
    public void The_flat_breakdowns_name_every_category_and_state_in_order()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            Entry(1, category: "Issue", state: "Closed"),
            Entry(2, category: "Issue", state: "Open")
        });

        Assert.Equal("0 Note · 0 Cosmetic · 2 Issue", WorkbookSummary.FormatCategoryBreakdown(totals));
        Assert.Equal("1 Open · 1 Closed", WorkbookSummary.FormatStateBreakdown(totals));
    }

    [Fact]
    public void The_attachment_breakdown_names_each_kind_with_its_count()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            Entry(1, hours: 1, comments: 2, links: 1, photos: 1, files: 0)
        });

        string text = WorkbookSummary.FormatAttachmentBreakdown(totals);

        Assert.Contains("2 comments", text);
        Assert.Contains("1 link", text);
        Assert.Contains("1 photo", text);
        Assert.Contains("0 files", text);
        Assert.Contains("1 work done", text);
    }

    // ###########################################################################################
    // Decimal hours and costs that do not sum exactly in binary floating point still PRINT
    // correctly.
    //
    // The totals are accumulated as doubles, so three 0.1 h rows really do sum to
    // 0.30000000000000004 and twenty 1.15 h rows to 22.999999999999993. That raw drift is real but
    // it never reaches a reader: every number in the summary and in the PDF goes through the
    // "0.##" format, which rounds to two decimals and absorbs an error many orders of magnitude
    // larger than anything reachable here.
    //
    // Pinned rather than "fixed" by rounding each addition, because the displayed figures are
    // already right and the same Totals object feeds both the on-screen strip and the exported
    // document - so what has to hold is that the two agree and that neither prints a long tail of
    // digits. These are the cases that would show it if it ever stopped holding.
    // ###########################################################################################
    [Fact]
    public void Repeated_decimal_hours_print_without_a_floating_point_tail()
    {
        var entries = Enumerable.Range(1, 3).Select(i => Entry(i, hours: 0.1)).ToList();

        var totals = WorkbookSummary.Summarize(entries);

        // Asserted on the RAW total rather than through the headline: the headline now says the
        // time in whole minutes, which would absorb the very tail this test exists to notice. That
        // rounding is a second layer of protection, not a replacement for the "0.##" one.
        Assert.Equal("0.3", totals.TotalHours.ToString("0.##", CultureInfo.InvariantCulture));

        // And what the reader actually sees is 18 minutes, with no tail of its own.
        Assert.Equal("18 minutes", WorklogDurationFormatter.Format(totals.TotalHours));
    }

    [Fact]
    public void Many_decimal_hours_rows_still_print_a_clean_total()
    {
        var entries = Enumerable.Range(1, 20).Select(i => Entry(i, hours: 1.15)).ToList();

        var totals = WorkbookSummary.Summarize(entries);

        Assert.Equal("23", totals.TotalHours.ToString("0.##", CultureInfo.InvariantCulture));
        Assert.Equal("23 hours", WorklogDurationFormatter.Format(totals.TotalHours));
    }

    // Cost is the same arithmetic on money, and matters more: a customer reads this figure.
    [Fact]
    public void Repeated_decimal_costs_print_without_a_floating_point_tail()
    {
        var entries = Enumerable.Range(1, 3).Select(i => Entry(i, cost: 0.1)).ToList();

        var totals = WorkbookSummary.Summarize(entries);
        // Found by its currency suffix rather than by index: these entries log no time at all, so
        // the duration contributes no part and the cost is not at a fixed position.
        var cost = WorkbookSummary.BuildHeadlineStats(totals, "USD").Single(s => s.Suffix == " USD");

        Assert.Equal("0.3", cost.Number);
    }
    // ---------------------------------------------------------------------------------------------
    // The summary strip's COMPONENT stats, which render side by side as
    // "5 components in scope · 0 components completed".
    //
    // The second one used to read just "0 completed", which said nothing about WHAT was completed -
    // beside a line already listing comments, links, photos, files and work-done rows, it read as
    // though it could be counting any of them.
    // ---------------------------------------------------------------------------------------------

    private static string Render(WorkbookSummary.Stat stat) => stat.Prefix + stat.Number + stat.Suffix;

    [Fact]
    public void The_completed_component_stat_says_what_it_counts()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            new WorklogEntryRecord
            {
                Id = 1,
                ComponentLabels = new List<string> { "U1", "U2", "C3" },
                CompletedComponentLabels = new List<string>()
            }
        });

        var stats = WorkbookSummary.BuildComponentStats(totals);

        Assert.Equal("3 components in scope", Render(stats[0]));
        Assert.Equal("0 components completed", Render(stats[1]));
    }

    // Through the same CountStat every other count uses, so the singular reads correctly rather than
    // saying "1 components completed".
    [Fact]
    public void One_completed_component_reads_in_the_singular()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            new WorklogEntryRecord
            {
                Id = 1,
                ComponentLabels = new List<string> { "U1" },
                CompletedComponentLabels = new List<string> { "U1" }
            }
        });

        var stats = WorkbookSummary.BuildComponentStats(totals);

        Assert.Equal("1 component in scope", Render(stats[0]));
        Assert.Equal("1 component completed", Render(stats[1]));
    }

    // Only the NUMBER is bold on screen, which is why these are Stat parts rather than finished
    // strings - the count has to stay in its own field for the UI to weight it differently.
    [Fact]
    public void The_component_stats_keep_their_number_separate_from_their_words()
    {
        var totals = WorkbookSummary.Summarize(new[]
        {
            new WorklogEntryRecord
            {
                Id = 1,
                ComponentLabels = new List<string> { "U1", "U2" },
                CompletedComponentLabels = new List<string> { "U1" }
            }
        });

        var stats = WorkbookSummary.BuildComponentStats(totals);

        Assert.Equal("2", stats[0].Number);
        Assert.Equal(" components in scope", stats[0].Suffix);

        // One completed, so the singular - the count and its words are separate fields, and the
        // words follow the count.
        Assert.Equal("1", stats[1].Number);
        Assert.Equal(" component completed", stats[1].Suffix);
    }

}
