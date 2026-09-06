using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What one workbook adds up to, across every entry in it: hours, cost, how many entries are in
    // each category and each state, and how many comments/links/photos/files/components the whole
    // workbook carries.
    //
    // The Workbooks tab shows these in the collapsible strip under its top-line, and the PDF/ZIP
    // export prints them as its opening section - two surfaces, one set of numbers, so an exported
    // document can never disagree with what the app showed when it was produced.
    //
    // Pure: takes the records it is given and returns a value. No file reads, no controls - the
    // caller has already read the entries (the tab from its within-pass cache, the exporter from
    // WorklogManager.GetEntries), and doing it again here would both slow the tab down and let the
    // two surfaces read different states of the same folder.
    // ###########################################################################################
    public static class WorkbookSummary
    {
        // The three categories and two states are fixed vocabulary (WorklogManager's model, as
        // opposed to the four-category one the old mockup drew), and the summary reports a count
        // for each of them INCLUDING the zeroes: "Issue 0" is information - it says this workbook
        // records no faults - whereas a row that silently omits the empty categories makes the
        // reader work out which are missing.
        public static readonly IReadOnlyList<string> Categories = new[] { "Note", "Cosmetic", "Issue" };

        public static readonly IReadOnlyList<string> States = new[] { "Open", "Closed" };

        public sealed class Totals
        {
            public int WorklogCount { get; init; }

            public double TotalHours { get; init; }

            public double TotalCost { get; init; }

            public int CommentCount { get; init; }

            public int LinkCount { get; init; }

            public int PhotoCount { get; init; }

            public int FileCount { get; init; }

            public int WorkDoneCount { get; init; }

            // How many entries carry at least one component in scope, and how many DISTINCT
            // components the workbook touches across all of its entries. Distinct rather than a
            // running sum: the same chip legitimately appears in several entries (checked in one,
            // replaced in another), and adding those up would report more components than the board
            // has - which reads as a data error to anyone who knows the board.
            public int ComponentCount { get; init; }

            public int CompletedComponentCount { get; init; }

            // Keyed by the vocabulary above, every key always present - see Categories/States.
            public IReadOnlyDictionary<string, int> EntriesByCategory { get; init; } =
                new Dictionary<string, int>();

            public IReadOnlyDictionary<string, int> EntriesByState { get; init; } =
                new Dictionary<string, int>();

            public int OpenWorklogCount =>
                this.EntriesByState.TryGetValue("Open", out int open) ? open : 0;

            public int ClosedWorklogCount =>
                this.EntriesByState.TryGetValue("Closed", out int closed) ? closed : 0;
        }

        // ###########################################################################################
        // Sums a workbook's worklogs. A null or empty list gives an all-zero Totals rather than null,
        // because every caller renders the result and a workbook with no worklogs yet is an ordinary
        // state, not a failure - "0 worklogs · 0 DKK · 0 open" is the correct thing to show for it.
        // (The hours contribute no part at all at zero - see BuildHeadlineStats.)
        // ###########################################################################################
        public static Totals Summarize(IEnumerable<WorklogEntryRecord>? entries)
        {
            var list = entries?.ToList() ?? new List<WorklogEntryRecord>();

            var byCategory = Categories.ToDictionary(c => c, _ => 0, StringComparer.OrdinalIgnoreCase);
            var byState = States.ToDictionary(s => s, _ => 0, StringComparer.OrdinalIgnoreCase);

            // Component labels are the user's own text from the board data, and the same component
            // reached from two entries must count once - hence a set, case-insensitively, matching
            // how every other component-label comparison in this app is done.
            var components = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var completedComponents = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            double totalHours = 0.0;
            double totalCost = 0.0;
            int comments = 0, links = 0, photos = 0, files = 0, workDone = 0;

            foreach (var entry in list)
            {
                if (entry == null)
                    continue;

                var (hours, cost) = WorklogEntryScope.GetWorkDoneTotals(entry);
                totalHours += hours;
                totalCost += cost;

                comments += entry.Comments.Count;
                links += entry.Links.Count;
                photos += entry.Photos.Count;
                files += entry.Files.Count;
                workDone += entry.WorkDoneItems.Count;

                foreach (var label in entry.ComponentLabels.Where(l => !string.IsNullOrWhiteSpace(l)))
                    components.Add(label.Trim());

                foreach (var label in entry.CompletedComponentLabels.Where(l => !string.IsNullOrWhiteSpace(l)))
                    completedComponents.Add(label.Trim());

                // An unrecognised category or state - a value from a future build, or an entry
                // edited by hand - is counted nowhere rather than being forced into one of the
                // known buckets. Miscounting it as "Note" would quietly misreport the workbook;
                // WorklogCount below still counts the entry itself, so nothing goes missing.
                if (byCategory.ContainsKey(entry.Category ?? string.Empty))
                    byCategory[entry.Category!]++;

                if (byState.ContainsKey(entry.State ?? string.Empty))
                    byState[entry.State!]++;
            }

            return new Totals
            {
                WorklogCount = list.Count,
                TotalHours = totalHours,
                TotalCost = totalCost,
                CommentCount = comments,
                LinkCount = links,
                PhotoCount = photos,
                FileCount = files,
                WorkDoneCount = workDone,
                ComponentCount = components.Count,
                CompletedComponentCount = completedComponents.Count,
                EntriesByCategory = byCategory,
                EntriesByState = byState
            };
        }

        // ###########################################################################################
        // ONE PIECE of a summary line: a number and the words around it.
        //
        // Exists because the numbers are rendered BOLD and the words are not, which a single string
        // cannot express - the UI needs to know where each number starts and ends. Splitting it
        // here rather than re-finding the digits in the UI keeps the "which part is the number"
        // decision beside the code that produced it; a regex over a finished string would have to
        // guess about "0.5 h" and about a workbook id inside a title.
        //
        // Prefix is what comes before the number (usually empty), Suffix the unit or noun after it.
        // ###########################################################################################
        // JoinedToPrevious says this part continues the one before it rather than being a new
        // stat, so a renderer must NOT put its " - " separator in front of it. It exists for
        // DURATIONS, which are the one stat carrying TWO bold numbers: "1 hour and 15 minutes" is
        // built as ("1", " hour and ") + ("15", " minutes"), and a separator between the two halves
        // would read as two separate figures. Everything else leaves it false.
        public readonly record struct Stat(string Prefix, string Number, string Suffix, bool JoinedToPrevious = false);

        // ###########################################################################################
        // The always-visible headline, as its parts: worklogs, hours, cost, and how many are open.
        //
        // "WORKLOGS", not "entries" - the app calls these worklogs everywhere the user can see
        // (the worklog bar, "Add worklog", the workbook cards), and "entry" was internal vocabulary
        // leaking out through this one line.
        //
        // Hours and cost are formatted exactly as WorklogEntryEditorWindow's own SummaryText and the
        // entry detail card's stats row do - "{0:0.##}" in InvariantCulture - and the cost carries
        // the user's chosen currency CODE as its suffix, so the figure says what it is rather than
        // being a bare number the reader has to guess at. It used to be bare because the app never
        // asked; the Configuration tab asks now (see WorklogCurrency).
        //
        // The code is a PARAMETER rather than a read of UserSettings, because this class is pure -
        // both callers (the Workbooks tab's strip and the PDF export) already hold it, and reading
        // a static setting here would make every test of these numbers depend on the user's
        // settings file. It rides in the SUFFIX so the currency stays unbolded alongside the words,
        // while the number keeps the bold the rest of the strip gives it.
        // ###########################################################################################
        public static IReadOnlyList<Stat> BuildHeadlineStats(Totals totals, string? currencyCode)
        {
            var stats = new List<Stat>
            {
                new Stat(string.Empty, totals.WorklogCount.ToString(CultureInfo.InvariantCulture),
                    totals.WorklogCount == 1 ? " worklog" : " worklogs")
            };

            // The time as WORDS - "45 minutes", "1 hour and 15 minutes" - not the decimal hours it
            // is stored as. See WorklogDurationFormatter for why; the short of it is that decimal
            // hours is a storage format, and "1.4 h" is read as an hour and forty by everyone who
            // has not just done the arithmetic. It contributes NO stat at all when the workbook has
            // no time logged, which is why the headline is built as a list rather than a fixed
            // array: a zero here would be the only "0" in a line of real figures.
            stats.AddRange(WorklogDurationFormatter.BuildStats(totals.TotalHours));

            // The cost is dropped when there is none, exactly as the time above is - "1 worklog .
            // 0 USD . 1 open" spends a column reporting the absence of a figure, which the absence
            // itself says better. Asked for directly.
            if (WorklogCurrency.FormatCostOrEmpty(totals.TotalCost, currencyCode).Length > 0)
            {
                stats.Add(new Stat(string.Empty, totals.TotalCost.ToString("0.##", CultureInfo.InvariantCulture),
                    " " + WorklogCurrency.NormalizeCode(currencyCode)));
            }

            // "open" STAYS at zero, unlike the two above. It is not a total that is missing - it is
            // a state count over worklogs that exist, and "0 open" on a finished job says something
            // a reader wants ("everything here is closed"), where "0 USD" says only that nobody has
            // billed anything yet.
            stats.Add(new Stat(string.Empty, totals.OpenWorklogCount.ToString(CultureInfo.InvariantCulture), " open"));

            return stats;
        }

        // The attachment counts as parts, so each number can be bolded. Same content as
        // FormatAttachmentBreakdown below, which the PDF still uses as plain text.
        public static IReadOnlyList<Stat> BuildAttachmentStats(Totals totals) => new[]
        {
            CountStat(totals.CommentCount, "comment", "comments"),
            CountStat(totals.LinkCount, "link", "links"),
            CountStat(totals.PhotoCount, "photo", "photos"),
            CountStat(totals.FileCount, "file", "files"),
            new Stat(string.Empty, totals.WorkDoneCount.ToString(CultureInfo.InvariantCulture), " work done")
        };

        public static IReadOnlyList<Stat> BuildComponentStats(Totals totals) => new[]
        {
            CountStat(totals.ComponentCount, "component in scope", "components in scope"),

            // "components completed", not the bare "completed" this used to say. The two stats are
            // rendered side by side as "5 components in scope · 0 completed", where the second one
            // read as though it could be counting anything - worklogs, work-done rows - rather than
            // components. Through CountStat like every other count here, so "1 component completed"
            // reads correctly too.
            CountStat(totals.CompletedComponentCount, "component completed", "components completed")
        };

        private static Stat CountStat(int count, string singular, string plural) =>
            new(string.Empty, count.ToString(CultureInfo.InvariantCulture), " " + (count == 1 ? singular : plural));

        // ###########################################################################################
        // The category and state counts, as (label, count) pairs for the UI to render as pills.
        //
        // Every category and every state is returned INCLUDING the zeroes - see Categories. A pill
        // reading "0 Issue" is information: it says this workbook records no faults. Dropping the
        // empty ones would leave the reader working out which are missing, and would make the row
        // change width as a workbook is worked on.
        // ###########################################################################################
        public static IReadOnlyList<(string Label, int Count)> BuildCategoryCounts(Totals totals) =>
            Categories.Select(c => (c, totals.EntriesByCategory.TryGetValue(c, out int n) ? n : 0)).ToList();

        public static IReadOnlyList<(string Label, int Count)> BuildStateCounts(Totals totals) =>
            States.Select(s => (s, totals.EntriesByState.TryGetValue(s, out int n) ? n : 0)).ToList();

        // ###########################################################################################
        // The same lines as flat strings, for the PDF export - which has no pills and no bold runs
        // to put a number in, and wants one plain sentence per line.
        //
        // These share the Stat/count builders above rather than formatting the numbers a second
        // time, so the document and the screen cannot disagree about what a workbook adds up to.
        // ###########################################################################################
        // THERE IS DELIBERATELY NO FormatHeadline HERE. The headline is the one line where BOTH
        // renderers - the summary strip and the PDF - bold the numbers and leave the words plain, so
        // both walk BuildHeadlineStats' parts directly and neither ever wanted a finished string.
        // One existed anyway, joined the parts back together, and had no caller outside the tests
        // that asserted on it - a public API shipping nothing, maintained through every future
        // change to the currency and duration formats for no user-visible effect. The parts are what
        // is tested now, which is what both renderers actually consume.

        public static string FormatCategoryBreakdown(Totals totals) =>
            string.Join(" · ", BuildCategoryCounts(totals).Select(c => $"{c.Count} {c.Label}"));

        public static string FormatStateBreakdown(Totals totals) =>
            string.Join(" · ", BuildStateCounts(totals).Select(s => $"{s.Count} {s.Label}"));

        public static string FormatAttachmentBreakdown(Totals totals) =>
            string.Join(" · ", BuildAttachmentStats(totals).Select(FormatStat));

        public static string FormatComponentBreakdown(Totals totals) =>
            string.Join(" · ", BuildComponentStats(totals).Select(FormatStat));

        private static string FormatStat(Stat stat) => stat.Prefix + stat.Number + stat.Suffix;
    }
}
