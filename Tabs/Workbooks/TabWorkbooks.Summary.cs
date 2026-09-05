using System;
using System.Collections.Generic;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Media;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Theming;

namespace CRT
{
    // ###########################################################################################
    // TabWorkbooks, part: the workbook SUMMARY STRIP under the top-line.
    //
    // One always-visible headline ("7 worklogs - 12.5 h - 430 - 4 open") with a chevron that expands
    // a breakdown by category, by state, by attachment count and by component scope. Everything it
    // shows comes from WorkbookSummary.Summarize (pure, in Handlers/, unit tested) - this file only
    // moves those strings onto controls and remembers whether the detail is open.
    //
    // WHY IT LIVES ON THE TOP-LINE and not, say, docked under the entry list: these are properties
    // of the WORKBOOK, and the workbook's identity, status, note and its Edit/Delete/Export buttons
    // are already on that row. A summary panel elsewhere would put "what this workbook is" in two
    // places on the same screen.
    //
    // WHY THE DETAIL COLLAPSES: the headline is four numbers and costs one line, which is worth it
    // on every visit. The breakdown is four more lines, which is not - so it starts collapsed and
    // the choice persists in UserSettings.WorkbooksSummaryExpanded, per user rather than per board:
    // a user who wants the detail wants it for their work, not for one particular board.
    //
    // Refreshed from ApplyHeaderForWorkbook, which is the one place the top-line is filled in, so
    // the summary cannot show one workbook's numbers under another workbook's title.
    // ###########################################################################################
    public partial class TabWorkbooks
    {
        // fa-solid chevron-right (collapsed) / chevron-down (expanded). Hex codepoints rather than
        // literal glyphs so this file stays plain ASCII, the same reason WorklogGlyphs spells its
        // padlocks out that way.
        private const int ChevronRightCodepoint = 0xF054;

        private const int ChevronDownCodepoint = 0xF078;

        // ###########################################################################################
        // Fills the summary strip for one workbook and shows it.
        //
        // Reads the entries through the within-pass cache rather than calling WorklogManager.GetEntries
        // directly: RefreshWorkbooks has usually already read this workbook's entries for the search
        // filter and the board pane, and GetEntries re-parses the whole file per call.
        // ###########################################################################################
        private void ApplySummaryForWorkbook(WorkbookRecord workbook)
        {
            var totals = WorkbookSummary.Summarize(this.GetEntriesForThisPass(workbook.Id));

            ApplyStatRuns(this.WorkbookSummaryHeadlineText, WorkbookSummary.BuildHeadlineStats(totals));
            ApplyStatRuns(this.WorkbookSummaryAttachmentsText, WorkbookSummary.BuildAttachmentStats(totals));

            // The category and state counts as PILLS rather than as a line of text - each is the
            // same non-selectable pill that category or state has everywhere else in the app, now
            // carrying its count. Asked for directly, and it also removes the odd position these
            // two lines were in: they named the app's own vocabulary in plain text while every
            // other surface drew the same words as pills.
            BuildCountPills(this.WorkbookSummaryCategoryPanel,
                WorkbookSummary.BuildCategoryCounts(totals),
                (label, count) => WorklogInfoPillBuilder.BuildCategoryChip(label, SummaryPillFontSize, count));

            BuildCountPills(this.WorkbookSummaryStatePanel,
                WorkbookSummary.BuildStateCounts(totals),
                (label, count) => WorklogInfoPillBuilder.BuildStatePill(label, SummaryPillFontSize, count));

            // Hidden outright rather than shown as "0 components in scope": a workbook of plain
            // notes scopes none, which is the common case, and a permanent zero row trains the eye
            // to skip the whole block.
            bool hasComponents = totals.ComponentCount > 0;
            this.WorkbookSummaryComponentsText.IsVisible = hasComponents;
            if (hasComponents)
            {
                ApplyStatRuns(this.WorkbookSummaryComponentsText, WorkbookSummary.BuildComponentStats(totals));
            }

            this.WorkbookSummaryPanel.IsVisible = true;

            this.ApplySummaryExpandedState(UserSettings.WorkbooksSummaryExpanded);
        }

        // The summary's pills are a touch smaller than an entry card's, which sit at the default
        // 11: this is a dense header line, and five state/category pills at full size crowd the
        // Edit/Delete/Export buttons beside them.
        private const double SummaryPillFontSize = 10.0;

        // ###########################################################################################
        // Renders one summary line with its NUMBERS BOLD and the words around them not.
        //
        // Through Inlines rather than Text, which is why WorkbookSummary hands back Stat parts
        // instead of finished strings - a TextBlock cannot mix weights within a single Text, and
        // re-finding the digits in a formatted string would have to guess about "0.5 h" and about
        // any number inside a workbook title.
        //
        // Text is cleared first: a TextBlock carrying BOTH renders the Text and silently ignores
        // the Inlines, the same trap TextLinkRenderer documents.
        // ###########################################################################################
        private static void ApplyStatRuns(TextBlock block, IReadOnlyList<WorkbookSummary.Stat> stats)
        {
            block.Text = null;
            block.Inlines?.Clear();
            block.Inlines ??= new InlineCollection();

            for (int i = 0; i < stats.Count; i++)
            {
                if (i > 0)
                    block.Inlines.Add(new Run(" · "));

                var stat = stats[i];

                if (!string.IsNullOrEmpty(stat.Prefix))
                    block.Inlines.Add(new Run(stat.Prefix));

                block.Inlines.Add(new Run(stat.Number) { FontWeight = FontWeight.Bold });

                if (!string.IsNullOrEmpty(stat.Suffix))
                    block.Inlines.Add(new Run(stat.Suffix));
            }
        }

        // Rebuilt rather than updated in place, matching how every other list on this tab refreshes:
        // the counts change with the workbook, and one build path is easier to keep right than a
        // build path plus an update path.
        private static void BuildCountPills(
            Panel panel,
            IReadOnlyList<(string Label, int Count)> counts,
            Func<string, int, Border> build)
        {
            panel.Children.Clear();

            foreach (var (label, count) in counts)
                panel.Children.Add(build(label, count));
        }

        // ###########################################################################################
        // Shows or hides the breakdown and points the chevron the right way.
        //
        // Separate from ApplySummaryForWorkbook because the toggle needs it without recomputing the
        // numbers, and ApplySummaryForWorkbook needs it to restore the saved state on every refresh
        // (the strip is rebuilt on every board change and every entry save).
        // ###########################################################################################
        private void ApplySummaryExpandedState(bool isExpanded)
        {
            this.WorkbookSummaryDetailPanel.IsVisible = isExpanded;

            int codepoint = isExpanded ? ChevronDownCodepoint : ChevronRightCodepoint;
            this.WorkbookSummaryChevron.Text = char.ConvertFromUtf32(codepoint);

            // Recomputed per state, not set once: the two chevrons overshoot the font's declared
            // ascent by different amounts, so one fixed padding clips one of them - see
            // FontAwesomeGlyphMetrics, and the same note on every other Font Awesome glyph here.
            this.WorkbookSummaryChevron.Padding = FontAwesomeGlyphMetrics
                .GetTopOverflowThickness(codepoint, this.WorkbookSummaryChevron.FontSize);
        }

        private void OnWorkbookSummaryToggleClick(object? sender, RoutedEventArgs e)
        {
            bool expanded = !UserSettings.WorkbooksSummaryExpanded;

            // Persisted BEFORE the visual update rather than after, so the state the strip shows is
            // always the state that was saved - the setter writes settings.json, and a failure
            // there must not leave the panel open claiming a preference that was not recorded.
            UserSettings.WorkbooksSummaryExpanded = expanded;

            this.ApplySummaryExpandedState(expanded);
        }

        // Lets the headless tests drive the toggle without a click, the same seam pattern
        // ApplySplitterWidthsForTests uses for this tab's splitters.
        internal void ToggleSummaryForTests() => this.OnWorkbookSummaryToggleClick(null, new RoutedEventArgs());
    }
}
