using Avalonia;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The pure parts of a worklog ENTRY that both tabs that render one need: which components its
    // marked area covers, and what its Work done rows add up to.
    //
    // Both were written out twice - once in TabSchematics.Worklog.cs and once in
    // TabWorkbooks.BoardPreviews.cs / WorklogEntryEditorWindow - and neither touches an Avalonia
    // control, so per CLAUDE.md they do not belong on a UserControl at all. The component-scope
    // computation in particular decides whether the "Mark components in scope" checklist appears in
    // the editor modal, and the two tabs are REQUIRED to open the same modal: a fix applied to one
    // copy and missed on the other is exactly the divergence already reported once.
    //
    // Avalonia's Rect is a plain value type with no display behind it, so this tests headlessly.
    // ###########################################################################################
    public static class WorklogEntryScope
    {
        // ###########################################################################################
        // The components an entry's marked area touches, as (BoardLabel, DisplayName) pairs for the
        // editor's "Mark components in scope" / "Mark components completed" checklist.
        //
        // NULL and EMPTY mean different things to the caller and must stay distinct: empty is "this
        // area genuinely covers no component", null is "the scope is unknown". Only null leaves the
        // entry's saved ComponentLabels untouched - returning an empty list for an unknown scope
        // would wipe the user's selection the first time they saved.
        //
        // Unknown means: no board data, no highlight-rect cache at all, or a cache with no entry for
        // this entry's own schematic. That last case is reachable on a region switch as well as
        // during a board load, since HighlightRectBuilder skips highlights failing IsVisibleByRegion.
        // ###########################################################################################
        public static List<(string BoardLabel, string DisplayName)>? BuildComponentsInScope(
            BoardData? boardData,
            IReadOnlyDictionary<string, Dictionary<string, List<Rect>>>? highlightRectsBySchematicAndLabel,
            WorklogEntryRecord entry)
        {
            if (boardData == null || highlightRectsBySchematicAndLabel == null || entry == null)
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(entry.SchematicName) ||
                !highlightRectsBySchematicAndLabel.TryGetValue(entry.SchematicName, out var rectsByLabel))
            {
                return null;
            }

            var area = new Rect(entry.AreaX, entry.AreaY, entry.AreaWidth, entry.AreaHeight);
            var touchedLabels = RectGeometry.FindKeysWithRectsIntersecting(rectsByLabel, area);

            return ComponentListBuilder.BuildComponentsInScope(boardData, touchedLabels)
                .Select(c => (c.BoardLabel, c.DisplayName))
                .ToList();
        }

        // ###########################################################################################
        // What an entry's "Work done" rows add up to: total hours and total cost.
        //
        // Both the full editor's own summary line and the Workbooks tab's entry-detail card show
        // these, and the two had the sums written out separately with a comment on one saying the
        // formatting had to match the other by hand. One place, and WorklogManagerTests can reach it.
        //
        // Zero for an entry with no rows, which is what both callers want to render - not a blank.
        // ###########################################################################################
        public static (double TotalHours, double TotalCost) GetWorkDoneTotals(WorklogEntryRecord? entry)
        {
            if (entry == null || entry.WorkDoneItems.Count == 0)
            {
                return (0.0, 0.0);
            }

            return (entry.WorkDoneItems.Sum(w => w.HoursSpent), entry.WorkDoneItems.Sum(w => w.Cost));
        }

        // ###########################################################################################
        // A workbook date as the app shows it: "2026-September-6".
        //
        // THE DAY CARRIES NO LEADING ZERO ("-d", not "-dd"), which is what this exists to pin. It is
        // used by the worklog bar above the tabs AND by the workbook cards in the Workbooks tab, and
        // those two must not disagree about the format - they had a copy each, under comments
        // asserting they matched.
        //
        // Invariant rather than the current culture on purpose: the month is always a NAME, never an
        // ambiguous number, so a date cannot be read as 06-09 or 09-06 depending on where the user
        // is. That also means the format string is a CUSTOM one - a lone "d" would be the standard
        // short-date specifier, but "yyyy-MMMM-d" has more than one character, so the "d" means
        // "day, unpadded" as intended.
        // ###########################################################################################
        public static string FormatWorkbookDate(DateTime value) =>
            value.ToString("yyyy-MMMM-d", CultureInfo.InvariantCulture);

        // ###########################################################################################
        // A workbook's date line as the app shows it: "started 2026-September-6" for an open
        // workbook, "ended 2026-September-6" for a closed one that recorded a finish date.
        //
        // THE WORD AND THE DATE ARE CHOSEN TOGETHER, WHICH IS WHY THIS IS ONE FUNCTION. The worklog
        // bar above the tabs and the workbook cards in the Workbooks tab both show this line for
        // the same workbook, often at the same time and inches apart - the bar said "ended
        // 2026-September-6" while the card underneath it said "started 2026-September-6" for that
        // one workbook, because the card only ever wrote "started" and only ever read StartDate.
        // Splitting the rule back into a label helper and a date helper would let exactly that
        // recur: a caller can pair the right word with the wrong date and nothing complains.
        //
        // "ended", NOT "closed": these are the two ends of one span, and "Closed" is already the
        // workbook's STATUS, showing in a pill beside this text in both places. Leave the status
        // word to say the status.
        //
        // A closed workbook with no EndDate falls back to "started" AND its start date - both
        // halves, together. That is every workbook closed before EndDate existed (there is
        // deliberately no migration - see WorkbookRecord.EndDate); reporting the start it really
        // has beats inventing a finish date, and beats showing nothing.
        // ###########################################################################################
        public static string FormatWorkbookDateLine(string? status, DateTime startDate, DateTime? endDate)
        {
            bool hasEnded = !WorklogManager.IsWorkbookStatusOpen(status) && endDate != null;

            return hasEnded
                ? $"ended {FormatWorkbookDate(endDate!.Value)}"
                : $"started {FormatWorkbookDate(startDate)}";
        }

        // ###########################################################################################
        // "0 comments" / "1 comment" / "2 comments" - a count with the right singular or plural word.
        //
        // There were three private copies of this (TabWorkbooks, WorklogEntryEditorWindow,
        // ContributionPackaging) plus several open-coded ternaries, and one site had no singular
        // branch at all and rendered a plain "1 entries".
        // ###########################################################################################
        public static string FormatCount(int count, string singular, string plural) =>
            $"{count} {(count == 1 ? singular : plural)}";
    }
}
