using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Ranks a workbook's entries for the "attach this capture to a worklog" picker, and formats
    // the label each one shows.
    //
    // The picker exists because a capture taken from the component popup (an oscilloscope image,
    // and later an IC test result) knows its WORKBOOK - ResolveActiveWorkbook settles that app-wide
    // - but not which ENTRY it belongs to. Asking the user is the only way to know, so the whole
    // job here is making the right answer the first one in the list.
    //
    // The ranking rule is deliberately just TWO levels:
    //
    //  1. Entries whose ComponentLabels contain the component being measured, by ascending id.
    //     Someone probing U8 while working a fault on U8 is the overwhelmingly common case, and
    //     this is what turns the picker into a single click for them.
    //  2. Everything else, by ascending id.
    //
    // Ascending id - #1, #2, #3 - because this renders as a plain dropdown list, and a LIST has to
    // look ordered or it looks broken. An earlier version sorted by newest-first within an
    // open-before-closed band, which is defensible per-criterion and produced "#2, #4, #3, #1" on
    // screen: three invisible criteria interleaving into what reads as no order at all. Reported as
    // exactly that. Entry ids are also what the user sees on the board pills ("#4"), so counting
    // order is the one ordering they can already follow.
    //
    // Open/closed is deliberately NOT a sort level any more, for the same reason. Closed entries
    // are still kept rather than hidden: re-measuring a repair you already finished is exactly what
    // someone does when a board comes back, and a picker that silently omits the entry describing
    // that repair sends the measurement to a new one instead.
    //
    // The component match survives as a level because it is the one that pays for itself - it puts
    // the right answer in the preselected slot - and because the dialog SAYS it is doing it
    // ("Worklogs covering [U8] are listed first"), so those rows leading is explained rather than
    // mysterious.
    //
    // Component matching is case-insensitive and trimmed, matching every other BoardLabel
    // comparison in the app (see ComponentImageQueries and BoardDataReader). A blank component
    // simply produces no matches rather than matching everything, so a capture whose component
    // cannot be resolved still gets a usable open-entries-first list.
    //
    // Pure - no Avalonia, no file access - so it is unit tested; the UI half is
    // WorklogAttachCaptureWindow.
    // ###########################################################################################
    public static class WorklogAttachTargets
    {
        // A single row in the picker: the entry it stands for, plus whether it earned its place
        // through the component match. The flag is what lets the UI mark those rows, rather than
        // the user having to guess why the order is what it is.
        public sealed record AttachTarget(WorklogEntryRecord Entry, bool IsComponentMatch);

        // ###########################################################################################
        // Orders the given entries for the picker. componentLabel is the board label of the
        // component the capture was taken on ("U8"), or blank when it is not known.
        //
        // A null entry list yields an empty result rather than throwing: the caller reads this
        // straight off GetEntries for a workbook that may legitimately have none yet.
        // ###########################################################################################
        public static List<AttachTarget> Rank(
            IReadOnlyList<WorklogEntryRecord>? entries,
            string? componentLabel)
        {
            if (entries == null || entries.Count == 0)
            {
                return new List<AttachTarget>();
            }

            string component = componentLabel?.Trim() ?? string.Empty;

            return entries
                .Where(entry => entry != null)
                .Select(entry => new AttachTarget(entry, EntryScopesComponent(entry, component)))
                .OrderByDescending(target => target.IsComponentMatch)
                .ThenBy(target => target.Entry.Id)
                .ToList();
        }

        // ###########################################################################################
        // Whether an entry names the given component among the components it scopes.
        //
        // Compared against ComponentLabels only - NOT CompletedComponentLabels, which is a subset of
        // it recording which of the scoped components have been dealt with. An entry that scopes U8
        // is a candidate whether or not U8 has been ticked off.
        // ###########################################################################################
        private static bool EntryScopesComponent(WorklogEntryRecord entry, string component)
        {
            if (component.Length == 0 || entry.ComponentLabels == null)
            {
                return false;
            }

            return entry.ComponentLabels.Any(label =>
                string.Equals(label?.Trim(), component, StringComparison.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // The label one picker row shows: "#7 - U8 gives no video".
        //
        // An entry with a blank title falls back to naming its id alone, since a bare "#7 - " with
        // nothing after it reads as a rendering fault. Title is a plain string in entries.json, so a
        // hand-edited or older-build record can carry an empty one - the same reason the PDF
        // exporter's PillLabel substitutes a fallback.
        // ###########################################################################################
        public static string FormatLabel(WorklogEntryRecord? entry)
        {
            if (entry == null)
            {
                return string.Empty;
            }

            string id = "#" + entry.Id.ToString(CultureInfo.InvariantCulture);
            string title = entry.Title?.Trim() ?? string.Empty;

            return title.Length == 0 ? id : $"{id} - {title}";
        }

        // ###########################################################################################
        // The read-only line naming which workbook a capture is about to be filed into:
        // "#3 - Dave's C64, breadbin".
        //
        // Shown because this dialog is opened from the component popup, which can be sitting over a
        // schematic while the user's attention has been on the oscilloscope - the worklog bar that
        // normally makes the active workbook obvious is not what they are looking at. Same blank
        // -title fallback and same shape as FormatLabel, so the two lines read as one family.
        // ###########################################################################################
        public static string FormatWorkbookLabel(WorkbookRecord? workbook)
        {
            if (workbook == null)
            {
                return string.Empty;
            }

            string id = "#" + workbook.Id.ToString(CultureInfo.InvariantCulture);
            string title = workbook.Title?.Trim() ?? string.Empty;

            return title.Length == 0 ? id : $"{id} - {title}";
        }
    }
}
