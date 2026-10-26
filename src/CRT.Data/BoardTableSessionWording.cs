using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THE DRAFTS TAB'S TABLE SAYS ABOUT A CHANGE IT HANDED THE DRAFT (BoardTableEditor's draft
    // session part, 2026-10-25).
    // ###########################################################################################
    public static class BoardTableSessionWording
    {
        // The draft history's name for one change made in the table ("Undo: Edit in the table, Components").
        public static string Step(string? sheetName) =>
            string.IsNullOrWhiteSpace(sheetName) ? "Edit in the table" : $"Edit in the table, {sheetName}";

        public static string Undone(string description) => $"Undone: {description}.";

        public static string Redone(string description) => $"Redone: {description}.";

        // ###########################################################################################
        // New components stay where they were typed while the table is open, and are saved into
        // their category in label order (BoardTableSheet.BuildRows) - said once, so the order the
        // table shows next time is no surprise.
        // ###########################################################################################
        public const string NewComponentsPlaced =
            "New components are saved into their category, in label order - the table shows them there the next time it is opened.";

        // ###########################################################################################
        // What was drawn on schematic images the change took off the board - their highlights and
        // KiCad calibrations, which go with them (DraftEditRules). There was a question before a SAVE
        // that would do this (code review, 2026-10-11); with no save to ask before, it is said once
        // done, with the way back. Null when nothing was drawn on them.
        // ###########################################################################################
        public static string? WentWith(DraftEditLoss loss)
        {
            ArgumentNullException.ThrowIfNull(loss);

            if (loss.Schematics.Count == 0 || (loss.Highlights == 0 && loss.Calibrations == 0))
            {
                return null;
            }

            var parts = new List<string>();

            if (loss.Highlights > 0)
            {
                parts.Add(loss.Highlights == 1 ? "1 highlight" : $"{loss.Highlights.ToString(CultureInfo.InvariantCulture)} highlights");
            }

            if (loss.Calibrations > 0)
            {
                parts.Add(loss.Schematics.Count == 1
                    ? "its KiCad calibration"
                    : loss.Calibrations == 1 ? "1 KiCad calibration" : $"{loss.Calibrations.ToString(CultureInfo.InvariantCulture)} KiCad calibrations");
            }

            string names = loss.Schematics.Count == 1
                ? $"[{loss.Schematics[0]}]"
                : string.Join(", ", loss.Schematics.Take(loss.Schematics.Count - 1).Select(name => $"[{name}]")) + $" and [{loss.Schematics[^1]}]";

            string with = loss.Schematics.Count == 1 ? "it" : "them";

            return $"What was drawn on {names} went with {with}: {string.Join(" and ", parts)}. Ctrl+Z brings it back.";
        }
    }
}
