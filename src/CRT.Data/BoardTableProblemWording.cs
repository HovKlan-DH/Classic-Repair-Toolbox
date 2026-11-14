using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THE TABLE SAYS ABOUT THE CHECKS' PROBLEMS, apart from the cells (owner request,
    // 2026-10-02). Pure, so the words are tested without a window.
    // ###########################################################################################
    public static class BoardTableProblemWording
    {
        // How many problems outside the sheets are spelled out before "and N more".
        public const int OutsideSheetsShown = 3;

        // ###########################################################################################
        // The problems in NO sheet - highlights, which are drawn on the schematics with the label
        // editor rather than typed into a row, so no cell can carry them. One line each, the worst
        // first, at most OutsideSheetsShown of them. Null when there are none.
        // ###########################################################################################
        public static string? OutsideSheets(IReadOnlyList<BoardDataProblem> problems)
        {
            ArgumentNullException.ThrowIfNull(problems);

            if (problems.Count == 0)
                return null;

            List<string> lines =
            [
                "On the schematics (highlights are not in any sheet - fix them with \"Mark components manually\" under \"Board: Schematic images\"):"
            ];

            lines.AddRange(problems
                .OrderByDescending(problem => problem.Level)
                .Take(BoardTableProblemWording.OutsideSheetsShown)
                .Select(problem => $"{BoardTableProblemWording.LevelWord(problem.Level)}: {problem.Message}"));

            int more = problems.Count - BoardTableProblemWording.OutsideSheetsShown;

            if (more > 0)
                lines.Add($"... and {more.ToString(CultureInfo.InvariantCulture)} more.");

            return string.Join(Environment.NewLine, lines);
        }

        // How many problems in the sheets the list above the table spells out before "and N more".
        public const int InSheetsShown = 40;

        // ###########################################################################################
        // *** THE PROBLEMS IN THE SHEETS, AS TEXT (code review, 2026-10-13). *** A cell's problem
        // was readable only by hovering a 9px triangle in its corner - nothing to copy, nothing a
        // keyboard user could reach, and easy to miss on a wide sheet. While the Errors or Warnings
        // pill is picked the table lists them above itself, one line each: the sheet, the row (its
        // own key - "U8"), the column, and the message - errors first, in sheet order. Null when
        // neither pill is picked or there is nothing to list.
        // ###########################################################################################
        public static string? InSheets(IReadOnlyList<BoardDataProblem> problems, bool errors, bool warnings)
        {
            ArgumentNullException.ThrowIfNull(problems);

            if (!errors && !warnings)
                return null;

            // The order the table's tabs stand in, which is the table's own (BoardTableTabs).
            var sheetOrder = BoardTableTabs.SheetsInOrder
                .Select((sheet, index) => (sheet.SheetName, index))
                .ToDictionary(pair => pair.SheetName, pair => pair.index, StringComparer.OrdinalIgnoreCase);

            List<BoardDataProblem> listed = problems
                .Where(problem => problem.IsInSheet)
                .Where(problem => problem.Level == BoardProblemLevel.Error ? errors : warnings)
                .OrderByDescending(problem => problem.Level)
                .ThenBy(problem => sheetOrder.TryGetValue(problem.Sheet!, out int order) ? order : int.MaxValue)
                .ThenBy(problem => problem.Index)
                .ToList();

            if (listed.Count == 0)
                return null;

            List<string> lines = listed
                .Take(BoardTableProblemWording.InSheetsShown)
                .Select(problem =>
                {
                    // By the names on the tabs and the Contribute tab's column, neither of which names a sheet.
                    string sheet = $"\"{BoardTableTabs.NameOf(problem.Sheet!)}\"";

                    string where = problem.Subject is { Length: > 0 } subject
                        ? $"{sheet}, {subject}"
                        : $"{sheet}, row {(problem.Index + 1).ToString(CultureInfo.InvariantCulture)}";

                    if (!string.IsNullOrWhiteSpace(problem.Column))
                        where += $", {problem.Column}";

                    return $"{BoardTableProblemWording.LevelWord(problem.Level)} - {where}: {problem.Message}";
                })
                .ToList();

            int more = listed.Count - BoardTableProblemWording.InSheetsShown;

            if (more > 0)
                lines.Add($"... and {more.ToString(CultureInfo.InvariantCulture)} more.");

            return string.Join(Environment.NewLine, lines);
        }

        // ###########################################################################################
        // Said when Submit is pressed on a draft with errors while the table is the view: nothing is
        // sent, and the table goes to the first error (TabDrafts) - its cell the cursor's when
        // `onFirst`, which it is unless the first is on a schematic, in no row. The rows are not
        // narrowed (owner request, 2026-11-13: "use the current view as-is"); the "Errors" pill does
        // that when wanted. Every error is something the server would refuse. `hiddenByFilter`: no error
        // in the rows a pill or the search lets through - so no cell to be on, and it says so (code
        // review, 2026-11-13: it said a red corner marks a cell nobody could see).
        // ###########################################################################################
        public static string SubmitBlocked(int errorCount, bool onFirst, bool hiddenByFilter = false) =>
            BoardTableProblemWording.ErrorsToFix(errorCount) +
            (onFirst ? (errorCount == 1 ? " The cursor is on it." : " The cursor is on the first.") : string.Empty) +
            (!onFirst && hiddenByFilter
                ? " The rows the pills or the search let through hide them - pick \"Errors\" to see them."
                : " Each error is listed above the table, and its cell has a red corner.");

        // The sentence Submit stopped on errors opens with, in the "Table" view (above) and the "UI" view (CRT's
        // ContributeProblems.SubmitBlocked) alike - ONE wording (code review, 2026-11-13: it was written twice).
        public static string ErrorsToFix(int errorCount) =>
            errorCount == 1
                ? "This draft has 1 error to fix before it can be submitted."
                : $"This draft has {errorCount.ToString(CultureInfo.InvariantCulture)} errors to fix before it can be submitted.";

        public static string LevelWord(BoardProblemLevel level) =>
            level == BoardProblemLevel.Error ? "Error" : "Warning";
    }
}
