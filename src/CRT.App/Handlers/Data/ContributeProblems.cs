using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // One problem of a draft, where it belongs on the Contribute tab: the entry of the left column that is
    // about it, and the component it is of - "" for one about the board, or a component the board does not
    // have (a pin left behind), which nothing could open.
    // Text: what is said of it, the row named when its message does not (TextOf).
    // Region: what its row says under "Applies to" - "" for a row of every region (RegionToOpen).
    internal sealed record EntryProblem(BoardDataProblem Problem, ContributeEntry Entry, string ComponentLabel, string Text = "", string Region = "");

    // ###########################################################################################
    // A DRAFT'S ERRORS AND WARNINGS, WHERE THEY ARE (owner request, 2026-11-04: "move the error and warning
    // badges to where the problem is. E.g. if there are errors in 'Component: Basics' then there should be a
    // badge there and so on, and then those errors and warnings should be clearly visible in there"). They
    // were two counts in the draft's header until then, saying nothing of where.
    //
    //   - EVERY PROBLEM HAS ONE ENTRY, by the sheet and row the checks say it is on (EntryOf); what no entry
    //     is about goes on "Overview", which says it in a panel of its own.
    //   - "OVERVIEW" HAS THEM ALL, THE ENTRIES BELOW IT THEIR OWN (owner request, 2026-11-04: "the 'Overview'
    //     entry should show all errors and warnings, and then you can see the details per entry in the
    //     list below").
    //   - WITH A COMPONENT SELECTED ITS PAGES HAVE THAT ONE'S ALONE (owner request, 2026-11-04: "I do not
    //     want to see any warnings or errors for other components") - "Overview" and each component's row
    //     in the list on the far left are how the others are found. With none selected, a component's
    //     entries count every component, each problem naming the one it is about.
    //
    // PURE: what the checks found in, where it goes out.
    // ###########################################################################################
    internal static class ContributeProblems
    {
        // How many a page's panel lists in full, the rest counted - the panel must leave room for the page.
        public const int ListedAtMost = 50;

        // ###########################################################################################
        // The entry about `problem`. "Component images" is two entries, told apart as everything tells them
        // apart (ComponentImageKinds); a highlight is on no sheet, and the images' as the Overview's panels
        // have it. A component's missing category is on its Basics, where its box says so.
        // ###########################################################################################
        public static ContributeEntry EntryOf(BoardDataProblem problem, BoardData? board)
        {
            ArgumentNullException.ThrowIfNull(problem);

            return problem.Sheet switch
            {
                BoardWorkbookSchema.SheetBoardSchematics => ContributeEntry.SchematicImages,
                BoardWorkbookSchema.SheetComponents => ContributeEntry.ComponentBasics,
                BoardWorkbookSchema.SheetComponentImages =>
                    ContributeProblems.RowOf(board?.ComponentImages, problem) is { } image && !ComponentImageKinds.IsScopeBaseline(image)
                        ? ContributeEntry.ComponentPictures
                        : ContributeEntry.ComponentBaselines,
                BoardWorkbookSchema.SheetComponentLocalFiles => ContributeEntry.ComponentFiles,
                BoardWorkbookSchema.SheetComponentLinks => ContributeEntry.ComponentLinks,
                BoardWorkbookSchema.SheetComponentPins => ContributeEntry.ComponentPinout,
                BoardWorkbookSchema.SheetComponentReplacements => ContributeEntry.ComponentReplacements,
                BoardWorkbookSchema.SheetBoardLocalFiles => ContributeEntry.Files,
                BoardWorkbookSchema.SheetBoardLinks => ContributeEntry.Links,
                BoardWorkbookSchema.SheetCredits => ContributeEntry.Credits,
                BoardWorkbookSchema.SheetKiCadImportantSignals => ContributeEntry.ImportantSignals,
                _ when problem.Code is "regions.unknown" or "regions.color" => ContributeEntry.Regions,
                _ when problem.Code.StartsWith("highlight.", StringComparison.Ordinal) => ContributeEntry.SchematicImages,
                _ => ContributeEntry.Overview,
            };
        }

        // The component the row of `problem` is of, as the board's rows spell it - "" for none.
        private static string LabelOf(BoardDataProblem problem, BoardData board) =>
            (problem.Sheet switch
            {
                BoardWorkbookSchema.SheetComponents => ContributeProblems.RowOf(board.Components, problem)?.BoardLabel,
                BoardWorkbookSchema.SheetComponentImages => ContributeProblems.RowOf(board.ComponentImages, problem)?.BoardLabel,
                BoardWorkbookSchema.SheetComponentLocalFiles => ContributeProblems.RowOf(board.ComponentLocalFiles, problem)?.BoardLabel,
                BoardWorkbookSchema.SheetComponentLinks => ContributeProblems.RowOf(board.ComponentLinks, problem)?.BoardLabel,
                BoardWorkbookSchema.SheetComponentPins => ContributeProblems.RowOf(board.ComponentPins, problem)?.BoardLabel,
                BoardWorkbookSchema.SheetComponentReplacements => ContributeProblems.RowOf(board.ComponentReplacements, problem)?.BoardLabel,
                _ => null,
            })?.Trim() ?? string.Empty;

        private static T? RowOf<T>(IReadOnlyList<T>? rows, BoardDataProblem problem)
            where T : class =>
            rows is not null && problem.Index >= 0 && problem.Index < rows.Count ? rows[problem.Index] : null;

        // ###########################################################################################
        // What the row of `problem` says under "Applies to" - "" for a row of every region, or a sheet with no
        // such column. And for a component not marked: told once, on its first row, but of the component in
        // every region - a mark belongs to the label (owner decision, 2026-11-04: "it is ONE physical component
        // only"), so "PAL:" before it, and "Open" switching to PAL, said what is not so.
        // ###########################################################################################
        private static string RegionOf(BoardDataProblem problem, BoardData? board) =>
            board is null || problem.Code == "component.no_highlight"
                ? string.Empty
                : (problem.Sheet switch
                {
                    BoardWorkbookSchema.SheetComponents => ContributeProblems.RowOf(board.Components, problem)?.Region,
                    BoardWorkbookSchema.SheetComponentImages => ContributeProblems.RowOf(board.ComponentImages, problem)?.Region,
                    BoardWorkbookSchema.SheetComponentLocalFiles => ContributeProblems.RowOf(board.ComponentLocalFiles, problem)?.Region,
                    BoardWorkbookSchema.SheetComponentLinks => ContributeProblems.RowOf(board.ComponentLinks, problem)?.Region,
                    BoardWorkbookSchema.SheetComponentPins => ContributeProblems.RowOf(board.ComponentPins, problem)?.Region,
                    BoardWorkbookSchema.SheetComponentReplacements => ContributeProblems.RowOf(board.ComponentReplacements, problem)?.Region,
                    BoardWorkbookSchema.SheetBoardLocalFiles => ContributeProblems.RowOf(board.BoardLocalFiles, problem)?.Region,
                    BoardWorkbookSchema.SheetBoardLinks => ContributeProblems.RowOf(board.BoardLinks, problem)?.Region,
                    _ => null,
                })?.Trim() ?? string.Empty;

        // ###########################################################################################
        // The region CRT has to show for the row of a problem to be on its page (code review, 2026-11-04: with
        // PAL shown an NTSC baseline off its case was counted on the entry and on no page): "" with nothing to
        // switch - a row of every region, "Files" and "Links", which list every row whatever CRT shows, and
        // any other word in the cell, which shows in neither region.
        // ###########################################################################################
        // `regions`: the board's (BoardRegions.Of - names of its own since 2026-11-09); null for PAL and NTSC.
        public static string RegionToOpen(EntryProblem item, IReadOnlyList<string>? regions = null)
        {
            ArgumentNullException.ThrowIfNull(item);

            if (item.Entry is ContributeEntry.ComponentFiles or ContributeEntry.ComponentLinks)
            {
                return string.Empty;
            }

            string wanted = item.Region?.Trim() ?? string.Empty;

            return (regions ?? BoardRegions.PalAndNtsc).FirstOrDefault(region => string.Equals(region, wanted, StringComparison.OrdinalIgnoreCase)) ?? string.Empty;
        }

        // Every problem of `read` where it belongs - errors first, each kind as the checks found them.
        public static List<EntryProblem> Placed(BoardProblemsRead read)
        {
            ArgumentNullException.ThrowIfNull(read);

            var labels = new HashSet<string>(
                read.Board?.Components.Select(row => row.BoardLabel?.Trim() ?? string.Empty) ?? [],
                StringComparer.OrdinalIgnoreCase);

            return read.Problems
                .OrderByDescending(problem => problem.Level)
                .Select(problem =>
                {
                    string label = read.Board is null ? string.Empty : ContributeProblems.LabelOf(problem, read.Board);

                    return new EntryProblem(
                        problem,
                        ContributeProblems.EntryOf(problem, read.Board),
                        labels.Contains(label) ? label : string.Empty,
                        ContributeProblems.TextOf(problem, read.Board),
                        ContributeProblems.RegionOf(problem, read.Board));
                })
                .ToList();
        }

        // ###########################################################################################
        // What is said of a problem away from its cell. 'This credit has no "Name or handle"' is plain beside
        // the cell the table marks, and says nothing in a list of three (owner report, 2026-11-04: "it is
        // not at all clear for me where those errors are"), so a message naming no row gets its row's name:
        // what identifies it, as the checks' own duplicate warning goes by, or its number with nothing
        // filled in to go by.
        //
        // *** ITS REGION IS SAID EITHER WAY (code review, 2026-11-04). *** A row's name holds it, but a message
        // naming its own row - "This pin of [U7] says pin 99" - said nothing of it, with the row on no page
        // while CRT shows the other region.
        // ###########################################################################################
        public static string TextOf(BoardDataProblem problem, BoardData? board)
        {
            ArgumentNullException.ThrowIfNull(problem);

            if (board is null || !problem.IsInSheet)
            {
                return problem.Message;
            }

            if (problem.Message.Contains('[', StringComparison.Ordinal))
            {
                string region = ContributeProblems.RegionOf(problem, board);

                return region.Length > 0 && !problem.Message.Contains(region, StringComparison.OrdinalIgnoreCase)
                    ? $"{region}: {problem.Message}"
                    : problem.Message;
            }

            object? row = problem.Sheet switch
            {
                BoardWorkbookSchema.SheetBoardSchematics => ContributeProblems.RowOf(board.Schematics, problem),
                BoardWorkbookSchema.SheetComponents => ContributeProblems.RowOf(board.Components, problem),
                BoardWorkbookSchema.SheetComponentImages => ContributeProblems.RowOf(board.ComponentImages, problem),
                BoardWorkbookSchema.SheetComponentLocalFiles => ContributeProblems.RowOf(board.ComponentLocalFiles, problem),
                BoardWorkbookSchema.SheetComponentLinks => ContributeProblems.RowOf(board.ComponentLinks, problem),
                BoardWorkbookSchema.SheetComponentPins => ContributeProblems.RowOf(board.ComponentPins, problem),
                BoardWorkbookSchema.SheetComponentReplacements => ContributeProblems.RowOf(board.ComponentReplacements, problem),
                BoardWorkbookSchema.SheetBoardLocalFiles => ContributeProblems.RowOf(board.BoardLocalFiles, problem),
                BoardWorkbookSchema.SheetBoardLinks => ContributeProblems.RowOf(board.BoardLinks, problem),
                BoardWorkbookSchema.SheetCredits => ContributeProblems.RowOf(board.Credits, problem),
                BoardWorkbookSchema.SheetKiCadImportantSignals => ContributeProblems.RowOf(board.KiCadImportantSignals, problem),
                _ => null,
            };

            if (row is null)
            {
                return problem.Message;
            }

            string name = BoardDraftNaturalKeys.Describe(BoardDraftNaturalKeys.ForRow(row));

            return name.Length > 0
                ? $"[{name}] {problem.Message}"
                : $"Row {(problem.Index + 1).ToString(CultureInfo.InvariantCulture)}: {problem.Message}";
        }

        // What a page shows of what is on it: on a component's page with components selected, theirs alone - one, or the
        // several edited as one (owner request, 2026-11-08).
        public static List<EntryProblem> OnEntry(IEnumerable<EntryProblem> placed, ContributeEntry entry, IReadOnlyCollection<string> selected)
        {
            ArgumentNullException.ThrowIfNull(placed);
            ArgumentNullException.ThrowIfNull(selected);

            bool ofSelected = selected.Count > 0 && ContributeEntries.PageOf(entry) is not null;

            // One lookup per problem, never a walk of the selected (code review, 2026-11-09: hundreds selected and thousands
            // of warnings, at every refresh of the left column).
            IReadOnlySet<string> lookup = ofSelected ? ContributeProblems.SetOf(selected) : new HashSet<string>();

            return placed
                .Where(item => item.Entry == entry)
                .Where(item => !ofSelected || lookup.Contains(item.ComponentLabel))
                .ToList();
        }

        // `labels` as a set ignoring capitals - the very set when it is one already.
        private static IReadOnlySet<string> SetOf(IReadOnlyCollection<string> labels) =>
            labels is HashSet<string> set && ReferenceEquals(set.Comparer, StringComparer.OrdinalIgnoreCase)
                ? set
                : new HashSet<string>(labels, StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // How many each entry's badges say, with `selected` the components selected (none for none): "Overview"
        // every problem of the draft, the rest what their page shows. An entry with none is not listed.
        // ###########################################################################################
        public static Dictionary<ContributeEntry, BoardProblemCounts> CountsShown(IReadOnlyCollection<EntryProblem> placed, IReadOnlyCollection<string> selected)
        {
            ArgumentNullException.ThrowIfNull(placed);

            var counts = new Dictionary<ContributeEntry, BoardProblemCounts>();

            // The set made once, for every entry (OnEntry takes it as it is).
            var picked = new HashSet<string>(selected ?? [], StringComparer.OrdinalIgnoreCase);

            foreach (ContributeEntry entry in placed.Select(item => item.Entry).Distinct().Where(entry => entry != ContributeEntry.Overview))
            {
                BoardProblemCounts shown = BoardProblemCounts.Of(ContributeProblems.OnEntry(placed, entry, picked).Select(item => item.Problem));

                if (shown.Errors > 0 || shown.Warnings > 0)
                {
                    counts[entry] = shown;
                }
            }

            if (placed.Count > 0)
            {
                counts[ContributeEntry.Overview] = BoardProblemCounts.Of(placed.Select(item => item.Problem));
            }

            return counts;
        }

        // How many each component has, on all its pages - what its row in the list says. By its label, capitals
        // aside; a problem of no component is on none.
        public static Dictionary<string, BoardProblemCounts> CountsByComponent(IEnumerable<EntryProblem> placed)
        {
            ArgumentNullException.ThrowIfNull(placed);

            return placed
                .Where(item => item.ComponentLabel.Length > 0)
                .GroupBy(item => item.ComponentLabel, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => BoardProblemCounts.Of(group.Select(item => item.Problem)), StringComparer.OrdinalIgnoreCase);
        }

        // ---- what is said ------------------------------------------------------------------------------

        public static string Counted(int count, string one, string many) =>
            count == 1 ? $"1 {one}" : $"{count.ToString(CultureInfo.InvariantCulture)} {many}";

        // "2 errors and 1 warning" - the heading of a panel's own errors and warnings on "Overview".
        public static string Said(BoardProblemCounts counts)
        {
            var parts = new List<string>(2);

            if (counts.Errors > 0)
            {
                parts.Add(ContributeProblems.Counted(counts.Errors, "error", "errors"));
            }

            if (counts.Warnings > 0)
            {
                parts.Add(ContributeProblems.Counted(counts.Warnings, "warning", "warnings"));
            }

            return string.Join(" and ", parts);
        }

        // The panel's heading: "2 errors and 1 warning here".
        public static string Heading(BoardProblemCounts counts) => $"{ContributeProblems.Said(counts)} here";

        // Under it: what they mean, and on a component's page that they are of every component.
        public static string Help(BoardProblemCounts counts, bool ofComponents)
        {
            string meaning = counts.Errors == 0
                ? "Worth a look, but a warning does not stop the draft being submitted."
                : "An error stops the draft being submitted; a warning does not.";

            return ofComponents ? $"Of all the board's components. {meaning}" : meaning;
        }

        public static string LevelWord(BoardProblemLevel level) => level == BoardProblemLevel.Error ? "Error" : "Warning";

        public static string OpenLabel(string component) => $"Open [{component}]";

        // Past what the panel lists.
        public static string More(int notListed) =>
            $"And {notListed.ToString(CultureInfo.InvariantCulture)} more - \"Table\" marks every one on its cell.";

        // What a badge in the left column says under the pointer.
        public static string ErrorsTip(int errors) => $"{ContributeProblems.Counted(errors, "error", "errors")} here - to fix before the draft can be submitted";

        public static string WarningsTip(int warnings) => $"{ContributeProblems.Counted(warnings, "warning", "warnings")} here - worth a look";
    }
}
