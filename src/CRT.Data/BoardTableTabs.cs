using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE TABS OF A TABLE THAT SHOWS ITS OWN (the Maintainer tab's; owner request, 2026-11-05: "group and order
    // into what is logical for the new setup. E.g. the 'Component images' have been split into two ('Pictures'
    // and 'Oscilloscope baseline') and the order of tabs is not correct either"). Under the names, and the
    // headings, of the Contribute tab's left column, so a maintainer reads what the contributor filled in by the
    // words the contributor saw. Only there does KiCad stand inside "Board": a table opens on its first tab, and
    // for a board with no KiCad data "Important signals" is an empty one.
    //
    // The order is the table's alone: every workbook is still written in BoardWorkbookSchema.AllSheets' order.
    // ###########################################################################################
    public sealed record BoardTableTab(string Group, string Name, string Sheet, BoardTableSheetPart Part);

    public static class BoardTableTabs
    {
        public const string BoardGroup = "Board";
        public const string KiCadGroup = "KiCad";
        public const string ComponentsGroup = "Components";

        public static readonly IReadOnlyList<BoardTableTab> All =
        [
            new(BoardGroup, "Schematic images", BoardWorkbookSchema.SheetBoardSchematics, BoardTableSheetPart.Whole),
            new(BoardGroup, "Files", BoardWorkbookSchema.SheetBoardLocalFiles, BoardTableSheetPart.Whole),
            new(BoardGroup, "Links", BoardWorkbookSchema.SheetBoardLinks, BoardTableSheetPart.Whole),
            new(BoardGroup, "Credits", BoardWorkbookSchema.SheetCredits, BoardTableSheetPart.Whole),

            new(KiCadGroup, "Important signals", BoardWorkbookSchema.SheetKiCadImportantSignals, BoardTableSheetPart.Whole),

            new(ComponentsGroup, "Basics", BoardWorkbookSchema.SheetComponents, BoardTableSheetPart.Whole),
            new(ComponentsGroup, "Replacement parts", BoardWorkbookSchema.SheetComponentReplacements, BoardTableSheetPart.Whole),
            new(ComponentsGroup, "Pinout", BoardWorkbookSchema.SheetComponentPins, BoardTableSheetPart.Whole),
            new(ComponentsGroup, "Pictures", BoardWorkbookSchema.SheetComponentImages, BoardTableSheetPart.Pictures),
            new(ComponentsGroup, "Oscilloscope baseline", BoardWorkbookSchema.SheetComponentImages, BoardTableSheetPart.Baselines),
            new(ComponentsGroup, "Files", BoardWorkbookSchema.SheetComponentLocalFiles, BoardTableSheetPart.Whole),
            new(ComponentsGroup, "Links", BoardWorkbookSchema.SheetComponentLinks, BoardTableSheetPart.Whole),
        ];

        // The sheets in the order the tabs have them - what a pill or a search going to the first it finds goes by.
        public static IReadOnlyList<BoardWorkbookSchema.SheetDefinition> SheetsInOrder { get; } =
            BoardTableTabs.All
                .Select(tab => tab.Sheet)
                .Distinct()
                .Select(name => BoardWorkbookSchema.AllSheets.First(sheet => sheet.SheetName == name))
                .ToList();

        // ###########################################################################################
        // A sheet as a person knows it: its tab under its group ("Components: Links"), and both tabs of a sheet in
        // two. For what is said beside the tabs, which name no sheet of the board file (code review, 2026-11-05).
        // ###########################################################################################
        public static string NameOf(string sheetName)
        {
            List<BoardTableTab> tabs = BoardTableTabs.All
                .Where(tab => string.Equals(tab.Sheet, sheetName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            return tabs.Count == 0
                ? sheetName
                : $"{tabs[0].Group}: {string.Join(" or ", tabs.Select(tab => tab.Name))}";
        }

        // The first tab holding a change - what a maintainer opening a submission wants to see. Null with none.
        public static BoardTableTab? FirstChanged(BoardTableDocument document) =>
            BoardTableTabs.All.FirstOrDefault(tab =>
                document.FindSheet(tab.Sheet) is { } sheet && BoardTableSheetParts.ChangeCountOf(sheet, tab.Part) > 0);
    }
}
