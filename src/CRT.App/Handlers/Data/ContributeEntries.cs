using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE LEFT COLUMN'S ENTRIES ON THE CONTRIBUTE TAB, AND WHICH TABLE EACH SHOWS (owner request,
    // 2026-11-03: "the table will show in the content area of the left-side 'Overview' + 'Board' +
    // 'Component' panel ... do we not have a direct 1:1 mapping for everything here?"). Every sheet of
    // the board has an entry, "Component images" two; five entries are no sheet at all - nothing of them
    // is a row (a category is a column of "Components", and the highlights never were a sheet) - and
    // show their page whichever view is picked.
    // ###########################################################################################
    internal enum ContributeEntry
    {
        Overview,

        Name,
        Regions,
        KiCadFiles,
        ImportantSignals,
        SchematicImages,
        Categories,
        Files,
        Links,
        Credits,

        ComponentBasics,
        ComponentReplacements,
        ComponentPinout,
        ComponentPictures,
        ComponentBaselines,
        ComponentFiles,
        ComponentLinks,
    }

    internal readonly record struct ContributeEntryTable(string Sheet, BoardTableSheetPart Part);

    internal static class ContributeEntries
    {
        private static readonly Dictionary<ContributeEntry, ContributeEntryTable> Tables = new()
        {
            [ContributeEntry.ImportantSignals] = new(BoardWorkbookSchema.SheetKiCadImportantSignals, BoardTableSheetPart.Whole),
            [ContributeEntry.SchematicImages] = new(BoardWorkbookSchema.SheetBoardSchematics, BoardTableSheetPart.Whole),
            [ContributeEntry.Files] = new(BoardWorkbookSchema.SheetBoardLocalFiles, BoardTableSheetPart.Whole),
            [ContributeEntry.Links] = new(BoardWorkbookSchema.SheetBoardLinks, BoardTableSheetPart.Whole),
            [ContributeEntry.Credits] = new(BoardWorkbookSchema.SheetCredits, BoardTableSheetPart.Whole),
            [ContributeEntry.ComponentBasics] = new(BoardWorkbookSchema.SheetComponents, BoardTableSheetPart.Whole),
            [ContributeEntry.ComponentReplacements] = new(BoardWorkbookSchema.SheetComponentReplacements, BoardTableSheetPart.Whole),
            [ContributeEntry.ComponentPinout] = new(BoardWorkbookSchema.SheetComponentPins, BoardTableSheetPart.Whole),
            [ContributeEntry.ComponentPictures] = new(BoardWorkbookSchema.SheetComponentImages, BoardTableSheetPart.Pictures),
            [ContributeEntry.ComponentBaselines] = new(BoardWorkbookSchema.SheetComponentImages, BoardTableSheetPart.Baselines),
            [ContributeEntry.ComponentFiles] = new(BoardWorkbookSchema.SheetComponentLocalFiles, BoardTableSheetPart.Whole),
            [ContributeEntry.ComponentLinks] = new(BoardWorkbookSchema.SheetComponentLinks, BoardTableSheetPart.Whole),
        };

        // The table of an entry - null for one that is no sheet.
        public static ContributeEntryTable? TableOf(ContributeEntry entry) =>
            ContributeEntries.Tables.TryGetValue(entry, out ContributeEntryTable table) ? table : null;

        // The entry showing a part of a sheet. A sheet in two asked for whole is its first part.
        public static ContributeEntry? EntryOfTable(string? sheet, BoardTableSheetPart part)
        {
            foreach ((ContributeEntry entry, ContributeEntryTable table) in ContributeEntries.Tables)
            {
                if (table.Sheet == sheet && (table.Part == part || part == BoardTableSheetPart.Whole))
                {
                    return entry;
                }
            }

            return null;
        }

        public static ContributeEntry OfPage(ComponentPage page) => page switch
        {
            ComponentPage.Basics => ContributeEntry.ComponentBasics,
            ComponentPage.Replacements => ContributeEntry.ComponentReplacements,
            ComponentPage.Pinout => ContributeEntry.ComponentPinout,
            ComponentPage.Pictures => ContributeEntry.ComponentPictures,
            ComponentPage.Baselines => ContributeEntry.ComponentBaselines,
            ComponentPage.Files => ContributeEntry.ComponentFiles,
            _ => ContributeEntry.ComponentLinks,
        };

        // The component's page of an entry - null for an entry about the board.
        public static ComponentPage? PageOf(ContributeEntry entry) => entry switch
        {
            ContributeEntry.ComponentBasics => ComponentPage.Basics,
            ContributeEntry.ComponentReplacements => ComponentPage.Replacements,
            ContributeEntry.ComponentPinout => ComponentPage.Pinout,
            ContributeEntry.ComponentPictures => ComponentPage.Pictures,
            ContributeEntry.ComponentBaselines => ComponentPage.Baselines,
            ContributeEntry.ComponentFiles => ComponentPage.Files,
            ContributeEntry.ComponentLinks => ComponentPage.Links,
            _ => null,
        };

        public static ContributeEntry OfPart(PaneBoardPart part) => part switch
        {
            PaneBoardPart.Files => ContributeEntry.Files,
            PaneBoardPart.Links => ContributeEntry.Links,
            PaneBoardPart.Credits => ContributeEntry.Credits,
            PaneBoardPart.Pictures => ContributeEntry.SchematicImages,
            PaneBoardPart.Regions => ContributeEntry.Regions,
            PaneBoardPart.Categories => ContributeEntry.Categories,
            _ => ContributeEntry.ImportantSignals,
        };

        // The board part of an entry shown in the pane - null for one that is not.
        public static PaneBoardPart? PartOf(ContributeEntry entry) => entry switch
        {
            ContributeEntry.Files => PaneBoardPart.Files,
            ContributeEntry.Links => PaneBoardPart.Links,
            ContributeEntry.Credits => PaneBoardPart.Credits,
            ContributeEntry.SchematicImages => PaneBoardPart.Pictures,
            ContributeEntry.Regions => PaneBoardPart.Regions,
            ContributeEntry.Categories => PaneBoardPart.Categories,
            ContributeEntry.ImportantSignals => PaneBoardPart.ImportantSignals,
            _ => null,
        };

        // ###########################################################################################
        // What a component's page holds, as the left column says it (owner request, 2026-11-03: "Basics
        // (show '1' when the region is not different, or '2' if it is regionalized)"): Basics its rows,
        // one per region it differs in, and the rest what the pane lists, by the pane's own rules - with
        // what Submit would refuse. For a component the pane is not on; on it, the pane's own rows also
        // know what it holds back (ComponentPane.PagesShown).
        // ###########################################################################################
        public static IReadOnlyDictionary<ComponentPage, ComponentPageFact> FactsOf(BoardData board, string label, string region)
        {
            List<ComponentEntry> rows = board.Components
                .Where(row => ComponentPaneEdits.IsOfComponent(row.BoardLabel, label))
                .ToList();

            // As the pane marks them (code review, 2026-11-04: without what a baseline needs, the entry went red
            // only while the pane was on the component, though Submit refuses it either way).
            bool Refused(ComponentImageEntry row) =>
                BoardDataChecks.PackageRefusing(board.Components, label, row.Region, row.Pin) is not null ||
                ComponentPaneWording.BaselineProblem(row.Pin, row.Name, row.File) is not null;

            List<ComponentImageEntry> baselines = ComponentPaneEdits.BaselinesOf(board, label, region);
            List<ComponentPinEntry> pins = ComponentPaneEdits.PinsOf(board, label, region);
            List<ComponentLocalFileEntry> files = ComponentPaneEdits.AllFilesOf(board, label);
            List<ComponentLinkEntry> links = ComponentPaneEdits.AllLinksOf(board, label);

            return new Dictionary<ComponentPage, ComponentPageFact>
            {
                [ComponentPage.Basics] = new(rows.Count, rows.Any(row => string.IsNullOrWhiteSpace(row.Category))),
                [ComponentPage.Replacements] = new(ComponentPaneEdits.ReplacementsOf(board, label, region).Count, false),
                [ComponentPage.Pinout] = new(pins.Count, pins.Any(pin => BoardDataChecks.PackageRefusing(board.Components, label, pin.Region, pin.Pin) is not null)),
                [ComponentPage.Pictures] = new(ComponentPaneEdits.PicturesOf(board, label, region).Count, false),
                [ComponentPage.Baselines] = new(baselines.Count, baselines.Any(Refused)),
                [ComponentPage.Files] = new(files.Count, false),
                [ComponentPage.Links] = new(links.Count, links.Any(link => BoardDataChecks.IsRefusedLink(link.Url))),
            };
        }
    }

    // How many a component's page holds, and whether something on it stops the draft being submitted.
    internal readonly record struct ComponentPageFact(int Count, bool HasError);
}
