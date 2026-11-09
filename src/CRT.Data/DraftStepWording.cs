using System.Collections.Generic;
using System.Globalization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT EACH EDITOR'S CHANGE IS CALLED IN A DRAFT'S HISTORY (DraftSession's step descriptions,
    // CRT 4.0 step 1, 2026-10-25) - "Undone: [U8] deleted." on the table's status line now, the
    // draft's Undo and Redo later. One place, so a change reads the same whichever screen undoes it.
    // The table's own steps are BoardTableSessionWording.Step.
    // ###########################################################################################
    public static class DraftStepWording
    {
        public static string ComponentDeleted(string label) => $"[{label}] deleted";

        // A new component, from the component pane's "Add component" (slice 3c).
        public static string ComponentAdded(string label) => $"[{label}] added";

        // A component's label changed in the component pane (owner report, 2026-11-01: "unable to rename").
        public static string ComponentRelabelled(string from, string to) => $"[{from}] renamed [{to}]";

        // A component told apart per region, or no longer (2026-11-01).
        public static string ComponentRegional(string label, bool regional) =>
            regional ? $"[{label}] told apart per region" : $"[{label}] the same in every region";

        public static string PinsRegional(string label, bool regional) =>
            regional ? $"Pins of [{label}] told apart per region" : $"Pins of [{label}] the same in every region";

        // What the board says about its regions (2026-11-01).
        public const string Regions = "The board's regions";

        // The board's categories (2026-11-01).
        public static string CategoryRenamed(string from, string to) =>
            from.Length == 0 ? $"Components with no category moved to [{to}]" : $"Category [{from}] renamed [{to}]";

        public static string CategoryDeleted(string name, string movedTo) => $"Category [{name}] deleted, its components moved to [{movedTo}]";

        public const string CategoriesOrdered = "Order of the categories";

        // A new board's components from its KiCad data, all in one go (step 3) - how many.
        public static string ComponentsFromKiCad(int count) =>
            count == 1 ? "1 component added from KiCad" : $"{count.ToString(System.Globalization.CultureInfo.InvariantCulture)} components added from KiCad";

        // "Mark all automatically" on a picture aligned with KiCad (step 3) - how many, and where.
        public static string ComponentsMarkedFromKiCad(string picture, int count) =>
            count == 1
                ? $"1 component marked on [{picture}] from KiCad"
                : $"{count.ToString(System.Globalization.CultureInfo.InvariantCulture)} components marked on [{picture}] from KiCad";

        // KiCad aligned again: the marks KiCad drew moved with it (owner request, 2026-11-01).
        public static string ComponentsFollowedAlignment(string picture, int count) =>
            count == 1
                ? $"1 component on [{picture}] moved with its KiCad alignment"
                : $"{count.ToString(System.Globalization.CultureInfo.InvariantCulture)} components on [{picture}] moved with their KiCad alignment";

        // "Also on other boards" (slice 5) - what the component took, and from where.
        public static string ComponentCopied(string label, string fromBoard) => $"[{label}] copied from {fromBoard}";

        // "Clone component from another board" (owner request, 2026-11-04).
        public static string ComponentCloned(string label, string fromLabel, string fromBoard) => $"[{label}] cloned from [{fromLabel}] of {fromBoard}";

        // "Also on other boards" for files - the files another board's draft took, by their row names.
        public static string FilesCopied(IReadOnlyList<string> rowNames, string fromBoard) =>
            rowNames.Count == 1 ? $"[{rowNames[0]}] copied from {fromBoard}" : $"Files copied from {fromBoard}";

        // One field of a component, from the component pane - "Short description of [U8]".
        public static string ComponentField(string label, string field) => $"{field} of [{label}]";

        public static string ComponentPins(string label) => $"Pins of [{label}]";

        // "Fill from KiCad" - the pins the board's KiCad data has, added and named.
        public static string ComponentPinsFromKiCad(string label) => $"Pins of [{label}] from KiCad";

        public static string ComponentReplacements(string label) => $"Replacement parts of [{label}]";

        public static string ComponentLinks(string label) => $"Links of [{label}]";

        public static string ComponentPictures(string label) => $"Pictures of [{label}]";

        public static string ComponentBaselines(string label) => $"Oscilloscope baselines of [{label}]";

        public static string ComponentFiles(string label) => $"Files of [{label}]";

        // "Make all like" with several components picked (ComponentBulk) - the page named as its entry is.
        public static string ComponentsMadeLike(string source, string page) => $"\"{page}\" made like [{source}]";

        // The board's own lists, from the component pane's board parts (slice 4a).
        public const string BoardFiles = "The board's files";

        public const string BoardLinks = "The board's links";

        public const string Credits = "Credits";

        // The rows a new board's credits start with (NewBoardCredits, 2026-11-02).
        public const string CreditsOfNewBoard = "Credits to fill in for the new board";

        // "Important KiCad signals" wherever the part is named since 2026-11-01 (owner request) - the
        // sheet keeps its stored name, "Important signals".
        public const string ImportantSignals = "Important KiCad signals";

        // The schematic images' highlight colours (slice 4b). The board's pictures are "schematic
        // images" wherever CRT names them since 2026-10-29 (owner request) - a COMPONENT's pictures
        // stay pictures (ComponentPictures).
        public const string PictureColours = "Colors of the schematic images";

        // "Undo changes to IC1" - the component back to its published values (ComponentRevert).
        public static string ComponentReverted(string label) => $"[{label}] back to its published values";

        public static string ComponentsMarked(string schematic) => $"Components marked on [{schematic}]";

        public static string Calibration(string schematic) => $"KiCad calibration of [{schematic}]";

        public static string SchematicsAdded(int count) =>
            count == 1 ? "Schematic image added" : $"{count.ToString(CultureInfo.InvariantCulture)} schematic images added";

        public static string SchematicRemoved(string name) => $"Schematic image [{name}] deleted";

        public static string SchematicRenamed(string from, string to) => $"Schematic image [{from}] renamed [{to}]";

        public const string SchematicOrder = "Order of the schematic images";

        public static string CadName(string schematic) => $"KiCad view of [{schematic}]";

        // Suggested matches taken on Board status (step 3) - the image, or how many.
        public static string PicturesMatched(IReadOnlyList<string> pictures) =>
            pictures.Count == 1
                ? DraftStepWording.CadName(pictures[0])
                : $"{pictures.Count.ToString(CultureInfo.InvariantCulture)} schematic images matched to KiCad views";

        // The file read back after a write that was more than the draft's content (DraftSession.WriteFileDirectly).
        public const string KiCadFileRemoved = "KiCad file removed";

        public const string BroughtUpToDate = "Brought up to date with the published board";
    }
}
