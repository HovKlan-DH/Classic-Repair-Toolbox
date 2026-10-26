using System.Collections.Generic;
using System.Globalization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT EACH EDITOR'S CHANGE IS CALLED IN A DRAFT'S HISTORY (DraftSession's step descriptions,
    // CRT 4.0 step 1, 2026-10-25) - "Undone: Changes to [U8]." on the table's status line now, the
    // draft's Undo and Redo later. One place, so a change reads the same whichever screen undoes it.
    // The table's own steps are BoardTableSessionWording.Step.
    // ###########################################################################################
    public static class DraftStepWording
    {
        public static string ComponentSaved(string label) => $"Changes to [{label}]";

        public static string ComponentDeleted(string label) => $"[{label}] deleted";

        // A new component, from the component pane's "Add component" (slice 3c).
        public static string ComponentAdded(string label) => $"[{label}] added";

        // "Copy from another board" (slice 5) - what the component took, and from where.
        public static string ComponentCopied(string label, string fromBoard) => $"[{label}] copied from {fromBoard}";

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

        public static string ComponentBaselines(string label) => $"Scope baselines of [{label}]";

        public static string ComponentFiles(string label) => $"Files of [{label}]";

        // The board's own lists, from the component pane's board parts (slice 4a).
        public const string BoardFiles = "The board's files";

        public const string BoardLinks = "The board's links";

        public const string Credits = "Credits";

        public const string ImportantSignals = "Important signals";

        // The pictures' highlight colours (slice 4b).
        public const string PictureColours = "Colours of the pictures";

        // "Undo changes to IC1" - the component back to its published values (ComponentRevert).
        public static string ComponentReverted(string label) => $"[{label}] back to its published values";

        public static string ComponentsMarked(string schematic) => $"Components marked on [{schematic}]";

        public static string Calibration(string schematic) => $"KiCad calibration of [{schematic}]";

        public static string SchematicsAdded(int count) =>
            count == 1 ? "Schematic image added" : $"{count.ToString(CultureInfo.InvariantCulture)} schematic images added";

        public static string SchematicRemoved(string name) => $"Schematic image [{name}] removed";

        public static string SchematicRenamed(string from, string to) => $"Schematic image [{from}] renamed [{to}]";

        public const string SchematicOrder = "Order of the schematic images";

        public static string CadName(string schematic) => $"CAD name of [{schematic}]";

        // The file read back after a write that was more than the draft's content (DraftSession.WriteFileDirectly).
        public const string KiCadFileRemoved = "KiCad file removed";

        public const string BroughtUpToDate = "Brought up to date with the published board";
    }
}
