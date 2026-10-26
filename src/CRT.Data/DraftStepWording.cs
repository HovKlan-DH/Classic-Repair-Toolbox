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

        // One field of a component, from the component pane - "Short description of [U8]".
        public static string ComponentField(string label, string field) => $"{field} of [{label}]";

        public static string ComponentPins(string label) => $"Pins of [{label}]";

        public static string ComponentReplacements(string label) => $"Replacement parts of [{label}]";

        public static string ComponentLinks(string label) => $"Links of [{label}]";

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
