using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What removing one schematic image from a draft takes with it: the components marked on it
    // (their highlight rectangles), and its KiCad calibration.
    // ###########################################################################################
    public sealed record SchematicRemovalLoss(int MarkedComponents, bool HasCalibration)
    {
        public static SchematicRemovalLoss Nothing { get; } = new(0, false);

        public bool IsNothing => this.MarkedComponents == 0 && !this.HasCalibration;
    }

    // ###########################################################################################
    // REMOVING A SCHEMATIC IMAGE ASKS FIRST WHEN SOMETHING IS LOST WITH IT (code review,
    // 2026-10-10).
    //
    // "Remove" in the "Schematic images" window never asked: only the workbook ROW went, the image
    // file stayed in the draft folder, and adding the image again put it back. Since every draft
    // write follows what hangs on a schematic's name (DraftWorkbookStore.EditCore), removing the row
    // also drops every highlight drawn on the image and its KiCad calibration - hours of work, and
    // adding the image again brings none of it back. So Remove asks when there is any; an image with
    // nothing drawn on it is still removed at once, as before.
    //
    // Nothing goes when another image keeps the name (two rows of one name, an error): the follow
    // drops only a name the board no longer has.
    // ###########################################################################################
    public static class SchematicRemoval
    {
        // What goes with the image `schematicName`, given the board's highlights and whether the
        // image has a calibration: the components with at least one rectangle on it.
        public static SchematicRemovalLoss LossOf(
            string schematicName,
            IEnumerable<ComponentHighlightEntry> highlights,
            bool hasCalibration)
        {
            int marked = highlights
                .Where(highlight => SchematicRename.SameName(highlight.SchematicName, schematicName))
                .Select(highlight => highlight.BoardLabel?.Trim() ?? string.Empty)
                .Where(label => label.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count();

            return new SchematicRemovalLoss(marked, hasCalibration);
        }

        // ###########################################################################################
        // The same, read from the draft's sidecar - the JSON beside `workbookPath`, which is where
        // both live. Nothing when there is no sidecar.
        // ###########################################################################################
        public static SchematicRemovalLoss Read(string workbookPath, string schematicName)
        {
            if (string.IsNullOrWhiteSpace(workbookPath) ||
                !File.Exists(BoardComponentHighlightStorage.GetJsonPath(workbookPath)))
            {
                return SchematicRemovalLoss.Nothing;
            }

            bool hasCalibration = BoardComponentHighlightStorage.TryLoadKiCadCalibration(
                workbookPath, schematicName, out _, out _, out _, out _, out _, out _, out _);

            return SchematicRemoval.LossOf(
                schematicName,
                BoardComponentHighlightStorage.LoadComponentHighlights(workbookPath),
                hasCalibration);
        }

        public static string Question(string schematicName) => $"Remove [{schematicName}] from your draft?";

        // What the question says goes with it. Empty for nothing - nothing is asked then.
        public static string WhatGoes(SchematicRemovalLoss loss)
        {
            if (loss.IsNothing)
            {
                return string.Empty;
            }

            string them = loss.MarkedComponents > 0 ? "them" : "it";

            return $"{SchematicRemoval.GoesWithIt(loss)} Adding the image again later does not bring {them} back - the image file itself stays in your draft folder.";
        }

        // "Its highlights for 3 components and its KiCad calibration go with it." - or the part of
        // that the image has.
        private static string GoesWithIt(SchematicRemovalLoss loss)
        {
            string components = loss.MarkedComponents == 1
                ? "1 component"
                : string.Create(CultureInfo.InvariantCulture, $"{loss.MarkedComponents} components");

            return (loss.MarkedComponents > 0, loss.HasCalibration) switch
            {
                (true, true) => $"Its highlights for {components} and its KiCad calibration go with it.",
                (true, false) => $"Its highlights for {components} go with it.",
                _ => "Its KiCad calibration goes with it.",
            };
        }
    }
}
