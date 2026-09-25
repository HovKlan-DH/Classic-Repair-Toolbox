using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Writes the `.json` SIDECAR that sits beside a board workbook
    // (NewContributeStrategy.md Phase 5, task 6).
    //
    // *** THIS EXISTS BECAUSE PUBLISHING WOULD OTHERWISE DESTROY DATA. *** A board is TWO files.
    // The workbook holds the rows; the sidecar holds every COMPONENT HIGHLIGHT - the rectangles
    // the whole app is built around - and every KiCad calibration, under two separate JSON roots.
    // `BoardWorkbookWriter` writes only the workbook, so a publish that stopped there would leave
    // a board whose highlights had silently vanished. With task 7 struck there is no retained
    // revision to restore from, so it would not be recoverable either.
    //
    // *** IT WRITES THE COMPLETE STATE, and that is why it is not
    // BoardComponentHighlightStorage.SaveComponentHighlights. *** That method replaces ONE
    // schematic's highlights and deliberately preserves every other schematic already in the file
    // - correct for the app's label editor, which edits one schematic at a time, and exactly wrong
    // here. A submission's manifest is BY CONTRACT the complete intended state, so a schematic
    // whose highlights it deleted must actually lose them. A preserving writer would leave them on
    // the published board forever and nothing would ever report it.
    //
    // *** BOTH ROOTS ARE WRITTEN IN ONE PASS, over one document. *** Writing them independently -
    // open, set one root, save, reopen, set the other, save - is the obvious shape and loses data:
    // the second write either reloads what the first just wrote and overwrites it, or replaces the
    // document outright. Highlights are the half nobody notices missing until they open the board.
    //
    // The FORMAT is the shipped one, matched field for field against
    // BoardComponentHighlightStorage's own readers and savers - integer rects under
    // "Component highlights", and the seven calibration fields under "KiCad calibration points".
    // It is the reader that decides what the format is; this file follows it.
    // ###########################################################################################
    public static class BoardSidecarWriter
    {
        public const string HighlightsRoot = "Component highlights";

        public const string CalibrationRoot = "KiCad calibration points";

        // ###########################################################################################
        // Writes the whole sidecar for a board.
        //
        // workbookPath is the `.xlsx`; the sidecar path is derived from it exactly as the READER
        // derives it, so the two cannot disagree about where the file lives.
        //
        // *** NULL LISTS ARE REFUSED RATHER THAN TREATED AS EMPTY. *** A caller passing null has
        // almost certainly failed to load something rather than genuinely meaning "this board has
        // no highlights", and treating the two alike is how a publish quietly wipes a board. An
        // EMPTY list is a real and supported answer - a submission may legitimately remove every
        // highlight - and is written as an empty root rather than skipped, because skipping would
        // leave the previous board's highlights in place.
        // ###########################################################################################
        public static void Write(
            string workbookPath,
            IReadOnlyList<ComponentHighlightEntry> highlights,
            IReadOnlyList<KiCadCalibrationEntry> calibrations)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(workbookPath);
            ArgumentNullException.ThrowIfNull(highlights);
            ArgumentNullException.ThrowIfNull(calibrations);

            // The same derivation the reader uses. A mismatch here means a file nobody ever loads
            // and a board that publishes with no highlights at all.
            string sidecarPath = BoardComponentHighlightStorage.GetJsonPath(workbookPath);

            if (string.IsNullOrWhiteSpace(sidecarPath))
                throw new InvalidOperationException($"Could not resolve a sidecar path for [{workbookPath}].");

            // ONE document carrying BOTH roots - see the class header on why this is not two
            // writes. Built from scratch rather than loaded: the published sidecar must be exactly
            // what the submission describes, and carrying an unknown root forward from whatever
            // was there before would publish data no reviewer ever saw.
            var root = new JsonObject
            {
                [BoardSidecarWriter.HighlightsRoot] = BoardSidecarWriter.BuildHighlights(highlights),
                [BoardSidecarWriter.CalibrationRoot] = BoardSidecarWriter.BuildCalibrations(calibrations)
            };

            string directory = Path.GetDirectoryName(sidecarPath) ?? string.Empty;

            // A brand-new system's folder does not exist before its first publish.
            if (!string.IsNullOrWhiteSpace(directory))
                Directory.CreateDirectory(directory);

            File.WriteAllText(
                sidecarPath,
                root.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
        }

        // ###########################################################################################
        // The "Component highlights" root: schematic -> board label -> an ARRAY of rects.
        //
        // THE ARRAY IS NOT INCIDENTAL. A component genuinely appears more than once on a schematic
        // - a connector drawn at both ends, a chip shown split across the sheet - and the shipped
        // format carries a list per label for exactly that. Writing one rect per label would drop
        // the second silently.
        //
        // Ordering is stable (schematic, then label, then position) so republishing an unchanged
        // board produces a byte-identical file. That matters beyond tidiness: the sync manifest
        // hashes these files, and an unstable order would make every client re-download every
        // board on every publish.
        // ###########################################################################################
        private static JsonObject BuildHighlights(IReadOnlyList<ComponentHighlightEntry> highlights)
        {
            var root = new JsonObject();

            IEnumerable<IGrouping<string, ComponentHighlightEntry>> bySchematic = highlights
                .Where(entry => !string.IsNullOrWhiteSpace(entry.SchematicName))
                .Where(entry => !string.IsNullOrWhiteSpace(entry.BoardLabel))
                .GroupBy(entry => entry.SchematicName.Trim(), StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase);

            foreach (IGrouping<string, ComponentHighlightEntry> schematic in bySchematic)
            {
                var labels = new JsonObject();

                IEnumerable<IGrouping<string, ComponentHighlightEntry>> byLabel = schematic
                    .GroupBy(entry => entry.BoardLabel.Trim(), StringComparer.OrdinalIgnoreCase)
                    .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase);

                foreach (IGrouping<string, ComponentHighlightEntry> label in byLabel)
                {
                    var rects = new JsonArray();

                    foreach (ComponentHighlightEntry entry in label)
                    {
                        if (!BoardSidecarWriter.TryBuildRect(entry, out JsonObject? rect))
                            continue;

                        rects.Add(rect);
                    }

                    // A label whose every rect was malformed contributes nothing rather than an
                    // empty array, which the reader would carry as a label with no rectangles.
                    if (rects.Count > 0)
                        labels[label.Key] = rects;
                }

                if (labels.Count > 0)
                    root[schematic.Key] = labels;
            }

            return root;
        }

        // ###########################################################################################
        // One rectangle, or nothing.
        //
        // *** A MALFORMED COORDINATE SKIPS THE HIGHLIGHT rather than writing a zero. *** Zero is a
        // real position - the board's top-left corner - so a defaulted value looks like a highlight
        // somebody mis-drew rather than like data that never parsed, and nobody would ever question
        // it.
        //
        // *** PARSED INVARIANT, ALWAYS. *** These are STRINGS in BoardData. Read under da-DK with
        // culture-sensitive parsing, "100.5" becomes 1005 - a highlight ten times across the board,
        // written into a PUBLISHED file that every user then downloads. This project has hit that
        // class of bug four times; this is the worst place for it, because the damage ships.
        //
        // ROUNDED TO INTEGERS, matching BoardComponentHighlightStorage's own save path
        // (MidpointRounding.AwayFromZero), so a board saved by the app and one saved by the server
        // are the same file.
        // ###########################################################################################
        private static bool TryBuildRect(ComponentHighlightEntry entry, out JsonObject? rect)
        {
            rect = null;

            if (!BoardSidecarWriter.TryParse(entry.X, out double x) ||
                !BoardSidecarWriter.TryParse(entry.Y, out double y) ||
                !BoardSidecarWriter.TryParse(entry.Width, out double width) ||
                !BoardSidecarWriter.TryParse(entry.Height, out double height))
            {
                return false;
            }

            rect = new JsonObject
            {
                ["X"] = BoardSidecarWriter.RoundToInt(x),
                ["Y"] = BoardSidecarWriter.RoundToInt(y),
                ["Width"] = BoardSidecarWriter.RoundToInt(width),
                ["Height"] = BoardSidecarWriter.RoundToInt(height)
            };

            return true;
        }

        // ###########################################################################################
        // The "KiCad calibration points" root: schematic -> the seven stored fields.
        //
        // The field names and their casing are the shipped ones - they are what
        // TryLoadKiCadCalibration reads, and a rename here is a calibration that silently stops
        // loading while the file still looks right.
        // ###########################################################################################
        private static JsonObject BuildCalibrations(IReadOnlyList<KiCadCalibrationEntry> calibrations)
        {
            var root = new JsonObject();

            IEnumerable<KiCadCalibrationEntry> usable = calibrations
                .Where(entry => !string.IsNullOrWhiteSpace(entry.SchematicName))
                .OrderBy(entry => entry.SchematicName.Trim(), StringComparer.OrdinalIgnoreCase);

            foreach (KiCadCalibrationEntry entry in usable)
            {
                root[entry.SchematicName.Trim()] = new JsonObject
                {
                    ["CadName"] = entry.CadName?.Trim() ?? string.Empty,
                    ["OffsetX"] = entry.OffsetX,
                    ["OffsetY"] = entry.OffsetY,
                    ["ScaleX"] = entry.ScaleX,
                    ["ScaleY"] = entry.ScaleY,
                    ["MirrorX"] = entry.MirrorX,
                    ["MirrorY"] = entry.MirrorY
                };
            }

            return root;
        }

        private static bool TryParse(string? text, out double value) =>
            double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);

        // Away from zero, matching BoardComponentHighlightStorage's own RoundToInt - so the app and
        // the server round a .5 the same way and produce the same file.
        private static int RoundToInt(double value) =>
            (int)Math.Round(value, MidpointRounding.AwayFromZero);
    }
}
