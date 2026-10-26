using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A BOARD'S KiCad CALIBRATIONS WITH ONE SAVED - one per schematic image, by its name (trimmed,
    // capitals ignored, as every schematic name is compared), kept in name order. The one rule for
    // the board file's writer (BoardComponentHighlightStorage.SaveKiCadCalibration) and the draft
    // held in memory (the Schematics tab's calibration through DraftSession, 2026-10-25).
    // ###########################################################################################
    public static class KiCadCalibrationList
    {
        public static List<KiCadCalibrationEntry> With(IEnumerable<KiCadCalibrationEntry> calibrations, KiCadCalibrationEntry saved)
        {
            ArgumentNullException.ThrowIfNull(calibrations);
            ArgumentNullException.ThrowIfNull(saved);

            string name = saved.SchematicName?.Trim() ?? string.Empty;

            if (name.Length == 0)
            {
                throw new InvalidOperationException("No schematic name was provided for KiCad calibration save.");
            }

            List<KiCadCalibrationEntry> kept = calibrations
                .Where(entry => !string.Equals(entry.SchematicName?.Trim(), name, StringComparison.OrdinalIgnoreCase))
                .ToList();

            kept.Add(new KiCadCalibrationEntry
            {
                SchematicName = name,
                CadName = saved.CadName?.Trim() ?? string.Empty,
                OffsetX = saved.OffsetX,
                OffsetY = saved.OffsetY,
                ScaleX = saved.ScaleX,
                ScaleY = saved.ScaleY,
                MirrorX = saved.MirrorX,
                MirrorY = saved.MirrorY,
            });

            return kept.OrderBy(entry => entry.SchematicName, StringComparer.OrdinalIgnoreCase).ToList();
        }
    }
}
