using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Layer filtering and zone polygon selection for the KiCad overlay, extracted from
    // TabSchematics so it can be unit tested.
    //
    // "Which side of the board am I looking at, and does this copper belong on it?" is the
    // question these answer. Getting it wrong shows top-side traces on a bottom-side view, which
    // looks plausible and is very hard to spot by eye.
    // ###########################################################################################
    public static class KiCadLayerGeometry
    {
        // ###########################################################################################
        // Returns true when a KiCad copper point is visible on the inspected PCB side.
        // Treats "*.Cu" as visible on both sides so through-hole pads and vias are included.
        // ###########################################################################################
        public static bool IsPointVisibleOnSide(IEnumerable<string> layers, string requiredLayer)
        {
            foreach (string layer in layers
                         .Where(layer => !string.IsNullOrWhiteSpace(layer))
                         .Select(layer => layer.Trim()))
            {
                if (string.Equals(layer, requiredLayer, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(layer, "*.Cu", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return !layers.Any();
        }

        // ###########################################################################################
        // Returns true when the supplied zone is visible on the inspected PCB side.
        // ###########################################################################################
        public static bool IsZoneVisibleOnSide(KiCadPcbZone zone, string requiredLayer)
        {
            return KiCadLayerGeometry.IsPointVisibleOnSide(zone.Layers, requiredLayer);
        }

        // ###########################################################################################
        // Returns the world-space polygons that should be used for one zone.
        // Filled polygons are preferred because they match the final poured copper area.
        // ###########################################################################################
        public static IReadOnlyList<IReadOnlyList<Point>> GetZoneWorldPolygons(KiCadPcbZone zone)
        {
            var sourcePolygons = zone.FilledPolygons.Count > 0
                ? zone.FilledPolygons
                : zone.OutlinePolygons;

            return sourcePolygons
                .Where(polygon => polygon.Points.Count >= 3)
                .Select(polygon => (IReadOnlyList<Point>)polygon.Points
                    .Select(point => new Point(point.X, point.Y))
                    .ToList())
                .ToList();
        }

        // ###########################################################################################
        // Compares KiCad pad designators so numeric pins sort numerically and non-numeric pins sort
        // alphabetically. This lets footprints like B/C/E choose B as the primary highlighted pin.
        // ###########################################################################################
        public static int ComparePadDesignators(string? left, string? right)
        {
            string leftValue = left?.Trim() ?? string.Empty;
            string rightValue = right?.Trim() ?? string.Empty;

            bool leftIsNumber = int.TryParse(leftValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int leftNumber);
            bool rightIsNumber = int.TryParse(rightValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int rightNumber);

            if (leftIsNumber && rightIsNumber)
            {
                return leftNumber.CompareTo(rightNumber);
            }

            if (leftIsNumber != rightIsNumber)
            {
                return leftIsNumber ? -1 : 1;
            }

            return StringComparer.OrdinalIgnoreCase.Compare(leftValue, rightValue);
        }

        // ###########################################################################################
        // Extracts the ordinal page number from names such as "Schematics #1 of 2".
        // ###########################################################################################
        public static bool TryExtractSchematicPageOrdinal(string schematicName, out int pageOrdinal)
        {
            pageOrdinal = 0;

            int hashIndex = schematicName.IndexOf('#');
            int ofIndex = schematicName.IndexOf(" of ", StringComparison.OrdinalIgnoreCase);

            if (hashIndex < 0 || ofIndex <= hashIndex)
            {
                return false;
            }

            string digits = new string(schematicName[(hashIndex + 1)..ofIndex].Where(char.IsDigit).ToArray());
            return int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out pageOrdinal) &&
                   pageOrdinal > 0;
        }

        // ###########################################################################################
        // Returns true when a schematic symbol reference is a generated internal KiCad helper symbol
        // that should not be used as a human-facing calibration candidate.
        // ###########################################################################################
        public static bool IsInternalSymbolReference(string reference)
        {
            string trimmed = reference?.Trim() ?? string.Empty;
            return trimmed.StartsWith("#", StringComparison.OrdinalIgnoreCase);
        }
    }
}
