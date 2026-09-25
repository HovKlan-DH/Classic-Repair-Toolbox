using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Selection and visibility rules for a component's images and entries in the popup window.
    //
    // Extracted from ComponentInfoWindow. Each rule here decides what the user is shown, and each
    // has an edge case that is invisible from the UI: a blank region means "shared" rather than
    // "no region", an image with no File is not displayable at all, and an entry is treated as an
    // oscilloscope baseline only when it has a pin AND at least one scope setting.
    // ###########################################################################################
    public static class ComponentImageQueries
    {
        // ###########################################################################################
        // Returns true when an image is visible for the requested region.
        // Empty image regions are treated as shared and count for both PAL and NTSC.
        // ###########################################################################################
        public static bool IsImageVisibleInRegion(ComponentImageEntry image, string region)
        {
            return string.IsNullOrWhiteSpace(image.Region) ||
                   string.Equals(image.Region.Trim(), region, StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // An image with no File cannot be shown, whatever else it declares.
        // ###########################################################################################
        public static bool HasDisplayableImageFile(ComponentImageEntry image)
        {
            return !string.IsNullOrWhiteSpace(image.File);
        }

        // ###########################################################################################
        // Counts how many displayable images belong to the given board label for the requested
        // region. Empty image regions are included in both counters; entries without a File are
        // excluded.
        // ###########################################################################################
        public static int CountImagesForRegion(
            IEnumerable<ComponentImageEntry> allComponentImages,
            string boardLabel,
            string region)
        {
            return allComponentImages.Count(img =>
                string.Equals(img.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase) &&
                HasDisplayableImageFile(img) &&
                IsImageVisibleInRegion(img, region));
        }

        // ###########################################################################################
        // Builds the caption shown under a component image: the pin number when the entry has one,
        // otherwise its name, otherwise nothing.
        // ###########################################################################################
        public static string BuildImageLabel(ComponentImageEntry entry)
        {
            if (!string.IsNullOrWhiteSpace(entry.Pin))
                return $"Pin {entry.Pin.Trim()}";

            if (!string.IsNullOrWhiteSpace(entry.Name))
                return entry.Name.Trim();

            return string.Empty;
        }

        // ###########################################################################################
        // Returns true when the entry carries an oscilloscope baseline: it must name a pin and
        // supply at least one of the three scope settings.
        // ###########################################################################################
        public static bool IsOscilloscopeImage(ComponentImageEntry? componentImageEntry)
        {
            return componentImageEntry != null &&
                   !string.IsNullOrWhiteSpace(componentImageEntry.Pin) &&
                   (!string.IsNullOrWhiteSpace(componentImageEntry.TimeDiv) ||
                    !string.IsNullOrWhiteSpace(componentImageEntry.VoltsDiv) ||
                    !string.IsNullOrWhiteSpace(componentImageEntry.TriggerLevelVolts));
        }

        // ###########################################################################################
        // Selects the best-fit ComponentEntry for the given region:
        // exact region match -> generic (empty region) -> first available -> null.
        // ###########################################################################################
        public static ComponentEntry? PickComponentEntry(
            IReadOnlyList<ComponentEntry> allComponentEntries,
            string region)
        {
            if (allComponentEntries.Count == 0)
                return null;

            var regionMatch = allComponentEntries.FirstOrDefault(e =>
                string.Equals(e.Region?.Trim(), region, StringComparison.OrdinalIgnoreCase));
            if (regionMatch != null)
                return regionMatch;

            var generic = allComponentEntries.FirstOrDefault(e =>
                string.IsNullOrWhiteSpace(e.Region));
            if (generic != null)
                return generic;

            return allComponentEntries[0];
        }
    }
}
