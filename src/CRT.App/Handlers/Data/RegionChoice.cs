using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHICH REGION CRT SHOWS, AND IN WHICH COLORS (2026-11-09: a board may come in regions of its own - BoardRegions).
    //
    //   - THE REGION SHOWN for a board: the one last picked on that board when it is not PAL or NTSC, else the one picked
    //     anywhere (UserSettings.Region - PAL or NTSC, the setting CRT 3.x reads too), else the board's first. A board
    //     coming in PAL and NTSC shows what it always did.
    //   - ITS COLORS: PAL and NTSC keep theirs, and the colors a person set for them; a region of the board's own takes one
    //     of a fixed set, by its place among the board's own regions - the same on the button, the label over the picture
    //     and the popup.
    // ###########################################################################################
    internal static class RegionChoice
    {
        // How many colors the board's own regions take turns with (App.axaml's Schematics_Region_Own1..4).
        public const int OwnColors = 4;

        public static string ToShow(IReadOnlyList<string> regions, string? picked, string? pickedOnBoard)
        {
            ArgumentNullException.ThrowIfNull(regions);

            if (regions.Count == 0)
            {
                return picked?.Trim() ?? string.Empty;
            }

            return RegionChoice.Spelt(regions, pickedOnBoard)
                ?? RegionChoice.Spelt(regions, picked)
                ?? regions[0];
        }

        // ###########################################################################################
        // The name the board's colors for `region` go by - "Schematics_Region_PAL", "Schematics_Region_Own2" - with "_Bg",
        // "_Fg" and "_Border" after it; "SchematicsRegion" for a region the board does not have.
        // ###########################################################################################
        public static string ColorsOf(IReadOnlyList<string> regions, string? region)
        {
            ArgumentNullException.ThrowIfNull(regions);

            string wanted = region?.Trim() ?? string.Empty;

            if (string.Equals(wanted, BoardRegions.Pal, StringComparison.OrdinalIgnoreCase))
            {
                return "Schematics_Region_PAL";
            }

            if (string.Equals(wanted, BoardRegions.Ntsc, StringComparison.OrdinalIgnoreCase))
            {
                return "Schematics_Region_NTSC";
            }

            return RegionChoice.OwnNumberOf(regions, wanted) is int number ? $"Schematics_Region_Own{number}" : "SchematicsRegion";
        }

        // 1 to OwnColors for a region of the board's own, by its place among them - null for PAL, NTSC or none of them.
        public static int? OwnNumberOf(IReadOnlyList<string> regions, string? region)
        {
            ArgumentNullException.ThrowIfNull(regions);

            List<string> own = regions.Where(each => !BoardRegions.IsPalOrNtsc(each)).ToList();
            int at = own.FindIndex(each => string.Equals(each, region?.Trim(), StringComparison.OrdinalIgnoreCase));

            return at < 0 ? null : (at % RegionChoice.OwnColors) + 1;
        }

        private static string? Spelt(IReadOnlyList<string> regions, string? region) =>
            regions.FirstOrDefault(each => string.Equals(each, region?.Trim(), StringComparison.OrdinalIgnoreCase));
    }
}
