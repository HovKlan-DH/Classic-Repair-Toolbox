using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Turns a board's contributed highlight rows into the per-schematic, per-component rectangles
    // the overlay draws, applying the active region filter (PAL / NTSC).
    //
    // Extracted from TabSchematics. The region rule is the subtle part: a component with no
    // declared region is visible everywhere, one with regions is visible only in those.
    // ###########################################################################################
    public static class HighlightRectBuilder
    {

    // ###########################################################################################
    // Builds per-schematic highlight rect lookups.
    // ###########################################################################################
    public static Dictionary<string, Dictionary<string, List<Rect>>> BuildHighlightRects(BoardData boardData, string region)
    {
        var componentRegionsByLabel = boardData.Components
            .Where(c => !string.IsNullOrWhiteSpace(c.BoardLabel))
            .GroupBy(c => c.BoardLabel, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                g => g.Key,
                g => g.Select(c => c.Region?.Trim() ?? string.Empty)
                      .Where(r => !string.IsNullOrWhiteSpace(r))
                      .Distinct(StringComparer.OrdinalIgnoreCase)
                      .ToList(),
                StringComparer.OrdinalIgnoreCase);

        bool IsVisibleByRegion(string boardLabel)
        {
            if (!componentRegionsByLabel.TryGetValue(boardLabel, out var regionsForLabel)) return true;
            if (regionsForLabel.Count == 0) return true;
            return regionsForLabel.Any(r => string.Equals(r, region, StringComparison.OrdinalIgnoreCase));
        }

        var result = new Dictionary<string, Dictionary<string, List<Rect>>>(StringComparer.OrdinalIgnoreCase);

        foreach (var h in boardData.ComponentHighlights)
        {
            if (string.IsNullOrWhiteSpace(h.SchematicName) || string.IsNullOrWhiteSpace(h.BoardLabel)) continue;
            if (!IsVisibleByRegion(h.BoardLabel)) continue;

            if (!RectGeometry.TryParseDouble(h.X, out var x) || !RectGeometry.TryParseDouble(h.Y, out var y) ||
                !RectGeometry.TryParseDouble(h.Width, out var w) || !RectGeometry.TryParseDouble(h.Height, out var hh))
                continue;

            if (w <= 0 || hh <= 0) continue;

            if (!result.TryGetValue(h.SchematicName, out var byLabel))
            {
                byLabel = new Dictionary<string, List<Rect>>(StringComparer.OrdinalIgnoreCase);
                result[h.SchematicName] = byLabel;
            }

            if (!byLabel.TryGetValue(h.BoardLabel, out var rects))
            {
                rects = new List<Rect>();
                byLabel[h.BoardLabel] = rects;
            }

            rects.Add(new Rect(x, y, w, hh));
        }

        return result;
    }
    }
}