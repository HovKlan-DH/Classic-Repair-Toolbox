using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE WIDTHS DRAGGED ON THE CONTRIBUTE TAB (owner request, 2026-10-29: "configurable (user
    // dragable) splitters ... Should of course be saved - probably per board, as I assume some boards
    // could look differently"): the left column beside the workspace, and the preview on a schematic
    // image's card beside its colours.
    //
    //   - KEPT PER BOARD ("<hardware>|<board>", the settings' board key - capitals aside), and a board
    //     never dragged on takes the width dragged LAST, on any board: one drag carries over to every
    //     board, and a board made wider for its long names keeps its own.
    //   - ALWAYS IN RANGE - a width written by hand or by an older screen is brought inside it, so a
    //     column can never be dragged, or read back, to nothing.
    //
    // PURE: the settings' dictionary and last width in, the width out. UserSettings stores them.
    // ###########################################################################################
    internal static class ContributeWidths
    {
        public const double DefaultNavWidth = 260.0;

        public const double DefaultPreviewWidth = 320.0;

        public static (double Min, double Max) NavRange { get; } = (180.0, 640.0);

        public static (double Min, double Max) PreviewRange { get; } = (200.0, 1000.0);

        // The board's width, else the one dragged last - inside `range`.
        public static double For(IReadOnlyDictionary<string, double> byBoard, double last, string boardKey, (double Min, double Max) range)
        {
            ArgumentNullException.ThrowIfNull(byBoard);

            string key = boardKey?.Trim() ?? string.Empty;
            double width = last;

            if (key.Length > 0 && ContributeWidths.Find(byBoard, key) is { } found)
            {
                width = found.Value;
            }

            return ContributeWidths.Within(width, range);
        }

        // ###########################################################################################
        // `width` dragged on `boardKey`: kept for that board (under the spelling it already has there,
        // so capitals never make two entries) - and returned, inside `range`, as the width dragged last.
        // ###########################################################################################
        public static double Remember(IDictionary<string, double> byBoard, string boardKey, double width, (double Min, double Max) range)
        {
            ArgumentNullException.ThrowIfNull(byBoard);

            double kept = ContributeWidths.Within(width, range);
            string key = boardKey?.Trim() ?? string.Empty;

            if (key.Length > 0)
            {
                byBoard[ContributeWidths.Find(byBoard, key)?.Key ?? key] = kept;
            }

            return kept;
        }

        // The board's entry - its own spelling first, else in other capitals.
        private static KeyValuePair<string, double>? Find(IEnumerable<KeyValuePair<string, double>> byBoard, string key)
        {
            KeyValuePair<string, double>? other = null;

            foreach (KeyValuePair<string, double> entry in byBoard)
            {
                if (string.Equals(entry.Key, key, StringComparison.Ordinal))
                {
                    return entry;
                }

                if (other is null && string.Equals(entry.Key, key, StringComparison.OrdinalIgnoreCase))
                {
                    other = entry;
                }
            }

            return other;
        }

        // ###########################################################################################
        // THE SETTINGS BESIDE THE PREVIEWS KEEP THEIR ROOM (owner request, 2026-10-30: the column "with
        // schematic name and color picker etc." could be dragged to 0px). A card is its padding and border,
        // the previews, the handle, and the settings - which keep at least SettingsMinWidth. The previews
        // are shown at most as wide as leaves them that, never under the range's least: `wanted` as
        // dragged, `cardsWidth` the width the cards have (0 or less while not laid out - then `wanted`).
        // ###########################################################################################
        public const double SettingsMinWidth = 380.0;

        // A card's padding (12 a side), its border (1 a side) and the handle's column (14).
        public const double CardChrome = 40.0;

        public static double PreviewShown(double wanted, double cardsWidth)
        {
            double width = ContributeWidths.Within(wanted, ContributeWidths.PreviewRange);

            if (!double.IsFinite(cardsWidth) || cardsWidth <= 0)
            {
                return width;
            }

            double room = cardsWidth - ContributeWidths.CardChrome - ContributeWidths.SettingsMinWidth;

            return Math.Max(ContributeWidths.PreviewRange.Min, Math.Min(width, room));
        }

        // A width inside the range; a value that is no number is the range's least.
        private static double Within(double width, (double Min, double Max) range) =>
            double.IsFinite(width) ? Math.Clamp(width, range.Min, range.Max) : range.Min;
    }
}
