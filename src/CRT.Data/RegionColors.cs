using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // One region's own colors on a board (owner request, 2026-11-11): its background, and its text - empty for the text color
    // that reads best on that background (RegionColors.ContrastText). Both "#RRGGBB".
    // ###########################################################################################
    public sealed class RegionColorEntry
    {
        public string Region { get; init; } = string.Empty;

        public string Background { get; init; } = string.Empty;

        public string Text { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // A BOARD'S REGIONS IN COLORS OF ITS OWN (owner request, 2026-11-11: "In "Board: Regions" I would like to be able to set
    // the background color, and then you should calculate the complementary text color, that gives the highest visiblity of
    // the text. However, I would also like to be able to set the text color, if your color is not to my liking"; asked, the
    // colors are "In the board data" - submitted and published with the board, so everybody sees the same).
    //
    //   - STORED in the board file's "regionColors" (written only when a region has any), one row per region: its background,
    //     and its text color or nothing - nothing reads as the color that stands out most on that background, black or white,
    //     whichever has the higher contrast ratio (WCAG 2's, the one the Web is measured by).
    //   - A REGION WITH NO ROW keeps the colors CRT gives it: PAL's and NTSC's (which a person may set in the "Configuration"
    //     tab), a region of the board's own one of four in turn (RegionChoice).
    //   - ONLY "#RRGGBB", both here and on the server (the error "regions.color"): what is sent is drawn on every CRT, and a
    //     color CRT cannot read must never reach one.
    //
    // PURE: boards in, a board out.
    // ###########################################################################################
    public static class RegionColors
    {
        public const string Black = "#000000";

        public const string White = "#FFFFFF";

        // "#RRGGBB", capitals for the digits - null for anything else.
        public static string? Normalised(string? text)
        {
            string trimmed = (text ?? string.Empty).Trim();

            return trimmed.Length == 7 && trimmed[0] == '#' && trimmed.Skip(1).All(char.IsAsciiHexDigit)
                ? trimmed.ToUpperInvariant()
                : null;
        }

        public static bool IsColor(string? text) => RegionColors.Normalised(text) is not null;

        // What a reviewer and the board's history are told of a change of them - and "Review changes" and the Submit dialog.
        public const string ChangedLine = "The colors of the board's regions changed.";

        // A new board's, which nothing compares with.
        public const string NewBoardLine = "The board gives its regions colors of their own.";

        // ###########################################################################################
        // The colors `region` is drawn in on `board` - null for none of its own (a region the board does not have, or a row
        // whose background CRT cannot read): the caller's defaults then. Text worked out when it is not set.
        // ###########################################################################################
        public static (string Background, string Text)? Of(BoardData? board, string? region) =>
            board is null ? null : RegionColors.Of(BoardRegions.Of(board), board.RegionColors, region);

        // The same from a board's regions and its rows of colors - what a window holding no board is handed.
        public static (string Background, string Text)? Of(IReadOnlyList<string> regions, IReadOnlyList<RegionColorEntry> rows, string? region)
        {
            ArgumentNullException.ThrowIfNull(regions);
            ArgumentNullException.ThrowIfNull(rows);

            string wanted = region?.Trim() ?? string.Empty;

            if (wanted.Length == 0 || !regions.Contains(wanted, StringComparer.OrdinalIgnoreCase) ||
                rows.FirstOrDefault(row => string.Equals(row.Region?.Trim(), wanted, StringComparison.OrdinalIgnoreCase)) is not { } row ||
                RegionColors.Normalised(row.Background) is not { } background)
            {
                return null;
            }

            return (background, RegionColors.Normalised(row.Text) ?? RegionColors.ContrastText(background));
        }

        // The row `board` keeps for `region`, capitals aside - null for none.
        public static RegionColorEntry? RowOf(BoardData? board, string? region)
        {
            string wanted = region?.Trim() ?? string.Empty;

            return wanted.Length == 0
                ? null
                : board?.RegionColors.FirstOrDefault(row => string.Equals(row.Region?.Trim(), wanted, StringComparison.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // The text color that reads best on `background`: black or white, whichever has the higher contrast ratio with it (the
        // "highest visibility" asked for - a complementary HUE is often as dark or as light as the background, and reads worse).
        // Black for a background CRT cannot read.
        // ###########################################################################################
        public static string ContrastText(string? background)
        {
            if (RegionColors.Normalised(background) is not { } color)
            {
                return RegionColors.Black;
            }

            return RegionColors.ContrastRatio(color, RegionColors.White) > RegionColors.ContrastRatio(color, RegionColors.Black)
                ? RegionColors.White
                : RegionColors.Black;
        }

        // ###########################################################################################
        // WCAG 2's contrast ratio of two "#RRGGBB" colors, 1 to 21 - (lighter + 0.05) / (darker + 0.05) of their relative
        // luminance. 1 for a color CRT cannot read.
        // ###########################################################################################
        public static double ContrastRatio(string? first, string? second)
        {
            if (RegionColors.Normalised(first) is not { } one || RegionColors.Normalised(second) is not { } other)
            {
                return 1;
            }

            double a = RegionColors.Luminance(one);
            double b = RegionColors.Luminance(other);

            return (Math.Max(a, b) + 0.05) / (Math.Min(a, b) + 0.05);
        }

        private static double Luminance(string color)
        {
            static double Channel(string hex)
            {
                double value = int.Parse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture) / 255.0;

                return value <= 0.03928 ? value / 12.92 : Math.Pow((value + 0.055) / 1.055, 2.4);
            }

            return (0.2126 * Channel(color.Substring(1, 2))) + (0.7152 * Channel(color.Substring(3, 2))) + (0.0722 * Channel(color.Substring(5, 2)));
        }

        // ###########################################################################################
        // `board` with `region`'s colors set: `background` "#RRGGBB" (empty or unreadable: the region's own colors go, back to
        // CRT's), `text` "#RRGGBB" or empty for the one ContrastText works out. The same board for a region it does not have,
        // or when that is what it holds already. Written as the board spells the region, in the board's order of regions.
        // ###########################################################################################
        public static BoardData WithColors(BoardData board, string? region, string? background, string? text)
        {
            ArgumentNullException.ThrowIfNull(board);

            if (BoardRegions.SpellingOf(board, region) is not { } spelt)
            {
                return board;
            }

            string? wantedBackground = RegionColors.Normalised(background);
            string wantedText = wantedBackground is null ? string.Empty : RegionColors.Normalised(text) ?? string.Empty;
            RegionColorEntry? now = RegionColors.RowOf(board, spelt);

            // Nothing to change: the very board, so no step - never one that only writes the list down below.
            if (wantedBackground is null ? now is null
                : now is not null &&
                  string.Equals(RegionColors.Normalised(now.Background), wantedBackground, StringComparison.Ordinal) &&
                  string.Equals(RegionColors.Normalised(now.Text) ?? string.Empty, wantedText, StringComparison.Ordinal))
            {
                return board;
            }

            // A board whose regions its rows decide (no list written) gets its list written down first: colors are given for a
            // region the board SAYS it has, the rule the checks hold a submission to. Its rows of colors are shared, so `now`
            // is still the row to replace.
            if (wantedBackground is not null && board.Regions.Count == 0)
            {
                board = board.WithRegions([.. BoardRegions.Of(board)]);
            }

            List<RegionColorEntry> rows = board.RegionColors.Where(row => !ReferenceEquals(row, now)).ToList();

            if (wantedBackground is not null)
            {
                rows.Add(new RegionColorEntry { Region = spelt, Background = wantedBackground, Text = wantedText });
            }

            return board.WithRegionColors(RegionColors.InRegionOrder(board, rows));
        }

        // The rows of the regions the board has, in its order of regions - a row of a region it no longer has goes.
        private static List<RegionColorEntry> InRegionOrder(BoardData board, IEnumerable<RegionColorEntry> rows)
        {
            List<RegionColorEntry> kept = rows.ToList();

            return BoardRegions.Of(board)
                .Select(region => kept.FirstOrDefault(row => string.Equals(row.Region?.Trim(), region, StringComparison.OrdinalIgnoreCase)))
                .OfType<RegionColorEntry>()
                .ToList();
        }

        // ###########################################################################################
        // `board` with the colors of the regions it no longer has let go of - the same board when there are none. A region
        // deleted, or every region taken away, takes its colors with it (BoardRegions.WithoutRegion, WithRegions).
        // ###########################################################################################
        public static BoardData Tidied(BoardData board)
        {
            ArgumentNullException.ThrowIfNull(board);

            if (board.RegionColors.Count == 0)
            {
                return board;
            }

            List<RegionColorEntry> kept = RegionColors.InRegionOrder(board, board.RegionColors);

            return kept.Count == board.RegionColors.Count && kept.SequenceEqual(board.RegionColors)
                ? board
                : board.WithRegionColors(kept);
        }

        // ###########################################################################################
        // Whether two boards differ in their regions' colors - a change a draft has to submit, and a reviewer is told of. The
        // colors as drawn: a row whose text is the one worked out anyway is the same as one with none.
        // ###########################################################################################
        public static bool Differ(BoardData? first, BoardData? second) =>
            !RegionColors.Said(first).SequenceEqual(RegionColors.Said(second), StringComparer.Ordinal);

        private static IEnumerable<string> Said(BoardData? board) =>
            (board?.RegionColors ?? [])
                .Select(row => (Region: row.Region?.Trim() ?? string.Empty, Background: RegionColors.Normalised(row.Background), Text: RegionColors.Normalised(row.Text)))
                .Where(row => row.Region.Length > 0 && row.Background is not null)
                .Select(row => $"{row.Region.ToUpperInvariant()}|{row.Background}|{row.Text ?? RegionColors.ContrastText(row.Background)}")
                .Order(StringComparer.Ordinal);

        // ###########################################################################################
        // What is wrong with a submission's region colors (`rows`) against the regions it says (`regions`) - one sentence per
        // row, nothing for none: a region it does not say, one said twice, a background or a text not "#RRGGBB".
        // ###########################################################################################
        public static List<(string Region, string Problem)> Problems(IReadOnlyList<RegionColorEntry>? rows, IReadOnlyList<string?>? regions)
        {
            var problems = new List<(string Region, string Problem)>();
            List<string> said = BoardRegions.Cleaned(regions ?? []);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (RegionColorEntry? row in rows ?? [])
            {
                string region = row?.Region?.Trim() ?? string.Empty;

                if (row is null || !said.Contains(region, StringComparer.OrdinalIgnoreCase))
                {
                    problems.Add((region, $"Colors are given for [{region}], which is not one of the board's regions."));
                }
                else if (!seen.Add(region))
                {
                    problems.Add((region, $"Colors are given for [{region}] more than once."));
                }
                else if (!RegionColors.IsColor(row.Background))
                {
                    problems.Add((region, $"The background color of [{region}], [{row.Background}], is not a color written as #RRGGBB."));
                }
                else if (!string.IsNullOrWhiteSpace(row.Text) && !RegionColors.IsColor(row.Text))
                {
                    problems.Add((region, $"The text color of [{region}], [{row.Text}], is not a color written as #RRGGBB."));
                }
            }

            return problems;
        }
    }
}
