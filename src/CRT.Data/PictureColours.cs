using System;
using System.Globalization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // HOW A MARKED COMPONENT IS DRAWN ON ONE OF THE BOARD'S PICTURES (CRT 4.0 step 2, slice 4b - the
    // plan: "Board > Pictures: a colour picker and opacity slider per picture, with a live preview on
    // the picture, plus 'Use for all pictures'"). Until now only the table, as text, reached them.
    //
    // As the board file keeps them, as text: a colour by its name or as "#RRGGBB" ("Red",
    // "DodgerBlue" - the live trees use names), an opacity as a percentage ("30%"). Blank is CRT's
    // own default, drawn as it always has been (RectGeometry.ParseColorOrDefault). The words are
    // compared trimmed: a colour picked again as it was is no change.
    // ###########################################################################################
    public sealed record PictureColours(
        string HighlightColor,
        string HighlightOpacity,
        string OppositeTraceColor,
        string ThumbnailHighlightColor,
        string ThumbnailHighlightOpacity)
    {
        // Every value trimmed, as it is written.
        public PictureColours Trimmed() => new(
            this.HighlightColor?.Trim() ?? string.Empty,
            this.HighlightOpacity?.Trim() ?? string.Empty,
            this.OppositeTraceColor?.Trim() ?? string.Empty,
            this.ThumbnailHighlightColor?.Trim() ?? string.Empty,
            this.ThumbnailHighlightOpacity?.Trim() ?? string.Empty);

        // ###########################################################################################
        // An opacity as the board keeps it: a whole percentage, "30%" - from a fraction (0.3) or a
        // percentage (30), clamped to 0..100. What the pane's slider writes.
        // ###########################################################################################
        public static string OpacityText(double value)
        {
            double percent = value <= 1.0 ? value * 100.0 : value;
            int whole = (int)Math.Round(Math.Clamp(percent, 0.0, 100.0), MidpointRounding.AwayFromZero);

            return whole.ToString(CultureInfo.InvariantCulture) + "%";
        }
    }
}
