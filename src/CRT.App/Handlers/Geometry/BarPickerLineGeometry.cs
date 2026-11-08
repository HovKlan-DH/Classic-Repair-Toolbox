using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // A LINE OF THE WORKLOG BAR'S DROP-DOWNS: ITS WORDS, THEN ITS BADGE - WHILE THERE IS ROOM FOR BOTH (owner
    // request, 2026-11-06: "Inside each of the drop-downs there should be a badge shown as the last in each
    // line"). The closed drop-down shows its line too, and on a narrow window the device's box was all badge -
    // "Working", no "C64" - as the badge kept its width and the words got what was left, nothing. So the words
    // keep at least LeastOfWords of themselves (all of them when shorter), and the badge gives way below that.
    // Pure: BarPickerLinePanel lays a line out by it.
    // ###########################################################################################
    internal static class BarPickerLineGeometry
    {
        // What the words keep before the badge goes - some five letters at the bar's 12px.
        public const double LeastOfWords = 40;

        // Whether a line `width` wide has room for the badge beside the words. An unbounded width (the open
        // list measuring its widest line) always has.
        public static bool ShowsBadge(double width, double wordsWidth, double badgeWidth) =>
            double.IsPositiveInfinity(width) || width >= badgeWidth + Math.Min(Math.Max(0, wordsWidth), LeastOfWords);

        // How wide the words are drawn: the whole line, less the badge while it shows.
        public static double WordsWidth(double width, double badgeWidth, bool showsBadge) =>
            Math.Max(0, showsBadge ? width - badgeWidth : width);

        // ###########################################################################################
        // A BADGE BEFORE THE WORDS - the device a workbook is on, while the "Workbook" drop-down lists every device's
        // (owner request, 2026-11-07: "show the device name first in a badge ... so I can separate the workbooks in the
        // list"). It takes its own width up to LeadingShare of the line, cut short beyond that, so a long device name
        // never leaves the workbook's own words no room. Unbounded, its whole width.
        // ###########################################################################################
        public const double LeadingShare = 0.4;

        public static double LeadingWidth(double width, double leadingWidth) =>
            double.IsPositiveInfinity(width)
                ? Math.Max(0, leadingWidth)
                : Math.Clamp(leadingWidth, 0, Math.Max(0, width) * LeadingShare);
    }
}
