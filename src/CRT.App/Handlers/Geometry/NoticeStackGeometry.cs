using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // How wide the notice cards over the tabs' bottom-right corner may be (Main.Notices.cs).
    //
    // At most HALF the tabs' width, so on the Schematics tab a card stays over the thumbnail list
    // at the default split. Rendered at CRT's smallest window, a full-width card ran over the
    // "Global settings" and "Labels visible" panels at the bottom of the schematic - and the label
    // editor's card cannot be closed, so they stayed covered for the whole edit. In a big window
    // the cap is Widest, which is plenty for a sentence and two buttons.
    // ###########################################################################################
    public static class NoticeStackGeometry
    {
        // The gap between the cards and the tabs' right and bottom edges.
        public const double EdgeMargin = 12;

        // A card is never wider than this, however wide the window.
        public const double Widest = 460;

        // ...and never narrower than this while the tabs have room for it - narrower, a card is a
        // column of single words. Tabs narrower still squeeze it anyway, as layout always does.
        public const double Narrowest = 240;

        // ###########################################################################################
        // The stack's MaxWidth for tabs `tabsWidth` wide: half of them less the edge margin,
        // within Narrowest..Widest.
        // ###########################################################################################
        public static double MaxWidth(double tabsWidth)
        {
            if (!double.IsFinite(tabsWidth))
            {
                return NoticeStackGeometry.Widest;
            }

            return Math.Clamp(
                tabsWidth / 2 - NoticeStackGeometry.EdgeMargin,
                NoticeStackGeometry.Narrowest,
                NoticeStackGeometry.Widest);
        }
    }
}
