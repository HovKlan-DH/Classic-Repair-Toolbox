using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // A PAGE THAT FILLS ITS SCROLL VIEWER WHILE THERE IS ROOM, AND SCROLLS ONCE THERE IS NOT (owner report, 2026-11-25:
    // on a low window the "Feedback" tab's text box was squeezed to nothing and "Send feedback" lay below the tab, with
    // nothing to scroll - "it makes the UI useless in these situation").
    //
    // One part of the page stretches - the "Feedback" tab's text box - and the rest is as high as it needs. Inside a
    // scroll viewer a page is measured with no limit, so a stretching part would grow with what it holds (a long text
    // would stretch the page, not scroll in its box); the page is given a height of its own instead: the viewport's
    // while that leaves the stretching part its least, else just enough for the rest and that least.
    // ###########################################################################################
    public static class ScrollingPageGeometry
    {
        // The height to give a page whose fixed parts are `fixedParts` high and whose stretching part needs at least
        // `stretchingLeast`, in a scroll viewer `viewport` high. An unknown viewport (not laid out yet) asks no more
        // than the page needs.
        public static double PageHeight(double viewport, double fixedParts, double stretchingLeast)
        {
            double least = Math.Max(0, fixedParts) + Math.Max(0, stretchingLeast);

            if (double.IsNaN(viewport) || double.IsInfinity(viewport) || viewport <= 0)
            {
                return least;
            }

            return Math.Max(viewport, least);
        }
    }
}
