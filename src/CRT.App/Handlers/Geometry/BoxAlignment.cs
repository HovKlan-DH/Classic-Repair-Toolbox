using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // A BOX STRETCHED TO END WHERE A BOX IN THE NEXT COLUMN ENDS (owner request, 2026-11-09: "For the new "Purpose or note"
    // text field - can it vertically align to the end-border of the "Other technical details" in the 1st column?";
    // ComponentPane.AlignPurposeWithDescription). Pure, so a test checks it.
    // ###########################################################################################
    internal static class BoxAlignment
    {
        // ###########################################################################################
        // The least height a box needs for its bottom to meet `otherBottom` - both measured from the box's own top - never
        // under `floor`, its height of its own. Its text may make it taller still: this is a least height, never a cap.
        // ###########################################################################################
        public static double MinHeightToMeet(double otherBottom, double floor)
        {
            if (double.IsNaN(otherBottom) || double.IsInfinity(otherBottom))
            {
                return floor;
            }

            return Math.Max(floor, Math.Round(otherBottom));
        }

        // Whether a least height already set is the one wanted - a layout pass changing nothing must not ask for another.
        public static bool IsAlready(double current, double wanted) => Math.Abs(current - wanted) < 0.5;
    }
}
