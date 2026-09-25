using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Small viewport and collection helpers extracted from TabSchematics so they can be unit
    // tested. Nothing here touches a control.
    // ###########################################################################################
    public static class ViewportMath
    {
        // ###########################################################################################
        // Converts one wheel event into a zoom factor.
        //
        // Linux and macOS deliver many small high-resolution wheel deltas where Windows delivers
        // one notch of magnitude 1.0; treating every event as a full step made zooming very coarse
        // and aggressive there. A delta magnitude of exactly 1.0 reduces this to exactly baseFactor.
        // ###########################################################################################
        public static double ComputeWheelZoomFactor(double deltaY, double baseFactor)
        {
            double magnitude = Math.Clamp(Math.Abs(deltaY), 0.1, 3.0);
            double factor = Math.Pow(baseFactor, magnitude);
            return deltaY > 0 ? factor : 1.0 / factor;
        }

        // ###########################################################################################
        // The range the schematic view's translation may take on one axis at a given zoom scale.
        //
        // Everything is in container coordinates, and the identity matrix is the fitted "first
        // view" - the image sitting at the top left of the container at scale 1, exactly as
        // Stretch="Uniform" laid it out. Two rules are combined and the wider of the two wins:
        //
        //  * The edge rule: the image may not be moved so far that empty space appears at an edge
        //    it could have covered. Where the scaled image is larger than the visible viewport
        //    this pins it to the viewport edges; where it is smaller it keeps it inside them.
        //    This is what lets an image be panned out from behind an edge-docked overlay panel,
        //    which narrows the viewport without changing the fitted image.
        //
        //  * The anchor rule: every position a cursor-anchored zoom can produce must be legal, or
        //    the point under the cursor cannot stay under the cursor. Zooming by `scale` about a
        //    container point `a` starting from the fitted view lands the translation on
        //    a * (1 - scale), and the cursor can be anywhere on the image, so everything between
        //    (1 - scale) * contentEnd and (1 - scale) * contentStart has to be reachable.
        //
        // The anchor rule only ever adds room on a *letterboxed* axis - the one where the fitted
        // image does not fill the container, leaving an empty band beside or below it. On the axis
        // that does fill it the two rules produce the same numbers. That asymmetry is exactly why
        // anchored zoom used to work horizontally on a landscape schematic and drift vertically:
        // the edge rule alone insisted the empty band stay put, which forced the translation back
        // towards zero and slid the image out from under the cursor.
        //
        // The far edge of a letterboxed image can therefore never rise above where the fitted view
        // put it, so zooming can never reveal more empty space than the first view already showed.
        // ###########################################################################################
        public static (double Min, double Max) ComputeAxisTranslationRange(
            double viewportStart,
            double viewportEnd,
            double contentStart,
            double contentEnd,
            double scale)
        {
            double startAligned = viewportStart - (scale * contentStart);
            double endAligned = viewportEnd - (scale * contentEnd);

            double anchoredOnStart = (1.0 - scale) * contentStart;
            double anchoredOnEnd = (1.0 - scale) * contentEnd;

            double min = Math.Min(Math.Min(startAligned, endAligned), Math.Min(anchoredOnStart, anchoredOnEnd));
            double max = Math.Max(Math.Max(startAligned, endAligned), Math.Max(anchoredOnStart, anchoredOnEnd));

            return (min, max);
        }

        // ###########################################################################################
        // Case-insensitive set equality for two string collections, without allocating when the
        // sizes already differ. Used to decide whether a selection actually changed.
        // ###########################################################################################
        public static bool SetEqualsOrdinalIgnoreCase(
            IReadOnlyCollection<string> left,
            IReadOnlyCollection<string> right)
        {
            if (ReferenceEquals(left, right))
            {
                return true;
            }

            if (left.Count != right.Count)
            {
                return false;
            }

            if (left.Count == 0)
            {
                return true;
            }

            var leftSet = new HashSet<string>(left, StringComparer.OrdinalIgnoreCase);
            return leftSet.SetEquals(right);
        }
    }
}
