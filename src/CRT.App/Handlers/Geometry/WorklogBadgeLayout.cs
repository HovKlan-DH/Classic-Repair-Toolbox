using Avalonia;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Where a worklog entry's "#N" badge goes on a schematic preview.
    //
    // Two cases, and getting them the wrong way round is a REPORTED BUG rather than a hypothetical
    // one: an entry with "show marked area" ticked has a rectangle on the board to anchor its badge
    // to, and one WITHOUT has nothing to anchor to and so parks in the image's top-right corner
    // instead, stacking with the other parked badges rather than overlapping them. An earlier
    // version anchored every badge to its marker regardless, so an entry meant to show only a parked
    // pill still appeared pinned to the spot it happened to be drawn at.
    //
    // Pure, and here rather than inside the tab because that is the difference between a rule
    // verified by a fast unit test and one verified only by driving a full headless layout pass and
    // reading pixel positions back off controls.
    // ###########################################################################################
    public static class WorklogBadgeLayout
    {
        // One badge's placement request: its measured size, and the pixel rect it anchors to, or
        // null when it has none and must be parked.
        public readonly record struct BadgePlacementRequest(Size DesiredSize, Rect? AnchorPixelRect);

        // ###########################################################################################
        // Lays out every badge for one preview against the image's drawn content rect.
        //
        // Returns one point per input, in the SAME order, so the caller can zip the results straight
        // back onto its controls without tracking which went where.
        //
        // Anchored badges are centred on their marked area, converted from bitmap pixels into the
        // image's own local coordinates. Parked ones are stacked in the top-right of that same
        // content rect - NOT of the control's bounds, which is a different rect whenever the
        // Uniform-stretched image is letterboxed inside its control, and using the latter put parked
        // badges outside the drawn image while anchored ones stayed correct.
        // ###########################################################################################
        public static List<Point> ArrangeBadges(
            IReadOnlyList<BadgePlacementRequest> requests,
            Rect contentRect,
            PixelSize bitmapPixelSize,
            double parkedMargin,
            double parkedSpacing)
        {
            var positions = new List<Point>(requests?.Count ?? 0);
            if (requests == null || requests.Count == 0)
            {
                return positions;
            }

            // Index within `requests` of each parked badge, so its arranged position can be written
            // back into the right slot of the result.
            var parkedIndexes = new List<int>();
            var parkedSizes = new List<Size>();

            for (int i = 0; i < requests.Count; i++)
            {
                var request = requests[i];

                if (request.AnchorPixelRect is { } pixelRect)
                {
                    var localRect = RectGeometry.PixelToLocalRect(pixelRect, contentRect, bitmapPixelSize);
                    var offset = BadgeGeometry.GetCenterScaledCentreOffset(request.DesiredSize);
                    positions.Add(new Point(localRect.Left + offset.X, localRect.Top + offset.Y));
                    continue;
                }

                // Placeholder, overwritten below - keeps the result aligned with the input order.
                positions.Add(default);
                parkedIndexes.Add(i);
                parkedSizes.Add(request.DesiredSize);
            }

            if (parkedIndexes.Count == 0)
            {
                return positions;
            }

            // The SAME geometry the real Schematics tab's own parked pills use. reservedRight is 0
            // for this pane: it has no equivalent of that view's "Netlist names" side panel to leave
            // room for.
            var parkedPositions = ParkedBadgeGeometry.ArrangeInTopRightBlock(
                parkedSizes, contentRect.Size, parkedMargin, parkedSpacing, reservedRight: 0.0);

            for (int i = 0; i < parkedIndexes.Count && i < parkedPositions.Count; i++)
            {
                positions[parkedIndexes[i]] = parkedPositions[i];
            }

            return positions;
        }
    }
}
