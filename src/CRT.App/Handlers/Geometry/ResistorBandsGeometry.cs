using Avalonia;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // WHERE THE RESISTOR PICTURE'S PARTS STAND (ResistorBandsDrawing, the component popup - owner request, 2026-11-07,
    // after the picture the owner makes by hand for every value): the body with its wider end caps and the lead stubs at
    // either end, each band across it, a faint stem from each band down to a label pill in the band's color. Pure, so a
    // test checks it: every position is a share of the width the picture is drawn at.
    //
    // The digits and the multiplier stand from the left cap evenly to the middle, the tolerance apart near the right
    // end, as a resistor is printed (and the owner's picture shows). The pills are centred under their bands, pushed
    // apart when a wide one ("x100 kΩ") would touch its neighbour, and kept inside the picture.
    // ###########################################################################################
    internal static class ResistorBandsGeometry
    {
        // The picture's height for its width - body, stems and pills.
        public const double BodyShare = 0.19;
        public const double StemShare = 0.09;
        public const double PillHeight = 18;
        public const double PillGap = 4;

        // Where on the body (0 its left edge, 1 its right) the first band and the band before the tolerance stand, and
        // the tolerance's.
        public const double FirstBand = 0.115;
        public const double LastBeforeTolerance = 0.55;
        public const double ToleranceBand = 0.88;

        public sealed record Layout(
            Rect Body,
            double CapWidth,
            double MiddleInset,
            Rect LeftLead,
            Rect RightLead,
            IReadOnlyList<Rect> Bands,
            double PillTop,
            double Height);

        // ###########################################################################################
        // The parts for `bandCount` bands at `width`; the last is the tolerance when `withTolerance`.
        // ###########################################################################################
        public static Layout For(double width, int bandCount, bool withTolerance)
        {
            width = Math.Max(1, width);

            double bodyLeft = width * 0.04;
            double bodyWidth = width * 0.92;
            double bodyHeight = width * BodyShare;
            var body = new Rect(bodyLeft, 0, bodyWidth, bodyHeight);

            double leadHeight = bodyHeight * 0.28;
            double leadTop = (bodyHeight - leadHeight) / 2;
            var leftLead = new Rect(0, leadTop, bodyLeft + 1, leadHeight);
            var rightLead = new Rect(bodyLeft + bodyWidth - 1, leadTop, width - (bodyLeft + bodyWidth) + 1, leadHeight);

            double bandWidth = bodyWidth * 0.06;
            var bands = new List<Rect>();

            foreach (double at in BandShares(bandCount, withTolerance))
            {
                double centre = bodyLeft + bodyWidth * at;
                bands.Add(new Rect(centre - bandWidth / 2, 0, bandWidth, bodyHeight));
            }

            double pillTop = bodyHeight + width * StemShare;

            return new Layout(body, bodyWidth * 0.16, bodyHeight * 0.11, leftLead, rightLead, bands, pillTop, pillTop + PillHeight + 1);
        }

        // ###########################################################################################
        // Each band's place on the body, 0 to 1: the digits and the multiplier from FirstBand evenly to
        // LastBeforeTolerance, the tolerance at ToleranceBand.
        // ###########################################################################################
        public static IReadOnlyList<double> BandShares(int bandCount, bool withTolerance)
        {
            int even = Math.Max(0, withTolerance ? bandCount - 1 : bandCount);
            var shares = new List<double>();

            for (int i = 0; i < even; i++)
            {
                shares.Add(even == 1 ? FirstBand : FirstBand + i * (LastBeforeTolerance - FirstBand) / (even - 1));
            }

            if (withTolerance && bandCount > 0)
            {
                shares.Add(ToleranceBand);
            }

            return shares;
        }

        // ###########################################################################################
        // The pills' left edges, left to right: each centred under its band at `centres`, `widths` wide, moved right
        // when it would come within PillGap of the one before, then the row moved left as a whole when it runs past
        // `width` - never past the left edge.
        // ###########################################################################################
        public static IReadOnlyList<double> PillLefts(IReadOnlyList<double> centres, IReadOnlyList<double> widths, double width)
        {
            ArgumentNullException.ThrowIfNull(centres);
            ArgumentNullException.ThrowIfNull(widths);

            var lefts = new List<double>(centres.Count);
            double nextFree = 0;

            for (int i = 0; i < centres.Count; i++)
            {
                double left = Math.Max(nextFree, centres[i] - widths[i] / 2);
                lefts.Add(left);
                nextFree = left + widths[i] + PillGap;
            }

            if (lefts.Count > 0)
            {
                double overrun = lefts[^1] + widths[^1] - width;

                if (overrun > 0)
                {
                    double shift = Math.Min(overrun, lefts[0]);
                    for (int i = 0; i < lefts.Count; i++)
                    {
                        lefts[i] -= shift;
                    }
                }
            }

            return lefts;
        }

        // Where a band's stem ends at the top: the body's foot at that band - a cap's foot is lower than the middle's.
        public static double StemTop(Layout layout, Rect band)
        {
            ArgumentNullException.ThrowIfNull(layout);

            double centre = band.Center.X;
            bool onCap = centre <= layout.Body.Left + layout.CapWidth || centre >= layout.Body.Right - layout.CapWidth;

            return onCap ? layout.Body.Bottom : layout.Body.Bottom - layout.MiddleInset;
        }

        // Whether two pills overlap - for the tests' own reading of PillLefts.
        internal static bool Overlap(IReadOnlyList<double> lefts, IReadOnlyList<double> widths) =>
            lefts.Zip(widths).Zip(lefts.Skip(1)).Any(pair => pair.First.First + pair.First.Second > pair.Second);
    }
}
