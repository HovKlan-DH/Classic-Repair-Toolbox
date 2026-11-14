using Avalonia;
using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // WHERE BadgeCloud's BADGES STAND - as the WrapPanel they replaced laid them out (2026-11-13): left to right
    // with `itemSpacing` between them, a new line when the next one does not fit in `width`, `lineSpacing` between
    // lines, each line as tall as its tallest badge and every badge of a line that tall. A badge wider than the width
    // stands alone on its line. An infinite width is one line.
    //
    // PURE: the badges' sizes in, their places and the size of all of them out.
    // ###########################################################################################
    internal static class BadgeCloudGeometry
    {
        public static (List<Rect> Places, Size Size) Place(IReadOnlyList<Size> badges, double width, double itemSpacing, double lineSpacing)
        {
            ArgumentNullException.ThrowIfNull(badges);

            var places = new List<Rect>(badges.Count);
            double room = double.IsNaN(width) || width <= 0 ? double.PositiveInfinity : width;

            double lineTop = 0;
            double widest = 0;
            int lineStart = 0;
            double x = 0;
            double lineHeight = 0;

            void EndLine(int end)
            {
                for (int index = lineStart; index < end; index++)
                {
                    places[index] = new Rect(places[index].X, lineTop, places[index].Width, lineHeight);
                }

                widest = Math.Max(widest, x);
            }

            for (int index = 0; index < badges.Count; index++)
            {
                Size badge = badges[index];
                double startX = index == lineStart ? 0 : x + itemSpacing;

                if (index > lineStart && startX + badge.Width > room)
                {
                    EndLine(index);
                    lineTop += lineHeight + lineSpacing;
                    lineStart = index;
                    startX = 0;
                    lineHeight = 0;
                }

                places.Add(new Rect(startX, lineTop, badge.Width, badge.Height));
                x = startX + badge.Width;
                lineHeight = Math.Max(lineHeight, badge.Height);
            }

            if (badges.Count == 0)
            {
                return (places, default);
            }

            EndLine(badges.Count);

            return (places, new Size(widest, lineTop + lineHeight));
        }
    }
}
