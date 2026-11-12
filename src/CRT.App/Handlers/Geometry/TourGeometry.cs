using Avalonia;
using System;

namespace Handlers.Geometry
{
    // Where a tour's card stands against what its step frames (Main/TourOverlay).
    internal enum TourCardSide
    {
        // No frame: the card in the middle of the window.
        Centre,
        Right,
        Left,
        Below,
        Above,

        // The frame is too large for the card to stand beside it (the picture): inside it, lower right.
        Inside,
    }

    // ###########################################################################################
    // Every position of an introduction tour (Main/TourOverlay, owner request 2026-11-11): the frame - the "hole" in
    // the dimmed window - round what a step is about, the card beside it, and the small pointer on the card's edge
    // towards it. Pure, in the overlay's own coordinates.
    //
    //   - the hole is the target a little larger, but never past the window's edge, so its frame is seen whole (the
    //     thumbnails stand against the right edge);
    //   - the card stands beside the hole where it fits whole - right, left, below, above, in that order - level with
    //     its top when the hole is the taller, centred on it otherwise; failing all four, inside it, lower right, clear of
    //     the buttons a picture has along its foot ("Draw trace");
    //   - nowhere is the card past the window's edge, however small the window.
    // ###########################################################################################
    internal static class TourGeometry
    {
        // How much larger than its target the hole is.
        public const double HolePadding = 6;

        // How far the hole keeps from the window's edge.
        public const double WindowEdge = 3;

        // How far the card keeps from the window's edge.
        public const double EdgeMargin = 12;

        // Between the hole and the card beside it - room for the pointer.
        public const double Gap = 18;

        // A card inside the hole: from its right edge, and from its foot.
        public const double InsideRight = 28;
        public const double InsideFoot = 64;

        // The pointer: half its base, how far it reaches, and how near a corner of the card it may stand.
        public const double PointerHalf = 9;
        public const double PointerDepth = 10;
        public const double PointerInset = 18;

        // The corners of the hole and its frame.
        public const double HoleRadius = 7;

        // ###########################################################################################
        // The hole round `target` in an overlay of `area` - null with no target, or one wholly outside the window.
        // ###########################################################################################
        public static Rect? HoleFor(Rect? target, Size area)
        {
            if (target is not Rect rect || area.Width <= 0 || area.Height <= 0)
                return null;

            Rect inside = new Rect(area).Deflate(TourGeometry.WindowEdge);
            Rect hole = rect.Inflate(TourGeometry.HolePadding).Intersect(inside);

            return hole.Width > 0 && hole.Height > 0 ? hole : null;
        }

        // ###########################################################################################
        // Where a card of `card` stands against `hole` in an overlay of `area`, and on which side.
        // ###########################################################################################
        public static (Rect Card, TourCardSide Side) PlaceCard(Rect? hole, Size card, Size area)
        {
            if (hole is not Rect frame)
            {
                return (TourGeometry.Kept(new Rect((area.Width - card.Width) / 2, (area.Height - card.Height) / 2.4, card.Width, card.Height), area), TourCardSide.Centre);
            }

            double beside = Math.Min(frame.Center.Y - card.Height / 2, frame.Y);
            double around = frame.Center.X - card.Width / 2;

            if (frame.Right + TourGeometry.Gap + card.Width + TourGeometry.EdgeMargin <= area.Width)
            {
                return (TourGeometry.Kept(new Rect(frame.Right + TourGeometry.Gap, beside, card.Width, card.Height), area), TourCardSide.Right);
            }

            if (frame.X - TourGeometry.Gap - card.Width >= TourGeometry.EdgeMargin)
            {
                return (TourGeometry.Kept(new Rect(frame.X - TourGeometry.Gap - card.Width, beside, card.Width, card.Height), area), TourCardSide.Left);
            }

            if (frame.Bottom + TourGeometry.Gap + card.Height + TourGeometry.EdgeMargin <= area.Height)
            {
                return (TourGeometry.Kept(new Rect(around, frame.Bottom + TourGeometry.Gap, card.Width, card.Height), area), TourCardSide.Below);
            }

            if (frame.Y - TourGeometry.Gap - card.Height >= TourGeometry.EdgeMargin)
            {
                return (TourGeometry.Kept(new Rect(around, frame.Y - TourGeometry.Gap - card.Height, card.Width, card.Height), area), TourCardSide.Above);
            }

            return (TourGeometry.Kept(
                new Rect(frame.Right - card.Width - TourGeometry.InsideRight, frame.Bottom - card.Height - TourGeometry.InsideFoot, card.Width, card.Height),
                area), TourCardSide.Inside);
        }

        // ###########################################################################################
        // The pointer on the card's edge facing the hole - its base on that edge, its tip towards the hole, as near the
        // hole's middle as the card's corners allow. Null for a card in the middle or inside the hole.
        // ###########################################################################################
        public static (Point A, Point Tip, Point B)? PointerFor(Rect hole, Rect card, TourCardSide side)
        {
            double Along(double middle, double from, double to) =>
                Math.Clamp(middle, from + TourGeometry.PointerInset, Math.Max(from + TourGeometry.PointerInset, to - TourGeometry.PointerInset));

            switch (side)
            {
                case TourCardSide.Right:
                {
                    double y = Along(hole.Center.Y, card.Y, card.Bottom);
                    return (new Point(card.X + 1, y - TourGeometry.PointerHalf), new Point(card.X - TourGeometry.PointerDepth, y), new Point(card.X + 1, y + TourGeometry.PointerHalf));
                }

                case TourCardSide.Left:
                {
                    double y = Along(hole.Center.Y, card.Y, card.Bottom);
                    return (new Point(card.Right - 1, y - TourGeometry.PointerHalf), new Point(card.Right + TourGeometry.PointerDepth, y), new Point(card.Right - 1, y + TourGeometry.PointerHalf));
                }

                case TourCardSide.Below:
                {
                    double x = Along(hole.Center.X, card.X, card.Right);
                    return (new Point(x - TourGeometry.PointerHalf, card.Y + 1), new Point(x, card.Y - TourGeometry.PointerDepth), new Point(x + TourGeometry.PointerHalf, card.Y + 1));
                }

                case TourCardSide.Above:
                {
                    double x = Along(hole.Center.X, card.X, card.Right);
                    return (new Point(x - TourGeometry.PointerHalf, card.Bottom - 1), new Point(x, card.Bottom + TourGeometry.PointerDepth), new Point(x + TourGeometry.PointerHalf, card.Bottom - 1));
                }

                default:
                    return null;
            }
        }

        // Inside the window's edge, as far as the card fits at all - its top left kept in sight first.
        private static Rect Kept(Rect card, Size area)
        {
            double x = Math.Max(TourGeometry.EdgeMargin, Math.Min(card.X, area.Width - card.Width - TourGeometry.EdgeMargin));
            double y = Math.Max(TourGeometry.EdgeMargin, Math.Min(card.Y, area.Height - card.Height - TourGeometry.EdgeMargin));

            return new Rect(x, y, card.Width, card.Height);
        }
    }
}
