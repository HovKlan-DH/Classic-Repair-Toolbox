using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Handlers.Geometry;
using Handlers.Theming;

namespace CRT
{
    // ###########################################################################################
    // What an introduction tour draws UNDER its card (Main/TourOverlay): the window dimmed but for a hole round what
    // the step is about, an amber frame with a soft glow round the hole, and the pointer from the card towards it. A
    // control of its own because Panel.Render is sealed in Avalonia 12. Colors are the theme's Tour_* resources.
    // ###########################################################################################
    internal sealed class TourShade : Control
    {
        private Rect? thisHole;
        private Rect thisCard;
        private TourCardSide thisSide = TourCardSide.Centre;

        internal Rect? Hole => this.thisHole;

        public TourShade()
        {
            this.IsHitTestVisible = false;
            this.ActualThemeVariantChanged += (_, _) => this.InvalidateVisual();
        }

        // The hole, the card and its side - drawn again only when one of them moved.
        internal void Show(Rect? hole, Rect card, TourCardSide side)
        {
            if (Equals(hole, this.thisHole) && card == this.thisCard && side == this.thisSide)
                return;

            this.thisHole = hole;
            this.thisCard = card;
            this.thisSide = side;
            this.InvalidateVisual();
        }

        public override void Render(DrawingContext context)
        {
            var whole = new Rect(this.Bounds.Size);
            IBrush dim = this.Brush("Tour_Dim", new SolidColorBrush(Color.FromArgb(150, 0, 0, 0)));

            if (this.thisHole is Rect hole)
            {
                Geometry shade = new CombinedGeometry(
                    GeometryCombineMode.Exclude,
                    new RectangleGeometry(whole),
                    new RectangleGeometry(hole, TourGeometry.HoleRadius, TourGeometry.HoleRadius));
                context.DrawGeometry(dim, null, shade);

                Color frame = this.Brush("Tour_Frame", Brushes.Orange) is ISolidColorBrush solid ? solid.Color : Colors.Orange;

                // A soft glow, wider and fainter outwards, then the frame itself.
                for (int ring = 4; ring >= 1; ring--)
                {
                    context.DrawRectangle(null, new Pen(new SolidColorBrush(frame, 0.10), 2 + ring * 3), hole, TourGeometry.HoleRadius, TourGeometry.HoleRadius);
                }

                context.DrawRectangle(null, new Pen(new SolidColorBrush(frame), 3), hole, TourGeometry.HoleRadius, TourGeometry.HoleRadius);

                if (TourGeometry.PointerFor(hole, this.thisCard, this.thisSide) is var (a, tip, b))
                {
                    var pointer = new StreamGeometry();

                    using (StreamGeometryContext path = pointer.Open())
                    {
                        path.BeginFigure(a, true);
                        path.LineTo(tip);
                        path.LineTo(b);
                        path.EndFigure(true);
                    }

                    context.DrawGeometry(this.Brush("Tour_Card_Bg", Brushes.White), null, pointer);
                }
            }
            else
            {
                context.DrawRectangle(dim, null, whole);
            }
        }

        // CRT's one theme lookup (code review, 2026-11-12: a copy of its own here, which ThemeResources' header warns of).
        private IBrush Brush(string key, IBrush fallback) => ThemeResources.ResolveForControl(this, key, fallback);
    }
}
