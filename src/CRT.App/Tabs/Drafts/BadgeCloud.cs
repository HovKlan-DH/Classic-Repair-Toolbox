using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Media.TextFormatting;
using Handlers.Geometry;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // An item a BadgeCloud draws with more than its text: an important KiCad signal's net, say (KiCadNetNameBadge).
    public interface IBadgeCloudItem
    {
        string Text { get; }

        // Drawn filled green, its text white and semibold - "already used".
        bool IsUsed { get; }

        // Said under the badge while the pointer is on it - null for nothing.
        string? Tip { get; }
    }

    // How a badge is drawn - which of the theme's "Draft_ReferenceBadge" colors.
    public enum BadgeCloudKind
    {
        Plain,
        Todo,
        Used,
    }

    // ###########################################################################################
    // A LIST OF BADGES DRAWN AS ONE CONTROL (2026-11-13, measured): "KiCad: Files" lists every component the KiCad data
    // matches, lacks or leaves unused, and every net - on the Open128-2 board some 1,150 badges, each a container, a Border
    // and a TextBlock in a WrapPanel, ~3,400 controls built as the page opened (~0.4 s). Here each is a rounded box and its
    // text drawn in place, laid out as that WrapPanel laid them out (BadgeCloudGeometry) - the look of KiCadDataPanel's
    // Border.ReferenceBadge until then: 1px edge, corners of 3, 6 by 1 inside, text 12, 4 apart, in the theme's
    // "Draft_ReferenceBadge" colors (the components left to label in the danger red; a net already used filled green, semibold).
    //
    // A badge with something to say (IBadgeCloudItem.Tip) says it under itself while the pointer is on it - the tooltip at
    // the control's bottom edge (ToolTipPlacement), moved up and across to stand under that badge, never under the pointer.
    // ###########################################################################################
    public sealed class BadgeCloud : Control
    {
        public static readonly StyledProperty<IEnumerable?> ItemsSourceProperty =
            AvaloniaProperty.Register<BadgeCloud, IEnumerable?>(nameof(ItemsSource));

        // Every badge in the danger red - "Components the KiCad data knows about that this board has no label for yet".
        public static readonly StyledProperty<bool> IsTodoProperty =
            AvaloniaProperty.Register<BadgeCloud, bool>(nameof(IsTodo));

        public const double TextSize = 12;
        public const double Spacing = 4;
        public const double Rounding = 3;
        private static readonly Thickness Inside = new(6, 1);
        private const double Edge = 1;

        private List<Badge> thisBadges = [];
        private List<Rect> thisPlaces = [];
        private double thisPlacedFor = double.NaN;
        private Size thisPlacedSize;
        private int thisPointedAt = -1;

        static BadgeCloud()
        {
            AffectsMeasure<BadgeCloud>(ItemsSourceProperty, IsTodoProperty);
        }

        public BadgeCloud()
        {
            // A badge's text color is drawn into it as it is built, so a change of the colors builds them again - of the
            // theme, and of the resources alone, as the "User preference" theme applies its colors (code review, 2026-11-13).
            this.ActualThemeVariantChanged += (_, _) => this.Forget();
            this.ResourcesChanged += (_, _) => this.Forget();
        }

        public IEnumerable? ItemsSource { get => this.GetValue(ItemsSourceProperty); set => this.SetValue(ItemsSourceProperty, value); }

        public bool IsTodo { get => this.GetValue(IsTodoProperty); set => this.SetValue(IsTodoProperty, value); }

        // What is drawn, in order, with where - for a test, after a layout pass.
        internal IReadOnlyList<(string Text, BadgeCloudKind Kind, Rect Place, IBrush? TextBrush)> BadgesForTests =>
            this.thisBadges.Select((badge, index) => (badge.Text, badge.Kind, index < this.thisPlaces.Count ? this.thisPlaces[index] : default, badge.TextBrush)).ToList();

        protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
        {
            base.OnPropertyChanged(change);

            if (change.Property == ItemsSourceProperty || change.Property == IsTodoProperty)
            {
                this.Forget();
            }
        }

        // What was built from the items and the theme: built again at the next measure. Each text layout is let go of as a
        // TextBlock lets go of its own (code review, 2026-11-13: ~1,150 left to the finalizer at every refresh).
        private void Forget()
        {
            foreach (Badge badge in this.thisBadges)
            {
                badge.Layout.Dispose();
            }

            this.thisBadges = [];
            this.thisPlaces = [];
            this.thisPlacedFor = double.NaN;
            this.InvalidateMeasure();
            this.InvalidateVisual();

            // And what the badge under the pointer said - another may stand there now (code review, 2026-11-13: an unused
            // net kept the "Already used" of the badge before it).
            if (this.thisPointedAt >= 0)
            {
                this.PointAt(null);
            }
        }

        protected override Size MeasureOverride(Size availableSize)
        {
            if (this.thisBadges.Count == 0)
            {
                this.thisBadges = this.BuildBadges();
                this.thisPlacedFor = double.NaN;
            }

            return this.PlaceFor(availableSize.Width);
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            this.PlaceFor(finalSize.Width);
            return finalSize;
        }

        // Placed again only for another width - each layout pass at the same one would place every badge again for nothing,
        // some 1,150 of them on the Open128-2 (code review, 2026-11-13).
        private Size PlaceFor(double width)
        {
            if (width.Equals(this.thisPlacedFor))
            {
                return this.thisPlacedSize;
            }

            (List<Rect> places, Size size) = BadgeCloudGeometry.Place(this.thisBadges.Select(badge => badge.Size).ToList(), width, Spacing, Spacing);

            this.thisPlaces = places;
            this.thisPlacedFor = width;
            this.thisPlacedSize = size;
            this.InvalidateVisual();

            return size;
        }

        private List<Badge> BuildBadges()
        {
            var typeface = new Typeface(this.GetValue(TextElement.FontFamilyProperty));
            var semibold = new Typeface(typeface.FontFamily, FontStyle.Normal, FontWeight.SemiBold);
            var badges = new List<Badge>();

            foreach (object? item in this.ItemsSource ?? Array.Empty<object>())
            {
                (string text, bool used, string? tip) = item is IBadgeCloudItem badge ? (badge.Text, badge.IsUsed, badge.Tip) : (item?.ToString() ?? string.Empty, false, null);
                BadgeCloudKind kind = used ? BadgeCloudKind.Used : this.IsTodo ? BadgeCloudKind.Todo : BadgeCloudKind.Plain;

                IBrush? textBrush = this.Brush(kind, "Fg");
                var layout = new TextLayout(text, kind == BadgeCloudKind.Used ? semibold : typeface, TextSize, textBrush);
                var size = new Size(
                    layout.WidthIncludingTrailingWhitespace + Inside.Left + Inside.Right + (2 * Edge),
                    layout.Height + Inside.Top + Inside.Bottom + (2 * Edge));

                badges.Add(new Badge(text, kind, tip, layout, size, textBrush));
            }

            return badges;
        }

        // The theme's color for a badge of `kind` - "Bg", "Border" or "Fg" (the plain badge's text is every text's, "Fg").
        private IBrush? Brush(BadgeCloudKind kind, string part)
        {
            string key = (kind, part) switch
            {
                (BadgeCloudKind.Plain, "Fg") => "Fg",
                (BadgeCloudKind.Plain, _) => $"Draft_ReferenceBadge_{part}",
                _ => $"Draft_ReferenceBadge_{kind}_{part}",
            };

            return this.TryFindResource(key, this.ActualThemeVariant, out object? found) ? found as IBrush : null;
        }

        public override void Render(DrawingContext context)
        {
            base.Render(context);

            var pens = new Dictionary<BadgeCloudKind, (IBrush? Fill, IPen? Pen)>();

            for (int index = 0; index < this.thisBadges.Count && index < this.thisPlaces.Count; index++)
            {
                Badge badge = this.thisBadges[index];
                Rect place = this.thisPlaces[index];

                if (!pens.TryGetValue(badge.Kind, out var paint))
                {
                    IBrush? edge = this.Brush(badge.Kind, "Border");
                    paint = (this.Brush(badge.Kind, "Bg"), edge is null ? null : new Pen(edge, Edge));
                    pens[badge.Kind] = paint;
                }

                // The edge drawn inside the box, as a Border draws its own.
                Rect box = place.Deflate(Edge / 2);
                context.DrawRectangle(paint.Fill, paint.Pen, box, Rounding, Rounding);

                double textTop = place.Y + ((place.Height - badge.Layout.Height) / 2);
                badge.Layout.Draw(context, new Point(place.X + Edge + Inside.Left, textTop));
            }
        }

        // ---- what a badge says under itself -----------------------------------------------------------

        protected override void OnPointerMoved(PointerEventArgs e)
        {
            base.OnPointerMoved(e);
            this.PointAt(e.GetPosition(this));
        }

        protected override void OnPointerExited(PointerEventArgs e)
        {
            base.OnPointerExited(e);
            this.PointAt(null);
        }

        // The badge under `point`, if it has something to say: its tip, under it.
        internal void PointAt(Point? point)
        {
            int at = point is { } where ? this.thisPlaces.FindIndex(place => place.Contains(where)) : -1;

            if (at >= 0 && this.thisBadges[at].Tip is null)
            {
                at = -1;
            }

            if (at == this.thisPointedAt)
            {
                return;
            }

            this.thisPointedAt = at;
            ToolTip.SetIsOpen(this, false);

            if (at < 0)
            {
                ToolTip.SetTip(this, null);
                return;
            }

            // ToolTipPlacement's Bottom: under the control, centred - moved to under this badge.
            Rect place = this.thisPlaces[at];
            ToolTip.SetHorizontalOffset(this, place.Center.X - (this.Bounds.Width / 2));
            ToolTip.SetVerticalOffset(this, place.Bottom - this.Bounds.Height);
            ToolTip.SetTip(this, this.thisBadges[at].Tip);
        }

        private sealed record Badge(string Text, BadgeCloudKind Kind, string? Tip, TextLayout Layout, Size Size, IBrush? TextBrush);
    }
}
