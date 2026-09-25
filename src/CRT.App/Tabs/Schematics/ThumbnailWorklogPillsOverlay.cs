using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Globalization;
using Handlers.Geometry;

namespace Tabs.TabSchematics
{
    // ###########################################################################################
    // Draws a small "#N" pill - no drawn bounds, no status - for every saved worklog entry that
    // belongs to this thumbnail's schematic. Colored by category, matching
    // TabSchematics.Worklog.cs's CreateWorklogEntriesListBadge outer badge (the main view's badge
    // also carries a state pill; this one deliberately does not - the thumbnail is just meant to
    // give an at-a-glance idea of where issues are marked, not their status). Used by the thumbnail
    // gallery; visibility follows the same "Show worklogs" checkbox as the full-size overlay
    // (WorklogEntriesOverlay) and the main view's badges.
    //
    // A pill is drawn one of two ways, MIRRORING the main schematic view's own branch exactly:
    //
    //   ShowMarkedArea ON   - centred on the entry's marked area, which is where that area is drawn
    //                         on the main view.
    //   ShowMarkedArea OFF  - PARKED in this thumbnail's own top-right corner instead. The entry
    //                         has no visible area, so a pill sitting where the area WOULD be points
    //                         at nothing and reads as a marked location that is not marked. This
    //                         was reported: the main view parked such a pill correctly while the
    //                         thumbnail still showed it at the marker, so the two disagreed about
    //                         the same entry.
    //
    // The parking geometry is ParkedBadgeGeometry.ArrangeInTopRightBlock - the SAME function the
    // main view's parked pills and the Workbooks tab's board previews use, so all three stack
    // identically rather than each having its own idea of a corner block.
    //
    // A self-contained Control that computes its own centered content rect from its arranged Bounds
    // (see RectGeometry.GetCenteredImageContentRect - the thumbnail Image is Stretch-aligned, unlike
    // the main SchematicsImage, so content is centered rather than anchored at the origin), the same
    // technique WorklogEntriesOverlay uses.
    // ###########################################################################################
    public sealed class ThumbnailWorklogPillsOverlay : Control
    {
        // ###########################################################################################
        // IsParked carries the entry's "Show marked area", inverted: an entry with no drawn area has
        // its pill parked in the corner. PixelCenter is then unused for that pill - the caller still
        // supplies it, since it costs nothing and keeps the record one shape.
        //
        // Defaulted to false so the pill is drawn at its marker unless a caller says otherwise -
        // matching WorklogEntryRecord.ShowMarkedArea, which defaults to true for exactly the same
        // reason (an entry written before the setting existed keeps showing its area).
        // ###########################################################################################
        public readonly record struct Pill(Point PixelCenter, Color Color, int EntryId, bool IsParked = false);

        // ###########################################################################################
        // Gap from the image's edges, and between stacked parked pills.
        //
        // Smaller than the main view's 10/6: a thumbnail is a fraction of the size, and the main
        // view's margins there ate a visible slice of a board already only ~150px across.
        // ###########################################################################################
        private const double ParkedMargin = 3.0;

        private const double ParkedSpacing = 2.0;

        private const double PillPaddingX = 5.0;

        private const double PillPaddingY = 2.0;

        public static readonly StyledProperty<IReadOnlyList<Pill>> PillsProperty =
            AvaloniaProperty.Register<ThumbnailWorklogPillsOverlay, IReadOnlyList<Pill>>(
                nameof(Pills), defaultValue: Array.Empty<Pill>());

        public static readonly StyledProperty<PixelSize> BitmapPixelSizeProperty =
            AvaloniaProperty.Register<ThumbnailWorklogPillsOverlay, PixelSize>(
                nameof(BitmapPixelSize), defaultValue: new PixelSize(0, 0));

        static ThumbnailWorklogPillsOverlay()
        {
            AffectsRender<ThumbnailWorklogPillsOverlay>(PillsProperty, BitmapPixelSizeProperty);
        }

        public IReadOnlyList<Pill> Pills
        {
            get => this.GetValue(PillsProperty);
            set => this.SetValue(PillsProperty, value ?? Array.Empty<Pill>());
        }

        public PixelSize BitmapPixelSize
        {
            get => this.GetValue(BitmapPixelSizeProperty);
            set => this.SetValue(BitmapPixelSizeProperty, value);
        }

        // ###########################################################################################
        // A little smaller than the main view's badge (FontSize 11, Padding 8,3, CornerRadius 10),
        // but not much smaller - thumbnails do not zoom, so unlike the main view's badges there is
        // no inverse-scale transform keeping this constant; it is just a fixed screen size.
        // ###########################################################################################
        public double FontSize { get; set; } = 10.0;

        protected override Size ArrangeOverride(Size finalSize)
        {
            var result = base.ArrangeOverride(finalSize);
            this.InvalidateVisual();
            return result;
        }

        public override void Render(DrawingContext context)
        {
            base.Render(context);

            foreach (var (pill, center, text) in this.LayOutPills())
            {
                DrawPill(context, center, pill.Color, pill.EntryId, text);
            }
        }

        // ###########################################################################################
        // Where every pill goes, in this control's own coordinate space - the placement decision on
        // its own, separated from the drawing so it can be asserted.
        //
        // This overlay renders straight to a DrawingContext rather than building controls, so there
        // is nothing on the visual tree for a test to read a position off - and "the parked pill
        // must not be at its marker's position" is exactly the rule that was reported broken here.
        // Splitting the decision out is what lets that rule be pinned; the alternative was pixel
        // sampling a RenderTargetBitmap, which needs a display.
        //
        // Returns each pill with the CENTRE it is drawn at, plus the laid-out text (the parked path
        // has to measure it to place the block, and re-laying it out to draw would be wasted work).
        //
        // The result is in the SAME ORDER as Pills, even though parked pills cannot be positioned
        // until every one of them is known (their block's layout depends on how many there are).
        // Rendering does not care, but a caller correlating this with its own list by index would -
        // and the obvious next thing to add here is click-to-open on a thumbnail pill, exactly as
        // the board pane already has, which is precisely that kind of caller. Appending the parked
        // ones at the end would map its click to the wrong entry, so their slots are reserved and
        // filled in place instead.
        // ###########################################################################################
        internal IReadOnlyList<(Pill Pill, Point Center, FormattedText Text)> LayOutPills()
        {
            var placed = new List<(Pill, Point, FormattedText)>();

            var pills = this.Pills;
            if (pills.Count == 0)
                return placed;

            if (this.Bounds.Width <= 0 || this.Bounds.Height <= 0)
                return placed;

            var bitmapPixelSize = this.BitmapPixelSize;
            if (bitmapPixelSize.Width <= 0 || bitmapPixelSize.Height <= 0)
                return placed;

            var contentRect = RectGeometry.GetCenteredImageContentRect(this.Bounds.Size, bitmapPixelSize);
            if (contentRect.Width <= 0 || contentRect.Height <= 0)
                return placed;

            double sx = contentRect.Width / bitmapPixelSize.Width;
            double sy = contentRect.Height / bitmapPixelSize.Height;

            List<Pill>? parked = null;
            List<int>? parkedSlots = null;

            foreach (var pill in pills)
            {
                if (pill.IsParked)
                {
                    // Collected rather than placed here: the block's layout depends on how many there
                    // are and how wide each is, so they cannot be positioned one at a time. Allocated
                    // lazily - most thumbnails have no parked pill at all, and this runs on every
                    // render pass of every thumbnail in the gallery.
                    //
                    // A placeholder holds this pill's slot so the finished list stays in Pills order -
                    // see the header. It is always overwritten below: parked and parkedSlots grow
                    // together, and every slot is filled from the arrangement.
                    (parked ??= new List<Pill>()).Add(pill);
                    (parkedSlots ??= new List<int>()).Add(placed.Count);
                    placed.Add(default);
                    continue;
                }

                var center = new Point(
                    contentRect.X + (pill.PixelCenter.X * sx),
                    contentRect.Y + (pill.PixelCenter.Y * sy));

                placed.Add((pill, center, this.BuildPillText(pill.EntryId)));
            }

            if (parked != null)
            {
                this.LayOutParkedPills(parked, parkedSlots!, contentRect, placed);
            }

            return placed;
        }

        // ###########################################################################################
        // Places the pills of entries with no marked area, stacked into the top-right corner.
        //
        // Anchored to the IMAGE's content rect, not to the control's Bounds: a thumbnail letterboxes
        // its image (Stretch="Uniform" in a fixed-size cell), so the control is wider or taller than
        // what is actually on screen, and parking against its Bounds would put the pills out in the
        // empty margin beside the board rather than on the corner of the board itself.
        //
        // ArrangeInTopRightBlock works in a viewport's own space and returns TOP-LEFT positions, so
        // the content rect's origin is added back on afterwards and the top-left converted to the
        // centre DrawPill wants. reservedRight is zero - a thumbnail has no side panel to step aside
        // for, unlike the main view's "Netlist names".
        // ###########################################################################################
        // parkedSlots names the index in "placed" each parked pill reserved, so the finished list
        // keeps the caller's own pill order - see LayOutPills' header.
        private void LayOutParkedPills(
            List<Pill> parked,
            List<int> parkedSlots,
            Rect contentRect,
            List<(Pill, Point, FormattedText)> placed)
        {
            var texts = new List<FormattedText>(parked.Count);
            var sizes = new List<Size>(parked.Count);

            foreach (var pill in parked)
            {
                var text = this.BuildPillText(pill.EntryId);
                texts.Add(text);
                sizes.Add(MeasurePill(text));
            }

            var positions = ParkedBadgeGeometry.ArrangeInTopRightBlock(
                sizes,
                contentRect.Size,
                ParkedMargin,
                ParkedSpacing,
                reservedRight: 0);

            for (int i = 0; i < parked.Count && i < positions.Count; i++)
            {
                var center = new Point(
                    contentRect.X + positions[i].X + (sizes[i].Width / 2.0),
                    contentRect.Y + positions[i].Y + (sizes[i].Height / 2.0));

                placed[parkedSlots[i]] = (parked[i], center, texts[i]);
            }

            // Any slot the arrangement did not reach still holds its placeholder, whose Text is null
            // and which Render would throw on. ArrangeInTopRightBlock returns one position per size
            // today, so this is unreachable - but the placeholders are only safe while that stays
            // true, and dropping them back-to-front keeps the surviving indices valid.
            for (int i = parkedSlots.Count - 1; i >= positions.Count; i--)
            {
                placed.RemoveAt(parkedSlots[i]);
            }
        }

        // ###########################################################################################
        // Draws one rounded, category-colored pill with a white bold "#N" label, centered at the
        // given point - the same look as CreateWorklogEntriesListBadge's outer badge, minus the
        // inner state pill (this overlay never shows status).
        //
        // The laid-out text comes from LayOutPills rather than being built here: the parked path has
        // to measure it anyway to place the block, and laying the same string out a second time to
        // draw it would be wasted work on every render pass of every thumbnail.
        // ###########################################################################################
        private static void DrawPill(
            DrawingContext context,
            Point center,
            Color color,
            int entryId,
            FormattedText formattedText)
        {
            var pillSize = MeasurePill(formattedText);
            var pillRect = new Rect(
                center.X - (pillSize.Width / 2.0),
                center.Y - (pillSize.Height / 2.0),
                pillSize.Width,
                pillSize.Height);

            context.DrawRectangle(
                new SolidColorBrush(color),
                new Pen(Brushes.White, 1.0),
                pillRect,
                pillSize.Height / 2.0,
                pillSize.Height / 2.0);

            var textOrigin = new Point(
                pillRect.X + PillPaddingX,
                pillRect.Y + PillPaddingY);
            context.DrawText(formattedText, textOrigin);
        }

        private FormattedText BuildPillText(int entryId) =>
            new(
                $"#{entryId.ToString(CultureInfo.InvariantCulture)}",
                CultureInfo.InvariantCulture,
                FlowDirection.LeftToRight,
                new Typeface(FontFamily.Default, FontStyle.Normal, FontWeight.Bold),
                this.FontSize,
                Brushes.White);

        private static Size MeasurePill(FormattedText text) =>
            new(text.Width + (PillPaddingX * 2), text.Height + (PillPaddingY * 2));
    }
}
