using Avalonia;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Pure rectangle and transform maths shared by the schematic viewer and its overlays,
    // extracted from TabSchematics and SchematicHighlightsOverlay so it can be unit tested.
    //
    // Uses Avalonia's Point/Rect/Matrix value types but touches no control and no instance state.
    // ###########################################################################################
    public static class RectGeometry
    {
        // ###########################################################################################
        // Tries to invert a 2D affine matrix. Returns false (and the identity) when it is singular.
        // ###########################################################################################
        public static bool TryInvert(Matrix m, out Matrix inv)
        {
            double a = m.M11, b = m.M12, c = m.M21, d = m.M22, e = m.M31, f = m.M32;
            double det = (a * d) - (b * c);

            if (Math.Abs(det) < 1e-12)
            {
                inv = Matrix.Identity;
                return false;
            }

            double idet = 1.0 / det;
            double na = d * idet, nb = -b * idet, nc = -c * idet, nd = a * idet;
            double ne = -((e * na) + (f * nc)), nf = -((e * nb) + (f * nd));

            inv = new Matrix(na, nb, nc, nd, ne, nf);
            return true;
        }

        // ###########################################################################################
        // Normalizes a rectangle so width and height are always positive regardless of drag direction.
        // ###########################################################################################
        public static Rect CreateNormalizedRect(Point start, Point end)
        {
            double x = Math.Min(start.X, end.X);
            double y = Math.Min(start.Y, end.Y);
            double width = Math.Abs(end.X - start.X);
            double height = Math.Abs(end.Y - start.Y);

            return new Rect(x, y, width, height);
        }

        // ###########################################################################################
        // Returns the area the bitmap occupies inside its control, preserving aspect ratio.
        // SchematicsImage uses HorizontalAlignment="Left" and VerticalAlignment="Top", so the
        // bitmap content always starts at (0, 0) - no centering offset is applied.
        // ###########################################################################################
        public static Rect GetImageContentRect(Size controlSize, PixelSize bitmapPixelSize)
        {
            // A zero bitmap dimension is reachable, not theoretical: a thumbnail whose image failed
            // to load keeps OriginalPixelSize at 0x0. Without this the aspect division yields
            // Infinity or NaN, and the NaN case returns a rect with NaN X/Width that poisons every
            // layout it reaches.
            if (controlSize.Width <= 0 || controlSize.Height <= 0 ||
                bitmapPixelSize.Width <= 0 || bitmapPixelSize.Height <= 0)
                return new Rect(controlSize);

            double containerAspect = controlSize.Width / controlSize.Height;
            double bitmapAspect = (double)bitmapPixelSize.Width / bitmapPixelSize.Height;

            if (bitmapAspect > containerAspect)
            {
                // Width-constrained - content starts at (0, 0), no vertical centering
                return new Rect(0, 0, controlSize.Width, controlSize.Width / bitmapAspect);
            }
            else
            {
                // Height-constrained - content starts at (0, 0), no horizontal centering
                return new Rect(0, 0, controlSize.Height * bitmapAspect, controlSize.Height);
            }
        }

        // ###########################################################################################
        // Returns the area the bitmap occupies inside its control when the control centers its
        // content instead of anchoring it top-left - the schematic thumbnail gallery's Image is
        // HorizontalAlignment/VerticalAlignment="Stretch" with Stretch="Uniform", so its rendered
        // content sits centered in the control box rather than starting at (0, 0) like
        // GetImageContentRect's Left/Top-aligned SchematicsImage.
        // ###########################################################################################
        public static Rect GetCenteredImageContentRect(Size controlSize, PixelSize bitmapPixelSize)
        {
            // Same zero-dimension guard as GetImageContentRect above - see the reasoning there.
            if (controlSize.Width <= 0 || controlSize.Height <= 0 ||
                bitmapPixelSize.Width <= 0 || bitmapPixelSize.Height <= 0)
                return new Rect(controlSize);

            double containerAspect = controlSize.Width / controlSize.Height;
            double bitmapAspect = (double)bitmapPixelSize.Width / bitmapPixelSize.Height;

            if (bitmapAspect > containerAspect)
            {
                // Width-constrained - full width, centered vertically
                double height = controlSize.Width / bitmapAspect;
                return new Rect(0, (controlSize.Height - height) / 2.0, controlSize.Width, height);
            }
            else
            {
                // Height-constrained - full height, centered horizontally
                double width = controlSize.Height * bitmapAspect;
                return new Rect((controlSize.Width - width) / 2.0, 0, width, controlSize.Height);
            }
        }

        // ###########################################################################################
        // Converts a rectangle in control-local coordinates into source-image pixel coordinates,
        // clamped to the bitmap bounds.
        // ###########################################################################################
        public static Rect LocalToPixelRect(Rect localRect, Rect contentRect, PixelSize pixelSize)
        {
            double sx = pixelSize.Width / contentRect.Width;
            double sy = pixelSize.Height / contentRect.Height;

            double x = (localRect.X - contentRect.X) * sx;
            double y = (localRect.Y - contentRect.Y) * sy;
            double w = localRect.Width * sx;
            double h = localRect.Height * sy;

            return new Rect(x, y, w, h).Intersect(new Rect(0, 0, pixelSize.Width, pixelSize.Height));
        }

        // ###########################################################################################
        // Converts a rectangle in source-image pixel coordinates into control-local coordinates.
        // ###########################################################################################
        public static Rect PixelToLocalRect(Rect pixelRect, Rect contentRect, PixelSize pixelSize)
        {
            double sx = contentRect.Width / pixelSize.Width;
            double sy = contentRect.Height / pixelSize.Height;

            double x = contentRect.X + (pixelRect.X * sx);
            double y = contentRect.Y + (pixelRect.Y * sy);
            double w = pixelRect.Width * sx;
            double h = pixelRect.Height * sy;

            return new Rect(x, y, w, h);
        }

        // ###########################################################################################
        // Parses an invariant-culture number. Board data is authored on machines with different
        // locales, so a comma decimal separator must never be accepted as a thousands separator.
        // ###########################################################################################
        public static bool TryParseDouble(string text, out double value)
            => double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);

        // ###########################################################################################
        // Parses a colour written in any form Avalonia accepts (#RRGGBB, #AARRGGBB, a named colour).
        // Falls back rather than throwing, because the text comes from contributed board data.
        // ###########################################################################################
        public static Color ParseColorOrDefault(string text, Color fallback)
        {
            if (string.IsNullOrWhiteSpace(text)) return fallback;
            try { return Color.Parse(text.Trim()); }
            catch { return fallback; }
        }

        // ###########################################################################################
        // Parses an opacity written either as 0..1, as 0..100, or with a trailing percent sign, and
        // clamps the result to 0..1. Falls back when the text cannot be parsed at all.
        // ###########################################################################################
        public static double ParseOpacityOrDefault(string text, double fallback)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return fallback;
            }

            string normalized = text.Trim();

            bool isPercent = normalized.EndsWith("%", StringComparison.Ordinal);
            if (isPercent)
            {
                normalized = normalized[..^1].Trim();
            }

            if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
            {
                return fallback;
            }

            if (isPercent || value > 1.0)
            {
                value /= 100.0;
            }

            return Math.Clamp(value, 0.0, 1.0);
        }
        // ###########################################################################################
        // Converts a pixel-space point into the overlay's local coordinate system.
        // ###########################################################################################
        public static Point PixelToLocalPoint(Point pixelPoint, Rect contentRect, PixelSize pixelSize)
        {
            double sx = contentRect.Width / pixelSize.Width;
            double sy = contentRect.Height / pixelSize.Height;

            return new Point(
                contentRect.X + (pixelPoint.X * sx),
                contentRect.Y + (pixelPoint.Y * sy));
        }

        // ###########################################################################################
        // Insets a rectangle by half the stroke thickness so the drawn border remains visually
        // inside the original bounds instead of growing outward. Width and height are floored at
        // zero so a rectangle thinner than its own stroke does not invert.
        // ###########################################################################################
        public static Rect InsetRectForStroke(Rect rect, double strokeThickness)
        {
            double inset = strokeThickness / 2.0;
            double width = Math.Max(0.0, rect.Width - strokeThickness);
            double height = Math.Max(0.0, rect.Height - strokeThickness);

            return new Rect(rect.X + inset, rect.Y + inset, width, height);
        }

        // ###########################################################################################
        // Returns the keys whose rect list contains at least one rect intersecting the target rect -
        // used to find which board labels a worklog entry area touches. A key with several rects
        // (a component can have more than one highlight) counts as touching if any one of them does.
        // ###########################################################################################
        public static HashSet<string> FindKeysWithRectsIntersecting(
            IReadOnlyDictionary<string, List<Rect>> rectsByKey,
            Rect targetRect)
        {
            var matches = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var entry in rectsByKey)
            {
                foreach (var rect in entry.Value)
                {
                    if (rect.Intersects(targetRect))
                    {
                        matches.Add(entry.Key);
                        break;
                    }
                }
            }

            return matches;
        }

        // ###########################################################################################
        // The rect a control actually occupies once a CENTERED ScaleTransform is applied to it -
        // the arrangement the "Show worklogs" pills use, where a per-zoom inverse scale keeps each
        // pill a constant size on screen (TabSchematics.Worklog.cs, PositionWorklogEntriesListBadge).
        //
        // layoutTopLeft/layoutSize are the control's placement and unscaled size. Because the
        // transform's origin is the middle of that layout box rather than its corner, the drawn
        // rect is NOT layoutTopLeft plus the scaled size: shrinking pulls all four edges inward
        // toward the center, so the top-left moves too. Getting that wrong leaves a hover rect
        // that is offset from what the user sees - fine at scale 1, drifting further out the more
        // the view is zoomed.
        //
        // Callers that can reach the live visual tree should prefer TranslatePoint, which applies
        // the real transforms; this is for reasoning about, and testing, the same placement.
        // ###########################################################################################
        public static Rect GetCenterScaledControlRect(Point layoutTopLeft, Size layoutSize, double scale)
        {
            if (layoutSize.Width <= 0 || layoutSize.Height <= 0)
            {
                return new Rect(layoutTopLeft, new Size(0, 0));
            }

            double safeScale = Math.Max(0.0, scale);

            var center = new Point(
                layoutTopLeft.X + (layoutSize.Width / 2.0),
                layoutTopLeft.Y + (layoutSize.Height / 2.0));

            double scaledWidth = layoutSize.Width * safeScale;
            double scaledHeight = layoutSize.Height * safeScale;

            return new Rect(
                center.X - (scaledWidth / 2.0),
                center.Y - (scaledHeight / 2.0),
                scaledWidth,
                scaledHeight);
        }
    }
}
