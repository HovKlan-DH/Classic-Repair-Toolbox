using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Layout;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Theming
{
    // ###########################################################################################
    // A DEVICE'S CONDITION AS STARS (owner requests, 2026-11-06: 'can we add a "Condition" and then have 0 to 5 stars?',
    // then 'have this show prettier - not sure how, but maybe more "depth" and coloring') - five stars, as many lit as
    // the device has. A lit star is gold, light at its top and deeper at its foot, with a darker edge and a soft shadow
    // under it, so it stands off the page; an unlit one a flat, quiet outline. Drawn as shapes (ConditionStarGeometry),
    // never a font's glyph, which can take neither a gradient nor a shadow. Colors Workbooks_Star_* in both themes.
    //
    // The one place they are drawn: a device's page, and the device form, whose stars take a click and light up to the
    // one under the pointer (OwnedDeviceWindow). Each star is a frame carrying its number as its Tag, with the class
    // "Lit" while lit - what is lit can be read without knowing how it was painted.
    // ###########################################################################################
    public static class ConditionStars
    {
        public static StackPanel Build(int stars, double size = 16.0)
        {
            var row = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = size / 5.0,
                VerticalAlignment = VerticalAlignment.Center,
            };

            Avalonia.Media.Geometry shape = StarShape(size, edge: EdgeOf(size));

            for (int star = 1; star <= OwnedDeviceCondition.MaxStars; star++)
            {
                row.Children.Add(new Border
                {
                    Tag = star,
                    Width = size,
                    Height = size,
                    Child = new Path { Data = shape, StrokeThickness = EdgeOf(size), StrokeJoin = PenLineJoin.Round },
                });
            }

            Show(row, stars);
            return row;
        }

        // ###########################################################################################
        // The row's stars painted for `stars` lit - in place, so a click on one never makes the row again under the
        // pointer.
        // ###########################################################################################
        public static void Show(Panel row, int stars)
        {
            int lit = OwnedDeviceCondition.Within(stars);

            foreach (Border star in row.Children.OfType<Border>())
            {
                if (star.Child is not Path path)
                {
                    continue;
                }

                bool isLit = star.Tag is int number && number <= lit;

                path.Fill = isLit ? LitFill() : ThemeResources.ResolveBrush("Workbooks_Star_Dim_Fill", Brushes.Transparent);
                path.Stroke = ThemeResources.ResolveBrush(isLit ? "Workbooks_Star_Lit_Edge" : "Workbooks_Star_Dim_Edge", Brushes.Gray);
                path.Effect = isLit
                    ? new DropShadowEffect { OffsetX = 0, OffsetY = 1, BlurRadius = 3, Opacity = 0.35, Color = Colors.Black }
                    : null;

                star.Classes.Set("Lit", isLit);
            }
        }

        // Gold, light at the top and deeper at the foot - the "depth".
        private static IBrush LitFill() => new LinearGradientBrush
        {
            StartPoint = new RelativePoint(0.5, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(0.5, 1, RelativeUnit.Relative),
            GradientStops =
            [
                new GradientStop(ThemeResources.ResolveColor("Workbooks_Star_Lit_Top", Colors.Gold), 0),
                new GradientStop(ThemeResources.ResolveColor("Workbooks_Star_Lit_Bottom", Colors.Orange), 1),
            ],
        };

        // The edge, a little heavier on a bigger star.
        private static double EdgeOf(double size) => size >= 18 ? 1.4 : 1.1;

        private static Avalonia.Media.Geometry StarShape(double size, double edge)
        {
            IReadOnlyList<Point> corners = ConditionStarGeometry.Corners(size, inset: edge / 2.0);
            var geometry = new StreamGeometry();

            using (StreamGeometryContext context = geometry.Open())
            {
                context.BeginFigure(corners[0], isFilled: true);

                foreach (Point corner in corners.Skip(1))
                {
                    context.LineTo(corner);
                }

                context.EndFigure(isClosed: true);
            }

            return geometry;
        }
    }
}
