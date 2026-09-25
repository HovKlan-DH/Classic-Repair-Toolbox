using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace Handlers.Theming
{
    // ###########################################################################################
    // The "#N" pill that marks a worklog entry on a schematic - a category-coloured badge carrying
    // the entry's id and a white disc with an open/closed padlock in it.
    //
    // ONE builder, called from both places it appears: the Schematics tab's own "Show worklogs" view
    // and the Workbooks tab's board pane. They were two separate ~45-line methods, line for line the
    // same, with a comment on each conceding the duplication and asserting the two "must look the
    // same" - which nothing enforced, and the parked-vs-anchored branch beside them had already
    // drifted once and been reported.
    //
    // Builds controls but reads no tab state, so it lives here rather than on either UserControl.
    // What the two callers genuinely differ on - the scale transform, the Tag, and the click handler
    // - they apply themselves afterwards; the Schematics tab's badges sit on a canvas carrying the
    // view matrix and cancel it out with an inverse scale, while the Workbooks pane never zooms.
    // ###########################################################################################
    public static class WorklogBadgeBuilder
    {
        // Font size of the padlock inside the white disc.
        private const double StateIconFontSize = 10.0;

        // One Hand cursor for every badge, rather than one per badge per rebuild: Cursor is
        // IDisposable and holds an HCURSOR on Win32, and these are built in bulk against controls the
        // next refresh throws away.
        private static readonly Cursor HandCursor = new(StandardCursorType.Hand);

        // ###########################################################################################
        // Builds one badge for an entry, in the given category colour.
        //
        // The disc stays WHITE rather than taking the state colour: the badge behind it is already
        // filled with the entry's category colour, and a state-coloured disc on that would put two
        // saturated colours against each other with the glyph lost between them. White separates the
        // two and lets the padlock itself carry the state.
        // ###########################################################################################
        public static Border Build(WorklogEntryRecord entry, Color categoryColor, Color stateColor)
        {
            var idText = new TextBlock
            {
                Text = $"#{entry.Id}",
                FontSize = 11,
                FontWeight = FontWeight.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center
            };

            bool isResolved = WorklogManager.IsResolvedState(entry.State);

            var stateIcon = new TextBlock
            {
                Text = WorklogGlyphs.GlyphFor(isResolved),
                FontFamily = ThemeResources.ResolveFontAwesomeSolid(),
                FontSize = StateIconFontSize,
                Foreground = new SolidColorBrush(stateColor),

                // The padlocks are drawn taller than the font's declared ascent, so without a
                // reserved row their top pixel row falls outside the line box and is clipped - see
                // FontAwesomeGlyphMetrics. Computed rather than hardcoded, so changing the font size
                // above cannot quietly reintroduce the clipping.
                Padding = FontAwesomeGlyphMetrics.GetTopOverflowThickness(
                    WorklogGlyphs.CodepointFor(isResolved), StateIconFontSize),

                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };

            var statePill = new Border
            {
                Width = 16,
                Height = 16,
                CornerRadius = new CornerRadius(8),
                Background = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Child = stateIcon
            };

            var content = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4 };
            content.Children.Add(idText);
            content.Children.Add(statePill);

            return new Border
            {
                Background = new SolidColorBrush(categoryColor),
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(8, 3),
                Cursor = HandCursor,
                Child = content
            };
        }
    }
}
