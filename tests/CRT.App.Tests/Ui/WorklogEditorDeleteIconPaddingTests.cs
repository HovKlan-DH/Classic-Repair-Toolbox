using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.Styling;
using CRT;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests.Ui;

// The worklog editor's five delete buttons render trash-can (U+F2ED), which overshoots the font's
// declared ascent by 16 design units and loses its top pixel row without extra room reserved for
// it - the same defect the padlocks had, in a glyph nobody had measured.
//
// The fix is a style Setter in WorklogEntryEditorWindow.axaml, and a Setter cannot call
// GetTopOverflowThickness, so the padding is a LITERAL there. That is the drift risk this file
// exists for: the literal is asserted against what FontAwesomeGlyphMetrics actually computes at
// the font size the markup uses, so changing one without the other fails here rather than shipping
// a clipped icon back.
[Collection("HeadlessUi")]
public class WorklogEditorDeleteIconPaddingTests
{
    // The FontSize the five delete TextBlocks are declared at in the markup.
    private const double DeleteIconFontSize = 13;

    private const int TrashCan = 0xF2ED;

    // The value hardcoded in the Button.WorklogRowIconButton > TextBlock style.
    private static readonly Thickness MarkupPadding = new(0, 1, 0, 0);

    // The whole point of the literal: it has to equal what the metrics helper would have returned.
    [Fact]
    public void The_markup_padding_matches_what_the_metrics_helper_computes()
    {
        Thickness computed = FontAwesomeGlyphMetrics.GetTopOverflowThickness(TrashCan, DeleteIconFontSize);

        Assert.Equal(computed, MarkupPadding);
    }

    // A guard on the premise rather than on the fix: if a font upgrade ever made trash-can sit
    // inside the ascent, the padding above would be reserving space for nothing and should go.
    [Fact]
    public void Trash_can_still_overshoots_and_so_still_needs_the_reservation()
    {
        Assert.True(
            FontAwesomeGlyphMetrics.OverflowsDeclaredAscent(TrashCan),
            "trash-can no longer overshoots - remove the padding style rather than reserving room for nothing");
    }

    // And that the rule is actually in the shipped window's markup, with the padding this file
    // pins. A correct constant that no style applies would pass every assertion above and still
    // clip.
    //
    // This reads the Style objects rather than a rendered icon: the delete rows live in
    // DataTemplates inside item lists, so an editor opened on no data realises none of them, and
    // the reconstruction alternative fails outright because an Avalonia Style cannot be
    // re-parented into a second window ("The Style already has a parent").
    [Fact]
    public void The_shipped_editor_markup_carries_the_padding_rule()
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();
            var editor = new WorklogEntryEditorWindow();

            var paddingSetters = editor.Styles
                .OfType<Style>()
                .Where(style => style.Selector?.ToString()?.Contains("WorklogRowIconButton") == true)
                .SelectMany(style => style.Setters)
                .OfType<Setter>()
                .Where(setter => setter.Property == TextBlock.PaddingProperty)
                .Select(setter => setter.Value)
                .ToList();

            Assert.Contains(MarkupPadding, paddingSetters.OfType<Thickness>());
        });
    }
}
