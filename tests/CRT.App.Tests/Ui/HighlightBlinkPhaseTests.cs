using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using CRT;
using Handlers.DataHandling;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "Blink selected" as an opacity, not a rebuild (code review, 2026-10-13). The blink timer used
// to call ApplyHighlightVisuals, which rendered every highlighted thumbnail again and rebuilt the
// KiCad overlay - twice a second. Now the highlights are built when the SELECTION changes, as a
// layer over each thumbnail, and ApplyBlinkPhase only sets opacities and the KiCad overlay's phase.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class HighlightBlinkPhaseTests
{
    private const string SchematicOne = "Schematic 1";
    private const string SchematicTwo = "Schematic 2";

    private static TabSchematics CreateTabWithThumbnails(out SchematicThumbnail one, out SchematicThumbnail two)
    {
        var tab = new TabSchematics();

        tab.highlightRectsBySchematicAndLabel =
            new Dictionary<string, Dictionary<string, List<Rect>>>(StringComparer.OrdinalIgnoreCase)
            {
                [SchematicOne] = new(StringComparer.OrdinalIgnoreCase) { ["U8"] = [new Rect(10, 10, 5, 5)] },
            };

        tab.schematicByName = new(StringComparer.OrdinalIgnoreCase)
        {
            [SchematicOne] = new BoardSchematicEntry { SchematicName = SchematicOne },
            [SchematicTwo] = new BoardSchematicEntry { SchematicName = SchematicTwo },
        };

        one = Thumbnail(SchematicOne);
        two = Thumbnail(SchematicTwo);
        tab.currentThumbnails.Add(one);
        tab.currentThumbnails.Add(two);

        return tab;
    }

    private static SchematicThumbnail Thumbnail(string name)
    {
        var bitmap = new RenderTargetBitmap(new PixelSize(80, 60), new Vector(96, 96));

        return new SchematicThumbnail
        {
            Name = name,
            BaseThumbnail = bitmap,
            ImageSource = bitmap,
            OriginalPixelSize = new PixelSize(800, 600),
        };
    }

    [Fact]
    public void A_selection_puts_a_highlight_layer_over_the_matching_thumbnail_only()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTabWithThumbnails(out var one, out var two);

            tab.UpdateHighlightsForComponents(["U8"]);

            Assert.NotNull(one.HighlightImage);
            Assert.Same(one.BaseThumbnail, one.ImageSource);
            Assert.Null(two.HighlightImage);
        });
    }

    [Fact]
    public void The_blink_only_sets_opacities_and_rebuilds_nothing()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTabWithThumbnails(out var one, out _);
            tab.UpdateHighlightsForComponents(["U8"]);
            var layer = one.HighlightImage;
            var overlay = tab.GetControl<SchematicHighlightsOverlay>("SchematicsHighlightsOverlay");
            var kiCad = tab.GetControl<KiCadOverlayRenderControl>("SchematicsKiCadOverlayCanvas");

            tab.ApplyBlinkPhase(0.0);

            Assert.Equal(0.0, one.HighlightLayerOpacity);
            Assert.Equal(0.0, overlay.Opacity);
            Assert.False(kiCad.BlinkPhaseVisible);
            Assert.Same(layer, one.HighlightImage);

            tab.ApplyBlinkPhase(1.0);

            Assert.Equal(1.0, one.HighlightLayerOpacity);
            Assert.Equal(1.0, overlay.Opacity);
            Assert.True(kiCad.BlinkPhaseVisible);
            Assert.Same(layer, one.HighlightImage);
        });
    }

    // The highlight strength itself carries no blink any more - that is the layer's opacity.
    [Fact]
    public void Applying_the_highlights_in_the_hidden_phase_keeps_their_strength_and_hides_the_layer()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTabWithThumbnails(out var one, out _);
            tab.UpdateHighlightsForComponents(["U8"]);

            tab.ApplyHighlightVisuals(hasSelection: true, blinkFactor: 0.0);

            Assert.NotNull(one.HighlightImage);
            Assert.Equal(0.0, one.HighlightLayerOpacity);
        });
    }

    // A selection that changes nothing on a thumbnail leaves its layer as it is - typing in the
    // component filter drew every highlighted thumbnail again on each keystroke (code review,
    // 2026-10-13). One that does change it draws it again.
    [Fact]
    public void A_thumbnail_whose_own_highlights_did_not_change_keeps_its_layer()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTabWithThumbnails(out var one, out _);
            tab.highlightRectsBySchematicAndLabel[SchematicOne]["U9"] = [new Rect(30, 30, 5, 5)];

            tab.UpdateHighlightsForComponents(["U8"]);
            var layer = one.HighlightImage;

            // U99 is on no schematic: thumbnail one shows exactly what it showed.
            tab.UpdateHighlightsForComponents(["U8", "U99"]);
            Assert.Same(layer, one.HighlightImage);

            // U9 is on it: drawn again.
            tab.UpdateHighlightsForComponents(["U8", "U9"]);
            Assert.NotSame(layer, one.HighlightImage);
            Assert.NotNull(one.HighlightImage);
        });
    }

    // A selection cleared takes the layer away (and disposes it).
    [Fact]
    public void Clearing_the_selection_removes_the_highlight_layer()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTabWithThumbnails(out var one, out _);
            tab.UpdateHighlightsForComponents(["U8"]);

            tab.UpdateHighlightsForComponents([]);

            Assert.Null(one.HighlightImage);
        });
    }
}
