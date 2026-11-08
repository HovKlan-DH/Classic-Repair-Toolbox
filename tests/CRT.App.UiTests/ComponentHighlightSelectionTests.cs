using Avalonia;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// Selecting and deselecting components in the always-visible component filter box, and what
// that does to the highlights on every schematic and thumbnail.
//
// The scenario these are modelled on: pick Commodore 64 / 250469, type "rp" in the filter box
// and five components come back selected, one of them "U8 | SuperPLA | 251715". Clicking U8
// once deselects it, and its highlight must vanish from the large image AND from every
// thumbnail it appeared on, while the other four stay exactly as they were.
//
// The path under test is Main.OnComponentFilterSelectionChanged calling
// TabSchematics.UpdateHighlightsForComponents(labels), which rebuilds highlightIndexBySchematic
// from scratch on every selection change. That dictionary is keyed by schematic name and is
// what both the main image and the thumbnails render from, so asserting on it covers "every
// image the component appears on" in one go.
//
// The fixture is synthetic rather than the real 250469 workbook on purpose: board data is
// contributed content that syncs from classic-repair-toolbox.dk independently of releases, so
// a test asserting "rp matches exactly five components" would break for reasons that are not
// bugs. What matters here is the add and remove behaviour, and that is board-independent.
// ###########################################################################################
[Collection("HeadlessUi")]
public class ComponentHighlightSelectionTests
{
    private const string SchematicOne = "Schematic 1";
    private const string SchematicTwo = "Schematic 2";
    private const string PcbTop = "PCB top";

    // The five components an "rp" search brings back, U8 among them.
    private const string U8 = "U8";
    private const string U1 = "U1";
    private const string U2 = "U2";
    private const string U17 = "U17";
    private const string CR4 = "CR4";

    // A distinct rectangle per label, so an assertion can say exactly whose highlight survived
    // rather than only counting them.
    private static readonly Rect U8OnSchematicOne = new(10, 10, 5, 5);
    private static readonly Rect U8OnPcbTop = new(20, 20, 5, 5);
    private static readonly Rect U8SecondOnPcbTop = new(30, 30, 5, 5);
    private static readonly Rect U1OnSchematicOne = new(40, 40, 5, 5);
    private static readonly Rect U2OnSchematicTwo = new(50, 50, 5, 5);
    private static readonly Rect U17OnSchematicTwo = new(60, 60, 5, 5);
    private static readonly Rect Cr4OnPcbTop = new(70, 70, 5, 5);

    [Fact]
    public void Selecting_components_highlights_them_on_every_schematic_they_appear_on()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTabWithBoardHighlights();

            tab.UpdateHighlightsForComponents(new List<string> { U8, U1, U2, U17, CR4 });

            // U8 spans two schematics, so both must light up. This is the "highlighted in the
            // large image and in all the thumbnails" requirement.
            Assert.Equal(
                new[] { U8OnSchematicOne, U1OnSchematicOne },
                RectsFor(tab, SchematicOne));

            Assert.Equal(
                new[] { U2OnSchematicTwo, U17OnSchematicTwo },
                RectsFor(tab, SchematicTwo));

            Assert.Equal(
                new[] { U8OnPcbTop, U8SecondOnPcbTop, Cr4OnPcbTop },
                RectsFor(tab, PcbTop));
        });
    }

    [Fact]
    public void Deselecting_one_component_removes_it_from_every_schematic_and_leaves_the_rest()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTabWithBoardHighlights();

            tab.UpdateHighlightsForComponents(new List<string> { U8, U1, U2, U17, CR4 });

            // The click that deselects U8: the same handler fires with the four survivors.
            tab.UpdateHighlightsForComponents(new List<string> { U1, U2, U17, CR4 });

            // Both of the U8 rectangles are gone from the PCB and its rectangle is gone from
            // schematic 1, but nothing else moved.
            Assert.Equal(new[] { U1OnSchematicOne }, RectsFor(tab, SchematicOne));
            Assert.Equal(new[] { U2OnSchematicTwo, U17OnSchematicTwo }, RectsFor(tab, SchematicTwo));
            Assert.Equal(new[] { Cr4OnPcbTop }, RectsFor(tab, PcbTop));

            Assert.DoesNotContain(U8OnSchematicOne, RectsFor(tab, SchematicOne));
            Assert.DoesNotContain(U8OnPcbTop, RectsFor(tab, PcbTop));
            Assert.DoesNotContain(U8SecondOnPcbTop, RectsFor(tab, PcbTop));
        });
    }

    [Fact]
    public void A_schematic_whose_only_highlight_was_deselected_drops_out_of_the_index_entirely()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTabWithBoardHighlights();

            // U2 and U17 are the only components on schematic 2.
            tab.UpdateHighlightsForComponents(new List<string> { U2, U17 });
            Assert.True(tab.highlightIndexBySchematic.ContainsKey(SchematicTwo));

            tab.UpdateHighlightsForComponents(new List<string> { U8 });

            // Leaving a zero-rect entry behind would make that thumbnail think it still has
            // something to draw.
            Assert.False(tab.highlightIndexBySchematic.ContainsKey(SchematicTwo));
            Assert.True(tab.highlightIndexBySchematic.ContainsKey(PcbTop));
        });
    }

    [Fact]
    public void Deselecting_the_last_component_clears_every_highlight()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTabWithBoardHighlights();

            tab.UpdateHighlightsForComponents(new List<string> { U8, U1, U2, U17, CR4 });
            Assert.NotEmpty(tab.highlightIndexBySchematic);

            tab.UpdateHighlightsForComponents(new List<string>());

            Assert.Empty(tab.highlightIndexBySchematic);
        });
    }

    [Fact]
    public void Board_labels_are_matched_regardless_of_case()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTabWithBoardHighlights();

            // Board labels come from contributed Excel files, where casing is not guaranteed.
            tab.UpdateHighlightsForComponents(new List<string> { "u8" });

            Assert.Equal(new[] { U8OnSchematicOne }, RectsFor(tab, SchematicOne));
            Assert.Equal(new[] { U8OnPcbTop, U8SecondOnPcbTop }, RectsFor(tab, PcbTop));
        });
    }

    [Fact]
    public void Blank_and_whitespace_labels_are_ignored_rather_than_matching_everything()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTabWithBoardHighlights();

            tab.UpdateHighlightsForComponents(new List<string> { string.Empty, "   ", U1 });

            Assert.Equal(new[] { U1OnSchematicOne }, RectsFor(tab, SchematicOne));
            Assert.False(tab.highlightIndexBySchematic.ContainsKey(PcbTop));
        });
    }

    [Fact]
    public void Reselecting_a_component_brings_its_highlights_back_everywhere()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTabWithBoardHighlights();

            tab.UpdateHighlightsForComponents(new List<string> { U8, U1 });
            tab.UpdateHighlightsForComponents(new List<string> { U1 });
            tab.UpdateHighlightsForComponents(new List<string> { U8, U1 });

            // A rebuild-from-scratch has to be symmetric: clicking U8 off and on again must
            // land back on exactly the original highlights, on both schematics.
            Assert.Equal(new[] { U8OnSchematicOne, U1OnSchematicOne }, RectsFor(tab, SchematicOne));
            Assert.Equal(new[] { U8OnPcbTop, U8SecondOnPcbTop }, RectsFor(tab, PcbTop));
        });
    }

    // ###########################################################################################
    // A tab holding the highlight rectangles a loaded board would have produced: three
    // schematics, with U8 deliberately present on two of them.
    // ###########################################################################################
    private static TabSchematics CreateTabWithBoardHighlights()
    {
        var tab = new TabSchematics();

        tab.highlightRectsBySchematicAndLabel =
            new Dictionary<string, Dictionary<string, List<Rect>>>(StringComparer.OrdinalIgnoreCase)
            {
                [SchematicOne] = new(StringComparer.OrdinalIgnoreCase)
                {
                    [U8] = new List<Rect> { U8OnSchematicOne },
                    [U1] = new List<Rect> { U1OnSchematicOne },
                },
                [SchematicTwo] = new(StringComparer.OrdinalIgnoreCase)
                {
                    [U2] = new List<Rect> { U2OnSchematicTwo },
                    [U17] = new List<Rect> { U17OnSchematicTwo },
                },
                [PcbTop] = new(StringComparer.OrdinalIgnoreCase)
                {
                    // Two rectangles for one label: a component can appear more than once.
                    [U8] = new List<Rect> { U8OnPcbTop, U8SecondOnPcbTop },
                    [CR4] = new List<Rect> { Cr4OnPcbTop },
                },
            };

        return tab;
    }

    private static Rect[] RectsFor(TabSchematics tab, string schematicName)
    {
        if (!tab.highlightIndexBySchematic.TryGetValue(schematicName, out var index))
        {
            return Array.Empty<Rect>();
        }

        var rects = new Rect[index.Count];
        for (int i = 0; i < index.Count; i++)
        {
            rects[i] = index.GetRect(i);
        }

        return rects;
    }
}
