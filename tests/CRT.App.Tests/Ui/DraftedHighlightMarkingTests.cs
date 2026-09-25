using Avalonia;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// TabSchematics.UpdateHighlightsForComponents marking a rebuilt HighlightSpatialIndex's rects
// as drafted (NewContributeStrategy.md Phase 2, session 2b, task 5 - "a tint on a drafted
// highlight"). The overlay's actual paint colour is not tested here (rendering is out of scope
// per CLAUDE.md); what is tested is that the right rect, by POSITION in the rebuilt index, ends
// up marked - the thing SchematicHighlightsOverlay.Render reads per visible rect to choose which
// brush to use.
//
// Modelled on ComponentHighlightSelectionTests' own fixture shape (same schematic/label
// dictionary), extended with draftedHighlightKeys - the field Main.SetComponentHighlightRects
// sets alongside highlightRectsBySchematicAndLabel on a real board load.
// ###########################################################################################
[Collection("HeadlessUi")]
public class DraftedHighlightMarkingTests
{
    private const string SchematicOne = "Schematic 1";
    private const string SchematicTwo = "Schematic 2";

    private const string U8 = "U8";
    private const string U1 = "U1";
    private const string U2 = "U2";

    private static readonly Rect U8OnSchematicOne = new(10, 10, 5, 5);
    private static readonly Rect U1OnSchematicOne = new(40, 40, 5, 5);
    private static readonly Rect U2OnSchematicTwo = new(50, 50, 5, 5);

    private static TabSchematics CreateTab(HashSet<string>? draftedHighlightKeys = null)
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
                },
            };

        if (draftedHighlightKeys != null)
        {
            tab.draftedHighlightKeys = draftedHighlightKeys;
        }

        return tab;
    }

    private static bool[] DraftedFlagsFor(TabSchematics tab, string schematicName)
    {
        Assert.True(tab.highlightIndexBySchematic.TryGetValue(schematicName, out var index));

        var flags = new bool[index!.Count];
        for (int i = 0; i < index.Count; i++)
        {
            flags[i] = index.GetIsDrafted(i);
        }

        return flags;
    }

    [Fact]
    public void With_no_drafted_keys_at_all_nothing_is_marked()
    {
        UiTest.Run(() =>
        {
            TabSchematics tab = CreateTab();

            tab.UpdateHighlightsForComponents(new List<string> { U8, U1 });

            Assert.All(DraftedFlagsFor(tab, SchematicOne), flag => Assert.False(flag));
        });
    }

    [Fact]
    public void A_drafted_labels_rect_is_marked_and_its_sibling_on_the_same_schematic_is_not()
    {
        UiTest.Run(() =>
        {
            var draftedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                BoardDraftNaturalKeys.ForComponentHighlight(SchematicOne, U8)
            };
            TabSchematics tab = CreateTab(draftedKeys);

            tab.UpdateHighlightsForComponents(new List<string> { U8, U1 });

            Assert.True(tab.highlightIndexBySchematic.TryGetValue(SchematicOne, out var index));

            bool foundDraftedU8 = false;
            for (int i = 0; i < index!.Count; i++)
            {
                bool isU8 = index.GetRect(i) == U8OnSchematicOne;
                Assert.Equal(isU8, index.GetIsDrafted(i));
                foundDraftedU8 |= isU8 && index.GetIsDrafted(i);
            }

            Assert.True(foundDraftedU8);
        });
    }

    [Fact]
    public void The_same_label_on_a_different_schematic_is_not_marked_drafted()
    {
        // U1's key is scoped to "Schematic 1" only - a highlight key naming a different schematic
        // must not leak a "drafted" mark onto an unrelated schematic's rect for the same label.
        UiTest.Run(() =>
        {
            var draftedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                BoardDraftNaturalKeys.ForComponentHighlight("A schematic nothing references", U2)
            };
            TabSchematics tab = CreateTab(draftedKeys);

            tab.UpdateHighlightsForComponents(new List<string> { U2 });

            Assert.All(DraftedFlagsFor(tab, SchematicTwo), flag => Assert.False(flag));
        });
    }

    [Fact]
    public void Deselecting_a_drafted_component_and_reselecting_it_still_marks_it()
    {
        // The index is rebuilt from scratch on every selection change (see
        // ComponentHighlightSelectionTests' own header comment) - the drafted mark must survive
        // that rebuild, not just the first one.
        UiTest.Run(() =>
        {
            var draftedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                BoardDraftNaturalKeys.ForComponentHighlight(SchematicOne, U8)
            };
            TabSchematics tab = CreateTab(draftedKeys);

            tab.UpdateHighlightsForComponents(new List<string> { U8, U1 });
            tab.UpdateHighlightsForComponents(new List<string> { U1 });
            tab.UpdateHighlightsForComponents(new List<string> { U8, U1 });

            Assert.True(tab.highlightIndexBySchematic.TryGetValue(SchematicOne, out var index));
            bool anyDrafted = false;
            for (int i = 0; i < index!.Count; i++)
            {
                anyDrafted |= index.GetRect(i) == U8OnSchematicOne && index.GetIsDrafted(i);
            }

            Assert.True(anyDrafted);
        });
    }
}
