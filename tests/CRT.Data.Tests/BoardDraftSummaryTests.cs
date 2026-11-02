using System.Collections.Generic;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Which rows the UI marks as drafted - the component list's "drafted" chip and the schematic
// overlay's tint.
//
// *** THE WAY IN CHANGED IN PHASE 6 (2026-09-23), THE OUTPUT DID NOT. *** This class used to be
// built from a BoardDraft's own delta rows (FromDraft). With the draft stored as a board
// workbook there is no delta list - and there must not be one, because the contributor can edit
// that workbook in Excel with the application closed - so the drafted rows are DERIVED by
// comparing the draft against the published board, and FromChanges turns that comparison into
// the same two sets the UI has always read.
//
// The FromDraft tests that lived here are gone with the method. These replace them, and they
// matter: FromChanges had NO coverage when it was written, and it is what decides whether a
// contributor can see which rows are theirs.
// ###########################################################################################
public sealed class BoardDraftSummaryTests
{
    private static BoardRowChange Change(
        string section,
        string naturalKey,
        BoardRowChangeKind kind = BoardRowChangeKind.Modified)
        => new()
        {
            Section = section,
            NaturalKey = naturalKey,
            DisplayLabel = naturalKey,
            Kind = kind,
        };

    // A component whose own rows only moved says "Draft" too (code review, 2026-11-02) - the mark the
    // Contribute tab's list and the component list read once the board is shown again.
    [Fact]
    public void A_component_with_its_own_rows_in_another_order_is_marked_though_no_row_changed()
    {
        BoardDraftSummary summary = BoardDraftSummary.FromChanges([], new HashSet<string> { "U8" });

        Assert.Contains("u8", summary.DraftedComponentBoardLabels);
        Assert.True(summary.HasAnyMarkers);
        Assert.Same(BoardDraftSummary.Empty, BoardDraftSummary.FromChanges([], new HashSet<string>()));
    }

    [Fact]
    public void A_changed_component_is_marked_by_its_board_label()
    {
        BoardDraftSummary summary = BoardDraftSummary.FromChanges(
        [
            BoardDraftSummaryTests.Change("Components", BoardDraftNaturalKeys.ForComponent("U8")),
        ]);

        Assert.Contains("U8", summary.DraftedComponentBoardLabels);
        Assert.True(summary.HasAnyMarkers);
    }

    // ###########################################################################################
    // The highlight key is "<schematic>U+241F<label>" - BoardDraftNaturalKeys built it, and it is
    // the exact shape the schematic overlay looks up. Carried through VERBATIM rather than rebuilt,
    // because a key rebuilt slightly differently here would silently match nothing and the tint
    // would simply never appear.
    // ###########################################################################################
    [Fact]
    public void A_changed_highlight_is_marked_by_its_schematic_and_label_key()
    {
        string key = BoardDraftNaturalKeys.ForComponentHighlight("Sheet 1", "U8");

        BoardDraftSummary summary = BoardDraftSummary.FromChanges(
        [
            BoardDraftSummaryTests.Change("Component highlights", key),
        ]);

        Assert.Contains(key, summary.DraftedHighlightKeys);
    }

    [Fact]
    public void An_ADDED_row_is_marked_exactly_like_a_modified_one()
    {
        // Both are on the board and both are the contributor's - the chip says "this row is yours",
        // not "this row was edited rather than created".
        BoardDraftSummary summary = BoardDraftSummary.FromChanges(
        [
            BoardDraftSummaryTests.Change(
                "Components",
                BoardDraftNaturalKeys.ForComponent("U9"),
                BoardRowChangeKind.Added),
        ]);

        Assert.Contains("U9", summary.DraftedComponentBoardLabels);
    }

    // ###########################################################################################
    // *** A DELETED ROW IS NOT MARKED, and that is the one rule here worth stating. ***
    //
    // A row absent from the draft is not ON the rendered board, so there is nothing on screen to
    // mark. Marking it would mean drawing a "drafted" chip for a component the contributor has
    // removed - a row that does not exist in the list the chip would appear in.
    //
    // FromDraft excluded Deleted tombstones for exactly the same reason.
    // ###########################################################################################
    [Fact]
    public void A_DELETED_row_is_never_marked()
    {
        BoardDraftSummary summary = BoardDraftSummary.FromChanges(
        [
            BoardDraftSummaryTests.Change(
                "Components",
                BoardDraftNaturalKeys.ForComponent("U8"),
                BoardRowChangeKind.Deleted),
        ]);

        Assert.Empty(summary.DraftedComponentBoardLabels);
        Assert.False(summary.HasAnyMarkers);
    }

    // ###########################################################################################
    // Only the two sections that HAVE a UI marker are collected. A changed credit or board link is
    // a real change and appears in the change list; it just has nothing on screen to tint, and
    // adding it here would put its natural key into a set the component list searches by label.
    // ###########################################################################################
    [Fact]
    public void A_section_with_NO_ui_marker_contributes_nothing()
    {
        BoardDraftSummary summary = BoardDraftSummary.FromChanges(
        [
            BoardDraftSummaryTests.Change("Credits", "Data␟Dennis"),
            BoardDraftSummaryTests.Change("Board links", "Docs␟Wiki"),
        ]);

        Assert.Same(BoardDraftSummary.Empty, summary);
    }

    // ###########################################################################################
    // *** THE COMPONENT'S OTHER ROWS MARK IT TOO (code review, 2026-11-02). *** Only its row on the
    // Components sheet did, so a pin, a part, a file, a link, a picture or a baseline changed took "Draft"
    // off the component as soon as Main showed the board again, and the Schematics tab's list never had
    // it - while "Review changes" listed the component. The label is the first part of each key
    // (BoardDraftNaturalKeys), and the second of a highlight's.
    // ###########################################################################################
    public static TheoryData<string, string> SectionsOfAComponent => new()
    {
        { "Component images", BoardDraftNaturalKeys.ForComponentImage("U8", "PAL", "6", "Clock") },
        { "Component local files", BoardDraftNaturalKeys.ForComponentLocalFile("U8", "Datasheet", "PAL") },
        { "Component links", BoardDraftNaturalKeys.ForComponentLink("U8", "Article") },
        { BoardWorkbookSchema.SheetComponentPins, BoardDraftNaturalKeys.ForComponentPin("U8", "NTSC", "1") },
        { BoardWorkbookSchema.SheetComponentReplacements, BoardDraftNaturalKeys.ForComponentReplacement("U8", "", "6510") },
        { "Component highlights", BoardDraftNaturalKeys.ForComponentHighlight("Sheet 1", "U8") },
    };

    [Theory]
    [MemberData(nameof(SectionsOfAComponent))]
    public void A_change_to_any_row_of_a_component_marks_the_component_by_its_label(string section, string key)
    {
        foreach (BoardRowChangeKind kind in Enum.GetValues<BoardRowChangeKind>())
        {
            BoardDraftSummary summary = BoardDraftSummary.FromChanges([BoardDraftSummaryTests.Change(section, key, kind)]);

            // A row taken OFF a component that stays is a change of it, as "Review changes" says.
            Assert.Equal(["U8"], summary.DraftedComponentBoardLabels);
        }
    }

    // Through the real comparison, with every section's row type.
    [Fact]
    public void The_comparison_marks_each_component_whose_rows_changed()
    {
        static BoardData Published() => new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1" }, new ComponentEntry { BoardLabel = "U2" }, new ComponentEntry { BoardLabel = "U3" },
                new ComponentEntry { BoardLabel = "U4" }, new ComponentEntry { BoardLabel = "U5" }, new ComponentEntry { BoardLabel = "U6" },
                new ComponentEntry { BoardLabel = "U7" }, new ComponentEntry { BoardLabel = "U8" },
            ],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U1", Pin = "1", Name = "A" }],
            ComponentReplacements = [new ComponentReplacementEntry { BoardLabel = "U2", PartNumber = "X", Kind = "Drop-in" }],
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U3", Name = "Datasheet", File = "a.pdf" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U4", Name = "Article", Url = "https://example.org/a" }],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U5", Name = "Top", File = "a.png" },
                new ComponentImageEntry { BoardLabel = "U6", Pin = "1", Name = "Clock", File = "b.png" },
            ],
            ComponentHighlights =
            [
                new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U7", X = "1", Y = "2", Width = "3", Height = "4" },
                new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U7", X = "10", Y = "2", Width = "3", Height = "4" },
            ],
        };

        BoardData draft = Published();
        draft = draft
            .WithComponentSections(
                componentPins: [new ComponentPinEntry { BoardLabel = "U1", Pin = "1", Name = "B" }],
                componentReplacements: [new ComponentReplacementEntry { BoardLabel = "U2", PartNumber = "X", Kind = "Close" }],
                componentLocalFiles: [new ComponentLocalFileEntry { BoardLabel = "U3", Name = "Datasheet", File = "b.pdf" }],
                componentLinks: [new ComponentLinkEntry { BoardLabel = "U4", Name = "Article", Url = "https://example.org/b" }],
                componentImages:
                [
                    new ComponentImageEntry { BoardLabel = "U5", Name = "Top", File = "c.png" },
                    new ComponentImageEntry { BoardLabel = "U6", Pin = "1", Name = "Clock", File = "d.png" },
                ])
            .WithComponentHighlights(
            [
                new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U7", X = "1", Y = "2", Width = "3", Height = "4" },
                new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U7", X = "20", Y = "2", Width = "3", Height = "4" },
            ]);

        BoardDraftSummary summary = BoardDraftSummary.FromChanges(BoardDataDiffer.Compare(Published(), draft));

        Assert.Equal(["U1", "U2", "U3", "U4", "U5", "U6", "U7"], summary.DraftedComponentBoardLabels.Order());
        Assert.Equal([BoardDraftNaturalKeys.ForComponentHighlight("Sheet 1", "U7")], summary.DraftedHighlightKeys);
    }

    // ###########################################################################################
    // A rectangle that only followed its schematic - the image renamed, or removed with what was marked on
    // it - is the board's change, as ComponentRevert has it: no "Undo changes to", so no "Draft" on every
    // component marked there. The rectangles on screen are tinted, as ever.
    // ###########################################################################################
    [Fact]
    public void A_schematic_renamed_or_removed_marks_no_component_its_rectangles_followed()
    {
        static BoardData Published() => new()
        {
            Schematics =
            [
                new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png" },
                new BoardSchematicEntry { SchematicName = "Back", SchematicImageFile = "back.png" },
            ],
            Components = [new ComponentEntry { BoardLabel = "U8" }, new ComponentEntry { BoardLabel = "U9" }],
            ComponentHighlights =
            [
                new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" },
                new ComponentHighlightEntry { SchematicName = "Back", BoardLabel = "U9", X = "1", Y = "2", Width = "3", Height = "4" },
            ],
        };

        BoardData draft = Published()
            .WithSchematics([new BoardSchematicEntry { SchematicName = "Front", SchematicImageFile = "top.png" }])
            .WithComponentHighlights([new ComponentHighlightEntry { SchematicName = "Front", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" }]);

        BoardDraftSummary summary = BoardDraftSummary.FromChanges(BoardDataDiffer.Compare(Published(), draft));

        Assert.Empty(summary.DraftedComponentBoardLabels);
        Assert.Equal([BoardDraftNaturalKeys.ForComponentHighlight("Front", "U8")], summary.DraftedHighlightKeys);

        // ComponentRevert says the same of each.
        Assert.False(ComponentRevert.HasChanges(draft, Published(), "U8"));
        Assert.False(ComponentRevert.HasChanges(draft, Published(), "U9"));

        // On a schematic renamed, a rectangle moved as well is its component's.
        BoardData moved = draft.WithComponentHighlights([new ComponentHighlightEntry { SchematicName = "Front", BoardLabel = "U8", X = "50", Y = "2", Width = "3", Height = "4" }]);

        Assert.Contains("U8", BoardDraftSummary.FromChanges(BoardDataDiffer.Compare(Published(), moved)).DraftedComponentBoardLabels);
    }

    // ###########################################################################################
    // *** A ROW GIVEN TO ANOTHER COMPONENT CHANGES BOTH (code review, 2026-11-02). *** A pin's Board label
    // changed from U1 to U2 in the table: a paired row carried only the draft's key, so U2 got the chip and U1
    // none - though U1 has its "Undo changes to" and a group "1 pin removed". The same for a rectangle.
    // ###########################################################################################
    private static BoardData Two(IEnumerable<ComponentPinEntry>? pins = null, IEnumerable<ComponentHighlightEntry>? rectangles = null) => new()
    {
        Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "a.png" }],
        Components = [new ComponentEntry { BoardLabel = "U1" }, new ComponentEntry { BoardLabel = "U2" }],
        ComponentPins = (pins ?? []).ToList(),
        ComponentHighlights = (rectangles ?? []).ToList(),
    };

    private static ComponentHighlightEntry RectangleOf(string label, string schematic = "Sheet 1") =>
        new() { SchematicName = schematic, BoardLabel = label, X = "1", Y = "2", Width = "3", Height = "4" };

    [Fact]
    public void A_pin_given_to_another_component_marks_both()
    {
        BoardData published = Two(pins: [new ComponentPinEntry { BoardLabel = "U1", Pin = "1", Name = "VCC" }]);
        BoardData draft = Two(pins: [new ComponentPinEntry { BoardLabel = "U2", Pin = "1", Name = "VCC" }]);

        BoardDraftSummary summary = BoardDraftSummary.FromChanges(BoardDataDiffer.Compare(published, draft));

        Assert.Equal(["U1", "U2"], summary.DraftedComponentBoardLabels.Order());

        // What each of them says itself.
        Assert.True(ComponentRevert.HasChanges(draft, published, "U1"));
        Assert.True(ComponentRevert.HasChanges(draft, published, "U2"));
    }

    [Fact]
    public void A_rectangle_given_to_another_component_marks_both()
    {
        BoardDraftSummary summary = BoardDraftSummary.FromChanges(BoardDataDiffer.Compare(
            Two(rectangles: [RectangleOf("U1")]),
            Two(rectangles: [RectangleOf("U2")])));

        Assert.Equal(["U1", "U2"], summary.DraftedComponentBoardLabels.Order());
        Assert.Equal([BoardDraftNaturalKeys.ForComponentHighlight("Sheet 1", "U2")], summary.DraftedHighlightKeys);
    }

    // ###########################################################################################
    // A rectangle whose schematic is spelt in other capitals, the schematic itself as it was, followed nothing:
    // ComponentRevert has it as a change, and "only followed its schematic" must not take its mark away. Handed
    // to another schematic the board still has, it followed nothing either.
    // ###########################################################################################
    [Fact]
    public void A_rectangle_that_followed_no_schematic_marks_its_component()
    {
        BoardData published = Two(rectangles: [RectangleOf("U1")]);
        BoardData capitals = Two(rectangles: [RectangleOf("U1", schematic: "SHEET 1")]);

        Assert.True(ComponentRevert.HasChanges(capitals, published, "U1"));
        Assert.Equal(["U1"], BoardDraftSummary.FromChanges(BoardDataDiffer.Compare(published, capitals)).DraftedComponentBoardLabels);

        BoardData twoSheets = Two(rectangles: [RectangleOf("U1")])
            .WithSchematics(
            [
                new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "a.png" },
                new BoardSchematicEntry { SchematicName = "Sheet 2", SchematicImageFile = "b.png" },
            ]);
        BoardData otherSheet = twoSheets.WithComponentHighlights([RectangleOf("U1", schematic: "Sheet 2")]);

        Assert.True(ComponentRevert.HasChanges(otherSheet, twoSheets, "U1"));
        Assert.Equal(["U1"], BoardDraftSummary.FromChanges(BoardDataDiffer.Compare(twoSheets, otherSheet)).DraftedComponentBoardLabels);
    }

    // A rectangle taken off: the component marked, the rectangle not - it is not on screen to tint.
    [Fact]
    public void A_rectangle_deleted_marks_its_component_and_tints_nothing()
    {
        BoardDraftSummary summary = BoardDraftSummary.FromChanges(
        [
            BoardDraftSummaryTests.Change("Component highlights", BoardDraftNaturalKeys.ForComponentHighlight("Sheet 1", "U8"), BoardRowChangeKind.Deleted),
        ]);

        Assert.Equal(["U8"], summary.DraftedComponentBoardLabels);
        Assert.Empty(summary.DraftedHighlightKeys);
    }

    [Fact]
    public void The_shared_Empty_instance_is_returned_when_nothing_is_marked()
    {
        // Most boards have no draft at all, so the common case must not allocate two HashSets.
        Assert.Same(BoardDraftSummary.Empty, BoardDraftSummary.FromChanges(null));
        Assert.Same(BoardDraftSummary.Empty, BoardDraftSummary.FromChanges([]));
    }

    [Fact]
    public void Board_labels_are_matched_case_insensitively_by_the_consumer()
    {
        // The component list looks up by whatever label its own row carries, which a hand-edited
        // workbook can case differently from the drafted row.
        BoardDraftSummary summary = BoardDraftSummary.FromChanges(
        [
            BoardDraftSummaryTests.Change("Components", BoardDraftNaturalKeys.ForComponent("U8")),
        ]);

        Assert.Contains("u8", summary.DraftedComponentBoardLabels);
    }

    [Fact]
    public void Several_sections_are_collected_in_one_pass()
    {
        string highlightKey = BoardDraftNaturalKeys.ForComponentHighlight("Sheet 1", "U8");

        BoardDraftSummary summary = BoardDraftSummary.FromChanges(
        [
            BoardDraftSummaryTests.Change("Components", BoardDraftNaturalKeys.ForComponent("U8")),
            BoardDraftSummaryTests.Change("Component highlights", highlightKey),
        ]);

        Assert.Contains("U8", summary.DraftedComponentBoardLabels);
        Assert.Contains(highlightKey, summary.DraftedHighlightKeys);
    }

    [Fact]
    public void A_changed_REGIONAL_row_marks_the_component_by_its_label_alone()
    {
        // A regionalised row's key is "label + region" since 2026-09-24; the chip marks the
        // component in the list, which shows the label - not "U8" plus a box character and a region.
        BoardDraftSummary summary = BoardDraftSummary.FromChanges(
        [
            BoardDraftSummaryTests.Change("Components", BoardDraftNaturalKeys.ForComponent("U8", "NTSC"), BoardRowChangeKind.Added),
        ]);

        Assert.Equal(["U8"], summary.DraftedComponentBoardLabels);
    }
}
