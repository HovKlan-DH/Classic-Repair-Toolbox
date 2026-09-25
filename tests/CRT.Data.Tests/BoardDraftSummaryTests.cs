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
