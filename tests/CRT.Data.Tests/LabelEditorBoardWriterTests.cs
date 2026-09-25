using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Saving a schematic label editor session into a draft's BOARD
// (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
//
// Replaces LabelEditorDraftWriterTests, which pinned the same behaviours expressed as BoardDraft
// row deltas. The behaviours are unchanged and are re-asserted here against the board itself;
// what has gone is everything that existed only to describe an overlay - Added-vs-Modified,
// deletion tombstones, and the "does this already exist officially" duplicate guard.
//
// The rule worth reading before touching this file is SCHEMATIC-SCOPED REPLACE: the editor hands
// over the set of rectangles it now holds for ONE schematic, so anything missing from that set
// has been deleted, and every other schematic must be left alone.
// ###########################################################################################
public sealed class LabelEditorBoardWriterTests
{
    private static LabelEditorSaveRow Row(
        string boardLabel,
        double x = 10,
        double y = 20,
        double width = 30,
        double height = 40,
        string category = "IC")
        => new()
        {
            BoardLabel = boardLabel,
            Category = category,
            X = x,
            Y = y,
            Width = width,
            Height = height,
        };

    private static ComponentHighlightEntry Highlight(string schematic, string label) => new()
    {
        SchematicName = schematic,
        BoardLabel = label,
        X = "1",
        Y = "2",
        Width = "3",
        Height = "4",
    };

    // ------------------------------------------------------------------ Highlights

    [Fact]
    public void A_drawn_rectangle_becomes_a_highlight_on_that_schematic()
    {
        var board = new BoardData();

        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("U8")], region: "");

        ComponentHighlightEntry highlight = Assert.Single(result.ComponentHighlights);

        Assert.Equal("Sheet 1", highlight.SchematicName);
        Assert.Equal("U8", highlight.BoardLabel);
    }

    [Fact]
    public void An_EXISTING_highlight_on_the_same_schematic_is_replaced_not_duplicated()
    {
        var board = new BoardData();
        board.ComponentHighlights.Add(LabelEditorBoardWriterTests.Highlight("Sheet 1", "U8"));

        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("U8", x: 99)], region: "");

        ComponentHighlightEntry highlight = Assert.Single(result.ComponentHighlights);
        Assert.Equal("99", highlight.X);
    }

    // ###########################################################################################
    // *** DELETING A RECTANGLE WORKS BECAUSE THE SAVE IS A REPLACE. ***
    //
    // The editor never says "delete this one" - it hands over the set it now holds, and anything
    // missing has been removed. Merging row by row instead would make a deleted rectangle
    // immortal, which is exactly the bug the original .xlsx writer avoided and the delta writer
    // had to reproduce with tombstones.
    // ###########################################################################################
    [Fact]
    public void A_rectangle_the_editor_no_longer_holds_is_DELETED()
    {
        var board = new BoardData();
        board.ComponentHighlights.Add(LabelEditorBoardWriterTests.Highlight("Sheet 1", "U8"));
        board.ComponentHighlights.Add(LabelEditorBoardWriterTests.Highlight("Sheet 1", "U9"));

        // The session now holds only U8.
        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("U8")], region: "");

        Assert.Single(result.ComponentHighlights);
        Assert.Equal("U8", result.ComponentHighlights.Single().BoardLabel);
    }

    // ###########################################################################################
    // *** AND EVERY OTHER SCHEMATIC IS LEFT COMPLETELY ALONE. ***
    //
    // The other half of the replace rule, and the more dangerous one to get wrong: a save that
    // rebuilt the whole highlight list would wipe every rectangle on every other page of the
    // board, from an editor session the contributor thought was scoped to one sheet.
    // ###########################################################################################
    [Fact]
    public void Highlights_on_ANOTHER_schematic_survive_untouched()
    {
        var board = new BoardData();
        board.ComponentHighlights.Add(LabelEditorBoardWriterTests.Highlight("Sheet 1", "U8"));
        board.ComponentHighlights.Add(LabelEditorBoardWriterTests.Highlight("Sheet 2", "U9"));

        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("U8")], region: "");

        Assert.Contains(result.ComponentHighlights, h => h.SchematicName == "Sheet 2" && h.BoardLabel == "U9");
    }

    [Fact]
    public void An_EMPTY_save_clears_that_schematic_and_only_that_schematic()
    {
        // Deleting the last rectangle on a page is a real thing to do, and it must not be mistaken
        // for "nothing to save".
        var board = new BoardData();
        board.ComponentHighlights.Add(LabelEditorBoardWriterTests.Highlight("Sheet 1", "U8"));
        board.ComponentHighlights.Add(LabelEditorBoardWriterTests.Highlight("Sheet 2", "U9"));

        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(board, "Sheet 1", [], region: "");

        Assert.Single(result.ComponentHighlights);
        Assert.Equal("Sheet 2", result.ComponentHighlights.Single().SchematicName);
    }

    [Fact]
    public void A_row_with_NO_board_label_is_skipped()
    {
        // The natural key for a highlight is schematic + label, so a blank one could never be
        // found, replaced or deleted again - it would be an unreachable row on the board.
        var board = new BoardData();

        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("   ")], region: "");

        Assert.Empty(result.ComponentHighlights);
    }

    // ###########################################################################################
    // *** COORDINATES ARE WRITTEN INVARIANTLY - the third face of a bug documented twice already
    // in this codebase. *** They are stored as strings and parsed back with an invariant parse, so
    // a culture-formatted "12,5" on a Danish machine reads back as 125 or fails outright, and the
    // rectangle lands somewhere else entirely or nowhere at all.
    // ###########################################################################################
    [Fact]
    public void Coordinates_are_written_the_same_in_every_culture()
    {
        CultureInfo original = Thread.CurrentThread.CurrentCulture;

        try
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("da-DK");

            BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
                new BoardData(), "Sheet 1", [LabelEditorBoardWriterTests.Row("U8", x: 12.5)], region: "");

            Assert.Equal("12.5", result.ComponentHighlights.Single().X);
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = original;
        }
    }

    // ------------------------------------------------------------------ Components

    [Fact]
    public void A_genuinely_NEW_board_label_gains_a_component_row()
    {
        var board = new BoardData();

        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("U8", category: "IC")], region: "PAL");

        ComponentEntry component = Assert.Single(result.Components);

        Assert.Equal("U8", component.BoardLabel);
        Assert.Equal("IC", component.Category);
        Assert.Equal("PAL", component.Region);
    }

    // ###########################################################################################
    // An EXISTING component is left completely alone - the label editor's job is rectangles.
    //
    // It knows nothing about a component's friendly name, part number or description, so writing
    // a row for one that already exists would blank all three.
    // ###########################################################################################
    [Fact]
    public void An_EXISTING_component_is_not_touched_or_duplicated()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry
        {
            BoardLabel = "U8",
            FriendlyName = "CPU",
            PartNumber = "6510",
            Region = "PAL",
        });

        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("U8")], region: "PAL");

        ComponentEntry component = Assert.Single(result.Components);

        Assert.Equal("CPU", component.FriendlyName);
        Assert.Equal("6510", component.PartNumber);
    }

    // ###########################################################################################
    // *** A BLANK REGION ON EITHER SIDE IS A WILDCARD. ***
    //
    // Carried over verbatim from both previous writers. Without it, a label already known under no
    // particular region would gain a duplicate row every time it was drawn on a region-scoped
    // schematic - and duplicates on a natural key are exactly what the differ has to guess about.
    // ###########################################################################################
    [Fact]
    public void A_component_with_NO_region_matches_a_region_scoped_save()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU" });

        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("U8")], region: "PAL");

        Assert.Single(result.Components);
    }

    [Fact]
    public void A_component_in_a_DIFFERENT_named_region_is_a_different_component()
    {
        // Both regions named and different - not a wildcard, so this really is a new component.
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U8", Region = "NTSC" });

        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("U8")], region: "PAL");

        Assert.Equal(2, result.Components.Count);
    }

    [Fact]
    public void Board_labels_are_matched_case_insensitively()
    {
        var board = new BoardData();
        board.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU" });

        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("u8")], region: "");

        Assert.Single(result.Components);
    }

    // ------------------------------------------------------------------ Everything else

    // ###########################################################################################
    // The writer returns a WHOLE board, so a section left out of that copy is a section erased.
    // Cheap to get wrong, and silent - the rows would simply be gone from the workbook on the next
    // save, with nothing to say why.
    // ###########################################################################################
    [Fact]
    public void Every_OTHER_section_of_the_board_survives_the_save()
    {
        var board = new BoardData
        {
            RevisionDate = "2026-09-01",
            Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1" }],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U8", Name = "Clock" }],
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U8", Name = "Ref" }],
            BoardLocalFiles = [new BoardLocalFileEntry { Category = "Docs", Name = "Manual" }],
            BoardLinks = [new BoardLinkEntry { Category = "Docs", Name = "Wiki" }],
            Credits = [new CreditEntry { Category = "Data", NameOrHandle = "Dennis" }],
            KiCadImportantSignals = [new KiCadImportantSignalEntry { DisplayName = "PHI2" }],
        };

        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("U8")], region: "");

        Assert.Equal("2026-09-01", result.RevisionDate);
        Assert.Single(result.Schematics);
        Assert.Single(result.ComponentImages);
        Assert.Single(result.ComponentLocalFiles);
        Assert.Single(result.ComponentLinks);
        Assert.Single(result.BoardLocalFiles);
        Assert.Single(result.BoardLinks);
        Assert.Single(result.Credits);
        Assert.Single(result.KiCadImportantSignals);
    }

    [Fact]
    public void The_INPUT_board_is_never_mutated()
    {
        // A failed write must leave the caller's copy untouched, and DraftWorkbookStore.Edit
        // relies on that: it hands over the board it read and writes whatever comes back.
        var board = new BoardData();
        board.ComponentHighlights.Add(LabelEditorBoardWriterTests.Highlight("Sheet 1", "U8"));

        LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("U9")], region: "");

        Assert.Single(board.ComponentHighlights);
        Assert.Equal("U8", board.ComponentHighlights.Single().BoardLabel);
    }

    // A component the label editor creates joins its category in label order (2026-09-24) - the
    // row's position is where it appears in the application's own list, so the bottom of the
    // sheet was the wrong place for a U2 on a board listing U1 and U3.
    [Fact]
    public void A_NEW_component_from_the_label_editor_joins_its_category_in_label_order()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Category = "IC" },
                new ComponentEntry { BoardLabel = "U3", Category = "IC" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
            ],
        };

        BoardData result = LabelEditorBoardWriter.ApplyLabelEditorSave(
            board, "Sheet 1", [LabelEditorBoardWriterTests.Row("U2", category: "IC")], region: "");

        Assert.Equal(["U1", "U2", "U3", "C1"], result.Components.Select(component => component.BoardLabel));
    }
}
