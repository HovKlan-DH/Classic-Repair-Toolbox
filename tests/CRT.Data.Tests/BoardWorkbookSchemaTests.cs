using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardWorkbookSchema's row mappers - the direction from cells back to BoardData, moved here
// from BoardDataReader on 2026-09-24 so the Drafts tab's table editor could save through the
// SAME mapping the reader loads through.
//
// The round-trip test is the one that matters: BuildRows and the mappers are two halves of one
// contract, and a column added to one and not the other would load or save as silently blank.
// It fills EVERY property of every entry type by reflection, so a new field is covered the day
// it exists rather than the day somebody remembers to add it here.
// ###########################################################################################
public sealed class BoardWorkbookSchemaTests
{
    // One entry of the given type with every string property set to something distinct.
    private static T Filled<T>(string seed) where T : new()
    {
        var entry = new T();

        foreach (PropertyInfo property in typeof(T).GetProperties().Where(p => p.PropertyType == typeof(string)))
        {
            // init-only setters are still ordinary setters to reflection.
            property.SetValue(entry, $"{seed}-{property.Name}");
        }

        return entry;
    }

    private static BoardData FullyPopulatedBoard() => new()
    {
        RevisionDate = "2026-09-24",
        HardwareName = "Commodore 64",
        BoardName = "250407",
        Schematics = [Filled<BoardSchematicEntry>("a"), Filled<BoardSchematicEntry>("b")],
        Components = [Filled<ComponentEntry>("a"), Filled<ComponentEntry>("b")],
        ComponentImages = [Filled<ComponentImageEntry>("a")],
        ComponentHighlights = [Filled<ComponentHighlightEntry>("a")],
        ComponentLocalFiles = [Filled<ComponentLocalFileEntry>("a")],
        ComponentLinks = [Filled<ComponentLinkEntry>("a")],
        BoardLocalFiles = [Filled<BoardLocalFileEntry>("a")],
        BoardLinks = [Filled<BoardLinkEntry>("a")],
        Credits = [Filled<CreditEntry>("a")],
        KiCadImportantSignals = [Filled<KiCadImportantSignalEntry>("a")],
    };

    private static IReadOnlyDictionary<string, string> Cells(params (string Column, string Value)[] cells)
    {
        var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach ((string column, string value) in cells)
        {
            row[column] = value;
        }

        return row;
    }

    [Fact]
    public void Every_sheet_round_trips_through_BuildRows_and_back_with_every_field_intact()
    {
        BoardData board = BoardWorkbookSchemaTests.FullyPopulatedBoard();

        foreach (BoardWorkbookSchema.SheetDefinition sheet in BoardWorkbookSchema.AllSheets)
        {
            IReadOnlyList<object> original = BoardWorkbookSchema.EntriesOf(sheet, board);
            IReadOnlyList<object> mapped = BoardWorkbookSchema.MapRows(sheet, BoardWorkbookSchema.BuildRows(sheet, board));

            Assert.Equal(original.Count, mapped.Count);

            for (int i = 0; i < original.Count; i++)
            {
                foreach (PropertyInfo property in original[i].GetType().GetProperties())
                {
                    Assert.True(
                        Equals(property.GetValue(original[i]), property.GetValue(mapped[i])),
                        $"[{sheet.SheetName}] row {i} lost [{property.Name}] on the round trip");
                }
            }
        }
    }

    [Fact]
    public void EntriesOf_lines_up_index_for_index_with_BuildRows()
    {
        // The table editor pairs row N of BuildRows with entry N of EntriesOf to build each
        // published row's key, so the two must never disagree about order or count.
        BoardData board = BoardWorkbookSchemaTests.FullyPopulatedBoard();

        IReadOnlyList<IReadOnlyDictionary<string, string>> rows =
            BoardWorkbookSchema.BuildRows(BoardWorkbookSchema.Components, board);
        IReadOnlyList<object> entries = BoardWorkbookSchema.EntriesOf(BoardWorkbookSchema.Components, board);

        Assert.Equal(2, entries.Count);
        Assert.Equal("a-BoardLabel", rows[0][BoardWorkbookSchema.ColBoardLabel]);
        Assert.Equal("a-BoardLabel", ((ComponentEntry)entries[0]).BoardLabel);
        Assert.Equal("b-BoardLabel", ((ComponentEntry)entries[1]).BoardLabel);
    }

    [Fact]
    public void A_column_missing_from_a_row_reads_as_empty_rather_than_throwing()
    {
        // An older workbook lacking a newer column must still load - see the schema's own header.
        List<ComponentEntry> mapped = BoardWorkbookSchema.MapComponents(
            [BoardWorkbookSchemaTests.Cells((BoardWorkbookSchema.ColBoardLabel, "U8"))]);

        ComponentEntry component = Assert.Single(mapped);
        Assert.Equal("U8", component.BoardLabel);
        Assert.Equal(string.Empty, component.PartNumber);
    }

    [Fact]
    public void Column_names_match_case_insensitively_through_the_row_dictionary()
    {
        List<ComponentEntry> mapped = BoardWorkbookSchema.MapComponents(
            [BoardWorkbookSchemaTests.Cells(("BOARD LABEL", "U8"))]);

        Assert.Equal("U8", Assert.Single(mapped).BoardLabel);
    }

    [Theory]
    [InlineData("CLK", "")]
    [InlineData("", "Net-(U1-Pad3)")]
    [InlineData("  ", "Net-(U1-Pad3)")]
    public void An_important_signal_missing_either_half_is_dropped(string displayName, string netName)
    {
        // The one mapper that drops rows. The table editor marks such a row "incomplete" by asking
        // this mapper, so the condition must live here and only here.
        List<KiCadImportantSignalEntry> mapped = BoardWorkbookSchema.MapKiCadImportantSignals(
        [
            BoardWorkbookSchemaTests.Cells(
                (BoardWorkbookSchema.ColDisplayName, displayName),
                (BoardWorkbookSchema.ColKiCadNetName, netName))
        ]);

        Assert.Empty(mapped);
    }

    [Fact]
    public void WithRows_replaces_only_the_named_section_and_carries_everything_else_across()
    {
        BoardData board = BoardWorkbookSchemaTests.FullyPopulatedBoard();

        BoardData updated = BoardWorkbookSchema.WithRows(
            board,
            BoardWorkbookSchema.Components,
            [BoardWorkbookSchemaTests.Cells((BoardWorkbookSchema.ColBoardLabel, "U99"))]);

        Assert.Equal("U99", Assert.Single(updated.Components).BoardLabel);

        // Shared, not copied - the same lists, per BoardData.WithRevisionDate's convention.
        Assert.Same(board.Schematics, updated.Schematics);
        Assert.Same(board.ComponentImages, updated.ComponentImages);
        Assert.Same(board.Credits, updated.Credits);
        Assert.Same(board.KiCadImportantSignals, updated.KiCadImportantSignals);

        // Not in any sheet, so no sheet edit may touch them.
        Assert.Same(board.ComponentHighlights, updated.ComponentHighlights);
        Assert.Equal("2026-09-24", updated.RevisionDate);
        Assert.Equal("Commodore 64", updated.HardwareName);
        Assert.Equal("250407", updated.BoardName);

        // And the input was not mutated.
        Assert.Equal(2, board.Components.Count);
    }

    [Fact]
    public void WithRows_can_replace_every_sheet_in_turn()
    {
        // A section left out of WithRows' copy would be erased by the table editor's save, which
        // replaces all nine. Each sheet is replaced with an empty one and must come back empty.
        BoardData board = BoardWorkbookSchemaTests.FullyPopulatedBoard();

        foreach (BoardWorkbookSchema.SheetDefinition sheet in BoardWorkbookSchema.AllSheets)
        {
            BoardData updated = BoardWorkbookSchema.WithRows(board, sheet, []);

            Assert.Empty(BoardWorkbookSchema.EntriesOf(sheet, updated));

            foreach (BoardWorkbookSchema.SheetDefinition other in BoardWorkbookSchema.AllSheets.Where(s => s != sheet))
            {
                Assert.Same(BoardWorkbookSchema.EntriesOf(other, board), BoardWorkbookSchema.EntriesOf(other, updated));
            }
        }
    }

    [Fact]
    public void An_unknown_sheet_is_refused_by_every_dispatcher()
    {
        // An unknown sheet is a programming error, not contributed data. Answering it with an
        // empty list would, on the save side, write a board with a whole section missing.
        var unknown = new BoardWorkbookSchema.SheetDefinition("Not a sheet", ["A"], ["A"]);

        Assert.Throws<ArgumentOutOfRangeException>(() => BoardWorkbookSchema.MapRows(unknown, []));
        Assert.Throws<ArgumentOutOfRangeException>(() => BoardWorkbookSchema.EntriesOf(unknown, new BoardData()));
        Assert.Throws<ArgumentOutOfRangeException>(() => BoardWorkbookSchema.WithRows(new BoardData(), unknown, []));
    }

    // ###########################################################################################
    // *** THE SHEETS COME IN THE PUBLISHED WORKBOOKS' OWN ORDER, "Credits" LAST (2026-09-24). ***
    // Every published board workbook that has an "Important signals" sheet - thirteen of them when
    // this was checked - puts it BEFORE "Credits". This list had them the other way round, so the
    // table editor's sheet tabs disagreed with the workbook the maintainer knows (reported), and
    // every workbook the app wrote ended in "Important signals". AllSheets is also the WRITE order
    // (BoardWorkbookWriter), which is what this pins for the file side.
    // ###########################################################################################
    [Fact]
    public void The_sheets_are_in_the_published_workbooks_order_with_Credits_last()
    {
        Assert.Equal(
            [
                BoardWorkbookSchema.SheetBoardSchematics,
                BoardWorkbookSchema.SheetComponents,
                BoardWorkbookSchema.SheetComponentImages,
                BoardWorkbookSchema.SheetComponentLocalFiles,
                BoardWorkbookSchema.SheetComponentLinks,
                BoardWorkbookSchema.SheetBoardLocalFiles,
                BoardWorkbookSchema.SheetBoardLinks,
                BoardWorkbookSchema.SheetKiCadImportantSignals,
                BoardWorkbookSchema.SheetCredits,
            ],
            BoardWorkbookSchema.AllSheets.Select(sheet => sheet.SheetName));
    }
}
