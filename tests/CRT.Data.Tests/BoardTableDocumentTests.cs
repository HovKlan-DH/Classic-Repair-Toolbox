using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The Drafts tab's table editor model - BoardTableDocument, BoardTableSheet and the row and
// cell types they hand to the grid (owner request, 2026-09-24).
//
// Colours are the project owner's: red = deleted (shown where the row used to be), orange =
// modified (on the changed cell only), green = added.
//
// *** THE TEST THAT MATTERS MOST is The_change_counts_agree_with_BoardDataDiffer_on_the_board_a_save_writes.
// *** The Drafts tab row says "N rows changed", counted by BoardDataDiffer, directly above the
// sheet tabs this table counts. If the two ever pair rows differently, the screen contradicts
// itself - and a count that is quietly wrong is believed. That test builds one board with every
// awkward case at once and holds the table to the differ, sheet by sheet.
// ###########################################################################################
public sealed class BoardTableDocumentTests
{
    private static ComponentEntry Component(string label, string friendlyName = "", string description = "") =>
        new() { BoardLabel = label, FriendlyName = friendlyName, TechnicalNameOrValue = "x", Description = description };

    private static BoardData Board(params ComponentEntry[] components) => new() { Components = [.. components] };

    private static BoardTableSheet Components(BoardTableDocument document) =>
        document.FindSheet(BoardWorkbookSchema.SheetComponents)!;

    private static int Column(BoardTableSheet sheet, string column) =>
        sheet.Columns.ToList().IndexOf(column);

    private static string Label(BoardTableRow row) =>
        row.Cells[Column(row.Sheet, BoardWorkbookSchema.ColBoardLabel)].Text;

    private static BoardTableRow Row(BoardTableSheet sheet, string label, bool deleted = false) =>
        sheet.Rows.Single(row => row.IsDeleted == deleted && Label(row) == label);

    private static BoardTableCell Cell(BoardTableRow row, string column) =>
        row.Cells[Column(row.Sheet, column)];

    // ------------------------------------------------------------------ The shape

    [Fact]
    public void The_table_has_one_sheet_per_schema_sheet_with_the_schemas_columns_in_order()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, new BoardData());

        Assert.Equal(
            BoardWorkbookSchema.AllSheets.Select(sheet => sheet.SheetName),
            document.Sheets.Select(sheet => sheet.Name));

        Assert.Equal(BoardWorkbookSchema.Components.ColumnOrder, Components(document).Columns);
    }

    [Fact]
    public void The_drafts_rows_are_shown_in_the_drafts_own_order()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            null,
            Board(Component("U3"), Component("U1"), Component("U2")));

        Assert.Equal(["U3", "U1", "U2"], Components(document).Rows.Select(Label));
    }

    // ------------------------------------------------------------------ Colours

    [Fact]
    public void An_unchanged_draft_shows_no_colour_and_counts_nothing()
    {
        BoardData published = Board(Component("U1", "CPU"), Component("U2", "VIC"));
        BoardTableDocument document = BoardTableDocument.Create(published, Board(Component("U1", "CPU"), Component("U2", "VIC")));

        Assert.Equal(0, document.TotalChangeCount);
        Assert.All(Components(document).Rows, row =>
        {
            Assert.Equal(BoardTableRowState.Unchanged, row.State);
            Assert.All(row.Cells, cell => Assert.Equal(BoardTableCellState.Unchanged, cell.State));
        });
        Assert.False(document.HasUnsavedChanges);
    }

    [Fact]
    public void A_changed_value_is_orange_on_that_cell_ONLY_and_carries_the_published_value()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU")),
            Board(Component("U1", "CPU 6510")));

        BoardTableRow row = Row(Components(document), "U1");

        Assert.Equal(BoardTableRowState.Modified, row.State);
        Assert.Equal("~", row.Marker);

        BoardTableCell changed = Cell(row, BoardWorkbookSchema.ColFriendlyName);
        Assert.Equal(BoardTableCellState.Modified, changed.State);
        Assert.Equal("CPU", changed.PublishedText);
        Assert.Equal("Published value: CPU", changed.ToolTip);

        // Every OTHER cell of the same row stays uncoloured - the project owner asked for the
        // changed cell, not the row.
        Assert.All(
            row.Cells.Where(cell => cell != changed),
            cell => Assert.Equal(BoardTableCellState.Unchanged, cell.State));

        Assert.Equal(1, Components(document).ChangeCount);
    }

    [Fact]
    public void A_value_changed_to_empty_says_so_in_its_tooltip()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "")),
            Board(Component("U1", "CPU")));

        Assert.Equal(
            "Published value: (empty)",
            Cell(Row(Components(document), "U1"), BoardWorkbookSchema.ColFriendlyName).ToolTip);
    }

    // ###########################################################################################
    // *** THE CALLER NAMES WHAT THE CELL IS COMPARED WITH (owner report, 2026-09-26). *** The
    // maintainer's NEW-SYSTEM table compares the submission with ITSELF as it arrived, so its
    // baseline is not a published board - and the tooltip still read "Published value: (empty)",
    // naming something that does not exist: "yes, it will always be empty, but it should state
    // [the contributor's value] if it was changed from empty to something".
    //
    // Note the baseline here is a real board, so HasBaseline is TRUE - which is exactly why the
    // label cannot be derived from it and is passed in instead.
    // ###########################################################################################
    [Fact]
    public void The_caller_can_name_what_a_changed_cell_is_compared_with()
    {
        // The maintainer's new system: the submission is its own baseline, and the maintainer has
        // since filled in a description that arrived empty.
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU")),
            Board(Component("U1", "CPU", "Added by the maintainer")),
            baselineLabel: "As submitted");

        Assert.True(document.HasBaseline);

        BoardTableCell filled = Cell(Row(Components(document), "U1"), BoardWorkbookSchema.ColDescription);

        Assert.Equal(BoardTableCellState.Modified, filled.State);
        Assert.Equal("As submitted: (empty)", filled.ToolTip);
    }

    // A draft names nothing, so it keeps the published wording - every existing caller is unchanged.
    [Fact]
    public void A_caller_that_names_nothing_still_says_published_value()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU")),
            Board(Component("U1", "Processor")));

        Assert.Equal(BoardTableDocument.DefaultBaselineLabel, document.BaselineLabel);
        Assert.Equal(
            "Published value: CPU",
            Cell(Row(Components(document), "U1"), BoardWorkbookSchema.ColFriendlyName).ToolTip);
    }

    [Fact]
    public void A_row_only_in_the_draft_is_green_on_every_cell()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1")),
            Board(Component("U1"), Component("U2", "New chip")));

        BoardTableRow added = Row(Components(document), "U2");

        Assert.Equal(BoardTableRowState.Added, added.State);
        Assert.Equal("+", added.Marker);
        Assert.All(added.Cells, cell => Assert.Equal(BoardTableCellState.Added, cell.State));
        Assert.Equal(1, Components(document).ChangeCount);
    }

    [Fact]
    public void A_row_only_published_is_a_red_ghost_shown_WHERE_IT_USED_TO_BE()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1"), Component("U2", "Gone"), Component("U3")),
            Board(Component("U1"), Component("U3")));

        BoardTableSheet sheet = Components(document);

        Assert.Equal(["U1", "U2", "U3"], sheet.Rows.Select(Label));

        BoardTableRow ghost = sheet.Rows[1];
        Assert.True(ghost.IsDeleted);
        Assert.Equal(BoardTableRowState.Deleted, ghost.State);
        Assert.Equal("-", ghost.Marker);
        Assert.Equal("Gone", Cell(ghost, BoardWorkbookSchema.ColFriendlyName).Text);
        Assert.All(ghost.Cells, cell => Assert.Equal(BoardTableCellState.Deleted, cell.State));
        Assert.Equal(1, sheet.ChangeCount);
    }

    [Fact]
    public void A_ghost_whose_earlier_neighbours_are_all_gone_is_shown_at_the_top()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1"), Component("U2")),
            Board(Component("U2")));

        Assert.Equal(["U1", "U2"], Components(document).Rows.Select(Label));
        Assert.True(Components(document).Rows[0].IsDeleted);
    }

    [Fact]
    public void Deleted_rows_in_a_run_keep_their_published_order()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1"), Component("U2"), Component("U3"), Component("U4"), Component("U5")),
            Board(Component("U1"), Component("U5")));

        BoardTableSheet sheet = Components(document);

        Assert.Equal(["U1", "U2", "U3", "U4", "U5"], sheet.Rows.Select(Label));
        Assert.Equal([false, true, true, true, false], sheet.Rows.Select(row => row.IsDeleted));
        Assert.Equal(3, sheet.ChangeCount);
    }

    [Fact]
    public void A_ghost_follows_its_surviving_neighbour_when_the_draft_reordered_the_rows()
    {
        // Rows reordered in Excel: the ghost for U2 goes after U1, wherever U1 now is. Order
        // itself is never a change, so nothing else is coloured.
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1"), Component("U2"), Component("U3")),
            Board(Component("U3"), Component("U1")));

        BoardTableSheet sheet = Components(document);

        Assert.Equal(["U3", "U1", "U2"], sheet.Rows.Select(Label));
        Assert.True(sheet.Rows[2].IsDeleted);
        Assert.Equal(1, sheet.ChangeCount);
    }

    [Fact]
    public void A_ghost_cannot_be_edited()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU")),
            Board());

        BoardTableRow ghost = Components(document).Rows.Single();
        BoardTableCell cell = Cell(ghost, BoardWorkbookSchema.ColFriendlyName);

        cell.Text = "Typed over";

        Assert.Equal("CPU", cell.Text);
        Assert.False(document.HasUnsavedChanges);
    }

    // ------------------------------------------------------------------ Keys and comparison

    [Fact]
    public void Changing_a_key_cell_reads_as_an_addition_plus_a_deleted_ghost()
    {
        // The table must describe the change the way BoardDataDiffer, the submission and the
        // maintainer all will - so a renamed Board label is NOT shown as a friendly "modified".
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U8")),
            Board(Component("U8")));

        BoardTableSheet sheet = Components(document);
        Cell(Row(sheet, "U8"), BoardWorkbookSchema.ColBoardLabel).Text = "U9";
        sheet.Refresh();

        Assert.Equal(BoardTableRowState.Added, Row(sheet, "U9").State);
        Assert.Equal(BoardTableRowState.Deleted, Row(sheet, "U8", deleted: true).State);
        Assert.Equal(2, sheet.ChangeCount);
    }

    [Fact]
    public void Keys_match_case_INSENSITIVELY_while_values_compare_case_SENSITIVELY()
    {
        // BoardDataDiffer's rule: "u8" is the same row as "U8", but that row's Board label cell
        // did change, so it is orange rather than an add-plus-delete.
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U8", "cpu")),
            Board(Component("u8", "CPU")));

        BoardTableRow row = Components(document).Rows.Single();

        Assert.False(row.IsDeleted);
        Assert.Equal(BoardTableRowState.Modified, row.State);
        Assert.Equal(BoardTableCellState.Modified, Cell(row, BoardWorkbookSchema.ColBoardLabel).State);
        Assert.Equal(BoardTableCellState.Modified, Cell(row, BoardWorkbookSchema.ColFriendlyName).State);
    }

    [Fact]
    public void Surrounding_whitespace_is_trimmed_on_the_way_in_and_is_never_a_change()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU")),
            Board(Component("U1", "CPU")));

        BoardTableCell cell = Cell(Components(document).Rows.Single(), BoardWorkbookSchema.ColFriendlyName);
        cell.Text = "  CPU \r\n";

        Assert.Equal("CPU", cell.Text);
        Assert.False(document.HasUnsavedChanges);
    }

    [Fact]
    public void A_second_row_with_the_same_key_is_flagged_and_NOT_counted()
    {
        // BoardDataDiffer pairs only the first row per key and ignores the rest; counting the
        // duplicate here would make the tab disagree with the draft row's own count.
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU")),
            Board(Component("U1", "CPU"), Component("U1", "Another CPU")));

        BoardTableSheet sheet = Components(document);
        BoardTableRow duplicate = sheet.Rows[1];

        Assert.Equal(BoardTableRowState.Unchanged, sheet.Rows[0].State);
        Assert.Equal(BoardTableRowState.Duplicate, duplicate.State);
        Assert.Equal("!", duplicate.Marker);
        Assert.NotNull(duplicate.MarkerToolTip);
        Assert.Equal(0, sheet.ChangeCount);

        // Coloured across the whole row and counted on its own (owner request, 2026-09-24:
        // the "!" alone was easy to miss) - but still not a change.
        Assert.All(duplicate.Cells, cell => Assert.Equal(BoardTableCellState.Flagged, cell.State));
        Assert.All(sheet.Rows[0].Cells, cell => Assert.Equal(BoardTableCellState.Unchanged, cell.State));
        Assert.Equal(1, sheet.FlaggedCount);
    }

    // ###########################################################################################
    // *** A BOARD COMPARED WITH ITSELF MARKS NOTHING - its duplicates included (2026-09-26). *** The
    // maintainer's table compares a NEW system's submission with itself as it was opened, so only
    // the maintainer's own edits are coloured. A duplicate in it is still flagged - it is one - but
    // its twin on the compared side must not come back as a red "deleted" ghost.
    // ###########################################################################################
    [Fact]
    public void A_board_compared_with_itself_marks_nothing_but_its_duplicates()
    {
        BoardData board = Board(Component("U1", "CPU"), Component("U1", "CPU again"), Component("U2", "VIC"));

        BoardTableSheet sheet = Components(BoardTableDocument.Create(board, Board(Component("U1", "CPU"), Component("U1", "CPU again"), Component("U2", "VIC"))));

        Assert.DoesNotContain(sheet.Rows, row => row.IsDeleted);
        Assert.Equal(0, sheet.ChangeCount);
        Assert.Equal(1, sheet.FlaggedCount);
        Assert.Equal(BoardTableRowState.Duplicate, sheet.Rows[1].State);
    }

    // ###########################################################################################
    // *** ONE DISPLAY NAME OVER SEVERAL NETS IS SEVERAL ROWS, NOT DUPLICATES (2026-09-26). ***
    // "9VAC" maps to both the 9VAC and the 9VAC~ net - the sheet exists to say exactly that. Keyed
    // on the name alone, every second one was flagged (reported from the maintainer's table: "the
    // uniqueness here is both columns"). Both columns the same is still a duplicate.
    // ###########################################################################################
    [Fact]
    public void Important_signals_sharing_a_display_name_but_not_a_net_are_not_flagged()
    {
        var board = new BoardData
        {
            KiCadImportantSignals =
            [
                new() { DisplayName = "9VAC", KiCadNetName = "9VAC" },
                new() { DisplayName = "9VAC", KiCadNetName = "9VAC~" },
                new() { DisplayName = "RESET", KiCadNetName = "~{RESET}" },
            ],
        };

        BoardTableSheet sheet = BoardTableDocument.Create(board, board).FindSheet(BoardWorkbookSchema.SheetKiCadImportantSignals)!;

        Assert.All(sheet.Rows, row => Assert.Equal(BoardTableRowState.Unchanged, row.State));
        Assert.Equal(0, sheet.FlaggedCount);
        Assert.False(sheet.HasChangeRows);

        BoardTableSheet doubled = BoardTableDocument.Create(
                board,
                new BoardData { KiCadImportantSignals = [.. board.KiCadImportantSignals, new() { DisplayName = "9VAC", KiCadNetName = "9VAC~" }] })
            .FindSheet(BoardWorkbookSchema.SheetKiCadImportantSignals)!;

        Assert.Equal(BoardTableRowState.Duplicate, doubled.Rows.Last().State);
    }

    [Fact]
    public void An_important_signal_missing_half_is_flagged_incomplete_and_its_published_row_shows_as_deleted()
    {
        // Saving drops such a row (the schema's own mapper does), so the published signal it came
        // from really will be gone - and the ghost says exactly that.
        var published = new BoardData { KiCadImportantSignals = [new() { DisplayName = "CLK", KiCadNetName = "Net-1" }] };
        BoardTableDocument document = BoardTableDocument.Create(published, published);

        BoardTableSheet sheet = document.FindSheet(BoardWorkbookSchema.SheetKiCadImportantSignals)!;
        sheet.Rows.Single().Cells[Column(sheet, BoardWorkbookSchema.ColKiCadNetName)].Text = string.Empty;
        sheet.Refresh();

        BoardTableRow incomplete = sheet.Rows.Single(row => !row.IsDeleted);
        Assert.Equal(BoardTableRowState.Incomplete, incomplete.State);
        Assert.Equal("!", incomplete.Marker);
        Assert.All(incomplete.Cells, cell => Assert.Equal(BoardTableCellState.Flagged, cell.State));

        Assert.Single(sheet.Rows, row => row.IsDeleted);
        Assert.Equal(1, sheet.ChangeCount);
        Assert.Equal(1, sheet.FlaggedCount);
        Assert.Equal(1, sheet.DeletedCount);
    }

    [Fact]
    public void A_flagged_rows_cells_say_why_when_hovered_and_follow_the_reason_as_it_changes()
    {
        // The reason is under the pointer wherever it lands on the violet row - not only on the
        // narrow "!" column at its start.
        var published = new BoardData { KiCadImportantSignals = [new() { DisplayName = "CLK", KiCadNetName = "Net-1" }] };
        BoardTableDocument document = BoardTableDocument.Create(
            published,
            new BoardData
            {
                KiCadImportantSignals =
                [
                    new() { DisplayName = "CLK", KiCadNetName = "Net-1" },
                    new() { DisplayName = "CLK", KiCadNetName = "Net-1" },
                ],
            });

        BoardTableSheet sheet = document.FindSheet(BoardWorkbookSchema.SheetKiCadImportantSignals)!;
        BoardTableRow second = sheet.Rows[1];
        BoardTableCell cell = second.Cells[0];

        Assert.Equal(BoardTableRowState.Duplicate, second.State);
        Assert.Equal(second.MarkerToolTip, cell.ToolTip);

        var raised = new List<string?>();
        cell.PropertyChanged += (_, e) => raised.Add(e.PropertyName);

        // Now missing a half: still flagged, for a different reason - and the cell says so.
        second.Cells[Column(sheet, BoardWorkbookSchema.ColKiCadNetName)].Text = string.Empty;
        sheet.Refresh();

        Assert.Equal(BoardTableRowState.Incomplete, second.State);
        Assert.Equal(BoardTableCellState.Flagged, cell.State);
        Assert.Contains("Incomplete", cell.ToolTip);
        Assert.Contains(nameof(BoardTableCell.ToolTip), raised);
    }

    [Fact]
    public void A_flagged_row_that_is_put_right_loses_its_flag()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU")),
            Board(Component("U1", "CPU"), Component("U1", "Another CPU")));
        BoardTableSheet sheet = Components(document);
        BoardTableRow duplicate = sheet.Rows[1];

        Cell(duplicate, BoardWorkbookSchema.ColBoardLabel).Text = "U2";
        sheet.Refresh();

        Assert.Equal(BoardTableRowState.Added, duplicate.State);
        Assert.All(duplicate.Cells, cell => Assert.Equal(BoardTableCellState.Added, cell.State));
        Assert.Equal(0, sheet.FlaggedCount);
        Assert.Equal(1, sheet.AddedCount);
    }

    // ------------------------------------------------------------------ Editing

    [Fact]
    public void A_cell_edit_marks_the_table_unsaved_and_asks_for_a_refresh_before_recolouring()
    {
        // Recolouring is deferred to Refresh on purpose - see BoardTableCell.Text.
        BoardTableDocument document = BoardTableDocument.Create(Board(Component("U1", "CPU")), Board(Component("U1", "CPU")));
        BoardTableSheet sheet = Components(document);
        BoardTableCell? edited = null;
        sheet.CellEdited += (_, cell) => edited = cell;

        BoardTableCell cell = Cell(sheet.Rows.Single(), BoardWorkbookSchema.ColFriendlyName);
        cell.Text = "CPU 6510";

        Assert.Same(cell, edited);
        Assert.True(document.HasUnsavedChanges);
        Assert.True(sheet.NeedsRefresh);
        Assert.Equal(BoardTableCellState.Unchanged, cell.State);

        sheet.Refresh();

        Assert.False(sheet.NeedsRefresh);
        Assert.Equal(BoardTableCellState.Modified, cell.State);
    }

    [Fact]
    public void An_inserted_row_lands_directly_after_the_given_row()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board(Component("U1"), Component("U2")));
        BoardTableSheet sheet = Components(document);

        BoardTableRow inserted = sheet.InsertRow(sheet.Rows[0]);

        Assert.Same(inserted, sheet.Rows[1]);
        Assert.True(document.HasUnsavedChanges);
    }

    [Fact]
    public void An_inserted_row_with_no_anchor_goes_at_the_end()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board(Component("U1")));
        BoardTableSheet sheet = Components(document);

        BoardTableRow inserted = sheet.InsertRow(null);

        Assert.Same(inserted, sheet.Rows[^1]);
    }

    [Fact]
    public void A_blank_inserted_row_is_green_but_neither_counted_nor_saved()
    {
        BoardTableDocument document = BoardTableDocument.Create(Board(Component("U1")), Board(Component("U1")));
        BoardTableSheet sheet = Components(document);

        BoardTableRow blank = sheet.InsertRow(null);

        Assert.Equal(BoardTableRowState.Blank, blank.State);
        Assert.Equal("+", blank.Marker);
        Assert.All(blank.Cells, cell => Assert.Equal(BoardTableCellState.Added, cell.State));
        Assert.Equal(0, sheet.ChangeCount);
        Assert.Single(document.ApplyTo(new BoardData()).Components);
    }

    [Fact]
    public void Typing_into_an_inserted_row_makes_it_an_added_row()
    {
        BoardTableDocument document = BoardTableDocument.Create(Board(Component("U1")), Board(Component("U1")));
        BoardTableSheet sheet = Components(document);

        BoardTableRow row = sheet.InsertRow(null);
        Cell(row, BoardWorkbookSchema.ColBoardLabel).Text = "C12";
        sheet.Refresh();

        Assert.Equal(BoardTableRowState.Added, row.State);
        Assert.Equal(1, sheet.ChangeCount);
    }

    [Fact]
    public void Deleting_a_published_row_turns_it_into_a_ghost_in_the_same_place()
    {
        BoardData board = Board(Component("U1"), Component("U2"), Component("U3"));
        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet sheet = Components(document);

        Assert.True(sheet.DeleteRow(Row(sheet, "U2")));

        Assert.Equal(["U1", "U2", "U3"], sheet.Rows.Select(Label));
        Assert.True(sheet.Rows[1].IsDeleted);
        Assert.Equal(1, sheet.ChangeCount);
        Assert.True(document.HasUnsavedChanges);
    }

    [Fact]
    public void Deleting_an_added_row_just_removes_it()
    {
        BoardTableDocument document = BoardTableDocument.Create(Board(Component("U1")), Board(Component("U1"), Component("U2")));
        BoardTableSheet sheet = Components(document);

        Assert.True(sheet.DeleteRow(Row(sheet, "U2")));

        Assert.Equal(["U1"], sheet.Rows.Select(Label));
        Assert.Equal(0, sheet.ChangeCount);
    }

    [Fact]
    public void Deleting_a_ghost_is_refused()
    {
        BoardTableDocument document = BoardTableDocument.Create(Board(Component("U1")), Board());
        BoardTableSheet sheet = Components(document);

        Assert.False(sheet.DeleteRow(sheet.Rows.Single()));
        Assert.Single(sheet.Rows);
    }

    [Fact]
    public void Restoring_a_ghost_brings_the_published_row_back_unchanged_in_its_own_position()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1"), Component("U2", "VIC"), Component("U3")),
            Board(Component("U1"), Component("U3")));
        BoardTableSheet sheet = Components(document);

        BoardTableRow? restored = sheet.RestoreRow(sheet.Rows[1]);

        Assert.NotNull(restored);
        Assert.Same(restored, sheet.Rows[1]);
        Assert.False(restored!.IsDeleted);
        Assert.Equal(BoardTableRowState.Unchanged, restored.State);
        Assert.Equal("VIC", Cell(restored, BoardWorkbookSchema.ColFriendlyName).Text);
        Assert.Equal(0, sheet.ChangeCount);
    }

    [Fact]
    public void Restoring_a_live_row_is_refused()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board(Component("U1")));
        BoardTableSheet sheet = Components(document);

        Assert.Null(sheet.RestoreRow(sheet.Rows.Single()));
    }

    [Fact]
    public void Reverting_a_changed_cell_puts_the_published_value_back()
    {
        BoardTableDocument document = BoardTableDocument.Create(Board(Component("U1", "CPU")), Board(Component("U1", "CPU 6510")));
        BoardTableSheet sheet = Components(document);
        BoardTableCell cell = Cell(sheet.Rows.Single(), BoardWorkbookSchema.ColFriendlyName);

        Assert.True(sheet.RevertCell(cell));

        Assert.Equal("CPU", cell.Text);
        Assert.Equal(BoardTableCellState.Unchanged, cell.State);
        Assert.Equal(0, sheet.ChangeCount);
    }

    [Fact]
    public void Reverting_is_refused_for_an_unchanged_cell_a_cell_on_an_added_row_and_a_ghost()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU"), Component("U3")),
            Board(Component("U1", "CPU"), Component("U2")));
        BoardTableSheet sheet = Components(document);

        Assert.False(sheet.RevertCell(Cell(Row(sheet, "U1"), BoardWorkbookSchema.ColFriendlyName)));
        Assert.False(sheet.RevertCell(Cell(Row(sheet, "U2"), BoardWorkbookSchema.ColFriendlyName)));
        Assert.False(sheet.RevertCell(Cell(Row(sheet, "U3", deleted: true), BoardWorkbookSchema.ColFriendlyName)));
    }

    [Fact]
    public void A_refresh_keeps_an_existing_ghost_as_the_same_object()
    {
        // A grid watching Rows would otherwise see every deleted row vanish and reappear on every
        // keystroke elsewhere in the sheet, losing its place.
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1"), Component("U2")),
            Board(Component("U1")));
        BoardTableSheet sheet = Components(document);
        BoardTableRow ghost = sheet.Rows[1];

        Cell(Row(sheet, "U1"), BoardWorkbookSchema.ColFriendlyName).Text = "Edited";
        sheet.Refresh();

        Assert.Same(ghost, sheet.Rows[1]);
    }

    // ------------------------------------------------------------------ No published board

    [Fact]
    public void With_no_published_board_nothing_is_coloured_but_duplicates_are_still_flagged()
    {
        // "Add a new system": every row would be green, which tells the contributor nothing. A
        // duplicate is still a duplicate, though, and is flagged violet all the same.
        BoardTableDocument document = BoardTableDocument.Create(
            null,
            Board(Component("U1"), Component("U2"), Component("U2")));
        BoardTableSheet sheet = Components(document);

        Assert.False(document.HasBaseline);
        Assert.Equal(BoardTableRowState.Unchanged, sheet.Rows[0].State);
        Assert.All(sheet.Rows[0].Cells, cell => Assert.Equal(BoardTableCellState.Unchanged, cell.State));
        Assert.Equal(BoardTableRowState.Duplicate, sheet.Rows[2].State);
        Assert.All(sheet.Rows[2].Cells, cell => Assert.Equal(BoardTableCellState.Flagged, cell.State));
        Assert.Equal(1, sheet.FlaggedCount);
        Assert.Equal(0, document.TotalChangeCount);
    }

    // ------------------------------------------------------------------ Saving

    [Fact]
    public void ApplyTo_replaces_the_sheets_and_keeps_the_highlights_revision_and_names_from_disk()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board(Component("U1")));
        Cell(Components(document).Rows.Single(), BoardWorkbookSchema.ColFriendlyName).Text = "CPU";

        var onDisk = new BoardData
        {
            RevisionDate = "2026-09-01",
            HardwareName = "Commodore 64",
            BoardName = "250407",
            Components = [Component("Stale")],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Main", BoardLabel = "U1" }],
        };

        BoardData saved = document.ApplyTo(onDisk);

        ComponentEntry component = Assert.Single(saved.Components);
        Assert.Equal("U1", component.BoardLabel);
        Assert.Equal("CPU", component.FriendlyName);

        Assert.Same(onDisk.ComponentHighlights, saved.ComponentHighlights);
        Assert.Equal("2026-09-01", saved.RevisionDate);
        Assert.Equal("Commodore 64", saved.HardwareName);
        Assert.Equal("250407", saved.BoardName);
    }

    [Fact]
    public void Saving_a_ghost_does_not_bring_the_deleted_row_back()
    {
        BoardTableDocument document = BoardTableDocument.Create(Board(Component("U1"), Component("U2")), Board(Component("U1")));

        Assert.Equal(["U1"], document.ApplyTo(new BoardData()).Components.Select(c => c.BoardLabel));
    }

    [Fact]
    public void MarkSaved_clears_the_unsaved_flag_and_says_so()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board(Component("U1")));
        Components(document).InsertRow(null);
        int changedEvents = 0;
        document.Changed += (_, _) => changedEvents++;

        document.MarkSaved();

        Assert.False(document.HasUnsavedChanges);
        Assert.Equal(1, changedEvents);
    }

    // ------------------------------------------------------------------ Agreement with the differ

    [Fact]
    public void The_change_counts_agree_with_BoardDataDiffer_on_the_board_a_save_writes()
    {
        BoardData published = new()
        {
            Schematics = [new() { SchematicName = "Main", SchematicImageFile = "main.png" }, new() { SchematicName = "Rear", SchematicImageFile = "rear.png" }],
            Components = [Component("U1", "CPU"), Component("U2", "VIC"), Component("U3", "SID"), Component("U4", "CIA")],
            Credits = [new() { Category = "Data", NameOrHandle = "Dennis" }],
            KiCadImportantSignals = [new() { DisplayName = "CLK", KiCadNetName = "Net-1" }, new() { DisplayName = "RESET", KiCadNetName = "Net-2" }],
        };

        BoardData draft = new()
        {
            Schematics = [new() { SchematicName = "Main", SchematicImageFile = "main-v2.png" }],
            Components = [Component("u1", "CPU"), Component("U3", "SID 8580"), Component("U3", "Duplicate SID"), Component("U9", "New")],
            Credits = [new() { Category = "Data", NameOrHandle = "Dennis" }, new() { Category = "Photos", NameOrHandle = "Someone" }],
            KiCadImportantSignals = [new() { DisplayName = "CLK", KiCadNetName = "Net-1b" }, new() { DisplayName = "RESET", KiCadNetName = "Net-2" }],
        };

        BoardTableDocument document = BoardTableDocument.Create(published, draft);

        // And some edits made in the table itself: an incomplete signal, a blank row, a key rename.
        BoardTableSheet signals = document.FindSheet(BoardWorkbookSchema.SheetKiCadImportantSignals)!;
        signals.Rows.Single(row => row.Cells[0].Text == "RESET").Cells[1].Text = string.Empty;
        signals.Refresh();

        Components(document).InsertRow(null);
        Cell(Row(Components(document), "U9"), BoardWorkbookSchema.ColBoardLabel).Text = "U10";
        Components(document).Refresh();

        BoardData saved = document.ApplyTo(new BoardData());
        IReadOnlyList<BoardRowChange> differ = BoardDataDiffer.Compare(published, saved);

        foreach (BoardTableSheet sheet in document.Sheets)
        {
            Assert.True(
                differ.Count(change => change.Section == sheet.Name) == sheet.ChangeCount,
                $"[{sheet.Name}] table counts {sheet.ChangeCount}, BoardDataDiffer counts {differ.Count(change => change.Section == sheet.Name)}");
        }

        Assert.Equal(differ.Count, document.TotalChangeCount);
    }

    // ------------------------------------------------------------------ Moving rows (2026-09-24)

    private static ComponentEntry InCategory(string label, string category) =>
        new() { BoardLabel = label, TechnicalNameOrValue = "x", Category = category };

    private static void Type(BoardTableRow row, string column, string text)
    {
        Cell(row, column).Text = text;
        row.Sheet.Refresh();
    }

    [Fact]
    public void Moving_a_row_down_puts_it_after_the_next_row_and_is_saved_in_that_order()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board(Component("U1"), Component("U2"), Component("U3")));
        BoardTableSheet sheet = Components(document);

        Assert.True(sheet.MoveRowDown(Row(sheet, "U1")));

        Assert.Equal(["U2", "U1", "U3"], sheet.Rows.Select(Label));
        Assert.True(document.HasUnsavedChanges);
        Assert.Equal(["U2", "U1", "U3"], document.ApplyTo(new BoardData()).Components.Select(c => c.BoardLabel));
    }

    [Fact]
    public void Moving_up_from_the_top_or_down_from_the_bottom_does_nothing()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board(Component("U1"), Component("U2")));
        BoardTableSheet sheet = Components(document);

        Assert.False(sheet.MoveRowUp(Row(sheet, "U1")));
        Assert.False(sheet.MoveRowDown(Row(sheet, "U2")));
        Assert.False(document.HasUnsavedChanges);
    }

    [Fact]
    public void A_move_is_not_a_counted_change_because_order_is_not_one()
    {
        // BoardDataDiffer pairs by key and ignores order - the table must say the same.
        BoardData board = Board(Component("U1"), Component("U2"));
        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet sheet = Components(document);

        sheet.MoveRowDown(Row(sheet, "U1"));

        Assert.Equal(0, sheet.ChangeCount);
        Assert.All(sheet.Rows, row => Assert.Equal(BoardTableRowState.Unchanged, row.State));
    }

    [Fact]
    public void Moving_skips_over_deleted_ghosts_so_every_press_moves_something_real()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1"), Component("U2"), Component("U3")),
            Board(Component("U1"), Component("U3")));
        BoardTableSheet sheet = Components(document);

        Assert.Equal(["U1", "U2", "U3"], sheet.Rows.Select(Label));

        Assert.True(sheet.MoveRowUp(Row(sheet, "U3")));

        // U3 went above U1; the ghost of U2 still follows U1, the published row before it.
        Assert.Equal(["U3", "U1", "U2"], sheet.Rows.Select(Label));
        Assert.True(sheet.Rows[2].IsDeleted);
    }

    [Fact]
    public void A_ghost_cannot_be_moved()
    {
        BoardTableDocument document = BoardTableDocument.Create(Board(Component("U1"), Component("U2")), Board(Component("U2")));
        BoardTableSheet sheet = Components(document);

        Assert.False(sheet.MoveRow(Row(sheet, "U1", deleted: true), 2));
        Assert.False(sheet.MoveRowDown(Row(sheet, "U1", deleted: true)));
    }

    [Fact]
    public void Moving_to_a_drop_position_uses_the_position_as_it_was_before_the_move()
    {
        BoardTableDocument document = BoardTableDocument.Create(null, Board(Component("A"), Component("B"), Component("C"), Component("D")));
        BoardTableSheet sheet = Components(document);

        // Drop A just before D (index 3 in the list as it stands).
        Assert.True(sheet.MoveRow(Row(sheet, "A"), 3));
        Assert.Equal(["B", "C", "A", "D"], sheet.Rows.Select(Label));

        // Drop D at the very top.
        Assert.True(sheet.MoveRow(Row(sheet, "D"), 0));
        Assert.Equal(["D", "B", "C", "A"], sheet.Rows.Select(Label));

        // Dropping a row on its own place changes nothing.
        Assert.False(sheet.MoveRow(Row(sheet, "B"), 1));
        Assert.False(sheet.MoveRow(Row(sheet, "B"), 2));
    }

    // ------------------------------------------------------------------ New components are placed on save

    [Fact]
    public void An_inserted_component_is_saved_into_its_category_in_label_order()
    {
        // The same rule as the Contribute window and the label editor (ComponentPlacement), so a
        // component lands in the same place however it was added.
        BoardTableDocument document = BoardTableDocument.Create(
            null,
            Board(InCategory("U1", "IC"), InCategory("U3", "IC"), InCategory("C1", "Capacitor")));
        BoardTableSheet sheet = Components(document);

        BoardTableRow inserted = sheet.InsertRow(null);
        Type(inserted, BoardWorkbookSchema.ColBoardLabel, "U2");
        Type(inserted, BoardWorkbookSchema.ColCategory, "IC");

        Assert.True(sheet.HasRowsToPlaceOnSave);
        Assert.Equal(["U1", "U2", "U3", "C1"], document.ApplyTo(new BoardData()).Components.Select(c => c.BoardLabel));
    }

    [Fact]
    public void An_inserted_component_the_contributor_MOVED_keeps_their_place()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            null,
            Board(InCategory("U1", "IC"), InCategory("U3", "IC")));
        BoardTableSheet sheet = Components(document);

        BoardTableRow inserted = sheet.InsertRow(null);
        Type(inserted, BoardWorkbookSchema.ColBoardLabel, "U2");
        Type(inserted, BoardWorkbookSchema.ColCategory, "IC");
        sheet.MoveRow(inserted, 0);

        Assert.False(sheet.HasRowsToPlaceOnSave);
        Assert.Equal(["U2", "U1", "U3"], document.ApplyTo(new BoardData()).Components.Select(c => c.BoardLabel));
    }

    [Fact]
    public void Rows_that_were_already_there_are_never_re_sorted_on_save()
    {
        // Only NEW rows are placed - the existing order is the project owner's.
        BoardTableDocument document = BoardTableDocument.Create(
            null,
            Board(InCategory("U9", "IC"), InCategory("C1", "Capacitor"), InCategory("U1", "IC")));

        Assert.Equal(["U9", "C1", "U1"], document.ApplyTo(new BoardData()).Components.Select(c => c.BoardLabel));
    }

    [Fact]
    public void Rows_inserted_on_OTHER_sheets_stay_where_they_were_put()
    {
        var draft = new BoardData { Credits = [new() { Category = "Data", NameOrHandle = "Z" }] };
        BoardTableDocument document = BoardTableDocument.Create(null, draft);
        BoardTableSheet credits = document.FindSheet(BoardWorkbookSchema.SheetCredits)!;

        BoardTableRow inserted = credits.InsertRow(null);
        Cell(inserted, BoardWorkbookSchema.ColCategory).Text = "Data";
        Cell(inserted, BoardWorkbookSchema.ColNameOrHandle).Text = "A";
        credits.Refresh();

        Assert.False(credits.HasRowsToPlaceOnSave);
        Assert.Equal(["Z", "A"], document.ApplyTo(new BoardData()).Credits.Select(c => c.NameOrHandle));
    }

    // ------------------------------------------------------------------ Regional rows (2026-09-24)

    [Fact]
    public void A_regional_variant_typed_in_beside_its_twin_is_GREEN_not_a_duplicate()
    {
        // The project owner's steps: insert a row under U1, give it the same label, then set the two
        // rows' regions to PAL and NTSC. The new one used to be flagged "!" and counted nowhere.
        BoardData board = new() { Components = [new ComponentEntry { BoardLabel = "U1", TechnicalNameOrValue = "x", Category = "IC" }] };
        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet sheet = Components(document);

        BoardTableRow original = sheet.Rows.Single();
        BoardTableRow added = sheet.InsertRow(original);
        Cell(added, BoardWorkbookSchema.ColBoardLabel).Text = "U1";
        Cell(added, BoardWorkbookSchema.ColRegion).Text = "NTSC";
        Cell(original, BoardWorkbookSchema.ColRegion).Text = "PAL";
        sheet.Refresh();

        Assert.Equal(BoardTableRowState.Added, added.State);
        Assert.Equal("+", added.Marker);

        // Giving the published U1 a region makes it a different row too: the region-less U1 is a
        // red ghost, and U1/PAL is added - the same as changing its label.
        Assert.Equal(BoardTableRowState.Added, original.State);
        Assert.Single(sheet.Rows, row => row.IsDeleted);
        Assert.Equal(3, sheet.ChangeCount);
    }

    [Fact]
    public void The_same_label_AND_region_twice_is_still_flagged_as_a_duplicate()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            null,
            new BoardData
            {
                Components =
                [
                    new ComponentEntry { BoardLabel = "U1", Region = "PAL", TechnicalNameOrValue = "x" },
                    new ComponentEntry { BoardLabel = "U1", Region = "pal", TechnicalNameOrValue = "x" },
                ],
            });

        Assert.Equal(BoardTableRowState.Duplicate, Components(document).Rows[1].State);
    }

    [Fact]
    public void A_regional_variant_inserted_with_NO_category_is_saved_right_beside_its_twin()
    {
        // The owner's report: the row "disappeared" after saving - it had gone to the end of
        // a long sheet, because a blank category is a category nobody else uses.
        BoardData board = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", TechnicalNameOrValue = "x", Category = "IC", Region = "PAL" },
                new ComponentEntry { BoardLabel = "U2", TechnicalNameOrValue = "x", Category = "IC" },
                new ComponentEntry { BoardLabel = "C1", TechnicalNameOrValue = "x", Category = "Capacitor" },
            ],
        };
        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet sheet = Components(document);

        BoardTableRow added = sheet.InsertRow(sheet.Rows[0]);
        Cell(added, BoardWorkbookSchema.ColBoardLabel).Text = "U1";
        Cell(added, BoardWorkbookSchema.ColRegion).Text = "NTSC";
        sheet.Refresh();

        Assert.Equal(
            ["U1/PAL", "U1/NTSC", "U2/", "C1/"],
            document.ApplyTo(new BoardData()).Components.Select(c => $"{c.BoardLabel}/{c.Region}"));
    }

    // ------------------------------------------------------------------ Legend counts (2026-09-24)

    [Fact]
    public void The_three_counts_split_the_sheets_change_count_by_kind()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU"), Component("U2"), Component("U3"), Component("U4")),
            Board(Component("U1", "CPU 6510"), Component("U2"), Component("U9"), Component("U10")));
        BoardTableSheet sheet = Components(document);

        Assert.Equal(2, sheet.AddedCount);
        Assert.Equal(1, sheet.ModifiedCount);
        Assert.Equal(2, sheet.DeletedCount);
        Assert.Equal(sheet.ChangeCount, sheet.AddedCount + sheet.ModifiedCount + sheet.DeletedCount);
    }

    [Fact]
    public void Only_rows_that_are_not_plainly_unchanged_count_as_change_rows()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1"), Component("U2")),
            Board(Component("U1"), Component("U3")));
        BoardTableSheet sheet = Components(document);
        BoardTableRow blank = sheet.InsertRow(null);

        Assert.False(BoardTableSheet.IsChangeRow(Row(sheet, "U1")));
        Assert.True(BoardTableSheet.IsChangeRow(Row(sheet, "U3")));
        Assert.True(BoardTableSheet.IsChangeRow(Row(sheet, "U2", deleted: true)));
        Assert.True(BoardTableSheet.IsChangeRow(blank));
    }

    // ------------------------------------------------------------------ A deleted component

    // ###########################################################################################
    // *** A DELETED COMPONENT TAKES EVERYTHING OF ITS OWN WITH IT (owner request, 2026-09-25):
    // "if a component really is deleted, then it should remove EVERYTHING related to this
    // component." *** Its rows on the image, local file and link sheets are deleted at once (red on
    // their own sheets), its highlights go when the table is saved, and one undo brings it all back.
    // It used to leave all of that behind - rows pointing at a component the board no longer had.
    // ###########################################################################################
    private static BoardData BoardWithU8AndU9()
    {
        var board = new BoardData
        {
            Components = [Component("U8"), Component("U9")],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U8", Name = "Clock", File = "a/u8.png" },
                new ComponentImageEntry { BoardLabel = "U9", Name = "Clock", File = "a/u9.png" }
            ],
            ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "a/u8.pdf" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U8", Name = "Ref", Url = "https://example.org" }],
            ComponentHighlights =
            [
                new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "1", Y = "1", Width = "5", Height = "5" },
                new ComponentHighlightEntry { SchematicName = "Sheet 2", BoardLabel = "u8", X = "1", Y = "1", Width = "5", Height = "5" },
                new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U9", X = "9", Y = "9", Width = "5", Height = "5" }
            ]
        };

        return board;
    }

    private static List<string> LiveLabels(BoardTableDocument document, string sheet) =>
        document.FindSheet(sheet)!.Rows.Where(row => !row.IsDeleted).Select(Label).ToList();

    [Fact]
    public void Deleting_a_component_deletes_its_rows_on_the_other_sheets_and_its_highlights_on_save()
    {
        BoardData board = BoardWithU8AndU9();
        BoardTableDocument document = BoardTableDocument.Create(board, board);

        Assert.True(Components(document).DeleteRow(Row(Components(document), "U8"), out BoardTableDeletedWith? deletedWith));

        // Shown as deleted on their own sheets, where the contributor or maintainer can see them.
        Assert.Equal(["U9"], LiveLabels(document, BoardWorkbookSchema.SheetComponentImages));
        Assert.True(Row(document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!, "U8", deleted: true).IsDeleted);
        Assert.Empty(LiveLabels(document, BoardWorkbookSchema.SheetComponentLocalFiles));
        Assert.Empty(LiveLabels(document, BoardWorkbookSchema.SheetComponentLinks));

        Assert.NotNull(deletedWith);
        Assert.Equal(
            [BoardWorkbookSchema.SheetComponentImages, BoardWorkbookSchema.SheetComponentLocalFiles, BoardWorkbookSchema.SheetComponentLinks],
            deletedWith!.Rows.Select(count => count.Sheet));
        Assert.Equal(2, deletedWith.Highlights);

        BoardData saved = document.ApplyTo(board);

        Assert.Equal("U9", Assert.Single(saved.Components).BoardLabel);
        Assert.Equal("U9", Assert.Single(saved.ComponentImages).BoardLabel);
        Assert.Empty(saved.ComponentLocalFiles);
        Assert.Empty(saved.ComponentLinks);
        Assert.Equal("U9", Assert.Single(saved.ComponentHighlights).BoardLabel);
    }

    // One Ctrl+Z undoes the component AND what went with it - on every sheet.
    [Fact]
    public void One_undo_brings_back_the_component_and_everything_it_took()
    {
        BoardData board = BoardWithU8AndU9();
        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableRow u8 = Row(Components(document), "U8");

        Components(document).DeleteRow(u8);

        BoardTableHistoryResult? undone = document.History.Undo();

        Assert.Same(Components(document), undone!.Sheet);
        Assert.Same(u8, undone.Row);
        Assert.Equal(0, document.TotalChangeCount);
        Assert.Equal(["U8", "U9"], LiveLabels(document, BoardWorkbookSchema.SheetComponentImages));
        Assert.Equal(["U8"], LiveLabels(document, BoardWorkbookSchema.SheetComponentLinks));
        Assert.Equal(3, document.ApplyTo(board).ComponentHighlights.Count);

        // And redo takes it all again.
        document.History.Redo();

        Assert.Empty(LiveLabels(document, BoardWorkbookSchema.SheetComponentLinks));
        Assert.Single(document.ApplyTo(board).ComponentHighlights);
    }

    // A component RENAMED in the table is the same component: it keeps its highlights. Only the
    // table can tell the two apart - the row object is still there.
    [Fact]
    public void A_renamed_component_keeps_its_highlights_and_its_rows_elsewhere()
    {
        BoardData board = BoardWithU8AndU9();
        BoardTableDocument document = BoardTableDocument.Create(board, board);

        Cell(Row(Components(document), "U8"), BoardWorkbookSchema.ColBoardLabel).Text = "U10";
        Components(document).Refresh();

        BoardData saved = document.ApplyTo(board);

        Assert.Equal(3, saved.ComponentHighlights.Count);
        Assert.Equal(2, saved.ComponentImages.Count);
    }

    // A regional variant deleted beside its twin takes only ITS region's images: the files, the
    // links, the region-less images and the highlights still belong to the twin.
    [Fact]
    public void Deleting_one_regional_variant_takes_only_that_regions_images()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U8", Region = "PAL", FriendlyName = "VIC", TechnicalNameOrValue = "6569" },
                new ComponentEntry { BoardLabel = "U8", Region = "NTSC", FriendlyName = "VIC", TechnicalNameOrValue = "6567" }
            ],
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U8", Region = "PAL", Name = "Clock", File = "a/pal.png" },
                new ComponentImageEntry { BoardLabel = "U8", Region = "NTSC", Name = "Clock", File = "a/ntsc.png" },
                new ComponentImageEntry { BoardLabel = "U8", Name = "Pinout", File = "a/pinout.png" }
            ],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U8", Name = "Ref", Url = "https://example.org" }],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "1", Y = "1", Width = "5", Height = "5" }]
        };

        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet components = Components(document);
        BoardTableRow pal = components.Rows.Single(row => components.CellText(row, BoardWorkbookSchema.ColRegion) == "PAL");

        components.DeleteRow(pal, out BoardTableDeletedWith? deletedWith);

        BoardData saved = document.ApplyTo(board);

        Assert.Equal(["a/ntsc.png", "a/pinout.png"], saved.ComponentImages.Select(image => image.File));
        Assert.Single(saved.ComponentLinks);
        Assert.Single(saved.ComponentHighlights);
        Assert.Equal(0, deletedWith!.Highlights);
    }

    // Deleting a row that is only a duplicate of another with the same label takes nothing else.
    [Fact]
    public void Deleting_a_duplicate_row_takes_nothing_else()
    {
        BoardData board = BoardWithU8AndU9();
        BoardData draft = BoardWithU8AndU9();
        draft.Components.Add(Component("U8", "second"));

        BoardTableDocument document = BoardTableDocument.Create(board, draft);
        BoardTableRow duplicate = Components(document).Rows.Last(row => Label(row) == "U8");

        Components(document).DeleteRow(duplicate, out BoardTableDeletedWith? deletedWith);

        Assert.Null(deletedWith);
        Assert.Equal(["U8", "U9"], LiveLabels(document, BoardWorkbookSchema.SheetComponentImages));
        Assert.Equal(3, document.ApplyTo(draft).ComponentHighlights.Count);
    }

    // ------------------------------------------------------------------ Which sheet tabs show (2026-09-26)

    // A board whose Components sheet has a change and whose Credits sheet has none.
    private static BoardTableDocument ComponentsChanged() =>
        BoardTableDocument.Create(Board(Component("U1", "CPU")), Board(Component("U1", "CPU 6510")));

    // ###########################################################################################
    // "Show changes only" hides the sheets it would show nothing of (owner request, 2026-09-26) -
    // and without it every sheet shows.
    // ###########################################################################################
    [Fact]
    public void With_show_changes_only_only_the_sheets_with_something_to_show_have_tabs()
    {
        BoardTableDocument document = ComponentsChanged();

        Assert.Equal(document.Sheets, document.SheetsShown(onlyChanges: false, current: null));
        Assert.Equal([BoardWorkbookSchema.SheetComponents], document.SheetsShown(onlyChanges: true, current: null).Select(sheet => sheet.Name));
    }

    // A flagged row (a duplicate, say) is not a change but IS shown by the filter - so its sheet keeps
    // its tab, or the one thing needing a second look would be out of reach.
    [Fact]
    public void A_sheet_with_only_a_flagged_row_keeps_its_tab()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1")),
            Board(Component("U1"), Component("U1")));

        Assert.Equal(0, Components(document).ChangeCount);
        Assert.True(Components(document).HasChangeRows);
        Assert.Contains(Components(document), document.SheetsShown(onlyChanges: true, current: null));
    }

    // The sheet being worked on keeps its tab when its last change is undone - it is not pulled
    // away mid-work; it goes once another sheet is chosen.
    [Fact]
    public void The_current_sheet_keeps_its_tab_while_it_is_on_screen()
    {
        BoardTableDocument document = ComponentsChanged();
        BoardTableSheet credits = document.FindSheet(BoardWorkbookSchema.SheetCredits)!;

        Assert.Contains(credits, document.SheetsShown(onlyChanges: true, current: credits));
        Assert.DoesNotContain(credits, document.SheetsShown(onlyChanges: true, current: null));
    }

    // Nothing to show anywhere: every tab stays, rather than all but one vanishing.
    [Fact]
    public void A_table_with_nothing_to_show_keeps_every_tab()
    {
        BoardTableDocument document = BoardTableDocument.Create(Board(Component("U1")), Board(Component("U1")));

        Assert.Equal(document.Sheets, document.SheetsShown(onlyChanges: true, current: null));
    }

    // ###########################################################################################
    // The sheet a table opens on: the one wanted (the sheet last looked at) while its tab shows,
    // else the first that shows.
    // ###########################################################################################
    [Fact]
    public void A_table_opens_on_the_wanted_sheet_while_its_tab_shows()
    {
        BoardTableDocument document = ComponentsChanged();

        Assert.Equal(BoardWorkbookSchema.SheetCredits, document.SheetToShow(BoardWorkbookSchema.SheetCredits, onlyChanges: false).Name);
        Assert.Equal(BoardWorkbookSchema.SheetComponents, document.SheetToShow(BoardWorkbookSchema.SheetCredits, onlyChanges: true).Name);
        Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, document.SheetToShow(null, onlyChanges: false).Name);
        Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, document.SheetToShow("No such sheet", onlyChanges: false).Name);
    }
}
