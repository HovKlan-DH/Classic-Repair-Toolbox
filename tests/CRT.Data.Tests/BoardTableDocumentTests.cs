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
    // Each component's technical value is its own, as real components' are. With one shared value,
    // every component deleted and every one added were identical but for their label - which reads
    // as a component relabelled (BoardDataDiffer.PairRenamedRows, 2026-10-04), not as the deletion
    // and addition those tests set up. A label edited in the table keeps the old value, so it does.
    private static ComponentEntry Component(string label, string friendlyName = "", string description = "") =>
        new() { BoardLabel = label, FriendlyName = friendlyName, TechnicalNameOrValue = $"part {label}", Description = description };

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
    // maintainer's NEW-BOARD table compares the submission with ITSELF as it arrived, so its
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
        // The maintainer's new board: the submission is its own baseline, and the maintainer has
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

    // ###########################################################################################
    // *** A TABLE COMPARED WITH A DATA SOURCE NAMES IT (owner request, 2026-10-05). *** A value the
    // BETA source already held, shown as "Published value", was first taken for an error. The two
    // names are the sources' own, as the rest of CRT writes them ("the BETA source", "the stable
    // source"), capitalised to start the tooltip.
    // ###########################################################################################
    [Theory]
    [InlineData(true, "BETA source value: CPU")]
    [InlineData(false, "Stable source value: CPU")]
    public void A_changed_cell_names_the_data_source_it_is_compared_with(bool betaSource, string expected)
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU")),
            Board(Component("U1", "Processor")),
            BoardTableDocument.SourceBaselineLabel(betaSource));

        Assert.Equal(expected, Cell(Row(Components(document), "U1"), BoardWorkbookSchema.ColFriendlyName).ToolTip);
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

    // ###########################################################################################
    // *** A KEY CELL CHANGED AND NOTHING ELSE IS ONE ROW MODIFIED (owner decision, 2026-10-04). ***
    // It was an addition plus a deleted ghost until then - changing a credit's "Name or handle" "did
    // remove one row and added a new row ... This seems weird to me". The changed key cell is orange
    // with the value it replaced, and nothing is red. BoardDataDiffer counts it the same way (its
    // tests, and the agreement test below).
    // ###########################################################################################
    [Fact]
    public void Changing_only_a_key_cell_reads_as_ONE_row_modified_with_the_old_value_in_the_cell()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U8", "CPU")),
            Board(Component("U8", "CPU")));

        BoardTableSheet sheet = Components(document);
        Cell(Row(sheet, "U8"), BoardWorkbookSchema.ColBoardLabel).Text = "U9";
        sheet.Refresh();

        BoardTableRow row = Row(sheet, "U9");
        BoardTableCell label = Cell(row, BoardWorkbookSchema.ColBoardLabel);

        Assert.Equal(BoardTableRowState.Modified, row.State);
        Assert.Equal(BoardTableCellState.Modified, label.State);
        Assert.Equal("U8", label.PublishedText);
        Assert.Equal(BoardTableCellState.Unchanged, Cell(row, BoardWorkbookSchema.ColFriendlyName).State);
        Assert.DoesNotContain(sheet.Rows, candidate => candidate.IsDeleted);
        Assert.Equal(1, sheet.ChangeCount);
        Assert.Equal((0, 1, 0), (sheet.AddedCount, sheet.ModifiedCount, sheet.DeletedCount));
    }

    // The other half of the owner's choice: a key cell AND another cell changed is no longer
    // recognisably the same row - an addition plus a deleted ghost, as before.
    [Fact]
    public void Changing_a_key_cell_AND_another_cell_still_reads_as_an_addition_plus_a_deleted_ghost()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U8", "CPU")),
            Board(Component("U8", "CPU")));

        BoardTableSheet sheet = Components(document);
        BoardTableRow row = Row(sheet, "U8");
        Cell(row, BoardWorkbookSchema.ColBoardLabel).Text = "U9";
        Cell(row, BoardWorkbookSchema.ColFriendlyName).Text = "VIC";
        sheet.Refresh();

        Assert.Equal(BoardTableRowState.Added, Row(sheet, "U9").State);
        Assert.Equal(BoardTableRowState.Deleted, Row(sheet, "U8", deleted: true).State);
        Assert.Equal(2, sheet.ChangeCount);
    }

    // The owner's own case: a credit's name given " 2". The name is part of a credit's key, beside
    // its category - one item can credit several people.
    [Fact]
    public void A_credits_name_changed_is_one_row_modified_and_agrees_with_BoardDataDiffer()
    {
        BoardData published = new()
        {
            Credits =
            [
                new CreditEntry { Category = "Board labelling", NameOrHandle = "Dennis", Contact = "dennis@example.org" },
                new CreditEntry { Category = "Board data", NameOrHandle = "Dennis", Contact = "dennis@example.org" }
            ]
        };

        BoardTableDocument document = BoardTableDocument.Create(published, published);
        BoardTableSheet credits = document.FindSheet(BoardWorkbookSchema.SheetCredits)!;
        BoardTableRow first = credits.Rows[0];

        first.Cells[credits.Columns.ToList().IndexOf(BoardWorkbookSchema.ColNameOrHandle)].Text = "Dennis 2";
        credits.Refresh();

        Assert.Equal(BoardTableRowState.Modified, first.State);
        Assert.Equal(BoardTableRowState.Unchanged, credits.Rows[1].State);
        Assert.DoesNotContain(credits.Rows, row => row.IsDeleted);
        Assert.Equal(1, credits.ChangeCount);

        BoardData saved = document.ApplyTo(new BoardData());
        Assert.Equal(credits.ChangeCount, BoardDataDiffer.CountChanges(published, saved));
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
    public void A_second_row_with_the_same_key_is_NOT_counted_and_is_the_servers_error()
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
        Assert.Equal(0, sheet.ChangeCount);

        // Not coloured and not marked since 2026-10-03 (it was violet, "!", "Flagged"): on
        // Components a duplicate is the server's error, so the error's corner mark says it.
        Assert.Equal(string.Empty, duplicate.Marker);
        Assert.Null(duplicate.MarkerToolTip);
        Assert.All(duplicate.Cells, cell => Assert.Equal(BoardTableCellState.Unchanged, cell.State));
        Assert.True(duplicate.HasErrors);

        // And the first row carries the error too, so the two read as one problem (owner
        // decision, 2026-10-03: "it should show all rows, and not only last") - while still the
        // row the differ pairs, so it stays Unchanged and counts nothing.
        Assert.True(sheet.Rows[0].HasErrors);
    }

    // ###########################################################################################
    // *** A BOARD COMPARED WITH ITSELF MARKS NOTHING - its duplicates included (2026-09-26). *** The
    // maintainer's table compares a NEW board's submission with itself as it was opened, so only
    // the maintainer's own edits are coloured. A duplicate in it is still told - it is one, the
    // server's error on Components - but its twin on the compared side must not come back as a
    // red "deleted" ghost.
    // ###########################################################################################
    [Fact]
    public void A_board_compared_with_itself_marks_nothing_but_its_duplicates()
    {
        BoardData board = Board(Component("U1", "CPU"), Component("U1", "CPU again"), Component("U2", "VIC"));

        BoardTableSheet sheet = Components(BoardTableDocument.Create(board, Board(Component("U1", "CPU"), Component("U1", "CPU again"), Component("U2", "VIC"))));

        Assert.DoesNotContain(sheet.Rows, row => row.IsDeleted);
        Assert.Equal(0, sheet.ChangeCount);
        Assert.Equal(2, sheet.ErrorRowCount);   // both U1 rows (every row of a duplicate since 2026-10-03)
        Assert.Equal(BoardTableRowState.Duplicate, sheet.Rows[1].State);
    }

    // ###########################################################################################
    // *** ONE DISPLAY NAME OVER SEVERAL NETS IS SEVERAL ROWS, NOT DUPLICATES (2026-09-26). ***
    // "9VAC" maps to both the 9VAC and the 9VAC~ net - the sheet exists to say exactly that. Keyed
    // on the name alone, every second one was flagged (reported from the maintainer's table: "the
    // uniqueness here is both columns"). Both columns the same is still a duplicate.
    // ###########################################################################################
    [Fact]
    public void Important_signals_sharing_a_display_name_but_not_a_net_are_not_duplicates()
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
        Assert.Equal(0, sheet.WarningRowCount);
        Assert.False(sheet.HasRowsShownBy(BoardTableRowFilter.Changes));

        BoardTableSheet doubled = BoardTableDocument.Create(
                board,
                new BoardData { KiCadImportantSignals = [.. board.KiCadImportantSignals, new() { DisplayName = "9VAC", KiCadNetName = "9VAC~" }] })
            .FindSheet(BoardWorkbookSchema.SheetKiCadImportantSignals)!;

        Assert.Equal(BoardTableRowState.Duplicate, doubled.Rows.Last().State);
    }

    [Fact]
    public void An_important_signal_missing_half_is_incomplete_and_its_published_row_shows_as_deleted()
    {
        // Saving drops such a row (the schema's own mapper does), so the published signal it came
        // from really will be gone - and the ghost says exactly that. The row itself looks like a
        // new row still being filled in, with the warning saying why it is not saved (2026-10-03;
        // it was violet and "!" until then).
        var published = new BoardData { KiCadImportantSignals = [new() { DisplayName = "CLK", KiCadNetName = "Net-1" }] };
        BoardTableDocument document = BoardTableDocument.Create(published, published);

        BoardTableSheet sheet = document.FindSheet(BoardWorkbookSchema.SheetKiCadImportantSignals)!;
        sheet.Rows.Single().Cells[Column(sheet, BoardWorkbookSchema.ColKiCadNetName)].Text = string.Empty;
        sheet.Refresh();

        BoardTableRow incomplete = sheet.Rows.Single(row => !row.IsDeleted);
        Assert.Equal(BoardTableRowState.Incomplete, incomplete.State);
        Assert.Equal("+", incomplete.Marker);
        Assert.All(incomplete.Cells, cell => Assert.Equal(BoardTableCellState.Added, cell.State));

        Assert.Single(sheet.Rows, row => row.IsDeleted);
        Assert.Equal(1, sheet.ChangeCount);
        Assert.Equal(1, sheet.WarningRowCount);
        Assert.Equal(1, sheet.DeletedCount);
    }

    [Fact]
    public void A_rows_warning_follows_its_reason_as_it_changes()
    {
        // A duplicate's warning sits on the row's first identity column; once the row is missing a
        // half instead, the warning moves to the empty cell - and the first cell is told its
        // tooltip changed.
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
        BoardTableCell cell = second.Cells[Column(sheet, BoardWorkbookSchema.ColDisplayName)];
        BoardTableCell net = second.Cells[Column(sheet, BoardWorkbookSchema.ColKiCadNetName)];

        Assert.Equal(BoardTableRowState.Duplicate, second.State);
        Assert.StartsWith("Warning: Another row on this sheet has the same Display name and KiCad net name.", cell.ToolTip, StringComparison.Ordinal);

        var raised = new List<string?>();
        cell.PropertyChanged += (_, e) => raised.Add(e.PropertyName);

        // Now missing a half: still warned about, for a different reason, on the empty cell.
        net.Text = string.Empty;
        sheet.Refresh();

        Assert.Equal(BoardTableRowState.Incomplete, second.State);
        Assert.Null(cell.ToolTip);
        Assert.Equal("Warning: This row is left out when saving, because KiCad net name is empty.", net.ToolTip);
        Assert.Contains(nameof(BoardTableCell.ToolTip), raised);
    }

    [Fact]
    public void A_duplicate_row_that_is_put_right_loses_its_error()
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
        Assert.Equal(0, sheet.ErrorRowCount);
        Assert.Equal(1, sheet.AddedCount);
    }

    // ------------------------------------------------------------------ Both rows of a duplicate (2026-10-02)

    // ###########################################################################################
    // *** EVERY ROW OF A DUPLICATE IS MARKED, NOT ONLY THE LATER ONES (owner request, 2026-10-02:
    // "if there is a duplicate here, then I think it should show both rows, and not only the last
    // one"; cases agreed with the project owner). *** Only the second R9 / Pinout was violet, so
    // the row it collided with had to be hunted for. Now the first is marked too - while which row
    // COUNTS stays BoardDataDiffer's: the first.
    //
    // *** THE MARK IS A WARNING SINCE 2026-10-03 *** (owner request: "Should flagged now be treated
    // as warnings?") - it was violet and "!". These cases hold the same rules for the warning; case
    // 1's own test became case 12 of 2026-10-03, below. The boards carry components R9 and R12
    // (ImagesWithComponents), so no image row is also warned about for naming a missing component.
    // ###########################################################################################

    private const string UpperRowWarning =
        "Warning: Another row on this sheet has the same Board label, Region, Pin and Name. Only this one, the first, is " +
        "compared with the published data, so a change in the other is not shown to a maintainer. Make them differ, or delete one.";

    private const string LowerRowWarning =
        "Warning: Another row on this sheet has the same Board label, Region, Pin and Name. Only the first is compared " +
        "with the published data, so a change in this one is not shown to a maintainer. Make them differ, or delete one.";

    private static ComponentImageEntry Image(string label, string name, string file, string note = "") =>
        new() { BoardLabel = label, Name = name, File = file, Note = note };

    private static BoardData ImagesWithComponents(params ComponentImageEntry[] images) => new()
    {
        Components = [Component("R9"), Component("R12")],
        ComponentImages = [.. images],
    };

    private static BoardTableSheet ImagesSheet(BoardTableDocument document) =>
        document.FindSheet(BoardWorkbookSchema.SheetComponentImages)!;

    private static bool IsWarnedAsDuplicate(BoardTableRow row) =>
        Cell(row, BoardWorkbookSchema.ColBoardLabel).Problems.Any(problem => problem.Code == "row.duplicate");

    // The owner's board: R9's pinout twice, where the second should have been "Pinout (secondary)".
    private static readonly ComponentImageEntry R9Pinout = Image("R9", "Pinout", "Generic shared files/Component images/resistor_10_5.png");
    private static readonly ComponentImageEntry R9Second = Image("R9", "Pinout", "Generic shared files/Component images/resistor.png");
    private static readonly ComponentImageEntry R12Pinout = Image("R12", "Pinout", "Generic shared files/Component images/resistor_1k_5.png");

    private static BoardTableDocument R9Twice() =>
        BoardTableDocument.Create(
            ImagesWithComponents(R9Pinout, R9Second, R12Pinout),
            ImagesWithComponents(R9Pinout, R9Second, R12Pinout));

    // Case 2.
    [Fact]
    public void Picking_Warnings_shows_both_rows_of_a_duplicate_and_nothing_else()
    {
        BoardTableSheet sheet = ImagesSheet(R9Twice());

        Assert.Equal(
            [true, true, false],
            sheet.Rows.Select(row => BoardTableRowFilter.Shows(BoardTableRowKinds.Warnings, row)));
    }

    // Case 3.
    [Fact]
    public void The_upper_row_says_it_is_the_one_that_counts_and_the_lower_row_that_it_is_not()
    {
        BoardTableSheet sheet = ImagesSheet(R9Twice());

        Assert.Equal(UpperRowWarning, Cell(sheet.Rows[0], BoardWorkbookSchema.ColBoardLabel).ToolTip);
        Assert.Equal(LowerRowWarning, Cell(sheet.Rows[1], BoardWorkbookSchema.ColBoardLabel).ToolTip);
    }

    // Case 4.
    [Fact]
    public void Three_rows_identifying_the_same_thing_are_all_warned_about()
    {
        ComponentImageEntry third = Image("R9", "Pinout", "Generic shared files/Component images/resistor_x.png");

        BoardTableSheet sheet = ImagesSheet(BoardTableDocument.Create(
            ImagesWithComponents(R9Pinout, R12Pinout),
            ImagesWithComponents(R9Pinout, R9Second, third, R12Pinout)));

        Assert.All(sheet.Rows.Take(3), row => Assert.True(IsWarnedAsDuplicate(row)));
        Assert.Equal(3, sheet.WarningRowCount);
    }

    // Case 5, renaming the lower row.
    [Fact]
    public void Renaming_the_lower_row_clears_both_warnings_and_undo_brings_them_back()
    {
        BoardTableDocument document = R9Twice();
        BoardTableSheet sheet = ImagesSheet(document);
        BoardTableRow upper = sheet.Rows[0];
        BoardTableRow lower = sheet.Rows[1];

        lower.Cells[Column(sheet, BoardWorkbookSchema.ColName)].Text = "Pinout (secondary)";
        sheet.Refresh();

        Assert.False(IsWarnedAsDuplicate(upper));
        Assert.Equal(0, sheet.WarningRowCount);

        document.History.Undo();

        Assert.True(IsWarnedAsDuplicate(upper));
        Assert.True(IsWarnedAsDuplicate(lower));
        Assert.Equal(2, sheet.WarningRowCount);
    }

    // Case 5, deleting the lower row.
    [Fact]
    public void Deleting_the_lower_row_clears_both_warnings_and_undo_brings_them_back()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            ImagesWithComponents(R9Pinout, R12Pinout),
            ImagesWithComponents(R9Pinout, R9Second, R12Pinout));
        BoardTableSheet sheet = ImagesSheet(document);
        BoardTableRow upper = sheet.Rows[0];

        sheet.DeleteRow(sheet.Rows[1]);

        Assert.False(IsWarnedAsDuplicate(upper));
        Assert.Equal(0, sheet.WarningRowCount);

        document.History.Undo();

        Assert.True(IsWarnedAsDuplicate(upper));
        Assert.Equal(2, sheet.WarningRowCount);
    }

    // Case 6.
    [Fact]
    public void A_duplicate_leaves_the_change_counts_as_BoardDataDiffer_has_them()
    {
        // The upper row is a real change (its Note edited); the lower one is never counted.
        ComponentImageEntry editedUpper = Image("R9", "Pinout", R9Pinout.File, note: "Checked");

        BoardData published = ImagesWithComponents(R9Pinout, R12Pinout);
        BoardData draft = ImagesWithComponents(editedUpper, R9Second, R12Pinout);

        BoardTableDocument document = BoardTableDocument.Create(published, draft);
        BoardTableSheet sheet = ImagesSheet(document);

        Assert.Equal(1, sheet.ChangeCount);
        Assert.Equal(BoardDataDiffer.CountChanges(published, document.ApplyTo(draft)), document.TotalChangeCount);
    }

    // Case 7.
    [Fact]
    public void An_upper_row_that_is_itself_changed_keeps_its_orange_cell_and_counts_as_modified_and_warned()
    {
        ComponentImageEntry editedUpper = Image("R9", "Pinout", R9Pinout.File, note: "Checked");

        BoardTableSheet sheet = ImagesSheet(BoardTableDocument.Create(
            ImagesWithComponents(R9Pinout, R12Pinout),
            ImagesWithComponents(editedUpper, R9Second, R12Pinout)));

        BoardTableRow upper = sheet.Rows[0];
        BoardTableCell note = upper.Cells[Column(sheet, BoardWorkbookSchema.ColNote)];

        Assert.Equal("~", upper.Marker);
        Assert.Equal(BoardTableCellState.Modified, note.State);
        Assert.Contains("(empty)", note.ToolTip);
        Assert.All(upper.Cells.Where(cell => cell != note), cell => Assert.Equal(BoardTableCellState.Unchanged, cell.State));

        Assert.Equal(1, sheet.ModifiedCount);
        Assert.Equal(2, sheet.WarningRowCount);
        Assert.True(BoardTableRowFilter.Shows(BoardTableRowKinds.Modified, upper));
        Assert.True(BoardTableRowFilter.Shows(BoardTableRowKinds.Warnings, upper));
    }

    // Case 8 - green since 2026-10-03 (violet until then): the colour says what changed.
    [Fact]
    public void An_upper_row_that_is_new_is_green_and_counts_as_added()
    {
        BoardTableSheet sheet = ImagesSheet(BoardTableDocument.Create(
            ImagesWithComponents(R12Pinout),
            ImagesWithComponents(R9Pinout, R9Second, R12Pinout)));

        BoardTableRow upper = sheet.Rows[0];

        Assert.Equal("+", upper.Marker);
        Assert.All(upper.Cells, cell => Assert.Equal(BoardTableCellState.Added, cell.State));
        Assert.Equal(1, sheet.AddedCount);
        Assert.Equal(1, sheet.ChangeCount);
        Assert.Equal(2, sheet.WarningRowCount);
    }

    // Case 9: an incomplete row has no identity, so it never pulls another row into a pair.
    [Fact]
    public void An_incomplete_row_is_warned_about_on_its_own()
    {
        var board = new BoardData
        {
            KiCadImportantSignals =
            [
                new() { DisplayName = "CLK", KiCadNetName = "Net-1" },
                new() { DisplayName = "CLK", KiCadNetName = "Net-2" },
            ],
        };

        BoardTableSheet sheet = BoardTableDocument.Create(board, board).FindSheet(BoardWorkbookSchema.SheetKiCadImportantSignals)!;

        sheet.Rows[1].Cells[Column(sheet, BoardWorkbookSchema.ColKiCadNetName)].Text = string.Empty;
        sheet.Refresh();

        BoardTableRow complete = sheet.Rows.First(row => !row.IsDeleted);
        BoardTableRow incomplete = sheet.Rows.Single(row => row.State == BoardTableRowState.Incomplete);

        Assert.Equal(string.Empty, complete.Marker);
        Assert.All(complete.Cells, cell => Assert.Equal(BoardTableCellState.Unchanged, cell.State));
        Assert.False(complete.HasWarnings);
        Assert.True(incomplete.HasWarnings);
        Assert.Equal(1, sheet.WarningRowCount);
    }

    // ------------------------------------------------------------------ Flagged is a warning (2026-10-03)

    // ###########################################################################################
    // *** NO MORE "FLAGGED" (owner request, 2026-10-03: "Should flagged now be treated as
    // warnings?"; cases agreed with the project owner). *** A duplicate row and a row the save
    // leaves out were violet, with a "!" and a pill of their own. They are WARNINGS now - the amber
    // corner mark on the cell, counted in the Warnings pill - and the colours are back to meaning
    // only what changed. These boards carry components R9 and R12 (ImagesWithComponents, above),
    // so no image row is also warned about for naming a component that is not there.
    // ###########################################################################################

    // Case 12, in the table.
    [Fact]
    public void Two_rows_identifying_the_same_thing_both_carry_a_warning_and_count_as_two_warnings()
    {
        BoardData board = ImagesWithComponents(R9Pinout, R9Second, R12Pinout);
        BoardTableSheet sheet = ImagesSheet(BoardTableDocument.Create(board, board));

        foreach (BoardTableRow r9 in sheet.Rows.Take(2))
        {
            BoardDataProblem problem = Assert.Single(Cell(r9, BoardWorkbookSchema.ColBoardLabel).Problems);
            Assert.Equal(BoardProblemLevel.Warning, problem.Level);
            Assert.Equal("row.duplicate", problem.Code);
        }

        Assert.Empty(sheet.Rows[2].Cells.SelectMany(cell => cell.Problems));
        Assert.Equal(2, sheet.WarningRowCount);

        Assert.StartsWith(
            "Warning: Another row on this sheet has the same Board label, Region, Pin and Name. Only the first is compared with the published data",
            Cell(sheet.Rows[1], BoardWorkbookSchema.ColBoardLabel).ToolTip,
            StringComparison.Ordinal);
    }

    // Case 14. (Agreed as "a Component images row with no Board label"; such a row is in fact
    // KEPT by the save - the Important signals sheet is the only one whose save leaves a row out.)
    [Fact]
    public void An_important_signal_with_no_net_name_is_warned_about_and_left_out_when_saving()
    {
        var published = new BoardData { KiCadImportantSignals = [new() { DisplayName = "CLK", KiCadNetName = "Net-1" }] };
        BoardTableDocument document = BoardTableDocument.Create(published, published);
        BoardTableSheet sheet = document.FindSheet(BoardWorkbookSchema.SheetKiCadImportantSignals)!;

        BoardTableCell net = sheet.Rows.Single().Cells[Column(sheet, BoardWorkbookSchema.ColKiCadNetName)];
        net.Text = string.Empty;
        sheet.Refresh();

        BoardDataProblem problem = Assert.Single(net.Problems);
        Assert.Equal(BoardProblemLevel.Warning, problem.Level);
        Assert.Equal("This row is left out when saving, because KiCad net name is empty.", problem.Message);
        Assert.Equal(1, sheet.WarningRowCount);

        // Saved: the row is gone, and with it the warning.
        BoardData saved = document.ApplyTo(published);
        Assert.Empty(saved.KiCadImportantSignals);
        Assert.Equal(0, BoardTableDocument.Create(published, saved).FindSheet(BoardWorkbookSchema.SheetKiCadImportantSignals)!.WarningRowCount);
    }

    // Case 15, the model's half: five kinds, and neither a duplicate nor a row the save leaves out
    // is coloured or marked "!" - the corner mark says it.
    [Fact]
    public void The_colour_key_has_five_kinds_and_a_duplicate_is_neither_violet_nor_marked()
    {
        Assert.Equal(
            [BoardTableRowKinds.Added, BoardTableRowKinds.Modified, BoardTableRowKinds.Deleted, BoardTableRowKinds.Errors, BoardTableRowKinds.Warnings],
            BoardTableRowFilter.Each);

        BoardData board = ImagesWithComponents(R9Pinout, R9Second, R12Pinout);
        BoardTableSheet sheet = ImagesSheet(BoardTableDocument.Create(board, board));

        Assert.All(sheet.Rows.Take(2), row =>
        {
            Assert.Equal(string.Empty, row.Marker);
            Assert.All(row.Cells, cell => Assert.Equal(BoardTableCellState.Unchanged, cell.State));
        });

        var signals = new BoardData { KiCadImportantSignals = [new() { DisplayName = "CLK", KiCadNetName = "Net-1" }] };
        BoardTableSheet signalSheet = BoardTableDocument.Create(signals, signals).FindSheet(BoardWorkbookSchema.SheetKiCadImportantSignals)!;
        signalSheet.Rows.Single().Cells[Column(signalSheet, BoardWorkbookSchema.ColKiCadNetName)].Text = string.Empty;
        signalSheet.Refresh();

        Assert.DoesNotContain(signalSheet.Rows, row => row.Marker == "!");
    }

    // Case 16: which row of a duplicate counts is BoardDataDiffer's, as before - the first.
    [Fact]
    public void The_first_row_of_a_duplicate_still_counts_as_added_or_modified_and_the_rows_after_it_count_as_nothing()
    {
        BoardTableSheet added = ImagesSheet(BoardTableDocument.Create(
            ImagesWithComponents(R12Pinout),
            ImagesWithComponents(R9Pinout, R9Second, R12Pinout)));

        Assert.Equal(BoardTableRowState.Added, added.Rows[0].State);
        Assert.Equal(BoardTableRowState.Duplicate, added.Rows[1].State);
        Assert.Equal(1, added.AddedCount);
        Assert.Equal(1, added.ChangeCount);

        BoardTableSheet modified = ImagesSheet(BoardTableDocument.Create(
            ImagesWithComponents(R9Pinout, R12Pinout),
            ImagesWithComponents(Image("R9", "Pinout", R9Pinout.File, note: "Checked"), R9Second, R12Pinout)));

        Assert.Equal(BoardTableRowState.Modified, modified.Rows[0].State);
        Assert.Equal(1, modified.ModifiedCount);
        Assert.Equal(1, modified.ChangeCount);
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
    public void With_no_published_board_nothing_is_coloured_but_duplicates_are_still_told()
    {
        // "Add a new board": every row would be green, which tells the contributor nothing. A
        // duplicate is still a duplicate, though - on Components the server's error, whose corner
        // mark says so (it was violet until 2026-10-03).
        BoardTableDocument document = BoardTableDocument.Create(
            null,
            Board(Component("U1"), Component("U2"), Component("U2")));
        BoardTableSheet sheet = Components(document);

        Assert.False(document.HasBaseline);
        Assert.All(sheet.Rows, row => Assert.All(row.Cells, cell => Assert.Equal(BoardTableCellState.Unchanged, cell.State)));
        Assert.Equal(BoardTableRowState.Unchanged, sheet.Rows[0].State);
        Assert.Equal(BoardTableRowState.Duplicate, sheet.Rows[2].State);
        Assert.True(sheet.Rows[2].HasErrors);
        Assert.True(sheet.Rows[1].HasErrors);   // both U2 rows (every row of a duplicate since 2026-10-03)
        Assert.Equal(2, sheet.ErrorRowCount);
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

        // Giving the published U1 a region changes its key, as changing its label does - and since
        // 2026-10-04 that is the same row MODIFIED, its region cell orange, when nothing else of it
        // changed (it was a red ghost plus U1/PAL added until then).
        Assert.Equal(BoardTableRowState.Modified, original.State);
        Assert.Equal(BoardTableCellState.Modified, Cell(original, BoardWorkbookSchema.ColRegion).State);
        Assert.DoesNotContain(sheet.Rows, row => row.IsDeleted);
        Assert.Equal(2, sheet.ChangeCount);
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

    // What "Show changes only" showed - every row added, changed or deleted, and a new empty one -
    // is what picking Added, Modified and Deleted together shows (BoardTableRowFilter.Changes,
    // since the colour key became the filter, 2026-10-02; Flagged was a fourth until 2026-10-03).
    [Fact]
    public void Only_rows_that_are_not_plainly_unchanged_are_shown_by_the_changes_filter()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1"), Component("U2")),
            Board(Component("U1"), Component("U3")));
        BoardTableSheet sheet = Components(document);
        BoardTableRow blank = sheet.InsertRow(null);

        Assert.False(BoardTableRowFilter.Shows(BoardTableRowFilter.Changes, Row(sheet, "U1")));
        Assert.True(BoardTableRowFilter.Shows(BoardTableRowFilter.Changes, Row(sheet, "U3")));
        Assert.True(BoardTableRowFilter.Shows(BoardTableRowFilter.Changes, Row(sheet, "U2", deleted: true)));
        Assert.True(BoardTableRowFilter.Shows(BoardTableRowFilter.Changes, blank));
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

    // ------------------------------------------------------------------ Deleting several rows at once (2026-10-02)

    // ###########################################################################################
    // Owner request, 2026-10-02: "mark multiple rows and then delete those in one go" - the eight
    // Pinout rows of a Component images sheet, say. BoardTableSheet.DeleteRows; the cases are the
    // ones agreed before the code was written, numbered as they were then.
    // ###########################################################################################

    // Case 2: every selected row goes - a published one stays as a red row where it was, one the
    // contributor added simply goes.
    [Fact]
    public void Deleting_several_rows_deletes_all_of_them_published_ones_stay_red_in_place_and_added_ones_go()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1"), Component("U2"), Component("U3")),
            Board(Component("U1"), Component("U2"), Component("U3"), Component("U4")));
        BoardTableSheet sheet = Components(document);

        int deleted = sheet.DeleteRows([Row(sheet, "U1"), Row(sheet, "U3"), Row(sheet, "U4")], out _);

        Assert.Equal(3, deleted);
        Assert.Equal(["U1", "U2", "U3"], sheet.Rows.Select(Label));
        Assert.Equal([true, false, true], sheet.Rows.Select(row => row.IsDeleted));
        Assert.Equal(2, sheet.DeletedCount);
        Assert.True(document.HasUnsavedChanges);
    }

    // Case 4: red rows among the selection are skipped - already deleted - and a selection of only
    // red rows deletes nothing and leaves no undo step behind.
    [Fact]
    public void Red_rows_among_several_selected_are_skipped_and_only_red_rows_delete_nothing()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1"), Component("U2"), Component("U3")),
            Board(Component("U2"), Component("U3")));
        BoardTableSheet sheet = Components(document);
        BoardTableRow ghost = Row(sheet, "U1", deleted: true);

        Assert.Equal(1, sheet.DeleteRows([ghost, Row(sheet, "U2")], out _));
        Assert.Equal(["-U1", "-U2", "U3"], sheet.Rows.Select(row => (row.IsDeleted ? "-" : "") + Label(row)));

        int steps = document.History.UndoCount;
        Assert.Equal(0, sheet.DeleteRows(sheet.Rows.Where(row => row.IsDeleted).ToList(), out BoardTableDeletedWith? deletedWith));
        Assert.Null(deletedWith);
        Assert.Equal(steps, document.History.UndoCount);
    }

    // Case 5: on Components, deleting two components at once takes each one's rows on the image,
    // file and link sheets, as deleting one does - still ONE undo step, and what went with them
    // said in one sentence for the status line.
    [Fact]
    public void Deleting_two_components_at_once_takes_both_ones_rows_on_the_other_sheets_in_one_undo_step()
    {
        BoardData board = BoardWithU8AndU9();
        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet components = Components(document);

        Assert.Equal(2, components.DeleteRows([Row(components, "U8"), Row(components, "U9")], out BoardTableDeletedWith? deletedWith));

        Assert.Empty(LiveLabels(document, BoardWorkbookSchema.SheetComponentImages));
        Assert.Empty(LiveLabels(document, BoardWorkbookSchema.SheetComponentLocalFiles));
        Assert.Empty(LiveLabels(document, BoardWorkbookSchema.SheetComponentLinks));
        Assert.Empty(document.ApplyTo(board).ComponentHighlights);

        Assert.Equal(
            "Also deleted with U8 and U9: 2 rows on Component images, 1 on Component local files and 1 on Component links. " +
            "Their 3 highlights on the schematics go when you save. Undo (Ctrl+Z) brings it all back.",
            deletedWith!.Describe());

        Assert.Equal(1, document.History.UndoCount);
        document.History.Undo();

        Assert.Equal(["U8", "U9"], LiveLabels(document, BoardWorkbookSchema.SheetComponents));
        Assert.Equal(["U8", "U9"], LiveLabels(document, BoardWorkbookSchema.SheetComponentImages));
        Assert.Equal(3, document.ApplyTo(board).ComponentHighlights.Count);
    }

    // ###########################################################################################
    // Every sheet refresh checks the whole board, and deleting components refreshes Components and
    // then up to three sheets per component - up to 25 whole-board checks for 8 components, on the
    // UI thread (code review, 2026-10-04). The operation, and its undo, now check it ONCE - and the
    // problems and the counts still come out as a check after every refresh would leave them.
    // ###########################################################################################
    [Fact]
    public void Deleting_components_with_rows_on_other_sheets_checks_the_board_once_and_so_does_its_undo()
    {
        BoardData board = BoardWithU8AndU9();
        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet components = Components(document);
        int before = document.ProblemChecksRun;

        components.DeleteRows([Row(components, "U8"), Row(components, "U9")], out _);

        Assert.Equal(before + 1, document.ProblemChecksRun);

        document.History.Undo();

        Assert.Equal(before + 2, document.ProblemChecksRun);

        // And what the counts say is what a fresh check says.
        int errors = document.ErrorRowCount;
        int warnings = document.WarningRowCount;
        document.RefreshProblems();
        Assert.Equal(errors, document.ErrorRowCount);
        Assert.Equal(warnings, document.WarningRowCount);
    }

    // The checks still follow a change made inside the deferral, and Changed is raised once they
    // have run, so a toolbar repainting for it reads the new counts.
    [Fact]
    public void A_deferred_check_runs_when_the_scope_ends_and_raises_Changed_after_it()
    {
        BoardData board = BoardWithU8AndU9();
        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet components = Components(document);
        int before = document.ProblemChecksRun;
        int problemRowsBefore = document.ErrorRowCount + document.WarningRowCount;
        int? problemRowsAtChanged = null;

        using (document.DeferProblems())
        {
            // U9 renamed U8: a duplicate label, which is a problem on the Components sheet.
            Row(components, "U9").Cells[components.Columns.ToList().IndexOf(BoardWorkbookSchema.ColBoardLabel)].Text = "U8";
            components.Refresh();

            Assert.Equal(before, document.ProblemChecksRun);
            document.Changed += (_, _) => problemRowsAtChanged ??= document.ErrorRowCount + document.WarningRowCount;
        }

        Assert.Equal(before + 1, document.ProblemChecksRun);
        Assert.True(document.ErrorRowCount + document.WarningRowCount > problemRowsBefore);
        Assert.Equal(document.ErrorRowCount + document.WarningRowCount, problemRowsAtChanged);
    }

    // ###########################################################################################
    // *** ONE EDIT MAPS ONE ROW AGAIN (code review, 2026-10-04). *** The checks run after every
    // cell edit, and each run mapped every live row of all nine sheets afresh - thousands of rows on
    // the largest boards for one typed cell, on the UI thread. A row's mapping is remembered until
    // one of ITS cells changes: an edit and its refresh map that row and nothing else.
    // ###########################################################################################
    [Fact]
    public void An_edit_maps_only_its_own_row_again_and_the_checks_still_follow_it()
    {
        BoardData board = BoardWithU8AndU9();
        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet components = Components(document);
        int label = components.Columns.ToList().IndexOf(BoardWorkbookSchema.ColBoardLabel);
        int mappedBefore = document.RowsMapped;
        int problemRowsBefore = document.ErrorRowCount + document.WarningRowCount;

        // U9 renamed U8: a duplicate label.
        Row(components, "U9").Cells[label].Text = "U8";
        components.Refresh();

        Assert.Equal(mappedBefore + 1, document.RowsMapped);
        Assert.True(document.ErrorRowCount + document.WarningRowCount > problemRowsBefore);

        // Undone, the text goes back through the other writer - and the problem goes with it.
        document.History.Undo();

        Assert.Equal(problemRowsBefore, document.ErrorRowCount + document.WarningRowCount);

        // A check with nothing changed maps nothing at all.
        int mappedNow = document.RowsMapped;
        document.RefreshProblems();
        Assert.Equal(mappedNow, document.RowsMapped);
    }

    // ------------------------------------------------------------------ The colour key counts the whole draft (2026-10-02)

    // ###########################################################################################
    // Owner decision, 2026-10-02 (question E): "Added, Modified and Deleted should work per
    // system like Error and Warning" - the colour key counts the WHOLE draft, every sheet added
    // together, for all five kinds. It counted the sheet on screen, which read "0 Errors" on a
    // sheet without any while the draft's row said "8 errors" - once the sheet tabs stopped
    // carrying their own counts, nothing else on screen said where they were.
    // ###########################################################################################
    [Fact]
    public void The_colour_key_counts_are_the_whole_drafts_every_sheet_added_together()
    {
        BoardData published = BoardWithU8AndU9();
        var draft = new BoardData
        {
            Components = [Component("U8", "changed"), Component("U10"), Component("U10")],
            ComponentImages = [published.ComponentImages[0], new ComponentImageEntry { BoardLabel = "U11", Name = "Clock", File = "a/u11.png" }],
            ComponentLocalFiles = published.ComponentLocalFiles,
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U8", Name = "Ref", Url = "not a link" }],
            ComponentHighlights = published.ComponentHighlights
        };

        BoardTableDocument document = BoardTableDocument.Create(published, draft);

        foreach ((System.Func<BoardTableSheet, int> perSheet, int total) in new (System.Func<BoardTableSheet, int>, int)[]
        {
            (sheet => sheet.AddedCount, document.AddedCount),
            (sheet => sheet.ModifiedCount, document.ModifiedCount),
            (sheet => sheet.DeletedCount, document.DeletedCount),
            (sheet => sheet.ErrorRowCount, document.ErrorRowCount),
            (sheet => sheet.WarningRowCount, document.WarningRowCount)
        })
        {
            Assert.Equal(document.Sheets.Sum(perSheet), total);
        }

        // And the board really has each kind on more than one sheet - or a count of the sheet on
        // screen could pass this test.
        Assert.True(document.Sheets.Count(sheet => sheet.AddedCount > 0) > 1);
        Assert.True(document.Sheets.Count(sheet => sheet.DeletedCount > 0) > 1);
        Assert.True(document.ModifiedCount > 0);
        Assert.True(document.ErrorRowCount > 0);
        Assert.True(document.WarningRowCount > 0);
    }

    // ------------------------------------------------------------------ Which sheet tabs show (2026-09-26)

    // A board whose Components sheet has a change and whose Credits sheet has none.
    private static BoardTableDocument ComponentsChanged() =>
        BoardTableDocument.Create(Board(Component("U1", "CPU")), Board(Component("U1", "CPU 6510")));

    // ###########################################################################################
    // A filter hides the sheets it would show nothing of (owner request, 2026-09-26, when it was
    // "Show changes only"; the colour key's pills since 2026-10-02) - and with nothing picked every
    // sheet shows.
    // ###########################################################################################
    [Fact]
    public void With_a_filter_only_the_sheets_with_something_to_show_have_tabs()
    {
        BoardTableDocument document = ComponentsChanged();

        Assert.Equal(document.Sheets, document.SheetsShown(BoardTableRowKinds.None, current: null));
        Assert.Equal([BoardWorkbookSchema.SheetComponents], document.SheetsShown(BoardTableRowFilter.Changes, current: null).Select(sheet => sheet.Name));
    }

    // A duplicate row is not a change, but it IS shown by the pill that tells it - Errors on
    // Components, where the server refuses it - so its sheet keeps its tab, or the one thing
    // needing a second look would be out of reach. (It was shown by "Show changes only" while it
    // was "Flagged", until 2026-10-03; the changes alone no longer show it.)
    [Fact]
    public void A_sheet_with_only_a_duplicate_row_keeps_its_tab_while_its_problem_is_picked()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1")),
            Board(Component("U1"), Component("U1")));

        Assert.Equal(0, Components(document).ChangeCount);
        Assert.True(Components(document).HasRowsShownBy(BoardTableRowKinds.Errors));
        Assert.Contains(Components(document), document.SheetsShown(BoardTableRowKinds.Errors, current: null));
        Assert.False(Components(document).HasRowsShownBy(BoardTableRowFilter.Changes));
    }

    // The sheet being worked on keeps its tab when its last change is undone - it is not pulled
    // away mid-work; it goes once another sheet is chosen.
    [Fact]
    public void The_current_sheet_keeps_its_tab_while_it_is_on_screen()
    {
        BoardTableDocument document = ComponentsChanged();
        BoardTableSheet credits = document.FindSheet(BoardWorkbookSchema.SheetCredits)!;

        Assert.Contains(credits, document.SheetsShown(BoardTableRowFilter.Changes, current: credits));
        Assert.DoesNotContain(credits, document.SheetsShown(BoardTableRowFilter.Changes, current: null));
    }

    // Nothing to show anywhere: every tab stays, rather than all but one vanishing.
    [Fact]
    public void A_table_with_nothing_to_show_keeps_every_tab()
    {
        BoardTableDocument document = BoardTableDocument.Create(Board(Component("U1")), Board(Component("U1")));

        Assert.Equal(document.Sheets, document.SheetsShown(BoardTableRowFilter.Changes, current: null));
    }

    // ###########################################################################################
    // The sheet a table opens on: the one wanted (the sheet last looked at) while its tab shows,
    // else the first that shows.
    // ###########################################################################################
    [Fact]
    public void A_table_opens_on_the_wanted_sheet_while_its_tab_shows()
    {
        BoardTableDocument document = ComponentsChanged();

        Assert.Equal(BoardWorkbookSchema.SheetCredits, document.SheetToShow(BoardWorkbookSchema.SheetCredits, BoardTableRowKinds.None).Name);
        Assert.Equal(BoardWorkbookSchema.SheetComponents, document.SheetToShow(BoardWorkbookSchema.SheetCredits, BoardTableRowFilter.Changes).Name);
        Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, document.SheetToShow(null, BoardTableRowKinds.None).Name);
        Assert.Equal(BoardWorkbookSchema.SheetBoardSchematics, document.SheetToShow("No such sheet", BoardTableRowKinds.None).Name);
    }

    // ###########################################################################################
    // Agreed case 6 (2026-10-03), the rule half: a pill counting nothing cannot be picked - it
    // could only show an empty table - but a picked one can always be put back, also once its
    // count has dropped to 0. CountOf is each pill's own count.
    // ###########################################################################################
    [Fact]
    public void A_pill_counting_nothing_cannot_be_picked_but_a_picked_one_can_always_be_put_back()
    {
        BoardTableDocument document = ComponentsChanged();

        Assert.Equal(1, document.CountOf(BoardTableRowKinds.Modified));
        Assert.Equal(0, document.CountOf(BoardTableRowKinds.Added));

        Assert.True(document.CanToggle(BoardTableRowKinds.None, BoardTableRowKinds.Modified));
        Assert.False(document.CanToggle(BoardTableRowKinds.None, BoardTableRowKinds.Added));
        Assert.False(document.CanToggle(BoardTableRowKinds.Modified, BoardTableRowKinds.Added));

        // Picked while it had rows, then the last one put back: still clickable, to turn it off.
        Assert.True(document.CanToggle(BoardTableRowKinds.Added, BoardTableRowKinds.Added));
        Assert.True(document.CanToggle(BoardTableRowKinds.Added | BoardTableRowKinds.Modified, BoardTableRowKinds.Added));
    }

    // Each kind's count is the colour key's own.
    [Fact]
    public void CountOf_is_the_colour_keys_count_for_each_kind()
    {
        BoardTableDocument document = BoardTableDocument.Create(
            Board(Component("U1", "CPU"), Component("U2")),
            Board(Component("U1", "CPU 6510"), Component("U3"), Component("U3")));

        Assert.Equal(document.AddedCount, document.CountOf(BoardTableRowKinds.Added));
        Assert.Equal(document.ModifiedCount, document.CountOf(BoardTableRowKinds.Modified));
        Assert.Equal(document.DeletedCount, document.CountOf(BoardTableRowKinds.Deleted));
        Assert.Equal(document.ErrorRowCount, document.CountOf(BoardTableRowKinds.Errors));
        Assert.Equal(document.WarningRowCount, document.CountOf(BoardTableRowKinds.Warnings));
        Assert.True(document.ErrorRowCount > 0 && document.DeletedCount > 0, "the test needs an error and a deleted row");

        Assert.Throws<ArgumentOutOfRangeException>(() => document.CountOf(BoardTableRowKinds.Added | BoardTableRowKinds.Deleted));
    }
}
