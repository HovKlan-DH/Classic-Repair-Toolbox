using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The checks in the table (owner request, 2026-10-02: "integrate the existing validation check
// into the "Draft" system table view ... either "Error" or "Warning"") - BoardTableDocument
// .RefreshProblems placing BoardDataChecks' problems on the cells they are about.
//
// What must hold: the problem lands on the RIGHT row - the table interleaves deleted ghosts and
// new empty rows, which the checks never see - and in the right column; it follows edits; and a
// problem no row can carry (a highlight) is kept apart rather than lost.
// ###########################################################################################
public sealed class BoardTableProblemsTests
{
    private const string Board = "Commodore/C64/250407";

    private static ComponentEntry Component(string label, string friendlyName = "") =>
        new() { BoardLabel = label, FriendlyName = friendlyName };

    private static ComponentHighlightEntry Highlight(string schematic, string label) =>
        new() { SchematicName = schematic, BoardLabel = label, X = "1", Y = "1", Width = "10", Height = "10" };

    private static readonly BoardSchematicEntry Main = new() { SchematicName = "Main", SchematicImageFile = $"{Board}/main.png" };

    private static int Column(BoardTableSheet sheet, string column) => sheet.Columns.ToList().IndexOf(column);

    private static BoardTableRow Row(BoardTableSheet sheet, string label, bool deleted = false) =>
        sheet.Rows.First(row => row.IsDeleted == deleted && row.Cells[Column(sheet, BoardWorkbookSchema.ColBoardLabel)].Text == label);

    // ###########################################################################################
    // A component nothing marks gets its warning on its Board label - on its OWN row, with a deleted
    // ghost and a new empty row in between, neither of which the checks see.
    // ###########################################################################################
    [Fact]
    public void A_problem_lands_on_its_own_row_past_ghosts_and_empty_rows()
    {
        BoardData published = new() { Schematics = [Main], Components = [Component("U1"), Component("U2"), Component("U3")] };
        BoardData draft = new()
        {
            Schematics = [Main],
            Components = [Component("U1"), Component("U3"), Component("U4")],
            ComponentHighlights = [Highlight("Main", "U1"), Highlight("Main", "U3")]
        };

        BoardTableDocument document = BoardTableDocument.Create(published, draft);
        BoardTableSheet sheet = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        sheet.InsertRow(Row(sheet, "U1"));

        BoardTableRow u4 = Row(sheet, "U4");
        BoardTableCell label = u4.Cells[Column(sheet, BoardWorkbookSchema.ColBoardLabel)];

        Assert.Equal(BoardProblemLevel.Warning, label.ProblemLevel);
        Assert.Equal("component.no_highlight", Assert.Single(label.Problems).Code);
        Assert.True(u4.HasWarnings);
        Assert.False(u4.HasErrors);

        Assert.All(sheet.Rows.Where(row => !ReferenceEquals(row, u4)), row => Assert.False(row.HasWarnings || row.HasErrors));
        Assert.Equal(1, sheet.WarningRowCount);
        Assert.Equal(0, sheet.ErrorRowCount);
        Assert.Equal(1, document.WarningCount);
    }

    // An error follows the edit that made it, and goes with the edit that fixes it.
    [Fact]
    public void An_error_appears_with_the_edit_that_makes_it_and_goes_with_the_fix()
    {
        BoardData board = new()
        {
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U1", Name = "Datasheet", Url = "https://example.com" }],
            Components = [Component("U1")],
            Schematics = [Main],
            ComponentHighlights = [Highlight("Main", "U1")]
        };

        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet links = document.FindSheet(BoardWorkbookSchema.SheetComponentLinks)!;
        BoardTableCell url = links.Rows[0].Cells[Column(links, BoardWorkbookSchema.ColUrl)];

        Assert.Equal(BoardProblemLevel.None, url.ProblemLevel);

        url.Text = "ftp://example.com";
        links.Refresh();

        Assert.Equal(BoardProblemLevel.Error, url.ProblemLevel);
        Assert.Equal(1, document.ErrorCount);
        Assert.StartsWith("Error: A link for [U1] is [ftp://example.com]", url.ToolTip, StringComparison.Ordinal);
        Assert.True(document.HasRowsShownBy(BoardTableRowKinds.Errors));

        url.Text = "https://example.com";
        links.Refresh();

        Assert.Equal(BoardProblemLevel.None, url.ProblemLevel);
        Assert.Null(url.ToolTip);
        Assert.Equal(0, document.ErrorCount);
    }

    // A changed cell with a problem says both - what is wrong, then the value it replaced.
    [Fact]
    public void A_changed_cell_with_a_problem_says_both()
    {
        BoardData published = new()
        {
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U1", Name = "Datasheet", Url = "https://example.com" }],
            Components = [Component("U1")],
            Schematics = [Main],
            ComponentHighlights = [Highlight("Main", "U1")]
        };

        BoardData draft = new()
        {
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U1", Name = "Datasheet", Url = "javascript:alert(1)" }],
            Components = [Component("U1")],
            Schematics = [Main],
            ComponentHighlights = [Highlight("Main", "U1")]
        };

        BoardTableSheet links = BoardTableDocument.Create(published, draft).FindSheet(BoardWorkbookSchema.SheetComponentLinks)!;
        BoardTableCell url = links.Rows[0].Cells[Column(links, BoardWorkbookSchema.ColUrl)];

        Assert.Equal(BoardTableCellState.Modified, url.State);
        string[] lines = url.ToolTip!.Split(Environment.NewLine);
        Assert.StartsWith("Error: ", lines[0], StringComparison.Ordinal);
        Assert.Equal("Published value: https://example.com", lines[1]);
        Assert.Equal(lines[0], url.ProblemToolTip);
    }

    // A highlight is in no sheet: its problem is kept apart, and still counted.
    [Fact]
    public void A_highlight_problem_is_kept_apart_and_counted()
    {
        BoardData board = new()
        {
            Schematics = [Main],
            Components = [Component("U1")],
            ComponentHighlights = [Highlight("Main", "U1"), Highlight("Renamed away", "U1")]
        };

        BoardTableDocument document = BoardTableDocument.Create(board, board);

        BoardDataProblem outside = Assert.Single(document.ProblemsOutsideSheets);
        Assert.Equal("highlight.unknown_schematic", outside.Code);
        Assert.Equal(1, document.ErrorCount);
        Assert.False(document.HasRowsShownBy(BoardTableRowKinds.Errors));
    }

    // ###########################################################################################
    // RENAMING A SCHEMATIC IN THE TABLE leaves its highlights naming the old name - which the
    // server refuses. So the rename is an error at once, kept apart since highlights are in no row.
    // ###########################################################################################
    [Fact]
    public void Renaming_a_schematic_its_highlights_are_on_is_an_error()
    {
        BoardData board = new() { Schematics = [Main], Components = [Component("U1")], ComponentHighlights = [Highlight("Main", "U1")] };

        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet schematics = document.FindSheet(BoardWorkbookSchema.SheetBoardSchematics)!;
        Assert.Equal(0, document.ErrorCount);

        schematics.Rows[0].Cells[Column(schematics, BoardWorkbookSchema.ColSchematicName)].Text = "Main board";
        schematics.Refresh();

        Assert.Equal("highlight.unknown_schematic", Assert.Single(document.ProblemsOutsideSheets).Code);
    }

    // Deleting a component in the table drops its highlights at save - so they are not reported
    // as highlights for a component that is gone.
    [Fact]
    public void A_component_deleted_in_the_table_takes_its_highlight_warnings_with_it()
    {
        BoardData board = new()
        {
            Schematics = [Main],
            Components = [Component("U1"), Component("U2")],
            ComponentHighlights = [Highlight("Main", "U1"), Highlight("Main", "U2")]
        };

        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet components = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;

        components.DeleteRow(Row(components, "U2"));

        Assert.Empty(document.ProblemsOutsideSheets);
        Assert.Equal(0, document.WarningCount);
    }

    // With a lookup, a file that is not there is an error on its file cell.
    [Fact]
    public void A_file_the_lookup_cannot_find_is_an_error_on_its_cell()
    {
        BoardData board = new()
        {
            Schematics = [Main, new BoardSchematicEntry { SchematicName = "Other", SchematicImageFile = $"{Board}/gone.png" }],
            Components = [Component("U1")],
            ComponentHighlights = [Highlight("Main", "U1")]
        };

        BoardTableDocument document = BoardTableDocument.Create(board, board, files: new SuppliedFileLookup([$"{Board}/main.png"]));
        BoardTableSheet schematics = document.FindSheet(BoardWorkbookSchema.SheetBoardSchematics)!;

        Assert.Equal(BoardProblemLevel.None, schematics.Rows[0].Cells[Column(schematics, BoardWorkbookSchema.ColSchematicImageFile)].ProblemLevel);
        Assert.Equal("file.missing", Assert.Single(schematics.Rows[1].Cells[Column(schematics, BoardWorkbookSchema.ColSchematicImageFile)].Problems).Code);
        Assert.Equal(1, schematics.ErrorRowCount);
    }
}
