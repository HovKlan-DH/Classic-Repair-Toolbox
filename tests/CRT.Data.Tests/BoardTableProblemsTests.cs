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

    // With a category - a component without one is a warning of its own (component.no_category).
    private static ComponentEntry Component(string label, string friendlyName = "") =>
        new() { BoardLabel = label, FriendlyName = friendlyName, Category = "IC" };

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
    // *** THE TABLE SAYS WHAT THE BOARD WILL HOLD (code review, 2026-11-02). *** Renaming a schematic
    // used to be an error at once, "which the server refuses" - but the highlights follow it now, in
    // a draft's session (DraftEditRules) and in a maintainer's amendment (WithTableSections), so the
    // table told the contributor to fix what was not broken, in words pointing at "Mark components
    // manually". The checks see the highlights as the rename leaves them.
    // ###########################################################################################
    [Fact]
    public void Renaming_a_schematic_its_highlights_are_on_is_no_problem_as_they_follow_it()
    {
        BoardData board = new() { Schematics = [Main], Components = [Component("U1")], ComponentHighlights = [Highlight("Main", "U1")] };

        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet schematics = document.FindSheet(BoardWorkbookSchema.SheetBoardSchematics)!;
        Assert.Equal(0, document.ErrorCount);

        schematics.Rows[0].Cells[Column(schematics, BoardWorkbookSchema.ColSchematicName)].Text = "Main board";
        schematics.Refresh();

        Assert.Empty(document.ProblemsOutsideSheets);
        Assert.Equal(0, document.ErrorCount);
    }

    // A marked schematic deleted takes what was drawn on it - said beside the table - so no error
    // about a highlight on it, and the component marked nowhere now is a warning that is true.
    [Fact]
    public void Deleting_a_schematic_its_highlights_are_on_takes_them_along()
    {
        BoardData board = new() { Schematics = [Main], Components = [Component("U1")], ComponentHighlights = [Highlight("Main", "U1")] };

        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet schematics = document.FindSheet(BoardWorkbookSchema.SheetBoardSchematics)!;

        schematics.DeleteRow(schematics.Rows[0]);

        Assert.Empty(document.ProblemsOutsideSheets);
        Assert.Equal(0, document.ErrorCount);
        Assert.Equal("component.no_highlight", Assert.Single(document.Problems).Code);
    }

    // A component renamed: neither "a highlight for a component that is not there" above the table
    // nor "not marked on any schematic" on its row.
    [Fact]
    public void Renaming_a_component_its_highlights_follow_it_says_nothing_about_them()
    {
        BoardData board = new() { Schematics = [Main], Components = [Component("U1")], ComponentHighlights = [Highlight("Main", "U1")] };

        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet components = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;

        Row(components, "U1").Cells[Column(components, BoardWorkbookSchema.ColBoardLabel)].Text = "U01";
        components.Refresh();

        Assert.Empty(document.Problems);
    }

    // What holds a highlight back from following is said as before: a label AND another cell changed
    // at once is a new component and a deleted one (ComponentRelabel), so the rectangle has no owner -
    // as the draft will find too (below) - and a highlight on a name nothing was renamed from stays.
    [Fact]
    public void A_relabel_the_rules_do_not_follow_still_says_the_highlight_has_no_component()
    {
        BoardData board = new() { Schematics = [Main], Components = [Component("U1")], ComponentHighlights = [Highlight("Main", "U1")] };

        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet components = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        BoardTableRow row = Row(components, "U1");

        row.Cells[Column(components, BoardWorkbookSchema.ColBoardLabel)].Text = "U01";
        row.Cells[Column(components, BoardWorkbookSchema.ColFriendlyName)].Text = "Another name";
        components.Refresh();

        Assert.Contains(document.ProblemsOutsideSheets, problem => problem.Code == "highlight.orphan");
    }

    // ###########################################################################################
    // *** STEP BY STEP, AS THE DRAFT HOLDS IT. *** A draft's session follows each hand-over in turn, so
    // a component relabelled and THEN given another name keeps its rectangles - though the rows as
    // opened against the rows now look like a new component. After every step of a sitting the table
    // and the checks on the board the session holds must say the same of the highlights.
    // ###########################################################################################
    [Fact]
    public void The_table_and_the_drafts_own_board_agree_on_the_highlights_after_every_hand_over()
    {
        BoardData board = new()
        {
            Schematics = [Main, new BoardSchematicEntry { SchematicName = "Back", SchematicImageFile = $"{Board}/back.png" }],
            Components = [Component("U1"), Component("U2")],
            ComponentHighlights = [Highlight("Main", "U1"), Highlight("Back", "U1"), Highlight("Main", "U2")]
        };

        var content = new BoardContent(board, [], []);
        BoardTableDocument document = BoardTableDocument.Create(board, board);
        BoardTableSheet components = document.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        BoardTableSheet schematics = document.FindSheet(BoardWorkbookSchema.SheetBoardSchematics)!;
        int label = Column(components, BoardWorkbookSchema.ColBoardLabel);
        int friendly = Column(components, BoardWorkbookSchema.ColFriendlyName);
        int name = Column(schematics, BoardWorkbookSchema.ColSchematicName);

        void Step(BoardTableSheet sheet, Action change)
        {
            change();
            sheet.Refresh();

            content = HandOver(document, content);

            List<string> onBoard = BoardDataChecks.Check(BoardCheckRows.From(content.Board), null, BoardCheckScope.Everything)
                .Where(problem => problem.Code is "highlight.unknown_schematic" or "highlight.orphan" or "component.no_highlight")
                .Select(problem => $"{problem.Code} {problem.Subject}")
                .Order()
                .ToList();

            List<string> inTable = document.Problems
                .Where(problem => problem.Code is "highlight.unknown_schematic" or "highlight.orphan" or "component.no_highlight")
                .Select(problem => $"{problem.Code} {problem.Subject}")
                .Order()
                .ToList();

            Assert.Equal(onBoard, inTable);
        }

        Step(components, () => Row(components, "U1").Cells[label].Text = "U01");
        Assert.Equal(["U01", "U01", "U2"], content.Board.ComponentHighlights.Select(highlight => highlight.BoardLabel));

        Step(components, () => Row(components, "U01").Cells[friendly].Text = "Another name");
        Assert.Empty(document.ProblemsOutsideSheets);

        Step(schematics, () => schematics.Rows[0].Cells[name].Text = "Main board");
        Step(schematics, () => schematics.Rows[0].Cells[name].Text = "Top");
        Assert.Equal(["Top", "Back", "Top"], content.Board.ComponentHighlights.Select(highlight => highlight.SchematicName));
        Assert.Empty(document.ProblemsOutsideSheets);

        // A label and another cell at once: nothing followed, and both sides say so.
        Step(components, () =>
        {
            Row(components, "U2").Cells[label].Text = "U02";
            Row(components, "U02").Cells[friendly].Text = "Yet another";
        });
        Assert.Contains(document.ProblemsOutsideSheets, problem => problem.Code == "highlight.orphan");

        Step(schematics, () => schematics.DeleteRow(schematics.Rows[1]));
        Assert.Equal(["Top", "Top"], content.Board.ComponentHighlights.Select(highlight => highlight.SchematicName));

        Step(components, () => components.DeleteRow(Row(components, "U01")));
        Assert.Equal(["U2"], content.Board.ComponentHighlights.Select(highlight => highlight.BoardLabel));
    }

    // The table's change put into the board as a draft's session does it, what follows included.
    private static BoardContent HandOver(BoardTableDocument document, BoardContent content)
    {
        Func<BoardContent, BoardContent?> edit = document.TakeSessionEdit();
        DraftPreparedEdit? prepared = DraftSession.Prepare(content, before => (edit(before), null));
        document.MarkSaved();

        return prepared?.After ?? content;
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
