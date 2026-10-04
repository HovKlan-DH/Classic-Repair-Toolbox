using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE CHECKS IN THE TABLE AS PAINTED (owner request, 2026-10-02: "integrate the existing validation
// check into the "Draft" system table view ... either "Error" or "Warning" next to the "Flagged"
// one ... ideally the user should be able to see/focus only these") - BoardTableEditor over
// BoardDataChecks' problems, in a SHOWN window, reading what the grid really draws.
//
// What is pinned: a problem cell's corner mark in its level's colour OVER its own wash; the
// Errors and Warnings pills counting the sheet on screen and filtering like the others; the sheet
// tab's pill for a sheet not on screen; and the line for problems in no row.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardTableEditorChecksTests
{
    private const string Board = "Commodore/C64/250407";

    private static readonly BoardSchematicEntry Main = new() { SchematicName = "Main", SchematicImageFile = $"{Board}/main.png" };

    private static ComponentEntry Component(string label, string friendlyName = "") =>
        new() { BoardLabel = label, FriendlyName = friendlyName };

    private static ComponentHighlightEntry Highlight(string schematic, string label) =>
        new() { SchematicName = schematic, BoardLabel = label, X = "1", Y = "1", Width = "10", Height = "10" };

    // U1 marked on Main; U2 not marked (a warning); a link that is not a web address (an error).
    private static BoardData Checked() => new()
    {
        Schematics = [Main],
        Components = [Component("U1"), Component("U2")],
        ComponentHighlights = [Highlight("Main", "U1")],
        ComponentLinks =
        [
            new ComponentLinkEntry { BoardLabel = "U1", Name = "Good", Url = "https://example.com" },
            new ComponentLinkEntry { BoardLabel = "U1", Name = "Bad", Url = "ftp://example.com" }
        ]
    };

    private static (Window Window, BoardTableEditor Editor) Shown(BoardData board, BoardData? published)
    {
        var editor = new BoardTableEditor();
        editor.Open(BoardTableDocument.Create(published, board));

        var window = new Window { Content = editor, Width = 1300, Height = 600 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return (window, editor);
    }

    private static Color? ThemeColor(string key)
    {
        Application app = Application.Current!;
        return app.TryGetResource(key, app.ActualThemeVariant, out object? resource) && resource is ISolidColorBrush brush
            ? brush.Color
            : null;
    }

    private static DataGridCell CellOnScreen(Window window, BoardTableRow row, int columnIndex) =>
        window.GetVisualDescendants()
            .OfType<DataGridCell>()
            .Single(cell => ReferenceEquals(cell.DataContext, row) && cell.OwningColumn?.Tag is int tag && tag == columnIndex);

    private static int Column(BoardTableSheet sheet, string column) => sheet.Columns.ToList().IndexOf(column);

    private static List<BoardTableRow> ShownRows(BoardTableEditor editor) =>
        ((System.Collections.IEnumerable)editor.GetControl<DataGrid>("TableGrid").ItemsSource!).Cast<BoardTableRow>().ToList();

    private static TabItem SheetTab(BoardTableEditor editor, string sheet) =>
        editor.GetControl<TabControl>("SheetTabs").Items.OfType<TabItem>().Single(tab => ((BoardTableSheet)tab.Tag!).Name == sheet);

    private static void Select(BoardTableEditor editor, string sheet) =>
        editor.SelectSheet(editor.CommitAndGetDocument()!.FindSheet(sheet)!);

    // ###########################################################################################
    // *** A FILE PUT IN PLACE CLEARS ITS MARK WHEN THE TABLE IS LOOKED AT AGAIN (code review,
    // 2026-10-04). *** The schematic picture is missing, so its cell carries an error. The picture is
    // then copied into the draft folder with nothing edited - which the table's file watch cannot
    // see, the workbook being untouched. RecheckFiles (the Drafts tab calls it when it is shown again
    // or CRT's window comes back) runs the checks again, and the error goes - as it already had from
    // the draft's row and from Submit.
    // ###########################################################################################
    [Fact]
    public void A_missing_file_put_in_the_draft_folder_loses_its_error_when_the_files_are_checked_again()
    {
        using var workspace = new TempWorkspace();
        string draftFolder = workspace.Path_("Drafts", "Commodore", "C64", "250407");
        Directory.CreateDirectory(draftFolder);
        Directory.CreateDirectory(workspace.Path_("Data"));

        UiTest.Run(() =>
        {
            BoardData board = new() { Schematics = [Main], Components = [Component("U1")], ComponentHighlights = [Highlight("Main", "U1")] };

            var editor = new BoardTableEditor();
            editor.Open(BoardTableDocument.Create(board, board, files: new DiskFileLookup(workspace.Path_("Data"), draftFolder)));

            var window = new Window { Content = editor, Width = 1300, Height = 600 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            BoardTableDocument document = editor.CommitAndGetDocument()!;
            Assert.Equal(1, document.ErrorCount);

            workspace.WriteFile("Drafts/Commodore/C64/250407/main.png", "png");
            editor.RecheckFiles();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(0, document.ErrorCount);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A PROBLEM IS A TRIANGLE IN THE CELL'S TOP-LEFT CORNER. *** The cell's own wash is kept
    // under it - so the mark is a gradient: the level's colour to half way along the diagonal to
    // (10, 10), then the wash. On an empty cell too (the unchanged wash is transparent).
    // ###########################################################################################
    [Fact]
    public void A_cell_with_an_error_carries_a_red_corner_over_its_own_wash()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(Checked(), Checked());
            Select(editor, BoardWorkbookSchema.SheetComponentLinks);
            Dispatcher.UIThread.RunJobs();

            BoardTableSheet links = editor.CurrentSheet!;
            int url = Column(links, BoardWorkbookSchema.ColUrl);

            var brush = Assert.IsType<LinearGradientBrush>(CellOnScreen(window, links.Rows[1], url).Background);

            Assert.Equal(ThemeColor("BoardTable_Error_Mark"), brush.GradientStops[0].Color);
            Assert.Equal(0.5, brush.GradientStops[1].Offset);
            Assert.Equal(Colors.Transparent, brush.GradientStops[3].Color);
            Assert.Equal(new RelativePoint(BoardTableEditor.CornerMarkSize, BoardTableEditor.CornerMarkSize, RelativeUnit.Absolute), brush.EndPoint);

            // The good link's cell is its plain wash.
            Assert.IsNotType<LinearGradientBrush>(CellOnScreen(window, links.Rows[0], url).Background);

            window.Close();
        });
    }

    // An amber corner for a warning, over a CHANGED cell's orange - both told at once.
    [Fact]
    public void A_changed_cell_with_a_warning_keeps_its_orange_under_an_amber_corner()
    {
        UiTest.Run(() =>
        {
            BoardData published = Checked();
            BoardData draft = Checked();
            draft = new BoardData
            {
                Schematics = draft.Schematics,
                Components = [Component("U1"), Component("U2", "Changed")],
                ComponentHighlights = draft.ComponentHighlights,
                ComponentLinks = draft.ComponentLinks
            };

            (Window window, BoardTableEditor editor) = Shown(draft, published);
            Select(editor, BoardWorkbookSchema.SheetComponents);
            Dispatcher.UIThread.RunJobs();

            BoardTableSheet components = editor.CurrentSheet!;
            BoardTableRow u2 = components.Rows[1];

            // The warning is on U2's Board label; the change on its Friendly name.
            var label = Assert.IsType<LinearGradientBrush>(CellOnScreen(window, u2, Column(components, BoardWorkbookSchema.ColBoardLabel)).Background);
            Assert.Equal(ThemeColor("BoardTable_Warning_Mark"), label.GradientStops[0].Color);

            Assert.Equal(
                ThemeColor("BoardTable_Modified_Bg"),
                (CellOnScreen(window, u2, Column(components, BoardWorkbookSchema.ColFriendlyName)).Background as ISolidColorBrush)?.Color);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE COLOUR KEY COUNTS THE WHOLE DRAFT (owner decision, 2026-10-02, question E: "Added,
    // Modified and Deleted should work per system like Error and Warning"). *** All six pills, on
    // whichever sheet is on screen. They counted the sheet on screen until then - this test used to
    // pin exactly that - which read "0 Errors" on Components while the draft's row said "1 error".
    // A pill is outlined, with no fill, only when the whole draft has none of its kind.
    // ###########################################################################################
    [Fact]
    public void The_colour_key_counts_the_whole_draft_whichever_sheet_is_on_screen()
    {
        UiTest.Run(() =>
        {
            BoardData published = Checked();
            published.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U1", Name = "Old", Url = "https://example.org" });

            BoardData draft = Checked();
            draft.Components[0] = Component("U1", "CPU");
            draft.Components.Add(Component("U3"));
            draft.Components.Add(Component("U3"));
            draft.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U2", Name = "New", Url = "https://example.net" });

            (Window window, BoardTableEditor editor) = Shown(draft, published);
            BoardTableDocument document = editor.CommitAndGetDocument()!;

            var pills = new (string Count, string Pill, Func<BoardTableSheet, int> PerSheet)[]
            {
                ("AddedCountText", "AddedPill", sheet => sheet.AddedCount),
                ("ModifiedCountText", "ModifiedPill", sheet => sheet.ModifiedCount),
                ("DeletedCountText", "DeletedPill", sheet => sheet.DeletedCount),
                ("ErrorsCountText", "ErrorsPill", sheet => sheet.ErrorRowCount),
                ("WarningsCountText", "WarningsPill", sheet => sheet.WarningRowCount)
            };

            foreach (string sheetName in new[] { BoardWorkbookSchema.SheetComponents, BoardWorkbookSchema.SheetComponentLinks, BoardWorkbookSchema.SheetCredits })
            {
                Select(editor, sheetName);

                foreach ((string count, string pill, Func<BoardTableSheet, int> perSheet) in pills)
                {
                    int total = document.Sheets.Sum(perSheet);
                    Assert.True(total > 0, $"the draft has no {pill} rows - the test needs every kind");

                    Assert.Equal(total.ToString(System.Globalization.CultureInfo.InvariantCulture), editor.GetControl<TextBlock>(count).Text);
                    Assert.DoesNotContain("Empty", editor.GetControl<Border>(pill).Classes);
                }
            }

            // The Credits sheet has none of anything - or the per-sheet count would have passed.
            Assert.All(pills, pill => Assert.Equal(0, pill.PerSheet(document.FindSheet(BoardWorkbookSchema.SheetCredits)!)));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** "ERRORS" PICKED GOES TO THE ERRORS. *** From a sheet with none, the table moves to the
    // first sheet that has some, and shows only those rows - the other sheets' tabs hidden. The
    // pill is outlined while picked, and does not fade although it counted nothing where it was.
    // ###########################################################################################
    [Fact]
    public void Picking_Errors_goes_to_the_rows_with_errors()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(Checked(), Checked());
            Select(editor, BoardWorkbookSchema.SheetComponents);

            editor.TogglePill(BoardTableRowKinds.Errors);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(BoardWorkbookSchema.SheetComponentLinks, editor.CurrentSheet!.Name);
            Assert.Equal("Bad", Assert.Single(ShownRows(editor)).Cells[Column(editor.CurrentSheet, BoardWorkbookSchema.ColName)].Text);
            Assert.False(SheetTab(editor, BoardWorkbookSchema.SheetComponents).IsVisible);

            Border pill = editor.GetControl<Border>("ErrorsPill");
            Assert.Contains("Selected", pill.Classes);
            Assert.Equal(new Thickness(2), pill.BorderThickness);
            Assert.Equal(1, pill.Opacity, 3);

            // Fixed: the row has no error any more, and leaves the view.
            editor.CurrentSheet.Rows[1].Cells[Column(editor.CurrentSheet, BoardWorkbookSchema.ColUrl)].Text = "https://example.org";
            editor.RefreshPendingForTests();

            Assert.Empty(ShownRows(editor));

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A SHEET TAB SAYS NOTHING ABOUT ITS PROBLEMS (owner request, 2026-10-02: "the tabs ...
    // should not show "2 flagged" and "8 error" - the tabs will be obvious when you click those
    // badges"). *** It carried a red and an amber pill from earlier the same day; picking the
    // Errors or Warnings pill hides every tab without such rows, which says the same thing. The
    // tab with an error and the tab with a warning draw their name and nothing else.
    // ###########################################################################################
    [Fact]
    public void A_sheet_tab_with_errors_or_warnings_shows_its_name_alone()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(Checked(), Checked());

            TabItem links = SheetTab(editor, BoardWorkbookSchema.SheetComponentLinks);
            TabItem components = SheetTab(editor, BoardWorkbookSchema.SheetComponents);

            // The rows really have problems - or a tab without pills would prove nothing.
            Assert.Equal(1, ((BoardTableSheet)links.Tag!).ErrorRowCount);
            Assert.Equal(1, ((BoardTableSheet)components.Tag!).WarningRowCount);

            foreach (TabItem tab in new[] { links, components })
            {
                Assert.Equal([tab.Header?.ToString()], tab.GetVisualDescendants().OfType<TextBlock>().Select(text => text.Text));
                Assert.DoesNotContain(tab.GetVisualDescendants().OfType<Border>(), border => border.Classes.Count > 0 && border.Child is TextBlock);
            }

            window.Close();
        });
    }

    // A highlight is in no row: its problem is a line above the table, in its level's colour.
    [Fact]
    public void A_problem_in_no_row_is_a_line_above_the_table()
    {
        UiTest.Run(() =>
        {
            BoardData board = Checked();
            board = new BoardData
            {
                Schematics = board.Schematics,
                Components = board.Components,
                ComponentHighlights = [.. board.ComponentHighlights, Highlight("Renamed away", "U1")],
                ComponentLinks = board.ComponentLinks
            };

            (Window window, BoardTableEditor editor) = Shown(board, board);
            TextBlock line = editor.GetControl<TextBlock>("OutsideProblemsText");

            Assert.True(line.IsVisible);
            Assert.Contains("Renamed away", line.Text, StringComparison.Ordinal);
            Assert.Equal(ThemeColor("BoardTable_Error_Mark"), (line.Foreground as ISolidColorBrush)?.Color);

            window.Close();
        });
    }

    // ###########################################################################################
    // The Errors and Warnings pills follow the other four (owner request, 2026-10-03): filled with
    // their own pale wash while they count rows, outlined in their corner mark's colour, and with
    // no fill at all when there are none.
    // ###########################################################################################
    [Fact]
    public void The_Errors_and_Warnings_pills_are_filled_while_they_count_rows_and_outlined_when_not()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(Checked(), Checked());

            Border errors = editor.GetControl<Border>("ErrorsPill");
            Border warnings = editor.GetControl<Border>("WarningsPill");

            Assert.Equal(ThemeColor("BoardTable_Error_Bg"), (errors.Background as ISolidColorBrush)?.Color);
            Assert.Equal(ThemeColor("BoardTable_Error_Mark"), (errors.BorderBrush as ISolidColorBrush)?.Color);
            Assert.Equal(ThemeColor("BoardTable_Warning_Bg"), (warnings.Background as ISolidColorBrush)?.Color);
            Assert.Equal(ThemeColor("BoardTable_Warning_Mark"), (warnings.BorderBrush as ISolidColorBrush)?.Color);

            // The bad link fixed: no error left, so the Errors pill is the outline alone.
            BoardTableSheet links = editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponentLinks)!;
            links.Rows[1].Cells[Column(links, BoardWorkbookSchema.ColUrl)].Text = "https://example.org";
            editor.RefreshPendingForTests();
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(ThemeColor("Table_Bg"), (errors.Background as ISolidColorBrush)?.Color);
            Assert.Equal(ThemeColor("BoardTable_Error_Mark"), (errors.BorderBrush as ISolidColorBrush)?.Color);

            window.Close();
        });
    }

    // Every new key the checks paint with resolves in both themes - a missing DynamicResource
    // draws nothing, silently.
    [Theory]
    [InlineData("BoardTable_Error_Mark")]
    [InlineData("BoardTable_Error_Fg")]
    [InlineData("BoardTable_Warning_Mark")]
    [InlineData("BoardTable_Warning_Fg")]
    [InlineData("BoardTable_Pill_Selected_Border")]
    [InlineData("BoardTable_Error_Bg")]
    [InlineData("BoardTable_Warning_Bg")]
    [InlineData("BoardTable_Added_Edge")]
    [InlineData("BoardTable_Modified_Edge")]
    [InlineData("BoardTable_Deleted_Edge")]
    public void The_checks_colours_are_in_both_themes(string key)
    {
        UiTest.Run(() =>
        {
            Application app = Application.Current!;

            Assert.True(app.TryGetResource(key, Avalonia.Styling.ThemeVariant.Light, out _), $"{key} is missing from the light theme");
            Assert.True(app.TryGetResource(key, Avalonia.Styling.ThemeVariant.Dark, out _), $"{key} is missing from the dark theme");
        });
    }
}
