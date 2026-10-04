using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE TABLE'S SEARCH BOX (owner request, 2026-10-02: "how about the same search/filter as we have
// in the "Workbooks" tab, where it highlights what I search for"). Which rows it shows is
// BoardTableSearch, pinned in CRT.Data.Tests; these pin the box and what the grid paints - the
// rows on screen, the tabs, and the found text marked in the Workbooks tab's own search colour.
//
// Each test is one case agreed before the code was written, numbered as it was then; C and D were
// questions, answered "your assumption is correct" and "yes".
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardTableEditorSearchTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private static ComponentImageEntry Image(string label, string name = "Pinout", string? file = null) =>
        new() { BoardLabel = label, Name = name, File = file ?? $"a/{label.ToLowerInvariant()}.png" };

    private static BoardData Board(string clockFile = "a/u8.png", string cr13File = "a/cr13.png") => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U8", FriendlyName = "RAM", TechnicalNameOrValue = "4164" },
            new ComponentEntry { BoardLabel = "U9", FriendlyName = "ROM", TechnicalNameOrValue = "x" }
        ],
        ComponentImages =
        [
            Image("CR13", file: cr13File),
            Image("CR17"),
            Image("R307", "Pinout (secondary)"),
            Image("U8", "Clock", clockFile)
        ]
    };

    private static (Window Window, BoardTableEditor Editor) Shown(BoardData draft, BoardData? published, string sheet = BoardWorkbookSchema.SheetComponentImages)
    {
        var editor = new BoardTableEditor();
        BoardTableDocument document = BoardTableDocument.Create(published, draft);
        editor.Open(document);
        editor.SelectSheet(document.FindSheet(sheet)!);

        var window = new Window { Content = editor, Width = 1600, Height = 700 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return (window, editor);
    }

    private static TextBox SearchBox(BoardTableEditor editor) => editor.GetControl<TextBox>("SearchBox");

    // Types into the box as a contributor would, then lets the short pause after typing pass.
    private static void Search(Window window, BoardTableEditor editor, string text)
    {
        TextBox box = SearchBox(editor);
        box.Focus();
        box.Text = string.Empty;

        if (text.Length > 0)
            window.KeyTextInput(text);

        Dispatcher.UIThread.RunJobs();
        editor.ApplyPendingSearchForTests();
        Settle();
    }

    private static void Settle()
    {
        for (int i = 0; i < 3; i++)
        {
            Dispatcher.UIThread.RunJobs();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        }
    }

    private static string Label(BoardTableRow row) =>
        row.Cells[row.Sheet.Columns.ToList().IndexOf(BoardWorkbookSchema.ColBoardLabel)].Text;

    private static List<string> ShownLabels(BoardTableEditor editor) =>
        ((System.Collections.IEnumerable)editor.GetControl<DataGrid>("TableGrid").ItemsSource!).Cast<BoardTableRow>()
            .Select(Label)
            .ToList();

    private static List<string> ShownSheetTabs(BoardTableEditor editor) =>
        editor.GetControl<TabControl>("SheetTabs").Items.OfType<TabItem>()
            .Where(tab => tab.IsVisible)
            .Select(tab => ((BoardTableSheet)tab.Tag!).Name)
            .ToList();

    private static int ColumnOf(BoardTableEditor editor, string column) =>
        editor.CurrentSheet!.Columns.ToList().IndexOf(column);

    private static DataGridCell CellOnScreen(Window window, BoardTableRow row, int columnIndex) =>
        window.GetVisualDescendants()
            .OfType<DataGridCell>()
            .Single(cell => cell.IsEffectivelyVisible && ReferenceEquals(cell.DataContext, row) && cell.OwningColumn?.Tag is int tag && tag == columnIndex);

    // The runs of a cell's text drawn with a background - the marked ones - as "text:colour".
    private static List<string> MarkedRuns(DataGridCell cell) =>
        cell.GetVisualDescendants().OfType<TextBlock>()
            .SelectMany(text => text.Inlines?.OfType<Run>() ?? [])
            .Where(run => run.Background is ISolidColorBrush)
            .Select(run => $"{run.Text}:{((ISolidColorBrush)run.Background!).Color}")
            .ToList();

    private static Color ThemeColor(string key)
    {
        Application app = Application.Current!;
        Assert.True(app.TryGetResource(key, app.ActualThemeVariant, out object? resource));
        return ((ISolidColorBrush)resource!).Color;
    }

    private static BoardTableRow Row(BoardTableEditor editor, string label) =>
        editor.CurrentSheet!.Rows.First(row => !row.IsDeleted && Label(row) == label);

    // Case 15: typing shows only the matching rows, and the text found is marked in the Workbooks
    // tab's search colour - in that cell, in no other, and without hiding the cell's own colour.
    [Fact]
    public void Typing_in_the_search_box_shows_only_matching_rows_and_marks_the_found_text_in_the_search_colour()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(Board(), Board(cr13File: "a/old.png"));
            int name = ColumnOf(editor, BoardWorkbookSchema.ColName);
            int file = ColumnOf(editor, BoardWorkbookSchema.ColFile);
            Color? fileBefore = (CellOnScreen(window, Row(editor, "CR13"), file).Background as ISolidColorBrush)?.Color;
            Color? nameBefore = (CellOnScreen(window, Row(editor, "CR13"), name).Background as ISolidColorBrush)?.Color;

            Search(window, editor, "pinout");

            Assert.Equal(["CR13", "CR17", "R307"], ShownLabels(editor));

            DataGridCell cr13Name = CellOnScreen(window, Row(editor, "CR13"), name);
            Assert.Equal([$"Pinout:{ThemeColor("Workbooks_SearchHit_Bg")}"], MarkedRuns(cr13Name));
            Assert.Equal(
                ThemeColor("Workbooks_SearchHit_Fg"),
                (cr13Name.GetVisualDescendants().OfType<TextBlock>().SelectMany(text => text.Inlines?.OfType<Run>() ?? []).Single(run => run.Text == "Pinout").Foreground as ISolidColorBrush)?.Color);

            Assert.Equal([$"Pinout:{ThemeColor("Workbooks_SearchHit_Bg")}"], MarkedRuns(CellOnScreen(window, Row(editor, "R307"), name)));
            Assert.Empty(MarkedRuns(CellOnScreen(window, Row(editor, "CR13"), ColumnOf(editor, BoardWorkbookSchema.ColBoardLabel))));

            // Each cell keeps its own colour - the changed File cell its orange: the grid's own
            // "match" fill does not cover it.
            Assert.Equal(fileBefore, (CellOnScreen(window, Row(editor, "CR13"), file).Background as ISolidColorBrush)?.Color);
            Assert.Equal(nameBefore, (cr13Name.Background as ISolidColorBrush)?.Color);

            window.Close();
        });
    }

    // ###########################################################################################
    // Case 15, "and nothing else is": in a cell with a match, only the text found takes the mark's
    // colours - the rest of the cell's text keeps the ordinary text colour. Seen in a dark render
    // (2026-10-02): the grid gave the whole cell the mark's dark text colour, so "(secondary)"
    // beside a marked "Pinout" was dark grey on violet. In the light theme the two colours differ
    // only slightly (#1A1A1A against black), which is enough for this test to tell.
    // ###########################################################################################
    [Fact]
    public void In_a_matching_cell_only_the_text_found_takes_the_marks_colours()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(Board(), Board());
            int name = ColumnOf(editor, BoardWorkbookSchema.ColName);
            int file = ColumnOf(editor, BoardWorkbookSchema.ColFile);

            Search(window, editor, "pinout");

            static Color? ForegroundOf(Run run) => (run.Foreground as ISolidColorBrush)?.Color;

            List<Run> runs = CellOnScreen(window, Row(editor, "R307"), name)
                .GetVisualDescendants().OfType<TextBlock>().SelectMany(text => text.Inlines?.OfType<Run>() ?? []).ToList();
            Color? ordinary = (CellOnScreen(window, Row(editor, "R307"), file).GetVisualDescendants().OfType<TextBlock>().First().Foreground as ISolidColorBrush)?.Color;

            Assert.NotEqual(ThemeColor("Workbooks_SearchHit_Fg"), ordinary);
            Assert.Equal(ThemeColor("Workbooks_SearchHit_Fg"), ForegroundOf(runs.Single(run => run.Text == "Pinout")));
            Assert.Equal(ordinary, ForegroundOf(runs.Single(run => run.Text == " (secondary)")));

            window.Close();
        });
    }

    // Case 15: the box's clear button empties it - every row is back and nothing is marked.
    [Fact]
    public void Clearing_the_search_shows_every_row_and_marks_nothing()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(Board(), Board());
            Search(window, editor, "pinout");

            Button clear = editor.GetControl<Button>("ClearSearchButton");
            Assert.True(clear.IsVisible);

            Point point = clear.TranslatePoint(new Point(clear.Bounds.Width / 2, clear.Bounds.Height / 2), window)!.Value;
            window.MouseDown(point, MouseButton.Left);
            window.MouseUp(point, MouseButton.Left);
            editor.ApplyPendingSearchForTests();
            Settle();

            Assert.Equal(string.Empty, SearchBox(editor).Text ?? string.Empty);
            Assert.False(clear.IsVisible);
            Assert.Equal(["CR13", "CR17", "R307", "U8"], ShownLabels(editor));
            Assert.Empty(MarkedRuns(CellOnScreen(window, Row(editor, "CR13"), ColumnOf(editor, BoardWorkbookSchema.ColName))));

            window.Close();
        });
    }

    // Case 16: a search and a picked pill together show only the rows both show.
    [Fact]
    public void A_search_and_a_picked_pill_together_show_only_rows_both_show()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(Board(), Board(clockFile: "a/old.png", cr13File: "a/old.png"));

            editor.TogglePill(BoardTableRowKinds.Modified);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["CR13", "U8"], ShownLabels(editor));

            Search(window, editor, "pinout");
            Assert.Equal(["CR13"], ShownLabels(editor));

            window.Close();
        });
    }

    // Case 17: on a sheet with no match, the table moves to the first sheet with one, and the tabs of
    // sheets without a match are hidden.
    [Fact]
    public void The_table_moves_to_the_first_sheet_with_a_match_and_hides_the_tabs_of_sheets_without()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(Board(), Board(), BoardWorkbookSchema.SheetComponents);

            Search(window, editor, "pinout");

            Assert.Equal(BoardWorkbookSchema.SheetComponentImages, editor.CurrentSheet!.Name);
            Assert.Equal([BoardWorkbookSchema.SheetComponentImages], ShownSheetTabs(editor));

            // "u8" is on both sheets.
            Search(window, editor, "u8");
            Assert.Equal([BoardWorkbookSchema.SheetComponents, BoardWorkbookSchema.SheetComponentImages], ShownSheetTabs(editor));

            window.Close();
        });
    }

    // Case 20: rows cannot be moved while a search is on, as with a pill picked - a position among
    // hidden rows means nothing.
    [Fact]
    public void Rows_cannot_be_moved_while_a_search_is_on()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(Board(), Board());
            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");
            Assert.Contains("RowsDraggable", grid.Classes);

            Search(window, editor, "pinout");
            Assert.DoesNotContain("RowsDraggable", grid.Classes);

            editor.SelectCell(Row(editor, "CR13"), 0);
            editor.MoveRowDown();
            Assert.Equal(["CR13", "CR17", "R307", "U8"], editor.CurrentSheet!.Rows.Select(Label));

            Search(window, editor, "");
            Assert.Contains("RowsDraggable", grid.Classes);

            window.Close();
        });
    }

    // Case D (answered "yes"): the search stays when switching sheet tabs, and when leaving CRT's tab
    // and coming back (the table is detached and attached again).
    [Fact]
    public void The_search_stays_when_switching_sheets_and_leaving_the_tab_and_coming_back()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(Board(), Board());
            Search(window, editor, "u8");
            Assert.Equal(["U8"], ShownLabels(editor));

            editor.SelectSheet(editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!);
            Settle();
            Assert.Equal(["U8"], ShownLabels(editor));

            window.Content = null;
            Settle();
            window.Content = editor;
            Settle();

            Assert.Equal("u8", SearchBox(editor).Text);
            Assert.Equal(["U8"], ShownLabels(editor));

            window.Close();
        });
    }

    // Case D: closing the table empties the search, and so does opening another draft - while a
    // save, which reads the same draft back, keeps it.
    private const string C64 = "Commodore/C64/250407/Data C64 250407.xlsx";
    private const string C128 = "Commodore/C128/310378/Data C128 310378.xlsx";

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private void WriteDraft(string systemKey, BoardData board)
    {
        string path = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, systemKey);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        CachedWorkbooks.Write(path, board);
    }

    [Fact]
    public void Closing_the_table_or_opening_another_draft_empties_the_search_and_a_save_keeps_it()
    {
        UiTest.Run(() =>
        {
            this.WriteDraft(C64, Board());
            this.WriteDraft(C128, Board());

            var editor = new BoardTableEditor();
            var window = new Window { Content = editor, Width = 1600, Height = 700 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            // A save reads the draft back - the search stays, and still narrows the rows.
            Assert.True(editor.Load(this.DraftsRoot, C64, published: Board()));
            editor.SelectSheet(editor.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponentImages)!);
            Search(window, editor, "pinout");
            Row(editor, "CR17").Cells[ColumnOf(editor, BoardWorkbookSchema.ColNote)].Text = "checked";
            editor.RefreshPendingForTests();
            Assert.Equal(DraftWorkbookEditOutcome.Saved, editor.Save());
            Settle();

            Assert.Equal("pinout", SearchBox(editor).Text);
            Assert.Equal(["CR13", "CR17", "R307"], ShownLabels(editor));

            // Another draft: an empty search.
            Assert.True(editor.Load(this.DraftsRoot, C128, published: Board()));
            Settle();
            Assert.Equal(string.Empty, SearchBox(editor).Text ?? string.Empty);
            Assert.Equal(string.Empty, editor.SearchText);

            // Closing the table: an empty search too.
            Search(window, editor, "pinout");
            editor.Clear();
            Settle();
            Assert.Equal(string.Empty, SearchBox(editor).Text ?? string.Empty);
            Assert.Equal(string.Empty, editor.SearchText);

            window.Close();
        });
    }

    // ###########################################################################################
    // A SEARCH ACROSS EVERY SHEET (owner request, 2026-10-03: "it should work in all sheets, only
    // showing sheets with filtered data"). Cases 1 and 2 were already so; a search matching nothing
    // anywhere kept every tab, each sheet empty, which read as a search of the one sheet on screen.
    // Numbered as agreed that day, so "Case 3 (2026-10-03)" is not the earlier rounds' case 3.
    // ###########################################################################################

    // U1 on the four component sheets and nowhere else; "Pinout" only on Component images.
    private static BoardData U1Board() => new()
    {
        Schematics = [new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = "a/main.png" }],
        Components =
        [
            new ComponentEntry { BoardLabel = "U1", FriendlyName = "CIA" },
            new ComponentEntry { BoardLabel = "C5", FriendlyName = "Ceramic" }
        ],
        ComponentImages = [Image("U1"), Image("CR13", file: "a/cr13.png")],
        ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U1", Name = "Datasheet", File = "a/6526.pdf" }],
        ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U1", Name = "Wiki", Url = "https://example.com/cia" }],
        BoardLinks = [new BoardLinkEntry { Category = "Manuals", Name = "Service manual", Url = "https://example.com/manual" }],
        Credits = [new CreditEntry { Category = "Repair", SubCategory = "Photos", NameOrHandle = "Someone" }]
    };

    // Case 1 (2026-10-03): on Board schematics, "u1" moves the table to Components and leaves the
    // tabs of the four sheets with a match.
    [Fact]
    public void A_search_moves_to_the_first_sheet_with_a_match_and_shows_only_the_sheets_with_one()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(U1Board(), U1Board(), BoardWorkbookSchema.SheetBoardSchematics);

            Search(window, editor, "u1");

            Assert.Equal(BoardWorkbookSchema.SheetComponents, editor.CurrentSheet!.Name);
            Assert.Equal(
                [BoardWorkbookSchema.SheetComponents, BoardWorkbookSchema.SheetComponentImages, BoardWorkbookSchema.SheetComponentLocalFiles, BoardWorkbookSchema.SheetComponentLinks],
                ShownSheetTabs(editor));

            window.Close();
        });
    }

    // Case 2 (2026-10-03): "pinout" leaves Component images alone.
    [Fact]
    public void A_search_found_on_one_sheet_only_leaves_that_sheets_tab_alone()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(U1Board(), U1Board(), BoardWorkbookSchema.SheetComponents);

            Search(window, editor, "pinout");

            Assert.Equal(BoardWorkbookSchema.SheetComponentImages, editor.CurrentSheet!.Name);
            Assert.Equal([BoardWorkbookSchema.SheetComponentImages], ShownSheetTabs(editor));

            window.Close();
        });
    }

    // Case 3 (2026-10-03): "zzz", found nowhere, keeps only the sheet on screen - empty - and the
    // line above the table says why.
    [Fact]
    public void A_search_matching_nothing_anywhere_keeps_only_the_sheet_on_screen_and_says_so()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(U1Board(), U1Board(), BoardWorkbookSchema.SheetComponentImages);

            Search(window, editor, "zzz");

            Assert.Equal(BoardWorkbookSchema.SheetComponentImages, editor.CurrentSheet!.Name);
            Assert.Equal([BoardWorkbookSchema.SheetComponentImages], ShownSheetTabs(editor));
            Assert.Empty(ShownLabels(editor));

            TextBlock line = editor.GetControl<TextBlock>("NoSearchMatchText");
            Assert.True(line.IsVisible);
            Assert.Equal("Nothing in any sheet matches \"zzz\"", line.Text);

            window.Close();
        });
    }

    // Case 4 (2026-10-03): clearing that search brings every tab back, on the same sheet, and the
    // line goes.
    [Fact]
    public void Clearing_a_search_that_matched_nothing_brings_every_tab_back_on_the_same_sheet()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = Shown(U1Board(), U1Board(), BoardWorkbookSchema.SheetComponentImages);
            Search(window, editor, "zzz");
            Assert.Equal([BoardWorkbookSchema.SheetComponentImages], ShownSheetTabs(editor));

            Button clear = editor.GetControl<Button>("ClearSearchButton");
            Point point = clear.TranslatePoint(new Point(clear.Bounds.Width / 2, clear.Bounds.Height / 2), window)!.Value;
            window.MouseDown(point, MouseButton.Left);
            window.MouseUp(point, MouseButton.Left);
            editor.ApplyPendingSearchForTests();
            Settle();

            Assert.Equal(BoardWorkbookSchema.SheetComponentImages, editor.CurrentSheet!.Name);
            Assert.Equal(editor.CommitAndGetDocument()!.Sheets.Select(sheet => sheet.Name), ShownSheetTabs(editor));
            Assert.False(editor.GetControl<TextBlock>("NoSearchMatchText").IsVisible);
            Assert.Equal(["U1", "CR13"], ShownLabels(editor));

            window.Close();
        });
    }

    // Case 5 (2026-10-03): "u1" with "Deleted" picked, where the one deleted row is CR17 - no row
    // matches both, so it is case 3 again, and the line names the picked counts too.
    [Fact]
    public void A_search_and_a_picked_pill_with_no_row_in_common_keep_only_the_sheet_on_screen_and_say_so()
    {
        UiTest.Run(() =>
        {
            BoardData published = U1Board();
            published.ComponentImages.Add(Image("CR17", file: "a/cr17.png"));

            (Window window, BoardTableEditor editor) = Shown(U1Board(), published, BoardWorkbookSchema.SheetComponents);
            editor.TogglePill(BoardTableRowKinds.Deleted);
            Settle();
            Assert.Equal(BoardWorkbookSchema.SheetComponentImages, editor.CurrentSheet!.Name);

            Search(window, editor, "u1");

            Assert.Equal(BoardWorkbookSchema.SheetComponentImages, editor.CurrentSheet!.Name);
            Assert.Equal([BoardWorkbookSchema.SheetComponentImages], ShownSheetTabs(editor));
            Assert.Empty(ShownLabels(editor));

            TextBlock line = editor.GetControl<TextBlock>("NoSearchMatchText");
            Assert.True(line.IsVisible);
            Assert.Equal("Nothing in any sheet matches both \"u1\" and the counts picked", line.Text);

            window.Close();
        });
    }
}
