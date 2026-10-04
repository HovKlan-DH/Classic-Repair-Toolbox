using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Interactivity;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The table's file hover card (owner request, 2026-09-26): resting on a file cell shows the
// picture - the published and the new one side by side when it changed - or, for a PDF, a link
// that opens it. Shared by the Drafts tab and the Maintainer tab (Controls/BoardTable), with the bytes
// from each host's IBoardTableFileSource; a fake one here.
//
// WHICH file each side is, is pinned in CRT.Data.Tests' BoardTableFileCellsTests. This pins what
// the card SHOWS for it: the sides and their labels, the picture collapsing to one when the bytes
// are identical, the link opening the right file - and, with a real pointer in a shown window, that
// the card opens and closes AT ONCE as the pointer comes and goes (owner request, 2026-09-26),
// follows the pointer down a column, and lets the pointer onto itself so a link can be clicked.
// Headless drawing decodes no pixels, so a picture here proves a bitmap was made, not what it
// looks like.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardTableFilePreviewTests
{
    // Two different real one-pixel PNGs.
    private static readonly byte[] PixelA = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    private static readonly byte[] PixelB = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    private sealed class FakeFileSource : IBoardTableFileSource
    {
        public Dictionary<(string Path, BoardTableFileSide Side), byte[]> Files { get; } = new();

        public List<(string Path, BoardTableFileSide Side)> Opened { get; } = new();

        // When set, every read waits for it - a slow disk or a slow server.
        public TaskCompletionSource? Gate { get; set; }

        public string PublishedLabel => "Published";

        public string CurrentLabel => "Your draft";

        public bool SaysUnchanged { get; set; } = true;

        public async Task<byte[]?> ReadAsync(string path, BoardTableFileSide side)
        {
            if (this.Gate is not null)
            {
                await this.Gate.Task;
            }

            return this.Files.TryGetValue((path, side), out byte[]? bytes) ? bytes : null;
        }

        public Task<string?> OpenAsync(string path, BoardTableFileSide side)
        {
            this.Opened.Add((path, side));
            return Task.FromResult<string?>(null);
        }
    }

    private static ComponentImageEntry Image(string file) =>
        new() { BoardLabel = "U8", Region = "PAL", Pin = "1", Name = "Pinout", File = file };

    private static int ImageFileColumn =>
        BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFile);

    // The editor in document mode (the Maintainer tab's way - no draft file needed) on the
    // Component images sheet.
    private static BoardTableEditor Open(BoardData? published, BoardData draft, IBoardTableFileSource? source, string sheetName = BoardWorkbookSchema.SheetComponentImages)
    {
        var editor = new BoardTableEditor();
        BoardTableDocument document = BoardTableDocument.Create(published, draft);

        editor.Open(document);
        editor.FileSource = source;
        editor.SelectSheet(document.FindSheet(sheetName)!);

        return editor;
    }

    private static BoardTableRow OnlyRow(BoardTableEditor editor) => editor.CurrentSheet!.Rows.Single();

    private static IReadOnlyList<StackPanel> VisibleSides(BoardTableFilePreview preview) =>
        preview.Sides.OfType<StackPanel>().Where(side => side.IsVisible).ToList();

    private static string Label(StackPanel side) => ((TextBlock)side.Children[0]).Text!;

    private static bool LabelShown(StackPanel side) => side.Children[0].IsVisible;

    private static Image? PictureOf(StackPanel side) => side.Children.OfType<Image>().SingleOrDefault();

    private static string StatusOf(StackPanel side) =>
        side.Children.OfType<TextBlock>().Skip(2).FirstOrDefault(text => text.IsVisible)?.Text ?? string.Empty;

    private static Button LinkOf(StackPanel side) => side.Children.OfType<Button>().Single();

    private static string LinkText(Button link) => ((TextBlock)link.Content!).Text!;

    // ###########################################################################################
    // *** THE CASE ASKED FOR BY NAME. *** A picture changed to another file: the published one and
    // the new one, side by side, each labelled - an unlabelled pair of near-identical board scans
    // leaves the reader guessing which is the proposal.
    // ###########################################################################################
    [Fact]
    public async Task A_changed_picture_shows_the_published_and_the_new_one_side_by_side()
    {
        var source = new FakeFileSource();
        source.Files[("Commodore/C64/250407/old.png", BoardTableFileSide.Published)] = PixelA;
        source.Files[("Commodore/C64/250407/new.png", BoardTableFileSide.Current)] = PixelB;

        await UiTest.RunAsync(async () =>
        {
            BoardTableEditor editor = Open(
                new BoardData { ComponentImages = [Image("Commodore/C64/250407/old.png")] },
                new BoardData { ComponentImages = [Image("Commodore/C64/250407/new.png")] },
                source);

            BoardTableFilePreview preview = editor.BuildFilePreview(OnlyRow(editor), ImageFileColumn)!;
            await preview.Loading;

            IReadOnlyList<StackPanel> sides = VisibleSides(preview);

            Assert.Equal(new[] { "Published", "Your draft" }, sides.Select(Label));
            Assert.All(sides, side => Assert.True(LabelShown(side)));
            Assert.All(sides, side => Assert.NotNull(PictureOf(side)!.Source));
            Assert.Equal("Changed to another file", preview.HeadlineText);
        });
    }

    // ###########################################################################################
    // *** REPLACED UNDER ITS OWN NAME. *** The cell's text is unchanged, so the table does not colour
    // it - and it is the change most worth seeing. The bytes differ, so both pictures show.
    // ###########################################################################################
    [Fact]
    public async Task A_picture_replaced_under_the_same_name_shows_both()
    {
        const string Path = "Commodore/C64/250407/u8.png";

        var source = new FakeFileSource();
        source.Files[(Path, BoardTableFileSide.Published)] = PixelA;
        source.Files[(Path, BoardTableFileSide.Current)] = PixelB;

        await UiTest.RunAsync(async () =>
        {
            BoardTableEditor editor = Open(
                new BoardData { ComponentImages = [Image(Path)] },
                new BoardData { ComponentImages = [Image(Path)] },
                source);

            BoardTableFilePreview preview = editor.BuildFilePreview(OnlyRow(editor), ImageFileColumn)!;
            await preview.Loading;

            Assert.Equal(2, VisibleSides(preview).Count);
            Assert.Equal("Replaced - same file name, new content", preview.HeadlineText);
        });
    }

    // Identical bytes: one picture is the whole story.
    [Fact]
    public async Task An_unchanged_picture_shows_once()
    {
        const string Path = "Commodore/C64/250407/u8.png";

        var source = new FakeFileSource();
        source.Files[(Path, BoardTableFileSide.Published)] = PixelA;
        source.Files[(Path, BoardTableFileSide.Current)] = PixelA;

        await UiTest.RunAsync(async () =>
        {
            BoardTableEditor editor = Open(
                new BoardData { ComponentImages = [Image(Path)] },
                new BoardData { ComponentImages = [Image(Path)] },
                source);

            BoardTableFilePreview preview = editor.BuildFilePreview(OnlyRow(editor), ImageFileColumn)!;
            await preview.Loading;

            // The draft's side is the one left - and, on its own, unlabelled.
            StackPanel side = Assert.Single(VisibleSides(preview));
            Assert.Equal("Your draft", Label(side));
            Assert.False(LabelShown(side));
            Assert.Equal("Unchanged", preview.HeadlineText);
        });
    }

    // ###########################################################################################
    // *** A HOST CAN LEAVE "UNCHANGED" UNSAID (owner request, 2026-09-26). *** The maintainer
    // application does for a new system, whose two sides are both the submission: the picture
    // alone, no headline. Anything else keeps its headline.
    // ###########################################################################################
    [Fact]
    public async Task A_host_that_does_not_say_unchanged_shows_the_picture_alone()
    {
        const string Path = "Manu1/Hardware1/Board1/1N754A.png";

        var source = new FakeFileSource { SaysUnchanged = false };
        source.Files[(Path, BoardTableFileSide.Published)] = PixelA;
        source.Files[(Path, BoardTableFileSide.Current)] = PixelA;
        source.Files[("Manu1/Hardware1/Board1/other.png", BoardTableFileSide.Current)] = PixelB;

        await UiTest.RunAsync(async () =>
        {
            BoardTableEditor editor = Open(
                new BoardData { ComponentImages = [Image(Path)] },
                new BoardData { ComponentImages = [Image(Path)] },
                source);

            BoardTableFilePreview preview = editor.BuildFilePreview(OnlyRow(editor), ImageFileColumn)!;
            await preview.Loading;

            Assert.Single(VisibleSides(preview));
            Assert.Equal(string.Empty, preview.HeadlineText);

            // Another file named in the cell is still said to be one.
            OnlyRow(editor).Cells[ImageFileColumn].Text = "Manu1/Hardware1/Board1/other.png";

            BoardTableFilePreview changed = editor.BuildFilePreview(OnlyRow(editor), ImageFileColumn)!;
            await changed.Loading;

            Assert.Equal("Changed to another file", changed.HeadlineText);
        });
    }

    // ###########################################################################################
    // *** A PDF GETS A LINK THAT OPENS IT. *** It cannot be drawn, so the card offers to open it -
    // the host's viewer, through the host's source, for the side it belongs to.
    // ###########################################################################################
    [Fact]
    public async Task A_PDF_gets_a_link_that_opens_it()
    {
        const string Path = "Commodore/C64/250407/Service manual.pdf";
        var source = new FakeFileSource();

        await UiTest.RunAsync(async () =>
        {
            BoardTableEditor editor = Open(
                null,
                new BoardData { BoardLocalFiles = [new BoardLocalFileEntry { Category = "Manuals", Name = "Service", File = Path }] },
                source,
                BoardWorkbookSchema.SheetBoardLocalFiles);

            int fileColumn = BoardWorkbookSchema.BoardLocalFiles.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFile);

            BoardTableFilePreview preview = editor.BuildFilePreview(OnlyRow(editor), fileColumn)!;
            await preview.Loading;

            StackPanel side = Assert.Single(VisibleSides(preview));
            Assert.Null(PictureOf(side));

            Button link = LinkOf(side);
            Assert.Equal("Open PDF file", LinkText(link));

            link.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(new[] { (Path, BoardTableFileSide.Current) }, source.Opened);
        });
    }

    // ###########################################################################################
    // A picture on its own - removed, or new - carries no "Published" / "Your draft" label (owner
    // request, 2026-09-26): the headline says what happened, and which side it is is obvious. Only
    // two pictures side by side are labelled, to say which is which.
    // ###########################################################################################
    [Fact]
    public async Task A_removed_or_new_picture_on_its_own_is_not_labelled()
    {
        var source = new FakeFileSource();
        source.Files[("Commodore/C64/250407/old.png", BoardTableFileSide.Published)] = PixelA;
        source.Files[("Commodore/C64/250407/new.png", BoardTableFileSide.Current)] = PixelB;

        await UiTest.RunAsync(async () =>
        {
            BoardTableEditor editor = Open(
                new BoardData { ComponentImages = [Image("Commodore/C64/250407/old.png")] },
                new BoardData { ComponentImages = [new ComponentImageEntry { BoardLabel = "U9", Region = "PAL", Pin = "1", Name = "Pinout", File = "Commodore/C64/250407/new.png" }] },
                source);

            BoardTableRow removed = editor.CurrentSheet!.Rows.Single(row => row.IsDeleted);
            BoardTableRow added = editor.CurrentSheet!.Rows.Single(row => !row.IsDeleted);

            foreach ((BoardTableRow row, string headline) in new[] { (removed, "Removed"), (added, "New file") })
            {
                BoardTableFilePreview preview = editor.BuildFilePreview(row, ImageFileColumn)!;
                await preview.Loading;

                StackPanel side = Assert.Single(VisibleSides(preview));
                Assert.False(LabelShown(side));
                Assert.NotNull(PictureOf(side)!.Source);
                Assert.Equal(headline, preview.HeadlineText);
            }
        });
    }

    // A picture gets a link too - to see it full size.
    [Fact]
    public async Task A_picture_links_to_its_full_size()
    {
        var source = new FakeFileSource();
        source.Files[("Commodore/C64/250407/new.png", BoardTableFileSide.Current)] = PixelA;

        await UiTest.RunAsync(async () =>
        {
            BoardTableEditor editor = Open(null, new BoardData { ComponentImages = [Image("Commodore/C64/250407/new.png")] }, source);

            BoardTableFilePreview preview = editor.BuildFilePreview(OnlyRow(editor), ImageFileColumn)!;
            await preview.Loading;

            Assert.Equal("Open full size", LinkText(LinkOf(Assert.Single(VisibleSides(preview)))));
        });
    }

    // A path with no file behind it says so, rather than showing an empty frame.
    [Fact]
    public async Task A_missing_file_says_so()
    {
        await UiTest.RunAsync(async () =>
        {
            BoardTableEditor editor = Open(null, new BoardData { ComponentImages = [Image("Commodore/C64/250407/gone.png")] }, new FakeFileSource());

            BoardTableFilePreview preview = editor.BuildFilePreview(OnlyRow(editor), ImageFileColumn)!;
            await preview.Loading;

            StackPanel side = Assert.Single(VisibleSides(preview));
            Assert.Null(PictureOf(side)!.Source);
            Assert.Equal("There is no file at this path.", StatusOf(side));

            // Nothing to open, so no link offering to.
            Assert.False(LinkOf(side).IsVisible);
        });
    }

    // No source, no card - a host that supplies no files keeps the table as it was. And only a file
    // cell gets one.
    [Fact]
    public void Without_a_source_or_off_a_file_cell_there_is_no_card()
    {
        UiTest.Run(() =>
        {
            BoardData draft = new() { ComponentImages = [Image("Commodore/C64/250407/new.png")] };

            BoardTableEditor withoutSource = Open(null, draft, source: null);
            Assert.Null(withoutSource.BuildFilePreview(OnlyRow(withoutSource), ImageFileColumn));

            BoardTableEditor withSource = Open(null, draft, new FakeFileSource());
            int nameColumn = BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName);
            Assert.Null(withSource.BuildFilePreview(OnlyRow(withSource), nameColumn));
        });
    }

    // ###########################################################################################
    // A file column's text tooltip ("Published value: ...") gives way to the card - two popups over
    // one cell would cover each other. Every other column keeps it.
    //
    // EXCEPT what is wrong with the file
    // (2026-10-02): a file that is not there, or spelled differently, is a file cell's error, and
    // the card cannot say it. So a file cell's tooltip is its problems alone (ProblemToolTip, null
    // with none - no box at all on a good file), and every other cell keeps its full tooltip.
    // ###########################################################################################
    [Fact]
    public void A_file_column_gives_its_text_tooltip_to_the_card_but_keeps_its_problems()
    {
        UiTest.Run(() =>
        {
            BoardData draft = new() { ComponentImages = [Image("Commodore/C64/250407/new.png")] };
            int nameColumn = BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName);

            BoardTableEditor withSource = Open(null, draft, new FakeFileSource());
            Assert.Equal($"Cells[{ImageFileColumn}].ProblemToolTip", ToolTipPath(withSource, ImageFileColumn));
            Assert.Equal($"Cells[{nameColumn}].ToolTip", ToolTipPath(withSource, nameColumn));

            BoardTableEditor withoutSource = Open(null, draft, source: null);
            Assert.Equal($"Cells[{ImageFileColumn}].ToolTip", ToolTipPath(withoutSource, ImageFileColumn));
        });
    }

    private static string? ToolTipPath(BoardTableEditor editor, int columnIndex)
    {
        DataGridColumn column = editor.GetControl<DataGrid>("TableGrid").Columns
            .Single(candidate => candidate.Tag is int tag && tag == columnIndex);

        ControlTheme theme = ((DataGridBoundColumn)column).CellTheme!;

        return theme.Setters.OfType<Setter>()
            .Where(setter => setter.Property == ToolTip.TipProperty)
            .Select(setter => (setter.Value as Avalonia.Data.Binding)?.Path)
            .SingleOrDefault();
    }

    // ###########################################################################################
    // The card really opens over a shown table, and closing the table closes it - letting go of its
    // pictures first (an Image left holding a disposed bitmap fails on the render thread).
    // ###########################################################################################
    [Fact]
    public async Task The_card_opens_over_a_shown_table_and_closes_with_it()
    {
        var source = new FakeFileSource();
        source.Files[("Commodore/C64/250407/new.png", BoardTableFileSide.Current)] = PixelA;

        await UiTest.RunAsync(async () =>
        {
            BoardTableEditor editor = Open(null, new BoardData { ComponentImages = [Image("Commodore/C64/250407/new.png")] }, source);

            var window = new Window { Content = editor, Width = 1200, Height = 600 };
            window.Show();
            Dispatcher.UIThread.RunJobs();

            BoardTableFilePreview preview = editor.ShowFilePreview(OnlyRow(editor), ImageFileColumn, editor)!;
            await preview.Loading;
            Dispatcher.UIThread.RunJobs();

            Assert.NotNull(TopLevel.GetTopLevel(preview));
            Assert.NotNull(PictureOf(Assert.Single(VisibleSides(preview)))!.Source);

            editor.Clear();
            Dispatcher.UIThread.RunJobs();

            Assert.Null(TopLevel.GetTopLevel(preview));
            Assert.Null(PictureOf(Assert.Single(VisibleSides(preview)))!.Source);

            window.Close();
        });
    }

    // ###########################################################################################
    // A card whose pointer has moved on before its picture arrived decodes nothing - the card
    // follows the pointer from cell to cell, so sweeping down a column builds one per row passed.
    // ###########################################################################################
    [Fact]
    public async Task A_card_closed_before_its_picture_arrives_decodes_nothing()
    {
        var source = new FakeFileSource { Gate = new TaskCompletionSource() };
        source.Files[("Commodore/C64/250407/new.png", BoardTableFileSide.Current)] = PixelA;

        await UiTest.RunAsync(async () =>
        {
            BoardTableEditor editor = Open(null, new BoardData { ComponentImages = [Image("Commodore/C64/250407/new.png")] }, source);

            BoardTableFilePreview preview = editor.BuildFilePreview(OnlyRow(editor), ImageFileColumn)!;
            preview.ReleaseImages();

            source.Gate.SetResult();
            await preview.Loading;

            Assert.Null(PictureOf(Assert.Single(VisibleSides(preview)))!.Source);
        });
    }

    // ------------------------------------------------------------------ With a real pointer

    // Two rows, each with its own picture, in a shown window wide enough for every column.
    private static (Window Window, BoardTableEditor Editor) ShowTwoPictures()
    {
        var source = new FakeFileSource();
        source.Files[("Commodore/C64/250407/a.png", BoardTableFileSide.Current)] = PixelA;
        source.Files[("Commodore/C64/250407/b.png", BoardTableFileSide.Current)] = PixelB;

        BoardData draft = new()
        {
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U8", Region = "PAL", Pin = "1", Name = "Pinout", File = "Commodore/C64/250407/a.png" },
                new ComponentImageEntry { BoardLabel = "U9", Region = "PAL", Pin = "1", Name = "Pinout", File = "Commodore/C64/250407/b.png" }
            ]
        };

        BoardTableEditor editor = Open(null, draft, source);

        var window = new Window { Content = editor, Width = 1800, Height = 700 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return (window, editor);
    }

    private static DataGridCell CellOnScreen(Window window, BoardTableRow row, int columnIndex) =>
        window.GetVisualDescendants()
            .OfType<DataGridCell>()
            .Single(cell => ReferenceEquals(cell.DataContext, row) && cell.OwningColumn?.Tag is int tag && tag == columnIndex);

    private static Point CentreOf(Window window, Control control) =>
        control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;

    private static void MoveTo(Window window, Point point)
    {
        window.MouseMove(point);
        Dispatcher.UIThread.RunJobs();
    }

    private static int NameColumn =>
        BoardWorkbookSchema.ComponentImages.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColName);

    // The path the open card is showing.
    private static string? ShownPath(BoardTableEditor editor) =>
        editor.ShownFilePreviewForTests is { } preview
            ? ((TextBlock)Assert.Single(VisibleSides(preview)).Children[1]).Text
            : null;

    // ###########################################################################################
    // *** AT ONCE, BOTH WAYS - THE REQUEST. *** No timer runs between the move and the card: it is
    // open straight after the pointer reaches the file cell, and gone straight after it reaches a
    // cell that is not a file, or leaves the grid.
    // ###########################################################################################
    [Fact]
    public void The_card_opens_the_moment_the_pointer_is_on_a_file_cell_and_closes_the_moment_it_leaves()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = ShowTwoPictures();
            BoardTableRow first = editor.CurrentSheet!.Rows[0];

            MoveTo(window, CentreOf(window, CellOnScreen(window, first, ImageFileColumn)));
            Assert.Equal("Commodore/C64/250407/a.png", ShownPath(editor));

            MoveTo(window, CentreOf(window, CellOnScreen(window, first, NameColumn)));
            Assert.Null(editor.ShownFilePreviewForTests);

            MoveTo(window, CentreOf(window, CellOnScreen(window, first, ImageFileColumn)));
            Assert.NotNull(editor.ShownFilePreviewForTests);

            // Off the grid altogether - the sheet tabs above it.
            MoveTo(window, new Point(4, 4));
            Assert.Null(editor.ShownFilePreviewForTests);

            window.Close();
        });
    }

    // ###########################################################################################
    // The card moves with the pointer down the File column. It sits BESIDE its cell, not under it:
    // under it, it covered the next row's file cell, and the move landed on the card instead.
    // ###########################################################################################
    [Fact]
    public void Moving_down_the_file_column_moves_the_card_to_the_next_file()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = ShowTwoPictures();
            IReadOnlyList<BoardTableRow> rows = editor.CurrentSheet!.Rows;

            MoveTo(window, CentreOf(window, CellOnScreen(window, rows[0], ImageFileColumn)));
            Assert.Equal("Commodore/C64/250407/a.png", ShownPath(editor));

            MoveTo(window, CentreOf(window, CellOnScreen(window, rows[1], ImageFileColumn)));
            Assert.Equal("Commodore/C64/250407/b.png", ShownPath(editor));

            window.Close();
        });
    }

    // ###########################################################################################
    // The pointer may cross from the cell onto the card - that is how a PDF's link gets clicked -
    // and the card closes the moment the pointer leaves it for anywhere else.
    // ###########################################################################################
    [Fact]
    public void The_pointer_can_cross_onto_the_card_and_leaving_the_card_closes_it()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = ShowTwoPictures();
            BoardTableRow first = editor.CurrentSheet!.Rows[0];

            MoveTo(window, CentreOf(window, CellOnScreen(window, first, ImageFileColumn)));
            BoardTableFilePreview shown = editor.ShownFilePreviewForTests!;

            MoveTo(window, CentreOf(window, editor.FilePreviewFrameForTests));
            Assert.Same(shown, editor.ShownFilePreviewForTests);

            MoveTo(window, new Point(4, 4));
            Assert.Null(editor.ShownFilePreviewForTests);

            window.Close();
        });
    }

    // ###########################################################################################
    // A click in the grid is not spent on closing the card: the cell under the pointer becomes the
    // current cell, as it would with no card, and the card stays. (A light-dismissed popup swallows
    // that first click.)
    // ###########################################################################################
    [Fact]
    public void A_click_on_the_file_cell_selects_it_and_keeps_the_card()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = ShowTwoPictures();
            BoardTableRow second = editor.CurrentSheet!.Rows[1];
            Point cell = CentreOf(window, CellOnScreen(window, second, ImageFileColumn));

            MoveTo(window, cell);
            window.MouseDown(cell, Avalonia.Input.MouseButton.Left);
            window.MouseUp(cell, Avalonia.Input.MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            DataGrid grid = editor.GetControl<DataGrid>("TableGrid");
            Assert.Same(second, grid.SelectedItem);
            Assert.Equal(ImageFileColumn, grid.CurrentColumn?.Tag);
            Assert.NotNull(editor.ShownFilePreviewForTests);

            window.Close();
        });
    }

    // The wheel scrolls the rows under a still pointer, so the card - anchored to a row that has
    // moved - closes, and the next move opens the right one.
    [Fact]
    public void Scrolling_the_rows_closes_the_card()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardTableEditor editor) = ShowTwoPictures();
            Point cell = CentreOf(window, CellOnScreen(window, editor.CurrentSheet!.Rows[0], ImageFileColumn));

            MoveTo(window, cell);
            Assert.NotNull(editor.ShownFilePreviewForTests);

            window.MouseWheel(cell, new Vector(0, -1));
            Dispatcher.UIThread.RunJobs();

            Assert.Null(editor.ShownFilePreviewForTests);

            window.Close();
        });
    }
}
