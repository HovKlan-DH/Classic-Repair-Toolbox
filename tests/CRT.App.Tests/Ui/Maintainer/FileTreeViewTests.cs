using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Headless;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;
using ClassicRepairToolbox.Tests.Maintainer;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The file tree as drawn (owner request, 2026-09-28) - FileTreeView, on "Beta > Prod" and in the
// window "Files..." opens. What each row says is FileTree's and tested there; these pin what the
// control does with it: folders opening and closing (one at a time and all at once), the boxed
// plus and minus, and - with a real pointer in a shown window - the table's hover card on a file.
// Headless drawing decodes no pixels, so a picture here proves a bitmap was made, not its look.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class FileTreeViewTests
{
    // A real one-pixel PNG.
    private static readonly byte[] Pixel = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    private sealed class FakeFiles : IFileTreeFiles
    {
        public List<string> Read { get; } = [];

        public List<string> Opened { get; } = [];

        public Task<byte[]?> ReadAsync(SystemFileEntry file)
        {
            this.Read.Add(file.Path);
            return Task.FromResult<byte[]?>(file.Path.EndsWith(".png", StringComparison.Ordinal) ? FileTreeViewTests.Pixel : null);
        }

        public Task<string?> OpenAsync(SystemFileEntry file)
        {
            this.Opened.Add(file.Path);
            return Task.FromResult<string?>(null);
        }
    }

    private static SystemFileEntry E(string path, SystemFileChange change = SystemFileChange.Unchanged, SystemFileSource from = SystemFileSource.Beta) =>
        new(path, change, from);

    private static IReadOnlyList<SystemFileEntry> Board() =>
    [
        FileTreeViewTests.E("Commodore/C128/310378/Data.xlsx", SystemFileChange.Changed),
        FileTreeViewTests.E("Commodore/C128/310378/Images/a.png"),
        FileTreeViewTests.E("Commodore/C128/310378/Images/b.png"),
        FileTreeViewTests.E("Commodore/C128/310378/Manuals/m.pdf"),
        FileTreeViewTests.E("Commodore/C128/310378/Manuals/n.pdf")
    ];

    private static FileTreeView Shown(bool onlyChanged)
    {
        var view = new FileTreeView();
        view.FindControl<CheckBox>("OnlyChangedCheckBox")!.IsChecked = onlyChanged;
        view.Show(FileTreeViewTests.Board());
        return view;
    }

    private static void Click(FileTreeView view, string button) =>
        view.FindControl<Button>(button)!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    // ###########################################################################################
    // *** OPENING A FOLDER REPLACES ONLY ITS OWN ROWS. *** The list is updated in place, so the rows
    // around it stay the same objects and the list does not jump back to the top every time a
    // folder is clicked halfway down a board.
    // ###########################################################################################
    [Fact]
    public void Opening_a_folder_adds_its_rows_and_keeps_the_others()
    {
        UiTest.Run(() =>
        {
            FileTreeView view = FileTreeViewTests.Shown(onlyChanged: false);

            // At first: the folders on the way to the change are open, Images and Manuals closed.
            Assert.Equal(
                ["Commodore", "C128", "310378", "Images", "Manuals", "Data.xlsx"],
                view.RowsForTests.Select(row => row.Name));

            FileTreeRowView data = view.RowsForTests.Single(row => row.Name == "Data.xlsx");
            FileTreeRowView top = view.RowsForTests[0];

            view.ToggleForTests("Commodore/C128/310378/Manuals");

            Assert.Equal(
                ["Commodore", "C128", "310378", "Images", "Manuals", "m.pdf", "n.pdf", "Data.xlsx"],
                view.RowsForTests.Select(row => row.Name));
            Assert.Same(data, view.RowsForTests.Single(row => row.Name == "Data.xlsx"));
            Assert.Same(top, view.RowsForTests[0]);

            view.ToggleForTests("Commodore/C128/310378/Manuals");

            Assert.Equal(6, view.RowsForTests.Count);
        });
    }

    // ###########################################################################################
    // A folder's box is a plus when it is closed and a minus when it is open, in Font Awesome's
    // outlined face (owner request, 2026-09-28: "a plus and a minus (in a box, if possible)"); a
    // file keeps its file icon.
    // ###########################################################################################
    [Fact]
    public void A_folder_shows_a_boxed_minus_open_and_a_boxed_plus_closed()
    {
        UiTest.Run(() =>
        {
            FileTreeView view = FileTreeViewTests.Shown(onlyChanged: false);

            FileTreeRowView open = view.RowsForTests.Single(row => row.Name == "310378");
            FileTreeRowView closed = view.RowsForTests.Single(row => row.Name == "Images");
            FileTreeRowView file = view.RowsForTests.Single(row => row.Name == "Data.xlsx");

            Assert.Equal(FileTreeRowView.OpenFolderGlyph, open.Glyph);
            Assert.Equal(FileTreeRowView.ClosedFolderGlyph, closed.Glyph);
            Assert.Equal(FileTreeRowView.FileGlyph, file.Glyph);
            Assert.True(open.IsFolder);
            Assert.False(file.IsFolder);

            view.ToggleForTests("Commodore/C128/310378/Images");

            Assert.Equal(FileTreeRowView.OpenFolderGlyph, view.RowsForTests.Single(row => row.Name == "Images").Glyph);
        });
    }

    // Only the changes, under their folders - which close like any other there too.
    [Fact]
    public void Only_changed_shows_the_change_under_its_folders_and_they_still_close()
    {
        UiTest.Run(() =>
        {
            FileTreeView view = FileTreeViewTests.Shown(onlyChanged: true);

            Assert.Equal(["Commodore", "C128", "310378", "Data.xlsx"], view.RowsForTests.Select(row => row.Name));

            view.ToggleForTests("Commodore/C128/310378");

            Assert.Equal(["Commodore", "C128", "310378"], view.RowsForTests.Select(row => row.Name));
        });
    }

    // ###########################################################################################
    // "Collapse all" and "Expand all" (owner request, 2026-09-28). Expand opens every folder -
    // unchanged ones too when the whole system is shown; collapse leaves the top folder alone.
    // ###########################################################################################
    [Fact]
    public void Expand_all_and_collapse_all_open_and_close_every_folder()
    {
        UiTest.Run(() =>
        {
            FileTreeView view = FileTreeViewTests.Shown(onlyChanged: false);

            Click(view, "CollapseAllButton");
            Assert.Equal(["Commodore"], view.RowsForTests.Select(row => row.Name));

            Click(view, "ExpandAllButton");
            Assert.Equal(
                ["Commodore", "C128", "310378", "Images", "a.png", "b.png", "Manuals", "m.pdf", "n.pdf", "Data.xlsx"],
                view.RowsForTests.Select(row => row.Name));
        });
    }

    // Ticking the box again brings every change back into view, whatever was closed before.
    [Fact]
    public void Ticking_only_changed_opens_the_folders_of_every_change_again()
    {
        UiTest.Run(() =>
        {
            FileTreeView view = FileTreeViewTests.Shown(onlyChanged: true);
            CheckBox box = view.FindControl<CheckBox>("OnlyChangedCheckBox")!;

            Click(view, "CollapseAllButton");
            box.IsChecked = false;
            box.IsChecked = true;

            Assert.Equal(["Commodore", "C128", "310378", "Data.xlsx"], view.RowsForTests.Select(row => row.Name));
        });
    }

    // The box is the switch: unticked shows the rest of the system around the change.
    [Fact]
    public void Unticking_only_changed_shows_the_whole_system()
    {
        UiTest.Run(() =>
        {
            FileTreeView view = FileTreeViewTests.Shown(onlyChanged: true);

            view.FindControl<CheckBox>("OnlyChangedCheckBox")!.IsChecked = false;

            Assert.Contains(view.RowsForTests, row => row.Name == "Images");
            Assert.Contains(view.RowsForTests, row => row.Name == "Manuals");
        });
    }

    // A file the approval has not written yet has nothing to open: it says so instead of asking
    // the host for bytes that do not exist.
    [Fact]
    public async Task A_file_not_written_yet_explains_instead_of_opening()
    {
        await UiTest.RunAsync(async () =>
        {
            var files = new FakeFiles();
            var view = new FileTreeView { Files = files };

            view.Show([new SystemFileEntry("Amstrad/CPC/464/Data.xlsx", SystemFileChange.Added, SystemFileSource.NotWrittenYet, WrittenOnApproval: true)]);

            ListBox list = view.FindControl<ListBox>("RowsList")!;
            list.SelectedItem = view.RowsForTests.Single(row => row.Name == "Data.xlsx");
            view.FindControl<MenuItem>("OpenRowItem")!.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));

            await Task.Yield();

            Assert.Empty(files.Opened);
            Assert.Contains("nothing to open yet", view.FindControl<TextBlock>("MessageText")!.Text, StringComparison.Ordinal);
        });
    }

    // A removed file is struck through and marked, so it reads as going, not as there.
    [Fact]
    public void A_removed_file_is_struck_through_and_marked()
    {
        UiTest.Run(() =>
        {
            var view = new FileTreeView();
            view.Show([FileTreeViewTests.E("a/old.png", SystemFileChange.Removed, SystemFileSource.Production)]);

            FileTreeRowView row = view.RowsForTests.Single(candidate => candidate.Name == "old.png");

            Assert.Equal("removed", row.PillText);
            Assert.True(row.HasPill);
            Assert.NotNull(row.Decorations);
        });
    }

    // ------------------------------------------------------------------ The hover card

    private static (Window Window, FileTreeView View, FakeFiles Files) ShownInAWindow(params SystemFileEntry[] entries)
    {
        var files = new FakeFiles();
        var view = new FileTreeView { Files = files };
        view.FindControl<CheckBox>("OnlyChangedCheckBox")!.IsChecked = false;
        view.Show(entries);
        view.FindControl<Button>("ExpandAllButton")!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

        var window = new Window { Content = view, Width = 900, Height = 600 };
        window.Show();
        Dispatcher.UIThread.RunJobs();

        return (window, view, files);
    }

    private static ListBoxItem RowItem(Window window, string name) =>
        window.GetVisualDescendants()
            .OfType<ListBoxItem>()
            .Single(candidate => (candidate.DataContext as FileTreeRowView)?.Name == name);

    // What the row says - its icon, name and marker.
    private static StackPanel RowContent(Window window, string name) =>
        RowItem(window, name).GetVisualDescendants()
            .OfType<StackPanel>()
            .Single(panel => panel.Classes.Contains("RowContent"));

    private static Point CentreOfRow(Window window, string name)
    {
        StackPanel content = RowContent(window, name);
        return content.TranslatePoint(new Point(content.Bounds.Width / 2, content.Bounds.Height / 2), window)!.Value;
    }

    private static bool IsShaded(IBrush? brush) => brush is ISolidColorBrush { Color.A: > 0 };

    // A point on the same row, in the empty width far to the right of what it says.
    private static Point EmptyBesideRow(Window window, string name)
    {
        ListBoxItem item = RowItem(window, name);
        return item.TranslatePoint(new Point(item.Bounds.Width - 40, item.Bounds.Height / 2), window)!.Value;
    }

    // A point on the same row, in the indent before its icon.
    private static Point IndentOfRow(Window window, string name)
    {
        ListBoxItem item = RowItem(window, name);
        return item.TranslatePoint(new Point(10, item.Bounds.Height / 2), window)!.Value;
    }

    private static void MoveTo(Window window, Point point)
    {
        window.MouseMove(point);
        Dispatcher.UIThread.RunJobs();
    }

    private static IReadOnlyList<StackPanel> VisibleSides(CRT.BoardTableFilePreview preview) =>
        preview.Sides.OfType<StackPanel>().Where(side => side.IsVisible).ToList();

    // ###########################################################################################
    // *** THE TABLE'S CARD, SAYING NOTHING ABOUT THE CHANGE (owner request, 2026-09-28: "only the
    // relative path and then the other functionality from the table format"). *** Pointing at a
    // picture's row shows its path and the picture - no "New file", no "Replaced" - at once, and
    // moving onto a folder's row closes it.
    // ###########################################################################################
    [Fact]
    public async Task Pointing_at_a_picture_shows_its_path_and_the_picture_and_nothing_else()
    {
        await UiTest.RunAsync(async () =>
        {
            (Window window, FileTreeView view, FakeFiles files) = ShownInAWindow(
                FileTreeViewTests.E("Commodore/C128/310378/Images/a.png", SystemFileChange.Changed));

            MoveTo(window, CentreOfRow(window, "a.png"));

            CRT.BoardTableFilePreview preview = view.ShownFilePreviewForTests!;
            Assert.NotNull(preview);
            await preview.Loading;

            Assert.Equal(string.Empty, preview.HeadlineText);

            StackPanel side = Assert.Single(VisibleSides(preview));
            Assert.False(side.Children[0].IsVisible);
            Assert.Equal("Commodore/C128/310378/Images/a.png", ((TextBlock)side.Children[1]).Text);
            Assert.NotNull(side.Children.OfType<Image>().Single().Source);
            Assert.Equal(["Commodore/C128/310378/Images/a.png"], files.Read);

            MoveTo(window, CentreOfRow(window, "Images"));
            Assert.Null(view.ShownFilePreviewForTests);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** ONLY WHAT THE ROW SAYS REACTS (owner request, 2026-09-28: "It gets confusing when I am in
    // the middle of an empty screen, that it then reacts of the file on that row"). *** The empty
    // width beside a file's name and the indent before it show no card, and moving there from the
    // name closes the card.
    // ###########################################################################################
    [Fact]
    public void The_empty_width_beside_a_file_and_its_indent_show_no_card()
    {
        UiTest.Run(() =>
        {
            (Window window, FileTreeView view, FakeFiles files) = ShownInAWindow(
                FileTreeViewTests.E("Commodore/C128/310378/Images/a.png", SystemFileChange.Changed));

            ContentPresenter rowBackground = RowItem(window, "a.png").GetVisualDescendants()
                .OfType<ContentPresenter>()
                .First(presenter => presenter.Name == "PART_ContentPresenter");

            MoveTo(window, EmptyBesideRow(window, "a.png"));
            Assert.Null(view.ShownFilePreviewForTests);

            // Nothing lights up there either - not the row, not its name.
            Assert.False(IsShaded(rowBackground.Background));
            Assert.False(IsShaded(RowContent(window, "a.png").Background));

            MoveTo(window, IndentOfRow(window, "a.png"));
            Assert.Null(view.ShownFilePreviewForTests);

            // On the name: the card, and the name's own shading. A background set on the panel in
            // the template outranked the style and kept it from ever showing.
            MoveTo(window, CentreOfRow(window, "a.png"));
            Assert.NotNull(view.ShownFilePreviewForTests);
            Assert.True(IsShaded(RowContent(window, "a.png").Background));
            Assert.False(IsShaded(rowBackground.Background));

            MoveTo(window, EmptyBesideRow(window, "a.png"));
            Assert.Null(view.ShownFilePreviewForTests);

            window.Close();
        });
    }

    // A PDF is not read on hover - its card is a link, and the link opens it through the host.
    [Fact]
    public void A_pdfs_card_is_a_link_that_opens_it()
    {
        UiTest.Run(() =>
        {
            (Window window, FileTreeView view, FakeFiles files) = ShownInAWindow(
                FileTreeViewTests.E("Commodore/C128/310378/Manuals/m.pdf"));

            MoveTo(window, CentreOfRow(window, "m.pdf"));

            StackPanel side = Assert.Single(VisibleSides(view.ShownFilePreviewForTests!));
            Button link = side.Children.OfType<Button>().Single();

            Assert.Equal("Open PDF file", ((TextBlock)link.Content!).Text);
            Assert.Empty(files.Read);

            link.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Commodore/C128/310378/Manuals/m.pdf"], files.Opened);

            window.Close();
        });
    }

    // ###########################################################################################
    // The one thing a card says beside the path: that the workbook it opens is BETA's current copy
    // when the approval writes a new one. A file with nothing to open yet has no card at all.
    // ###########################################################################################
    [Fact]
    public void Only_a_file_the_approval_writes_carries_a_note_and_one_not_written_yet_has_no_card()
    {
        UiTest.Run(() =>
        {
            (Window window, FileTreeView view, _) = ShownInAWindow(
                new SystemFileEntry("Commodore/C128/310378/Data.xlsx", SystemFileChange.Changed, SystemFileSource.Beta, WrittenOnApproval: true),
                new SystemFileEntry("Commodore/C128/310378/Data.json", SystemFileChange.Added, SystemFileSource.NotWrittenYet, WrittenOnApproval: true));

            MoveTo(window, CentreOfRow(window, "Data.xlsx"));

            CRT.BoardTableFilePreview preview = view.ShownFilePreviewForTests!;
            TextBlock note = ((StackPanel)preview.Child!).Children.OfType<TextBlock>().First();
            Assert.Contains("BETA's copy as it is now", note.Text, StringComparison.Ordinal);

            MoveTo(window, CentreOfRow(window, "Data.json"));
            Assert.Null(view.ShownFilePreviewForTests);
            Assert.Contains("nothing to open yet", view.RowsForTests.Single(row => row.Name == "Data.json").ToolTip, StringComparison.Ordinal);

            window.Close();
        });
    }
}
