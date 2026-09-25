using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The full editor's header (id badge, title, category chips, state pills) belongs to the LEFT
// column, not to the window: it describes the entry that side edits, so it must end at the
// splitter rather than spanning the dialog.
//
// This is markup-only, which the project normally verifies by eye - but "which column is this
// control in" is a structural fact a test can hold, and the header spanning the full width is
// exactly the regression that would come back if someone re-parented it while tidying.
[Collection("HeadlessUi")]
public class WorklogEditorLayoutTests
{
    private static WorklogEntryRecord CreateEntry() => new()
    {
        Id = 1,
        SchematicName = "Sch",
        Title = "Bad cap",
        Category = "Issue",
        State = "Open",
        AreaX = 10, AreaY = 10, AreaWidth = 50, AreaHeight = 50,
    };

    private static void WithEditor(Action<WorklogEntryEditorWindow> body)
    {
        UiTest.Run(() =>
        {
            // Placement persistence off for this window only: it otherwise restores the size,
            // position and splitter ratio from the developer's REAL settings file, so every layout
            // assertion below would depend on how they last left the editor. Scoped rather than
            // assigned, so it cannot leak into other tests in the shared session.
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 900;
            window.Height = 700;

            using var bitmap = new WriteableBitmap(
                new PixelSize(400, 200), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);
            window.Initialize(1, CreateEntry(), bitmap);

            try
            {
                window.Show();
                window.Measure(new Size(900, 700));
                window.Arrange(new Rect(0, 0, 900, 700));
                Dispatcher.UIThread.RunJobs();
                body(window);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // The title box must not be wider than the left column it belongs to. With the header spanning
    // the window it stretched across both sides, which is what made the split look like it began
    // halfway down the dialog.
    [Fact]
    public void The_title_field_stays_within_the_left_column()
    {
        double titleWidth = 0;
        double windowWidth = 0;

        WithEditor(window =>
        {
            titleWidth = window.GetVisualDescendants()
                .OfType<TextBox>()
                .First(t => t.Name == "EditorTitleTextBox")
                .Bounds.Width;

            windowWidth = window.Bounds.Width;
        });

        Assert.True(titleWidth > 0, "title box was not laid out");

        // The split is 3*:2*, so the left column is 60% of the width. Anything approaching the full
        // window width means the header is spanning again.
        Assert.True(
            titleWidth < windowWidth * 0.75,
            $"title box spans too much of the dialog: {titleWidth} of {windowWidth}");
    }

    // The header sits above the left column's scroller, so it must be left of the splitter - the
    // same side as the Description field it heads.
    [Fact]
    public void The_header_sits_on_the_same_side_as_the_description()
    {
        double titleRight = 0;
        double splitterLeft = 0;

        WithEditor(window =>
        {
            var title = window.GetVisualDescendants()
                .OfType<TextBox>()
                .First(t => t.Name == "EditorTitleTextBox");

            var splitter = window.GetVisualDescendants().OfType<GridSplitter>().First();

            titleRight = title.TranslatePoint(new Point(title.Bounds.Width, 0), window)?.X ?? -1;
            splitterLeft = splitter.TranslatePoint(new Point(0, 0), window)?.X ?? -1;
        });

        Assert.True(titleRight > 0 && splitterLeft > 0, "controls were not laid out");
        Assert.True(
            titleRight <= splitterLeft,
            $"title box crosses the splitter: ends at {titleRight}, splitter starts at {splitterLeft}");
    }

    // The footer must still be visible after the row restructure - it moved from Row 2 to Row 1
    // when the header stopped occupying a row of its own, and a stale index would place it in a
    // row that no longer exists.
    [Fact]
    public void The_save_and_cancel_footer_is_still_laid_out()
    {
        double saveWidth = 0;

        WithEditor(window =>
        {
            saveWidth = window.GetVisualDescendants()
                .OfType<Button>()
                .First(b => b.Name == "EditorSaveButton")
                .Bounds.Width;
        });

        Assert.True(saveWidth > 0, "Save button was not laid out - check the footer's Grid.Row");
    }
}
