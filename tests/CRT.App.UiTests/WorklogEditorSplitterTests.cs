using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The editor's splitter position is persisted as the left column's SHARE of the two content
// columns, not as a pixel width, so reopening at a different window size keeps the proportion.
//
// These tests pin the measurement that feeds that setting. They do NOT touch UserSettings: this
// file is in the HeadlessUi collection, and mutating UserSettings here would race the tests in the
// "UserSettings" collection (see the note in CLAUDE.md about statics needing their own collection).
// What is verified here is that the grid reports the widths the save path reads, and that applying
// a ratio produces the split it claims.
[Collection("HeadlessUi")]
public class WorklogEditorSplitterTests
{
    private static void WithEditor(double windowWidth, Action<WorklogEntryEditorWindow, Grid> body)
    {
        UiTest.Run(() =>
        {
            // Placement persistence off for this window only: it otherwise restores the size,
            // position and splitter ratio from the developer's REAL settings file, so every layout
            // assertion below would depend on how they last left the editor. Scoped rather than
            // assigned, so it cannot leak into other tests in the shared session.
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = windowWidth;
            window.Height = 700;

            using var bitmap = new WriteableBitmap(
                new PixelSize(400, 200), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);

            window.Initialize(1, new WorklogEntryRecord
            {
                Id = 1, SchematicName = "Sch", Title = "t", Category = "Issue", State = "Open",
                AreaWidth = 10, AreaHeight = 10,
            }, bitmap);

            try
            {
                window.Show();
                window.Measure(new Size(windowWidth, 700));
                window.Arrange(new Rect(0, 0, windowWidth, 700));
                Dispatcher.UIThread.RunJobs();

                var grid = window.GetVisualDescendants().OfType<Grid>().First(g => g.Name == "EditorSplitGrid");
                body(window, grid);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ColumnDefinition.ActualWidth is what the save path measures. If it reported 0 (or the star
    // value rather than the rendered width) the saved ratio would be meaningless, and the window
    // would silently stop remembering the splitter.
    [Fact]
    public void The_split_grid_reports_real_column_widths()
    {
        double left = 0, right = 0;

        WithEditor(1200, (_, grid) =>
        {
            left = grid.ColumnDefinitions[0].ActualWidth;
            right = grid.ColumnDefinitions[2].ActualWidth;
        });

        Assert.True(left > 0, "left column reported no width");
        Assert.True(right > 0, "right column reported no width");
    }

    // The rendered split must match whatever ratio the column definitions ask for - the mechanism
    // both the XAML default and the restore path rely on.
    [Fact]
    public void The_rendered_split_matches_the_requested_column_ratio()
    {
        double requested = 0;
        double rendered = 0;

        WithEditor(1200, (_, grid) =>
        {
            double leftStar = grid.ColumnDefinitions[0].Width.Value;
            double rightStar = grid.ColumnDefinitions[2].Width.Value;
            requested = leftStar / (leftStar + rightStar);

            double left = grid.ColumnDefinitions[0].ActualWidth;
            double right = grid.ColumnDefinitions[2].ActualWidth;
            rendered = left / (left + right);
        });

        Assert.True(requested > 0, "columns did not report a star ratio");
        Assert.Equal(requested, rendered, 2);
    }

    // With persistence off the window uses the split its XAML declares, 3*:2* = 60% to the left.
    // That is also the default the settings carry, so a first run and a cleared setting agree.
    [Fact]
    public void The_default_split_gives_the_left_column_sixty_percent()
    {
        double ratio = 0;

        WithEditor(1200, (_, grid) =>
        {
            double left = grid.ColumnDefinitions[0].ActualWidth;
            double right = grid.ColumnDefinitions[2].ActualWidth;
            ratio = left / (left + right);
        });

        // The settings default is asserted in UserSettingsTests, against a temp settings file -
        // reading UserSettings here would put this test back at the mercy of the developer's own
        // saved layout, which is the whole reason persistence is disabled above.
        Assert.Equal(0.6, ratio, 2);
    }

    // Applying a ratio the way RestoreWindowPlacement does must actually produce that split, and
    // must survive a re-layout - this is the mechanism the restore relies on.
    //
    // The original widths are put back in a finally. Avalonia's XAML loader hands each window
    // ColumnDefinition instances derived from one compiled template, so a width written here leaks
    // into every editor window built later in this shared headless session: without the restore,
    // The_default_split_gives_the_left_column_sixty_percent saw 0.35 instead of 0.6 and failed -
    // but only when run after this test, never alone. That is precisely the order-dependent flake
    // that is miserable to diagnose from the symptom, so this test cleans up after itself.
    [Fact]
    public void Applying_a_ratio_produces_that_split()
    {
        double ratio = 0;

        WithEditor(1200, (window, grid) =>
        {
            GridLength originalLeft = grid.ColumnDefinitions[0].Width;
            GridLength originalRight = grid.ColumnDefinitions[2].Width;

            try
            {
                grid.ColumnDefinitions[0].Width = new GridLength(0.35, GridUnitType.Star);
                grid.ColumnDefinitions[2].Width = new GridLength(0.65, GridUnitType.Star);

                window.Measure(new Size(1200, 700));
                window.Arrange(new Rect(0, 0, 1200, 700));
                Dispatcher.UIThread.RunJobs();

                double left = grid.ColumnDefinitions[0].ActualWidth;
                double right = grid.ColumnDefinitions[2].ActualWidth;
                ratio = left / (left + right);
            }
            finally
            {
                grid.ColumnDefinitions[0].Width = originalLeft;
                grid.ColumnDefinitions[2].Width = originalRight;
            }
        });

        Assert.Equal(0.35, ratio, 2);
    }
}
