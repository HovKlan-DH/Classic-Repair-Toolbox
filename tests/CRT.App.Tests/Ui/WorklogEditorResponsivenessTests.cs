using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The full editor's header must stay usable as the splitter is dragged.
//
// The bug: the category chips and state pills sat in a horizontal StackPanel, which does not
// shrink or wrap - it simply overflows its bounds. Dragging the splitter left slid the "Closed"
// pill under the splitter and the right-hand panel drew on top of it, so a control the user needs
// to click became unreachable. Neither column had a MinWidth either, so there was nothing to stop
// the drag going arbitrarily far.
[Collection("HeadlessUi")]
public class WorklogEditorResponsivenessTests
{
    private static void WithEditor(double windowWidth, Action<WorklogEntryEditorWindow> body)
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
                Id = 1,
                SchematicName = "Sch",
                Title = "Bad cap",
                Category = "Issue",
                State = "Open",
                AreaWidth = 10,
                AreaHeight = 10,
            }, bitmap);

            try
            {
                window.Show();
                window.Measure(new Size(windowWidth, 700));
                window.Arrange(new Rect(0, 0, windowWidth, 700));
                Dispatcher.UIThread.RunJobs();
                body(window);
            }
            finally
            {
                window.Close();
            }
        });
    }

    private static Border Pill(WorklogEntryEditorWindow window, string name) =>
        window.GetVisualDescendants().OfType<Border>().First(b => b.Name == name);

    // The state pill furthest right is the one that used to disappear. At a narrow window it must
    // still be laid out and still sit left of the splitter, having wrapped onto its own line.
    // 900 is the window's declared MinWidth - narrower windows are unreachable, so asserting at
    // them would verify a state no user can produce. The squeeze that matters is the splitter,
    // covered by the left-column theory below.
    [Theory]
    [InlineData(1200.0)]
    [InlineData(900.0)]
    public void The_closed_pill_stays_left_of_the_splitter_at_every_width(double windowWidth)
    {
        double pillRight = 0;
        double splitterLeft = 0;
        double pillWidth = 0;

        WithEditor(windowWidth, window =>
        {
            var closed = Pill(window, "EditorStateClosedPill");
            var splitter = window.GetVisualDescendants().OfType<GridSplitter>().First();

            pillWidth = closed.Bounds.Width;
            pillRight = closed.TranslatePoint(new Point(closed.Bounds.Width, 0), window)?.X ?? -1;
            splitterLeft = splitter.TranslatePoint(new Point(0, 0), window)?.X ?? -1;
        });

        Assert.True(pillWidth > 0, $"the Closed pill was not laid out at {windowWidth}px");
        Assert.True(
            pillRight <= splitterLeft + 1,
            $"at {windowWidth}px the Closed pill runs to {pillRight}, past the splitter at {splitterLeft}");
    }

    // The squeeze a user can actually produce: drag the splitter left. 260 is the left column's
    // MinWidth, so this is the narrowest the header can ever be - and the pills must still wrap
    // rather than overflow under the splitter.
    //
    // Driving the splitter rather than the window size is what makes this reachable: the window's
    // own MinWidth is 900, so shrinking the WINDOW can never squeeze the header much.
    [Fact]
    public void The_state_pills_wrap_below_the_category_chips_at_the_narrowest_left_column()
    {
        double chipTopWide = 0, pillTopWide = 0;
        double chipTopNarrow = 0, pillTopNarrow = 0;

        WithEditor(1200, window =>
        {
            chipTopWide = Pill(window, "EditorCategoryNoteChip").TranslatePoint(new Point(0, 0), window)?.Y ?? -1;
            pillTopWide = Pill(window, "EditorStateOpenPill").TranslatePoint(new Point(0, 0), window)?.Y ?? -1;
        });

        WithEditor(900, window =>
        {
            SetLeftColumnWidth(window, 260);

            chipTopNarrow = Pill(window, "EditorCategoryNoteChip").TranslatePoint(new Point(0, 0), window)?.Y ?? -1;
            pillTopNarrow = Pill(window, "EditorStateOpenPill").TranslatePoint(new Point(0, 0), window)?.Y ?? -1;
        });

        // Wide: chips and pills share a line. Compared with a tolerance rather than for equality -
        // the pills are a couple of pixels taller than the chips (rounder corners, taller padding)
        // and are centred against them, so their tops differ by ~1px on the SAME line. The
        // difference that matters is a whole line height, checked below.
        Assert.True(
            Math.Abs(pillTopWide - chipTopWide) < 10,
            $"chips and pills should share a line when wide: chips at {chipTopWide}, pills at {pillTopWide}");

        // Narrow: the pills have dropped to their own line - a full row lower, not a pixel or two.
        Assert.True(
            pillTopNarrow - chipTopNarrow > 10,
            $"pills did not wrap: chips at {chipTopNarrow}, pills at {pillTopNarrow}");
    }

    // And the Closed pill still fits beside the splitter in that same narrowest state - the
    // original complaint was that it slid underneath and became unclickable.
    [Fact]
    public void The_closed_pill_stays_left_of_the_splitter_at_the_narrowest_left_column()
    {
        double pillRight = 0;
        double splitterLeft = 0;

        WithEditor(900, window =>
        {
            SetLeftColumnWidth(window, 260);

            var closed = Pill(window, "EditorStateClosedPill");
            var splitter = window.GetVisualDescendants().OfType<GridSplitter>().First();

            pillRight = closed.TranslatePoint(new Point(closed.Bounds.Width, 0), window)?.X ?? -1;
            splitterLeft = splitter.TranslatePoint(new Point(0, 0), window)?.X ?? -1;
        });

        Assert.True(
            pillRight <= splitterLeft + 1,
            $"the Closed pill runs to {pillRight}, past the splitter at {splitterLeft}");
    }

    // Squeezes the header to a given width by moving the splitter, then re-lays out - the same
    // thing dragging the splitter does.
    private static void SetLeftColumnWidth(WorklogEntryEditorWindow window, double width)
    {
        var grid = window.GetVisualDescendants().OfType<Grid>().First(g => g.Name == "EditorSplitGrid");
        grid.ColumnDefinitions[0].Width = new GridLength(width, GridUnitType.Pixel);

        window.Measure(new Size(window.Width, window.Height));
        window.Arrange(new Rect(0, 0, window.Width, window.Height));
        Dispatcher.UIThread.RunJobs();
    }
}
