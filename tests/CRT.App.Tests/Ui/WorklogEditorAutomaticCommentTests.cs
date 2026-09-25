using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// Clicking a category chip or a state pill in the full editor records what happened as an
// automatic comment, so an entry carries its own history rather than only its current state.
//
// These drive the real chips through a real mouse press, because the rule being tested lives in
// the pointer handlers - asserting it by calling a helper directly would prove the helper works
// while leaving the wiring (which is where a "clicked the already-selected chip" bug would sit)
// completely uncovered.
//
// The comment ROWS are what is asserted, not the file on disk. WorklogManager is a static pointed
// at the user's real Workbooks folder, and re-pointing it from this collection would race the
// "Worklog" collection that owns it - so these stop at the observable UI, and
// WorklogAutomaticCommentTests covers the wording and WorklogManagerTests the persistence.
[Collection("HeadlessUi")]
public class WorklogEditorAutomaticCommentTests
{
    private static Bitmap CreateBitmap() =>
        new WriteableBitmap(new PixelSize(400, 200), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);

    private static WorklogEntryRecord CreateEntry(string category, string state) => new()
    {
        Id = 7,
        SchematicName = "Sch",
        Title = "Bad cap",
        Category = category,
        State = state,
        AreaX = 10,
        AreaY = 10,
        AreaWidth = 50,
        AreaHeight = 50,
    };

    private static void WithEditor(string category, string state, Action<WorklogEntryEditorWindow> body)
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            using var bitmap = CreateBitmap();
            window.Initialize(1, CreateEntry(category, state), bitmap);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();
                body(window);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Clicks the named chip/pill the way a user does - a real press at its centre, so the whole
    // hit-testing and routing path is exercised rather than the handler being called directly.
    private static void Click(Window window, string name)
    {
        var target = window.FindControl<Border>(name)!;

        var centre = target.TranslatePoint(new Point(target.Bounds.Width / 2, target.Bounds.Height / 2), window);
        Assert.NotNull(centre);

        window.MouseDown(centre!.Value, MouseButton.Left);
        window.MouseUp(centre.Value, MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
    }

    private static List<string> CommentTexts(Window window)
    {
        var list = window.FindControl<ItemsControl>("EditorCommentsList")!;

        return list.ItemsSource!.Cast<WorklogCommentRow>().Select(r => r.Text).ToList();
    }

    // The comment list is sorted by the user's own newest-first/oldest-first preference, read from
    // the REAL settings file - so the order rows appear in is not this test's to assert, and a
    // test that pinned it would pass or fail depending on how the developer last left the toggle.
    // Ordering by id recovers the sequence the events actually happened in, whichever way the
    // list happens to be showing them - ids are allocated in order, and unlike the timestamps they
    // cannot tie when two clicks land inside the same second.
    private static List<string> CommentTextsInEventOrder(Window window)
    {
        var list = window.FindControl<ItemsControl>("EditorCommentsList")!;

        return list.ItemsSource!.Cast<WorklogCommentRow>()
            .OrderBy(r => r.Id)
            .Select(r => r.Text)
            .ToList();
    }

    // ------------------------------------------------------------- state

    [Fact]
    public void Closing_an_open_worklog_records_a_worklog_closed_comment()
    {
        WithEditor("Note", "Open", window =>
        {
            Assert.Empty(CommentTexts(window));

            Click(window, "EditorStateClosedPill");

            Assert.Equal(new[] { "Worklog closed" }, CommentTexts(window));
        });
    }

    [Fact]
    public void Reopening_a_closed_worklog_records_a_worklog_opened_comment()
    {
        WithEditor("Note", "Closed", window =>
        {
            Click(window, "EditorStateOpenPill");

            Assert.Equal(new[] { "Worklog opened" }, CommentTexts(window));
        });
    }

    // Clicking the pill that is ALREADY selected is not a change. Recording it would let a user
    // fill the comment list by clicking the same pill repeatedly, and would claim an event that
    // never happened.
    [Fact]
    public void Clicking_the_already_selected_state_records_nothing()
    {
        WithEditor("Note", "Open", window =>
        {
            Click(window, "EditorStateOpenPill");
            Click(window, "EditorStateOpenPill");

            Assert.Empty(CommentTexts(window));
        });
    }

    // Each real change records its own line, so the trail reads as a history rather than
    // collapsing to the latest state.
    [Fact]
    public void Flipping_the_state_back_and_forth_records_both_changes_in_order()
    {
        WithEditor("Note", "Open", window =>
        {
            Click(window, "EditorStateClosedPill");
            Click(window, "EditorStateOpenPill");

            Assert.Equal(new[] { "Worklog closed", "Worklog opened" }, CommentTextsInEventOrder(window));
        });
    }

    // ------------------------------------------------------------- category

    [Theory]
    [InlineData("EditorCategoryCosmeticChip", "Worklog changed to \"Cosmetic\"")]
    [InlineData("EditorCategoryIssueChip", "Worklog changed to \"Issue\"")]
    public void Changing_the_category_records_a_quoted_category_comment(string chipName, string expected)
    {
        WithEditor("Note", "Open", window =>
        {
            Click(window, chipName);

            Assert.Equal(new[] { expected }, CommentTexts(window));
        });
    }

    [Fact]
    public void Clicking_the_already_selected_category_records_nothing()
    {
        WithEditor("Note", "Open", window =>
        {
            Click(window, "EditorCategoryNoteChip");

            Assert.Empty(CommentTexts(window));
        });
    }

    // A category change and a state change are separate events and each gets its own line.
    [Fact]
    public void Category_and_state_changes_are_recorded_separately()
    {
        WithEditor("Note", "Open", window =>
        {
            Click(window, "EditorCategoryIssueChip");
            Click(window, "EditorStateClosedPill");

            Assert.Equal(new[] { "Worklog changed to \"Issue\"", "Worklog closed" }, CommentTextsInEventOrder(window));
        });
    }

    // The automatic comment appearing means the empty-state helper text must go, exactly as it
    // does for a comment the user adds by hand.
    [Fact]
    public void An_automatic_comment_hides_the_no_comments_helper_text()
    {
        WithEditor("Note", "Open", window =>
        {
            Assert.True(window.FindControl<TextBlock>("EditorNoCommentsText")!.IsVisible);

            Click(window, "EditorStateClosedPill");

            Assert.False(window.FindControl<TextBlock>("EditorNoCommentsText")!.IsVisible);
        });
    }
}
