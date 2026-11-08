using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The editor opened on a NEW entry - the "Add worklog" flow, where drawing an area on the
// schematic now comes straight here rather than through a small "New fault" quick card first.
//
// The card is gone, so this window is the only place a worklog entry is written, and it has to
// behave differently in two ways when the entry does not exist yet:
//
//   1. Save must be reachable. Initialize is written for a saved entry and deliberately ends with
//      the window CLEAN (Save disabled), since opening an entry is not editing it. A new entry is
//      the opposite: it is unsaved by definition, so typing a title has to be enough to save.
//   2. Nothing may be written before Save. A saved entry's sub-list changes write through to disk
//      at once, which is what lets Cancel be harmless there; for a draft that would leave a
//      half-made entry behind, so the draft is held entirely in memory.
//
// The FILE side of that - what AddEntryRecord writes, and the id it allocates - is
// WorklogManagerTests' job; these stop at the window, because WorklogManager is a static pointed
// at the user's real Workbooks folder and re-pointing it from this collection would race the
// "Worklog" collection that owns it.
[Collection("HeadlessUi")]
public class WorklogEditorNewEntryTests
{
    private static Bitmap CreateBitmap() =>
        new WriteableBitmap(new PixelSize(400, 200), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);

    private static void WithNewEntryEditor(Action<WorklogEntryEditorWindow> body)
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow { Width = 1000, Height = 700 };

            using var bitmap = CreateBitmap();

            // Workbook 0 does not exist, so PeekNextEntryId falls back to 1 and nothing is written -
            // which is the whole point: a draft touches no workbook until Save.
            window.InitializeForNewEntry(0, "Sheet 1", new Rect(10, 20, 30, 40), bitmap);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();
                body(window);
            }
            finally
            {
                window.Close();
                Dispatcher.UIThread.RunJobs();
            }
        });
    }

    private static void SetText(TextBox box, string? text)
    {
        box.Text = text;

        // Avalonia raises TextChanged from a posted dispatcher job rather than from the setter, so
        // the dirty flag and the Save gate have not run yet when the setter returns.
        Dispatcher.UIThread.RunJobs();
    }

    // A new entry opens blank, so there is nothing to save yet - the title is what makes it
    // identifiable, and an entry with none shows nothing but its "#N" in every list.
    [Fact]
    public void A_new_entry_opens_with_an_empty_title_and_save_disabled()
    {
        WithNewEntryEditor(window =>
        {
            Assert.True(string.IsNullOrEmpty(window.FindControl<TextBox>("EditorTitleTextBox")!.Text));
            Assert.False(window.FindControl<Button>("EditorSaveButton")!.IsEnabled);
        });
    }

    // THE REASON THIS MODE EXISTS AS A MODE. Initialize ends by clearing the dirty flag, so without
    // the draft's own re-raise, typing a title on a new entry left Save disabled and the entry
    // could never be saved at all.
    [Fact]
    public void Typing_a_title_enables_save_on_a_new_entry()
    {
        WithNewEntryEditor(window =>
        {
            SetText(window.FindControl<TextBox>("EditorTitleTextBox")!, "Dead VIC");

            Assert.True(window.FindControl<Button>("EditorSaveButton")!.IsEnabled);
        });
    }

    // Whitespace does not count: the title is Trim()ed before it is persisted, so a title of spaces
    // would be saved as an empty one and the gate has to agree with what the save does.
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("\t")]
    public void A_blank_or_whitespace_title_keeps_save_disabled_on_a_new_entry(string title)
    {
        WithNewEntryEditor(window =>
        {
            var titleBox = window.FindControl<TextBox>("EditorTitleTextBox")!;

            // Typed and then cleared, so this exercises the gate re-evaluating rather than the
            // untouched starting state the first test already covers.
            SetText(titleBox, "Something");
            SetText(titleBox, title);

            Assert.False(window.FindControl<Button>("EditorSaveButton")!.IsEnabled);
        });
    }

    // The drawn area and its schematic come from the drag on the board and are the one thing the
    // editor cannot ask for, so they must survive into the record it will write.
    [Fact]
    public void A_new_entry_carries_the_drawn_area_and_its_schematic()
    {
        WithNewEntryEditor(window =>
        {
            Assert.Equal("Sheet 1", window.FindControl<TextBlock>("EditorLocationSchematicNameText")!.Text);
        });
    }

    // Note/Open, the same defaults the quick card this flow replaced started at.
    [Fact]
    public void A_new_entry_starts_as_an_open_note()
    {
        WithNewEntryEditor(window =>
        {
            // The state pill and category chip are visual, so the selection is read off the audit
            // comment the entry starts with instead - it is written from the same defaults.
            var comments = window.FindControl<ItemsControl>("EditorCommentsList")!;
            var rows = Assert.IsAssignableFrom<IEnumerable<WorklogCommentRow>>(comments.ItemsSource).ToList();

            Assert.Equal(WorklogManager.CreatedCommentText, Assert.Single(rows).Text);
        });
    }

    // "Show marked area" starts ticked, matching WorklogEntryRecord's own default - an entry drawn
    // on the board shows where it was drawn unless the user says otherwise.
    [Fact]
    public void A_new_entry_shows_its_marked_area_by_default()
    {
        WithNewEntryEditor(window =>
        {
            Assert.True(window.FindControl<CheckBox>("EditorShowMarkedAreaCheckBox")!.IsChecked);
        });
    }

    // The window says it is new rather than reusing the generic "Worklog" title, so a user who has
    // several open can tell which one has not been saved yet.
    [Fact]
    public void A_new_entry_window_names_itself_as_new()
    {
        WithNewEntryEditor(window => Assert.Equal("New worklog", window.Title));
    }

    // Closing without saving must report NOTHING saved. For a saved entry WasSaved can be true
    // after a Cancel (an instant-saved sub-list change already reached disk); a draft writes
    // nothing at all, so the caller must not be told to refresh - and must certainly not find a
    // half-made entry on the board.
    [Fact]
    public void Cancelling_a_new_entry_reports_nothing_saved()
    {
        WithNewEntryEditor(window =>
        {
            SetText(window.FindControl<TextBox>("EditorTitleTextBox")!, "Abandoned");

            window.Close();
            Dispatcher.UIThread.RunJobs();

            Assert.False(window.WasSaved);
            Assert.Null(window.SavedNewEntry);
        });
    }

    // A new entry's component checklist starts fully ticked: the user drew the area around these
    // components, so all of them in scope is the sensible starting point and unticking one is
    // quicker than ticking eight. A SAVED entry restores the choice made last time instead, which
    // is why tickAll is a parameter rather than the rule.
    [Fact]
    public void A_new_entrys_component_scope_starts_fully_ticked()
    {
        WithNewEntryEditor(window =>
        {
            window.InitializeComponentScope(
                new[] { ("U1", "VIC-II"), ("C7", "Ceramic") },
                tickAll: true);
            Dispatcher.UIThread.RunJobs();

            var list = window.FindControl<ItemsControl>("EditorComponentList")!;
            var rows = Assert.IsAssignableFrom<IEnumerable<WorklogEntryComponentRow>>(list.ItemsSource).ToList();

            Assert.Equal(2, rows.Count);
            Assert.All(rows, row => Assert.True(row.IsChecked, $"{row.BoardLabel} did not start ticked"));
        });
    }

    // The contrast that makes the flag meaningful: without tickAll, a scope whose labels are not in
    // the entry's saved ComponentLabels starts UNticked - which is right for a saved entry and
    // would be wrong for a new one.
    [Fact]
    public void Without_tick_all_an_unsaved_scope_starts_unticked()
    {
        WithNewEntryEditor(window =>
        {
            window.InitializeComponentScope(new[] { ("U1", "VIC-II"), ("C7", "Ceramic") });
            Dispatcher.UIThread.RunJobs();

            var list = window.FindControl<ItemsControl>("EditorComponentList")!;
            var rows = Assert.IsAssignableFrom<IEnumerable<WorklogEntryComponentRow>>(list.ItemsSource).ToList();

            Assert.All(rows, row => Assert.False(row.IsChecked, $"{row.BoardLabel} started ticked"));
        });
    }

    // A new entry's button says "Add worklog" - it is being created, not updated. The counterpart
    // to WorklogEditorHeaderTests' "Update worklog" assertion for a saved entry.
    [Fact]
    public void A_new_entrys_save_button_says_add_worklog()
    {
        WithNewEntryEditor(window =>
            Assert.Equal("Add worklog", window.FindControl<Button>("EditorSaveButton")!.Content));
    }

    // THE REQUESTED CHANGE: a new entry opens with an empty title by definition, and showing "A
    // worklog needs a title before it can be saved." before the user has typed anything reads as
    // nagging rather than helpful - there is nothing on disk yet for the window to disagree with.
    // WorklogEditorHeaderTests' matching test shows the message DOES still appear for a SAVED
    // entry, where clearing the title is a real inconsistency risk against what is already on disk.
    [Fact]
    public void A_new_entrys_blank_title_shows_no_explanatory_message()
    {
        WithNewEntryEditor(window =>
        {
            Assert.False(window.FindControl<TextBlock>("EditorSaveFailedText")!.IsVisible);
        });
    }

    // Typing then clearing the title again must not surface the message either - the suppression
    // is not just a first-paint accident of the guard being up.
    [Fact]
    public void Clearing_a_new_entrys_typed_title_still_shows_no_message()
    {
        WithNewEntryEditor(window =>
        {
            var titleBox = window.FindControl<TextBox>("EditorTitleTextBox")!;

            SetText(titleBox, "Dead VIC");
            SetText(titleBox, "");

            Assert.False(window.FindControl<TextBlock>("EditorSaveFailedText")!.IsVisible);
        });
    }
}
