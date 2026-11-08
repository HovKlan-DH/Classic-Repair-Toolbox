using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The full editor's header: the mandatory title, and the layout that puts the "#N" badge to the
// right of the title with the category chips and state pills starting at the title's left edge.
//
// The title rule matters beyond the button: the sub-lists (links, comments, work done, photos,
// files) save themselves instantly through PersistEntrySilently, which writes the title too - so
// a blank title has a route to disk that never passes the Save button at all.
[Collection("HeadlessUi")]
public class WorklogEditorHeaderTests
{
    private static Bitmap CreateBitmap() =>
        new WriteableBitmap(new PixelSize(400, 200), new Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);

    private static WorklogEntryRecord CreateEntry(string title = "Bad cap") => new()
    {
        Id = 7,
        SchematicName = "Sch",
        Title = title,
        Category = "Note",
        State = "Open",
        AreaX = 10,
        AreaY = 10,
        AreaWidth = 50,
        AreaHeight = 50,
    };

    private static void WithEditor(Action<WorklogEntryEditorWindow> body)
    {
        UiTest.Run(() =>
        {
            // Placement persistence off: the window otherwise restores size and splitter ratio
            // from the developer's REAL settings file, so layout assertions would depend on how
            // they last left the editor.
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            using var bitmap = CreateBitmap();
            window.Initialize(1, CreateEntry(), bitmap);

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

    private static TextBox TitleBox(Window w) => w.FindControl<TextBox>("EditorTitleTextBox")!;

    private static Button SaveButton(Window w) => w.FindControl<Button>("EditorSaveButton")!;

    private static void SetTitle(Window w, string? text)
    {
        TitleBox(w).Text = text;

        // Avalonia raises TextChanged from a posted dispatcher job, not from the setter, so the
        // gate has not run yet at the point the setter returns.
        Dispatcher.UIThread.RunJobs();
    }

    // ------------------------------------------------------------- mandatory title

    [Fact]
    public void Clearing_the_title_disables_save_even_though_the_entry_is_dirty()
    {
        WithEditor(window =>
        {
            // Editing the title is itself what marks the window dirty, so by the time it is blank
            // there IS an unsaved change - which is exactly the state that must not be saveable.
            SetTitle(window, "Edited title");
            Assert.True(SaveButton(window).IsEnabled);

            SetTitle(window, string.Empty);

            Assert.False(SaveButton(window).IsEnabled);
        });
    }

    // Whitespace is not a title: SyncDirectFieldsToEntry trims before writing, so a title of
    // spaces would be persisted as an empty one and the gate has to agree with what the save does.
    [Theory]
    [InlineData(" ")]
    [InlineData("   ")]
    [InlineData("\t")]
    public void A_whitespace_only_title_disables_save(string title)
    {
        WithEditor(window =>
        {
            SetTitle(window, "Edited title");
            Assert.True(SaveButton(window).IsEnabled);

            SetTitle(window, title);

            Assert.False(SaveButton(window).IsEnabled);
        });
    }

    [Fact]
    public void Restoring_a_title_re_enables_save()
    {
        WithEditor(window =>
        {
            SetTitle(window, string.Empty);
            Assert.False(SaveButton(window).IsEnabled);

            SetTitle(window, "A real title");

            Assert.True(SaveButton(window).IsEnabled);
        });
    }

    // A valid title is necessary but not sufficient - an untouched entry still has nothing to
    // save, so opening a perfectly valid entry must not present an enabled Save button.
    [Fact]
    public void A_valid_but_unedited_entry_still_has_save_disabled()
    {
        WithEditor(window => Assert.False(SaveButton(window).IsEnabled));
    }

    // ------------------------------------------------------------- header layout

    // The badge sits to the RIGHT of the title box, not to its left.
    [Fact]
    public void The_id_badge_sits_to_the_right_of_the_title_box()
    {
        WithEditor(window =>
        {
            var badge = window.FindControl<Border>("EditorIdBadge")!;
            var titleBox = TitleBox(window);

            double badgeLeft = badge.TranslatePoint(new Point(0, 0), window)!.Value.X;
            double titleRight = titleBox.TranslatePoint(new Point(titleBox.Bounds.Width, 0), window)!.Value.X;

            Assert.True(badgeLeft >= titleRight - 0.5, $"badge at {badgeLeft} is not right of the title end {titleRight}");
        });
    }

    // With the badge no longer above them, the chips must start at the title box's left edge
    // rather than staying indented past where the badge used to be.
    [Fact]
    public void The_category_chips_line_up_with_the_title_box_left_edge()
    {
        WithEditor(window =>
        {
            var titleBox = TitleBox(window);
            var noteChip = window.FindControl<Border>("EditorCategoryNoteChip")!;

            double titleLeft = titleBox.TranslatePoint(new Point(0, 0), window)!.Value.X;
            double chipLeft = noteChip.TranslatePoint(new Point(0, 0), window)!.Value.X;

            Assert.True(Math.Abs(chipLeft - titleLeft) <= 1.0, $"chip left {chipLeft} does not line up with title left {titleLeft}");
        });
    }

    // The badge is on the title's own row, so it must sit beside the title rather than beside the
    // chips below it.
    [Fact]
    public void The_id_badge_shares_a_row_with_the_title_box()
    {
        WithEditor(window =>
        {
            var badge = window.FindControl<Border>("EditorIdBadge")!;
            var titleBox = TitleBox(window);

            double badgeCentreY = badge.TranslatePoint(new Point(0, badge.Bounds.Height / 2), window)!.Value.Y;
            double titleCentreY = titleBox.TranslatePoint(new Point(0, titleBox.Bounds.Height / 2), window)!.Value.Y;

            Assert.True(Math.Abs(badgeCentreY - titleCentreY) <= 4.0, $"badge centre {badgeCentreY} is not on the title row {titleCentreY}");
        });
    }

    // ------------------------------------------------------------- empty-state helper text

    // A fresh entry has no work done and no comments, so both helper lines must show - matching
    // the "No links added" line that was already there.
    [Fact]
    public void Empty_work_done_and_comments_sections_show_their_helper_text()
    {
        WithEditor(window =>
        {
            Assert.True(window.FindControl<TextBlock>("EditorNoWorkDoneText")!.IsVisible);
            Assert.True(window.FindControl<TextBlock>("EditorNoCommentsText")!.IsVisible);

            // The line that was already correct, asserted alongside so a regression in the shared
            // pattern is caught here too.
            Assert.True(window.FindControl<TextBlock>("EditorNoLinksText")!.IsVisible);
        });
    }

    [Fact]
    public void Populated_work_done_and_comments_sections_hide_their_helper_text()
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            var entry = CreateEntry();
            entry.Comments.Add(new WorklogCommentRecord { Id = 1, Text = "Checked it", Date = DateTime.Now });
            entry.WorkDoneItems.Add(new WorklogWorkDoneRecord { Id = 1, Text = "Replaced", Date = DateTime.Now, HoursSpent = 1, Cost = 2 });

            using var bitmap = CreateBitmap();
            window.Initialize(1, entry, bitmap);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();

                Assert.False(window.FindControl<TextBlock>("EditorNoCommentsText")!.IsVisible);
                Assert.False(window.FindControl<TextBlock>("EditorNoWorkDoneText")!.IsVisible);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ------------------------------------------------------------- "Show marked area"

    // The checkbox reflects the entry it was opened for, rather than always starting ticked.
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void The_show_marked_area_checkbox_reflects_the_entry(bool showMarkedArea)
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            var entry = CreateEntry();
            entry.ShowMarkedArea = showMarkedArea;

            using var bitmap = CreateBitmap();
            window.Initialize(1, entry, bitmap);

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();

                Assert.Equal(showMarkedArea, window.FindControl<CheckBox>("EditorShowMarkedAreaCheckBox")!.IsChecked);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Seeding the checkbox during Initialize must NOT mark the window dirty - the same deferred
    // event trap the title box falls into (IsCheckedChanged is raised for the initial value too).
    [Fact]
    public void Seeding_the_checkbox_does_not_mark_the_window_dirty()
    {
        WithEditor(window => Assert.False(SaveButton(window).IsEnabled));
    }

    // Toggling it IS an edit, so it must enable Save and reach disk with it.
    [Fact]
    public void Toggling_show_marked_area_marks_the_window_dirty()
    {
        WithEditor(window =>
        {
            Assert.False(SaveButton(window).IsEnabled);

            window.FindControl<CheckBox>("EditorShowMarkedAreaCheckBox")!.IsChecked = false;
            Dispatcher.UIThread.RunJobs();

            Assert.True(SaveButton(window).IsEnabled);
        });
    }

    // ------------------------------------------------------------- the initializing guard

    // Opening an entry must not leave it marked dirty, even transiently.
    //
    // Initialize raises the guard and posts the lift at Background priority precisely so its own
    // TextBox assignments' QUEUED TextChanged events run while it is still up. InitializeComponentScope
    // used to lower the guard synchronously in a finally block - and the caller runs it immediately
    // after Initialize - so those queued events arrived with the guard down and called MarkDirty.
    //
    // The Save button hid it: Initialize's posted job also reset the flag, so the button looked
    // right while the flag was wrong in between. This drives the dispatcher one priority at a time
    // to catch the state the button masked - verified against the old code, where it read dirty=True
    // with Save ENABLED at this exact point.
    [Fact]
    public void An_entry_is_never_transiently_dirty_while_opening()
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            using var bitmap = CreateBitmap();
            window.Initialize(1, CreateEntry(), bitmap);
            window.InitializeComponentScope(new[] { ("C1", "Ceramic"), ("C2", "Ceramic") });

            try
            {
                // Everything BELOW Background priority, so the queued TextChanged events run but
                // Initialize's guard-lift has not yet.
                Dispatcher.UIThread.RunJobs(DispatcherPriority.Render);

                Assert.False(
                    SaveButton(window).IsEnabled,
                    "the window was marked dirty by its own initialisation before the guard was lifted");

                Dispatcher.UIThread.RunJobs();
                Assert.False(SaveButton(window).IsEnabled);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ...and it is still editable afterwards, so the guard is genuinely lifted rather than stuck.
    [Fact]
    public void An_edit_after_opening_still_enables_save()
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow();
            window.Width = 1000;
            window.Height = 700;

            using var bitmap = CreateBitmap();
            window.Initialize(1, CreateEntry(), bitmap);
            window.InitializeComponentScope(new[] { ("C1", "Ceramic") });

            try
            {
                window.Show();
                Dispatcher.UIThread.RunJobs();

                SetTitle(window, "A different title");

                Assert.True(SaveButton(window).IsEnabled);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ------------------------------------------------------------- Save button label

    // A SAVED entry's button says "Update worklog" - the counterpart to "Add worklog" on a new
    // entry (see WorklogEditorNewEntryTests), and set explicitly in Initialize rather than left as
    // the markup's default so the two paths cannot drift from each other.
    [Fact]
    public void A_saved_entrys_save_button_says_update_worklog()
    {
        WithEditor(window => Assert.Equal("Update worklog", SaveButton(window).Content));
    }

    // ------------------------------------------------------------- blank title message

    // A SAVED entry's blank-title message explains a real inconsistency risk: the sub-lists write
    // through to disk instantly (PersistEntrySilently), so without an explanation the window and
    // the file could silently disagree about the title. This is the case
    // WorklogEditorNewEntryTests' matching test shows does NOT apply to a brand-new entry.
    [Fact]
    public void Clearing_a_saved_entrys_title_shows_why_save_is_disabled()
    {
        WithEditor(window =>
        {
            SetTitle(window, "");

            var message = window.FindControl<TextBlock>("EditorSaveFailedText")!;
            Assert.True(message.IsVisible);
            Assert.Equal("A worklog needs a title before it can be saved.", message.Text);
        });
    }

    // Retyping a title clears OUR message specifically - not any real save failure that happened
    // to be showing, which the message's own visibility toggle has to leave alone.
    [Fact]
    public void Retyping_a_title_clears_the_blank_title_message()
    {
        WithEditor(window =>
        {
            SetTitle(window, "");
            Assert.True(window.FindControl<TextBlock>("EditorSaveFailedText")!.IsVisible);

            SetTitle(window, "Restored");

            Assert.False(window.FindControl<TextBlock>("EditorSaveFailedText")!.IsVisible);
        });
    }
}
