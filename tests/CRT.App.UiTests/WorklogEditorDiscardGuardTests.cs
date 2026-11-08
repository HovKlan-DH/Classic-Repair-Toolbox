using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The "you have not saved this worklog yet" guard on Cancel / Escape / the title-bar close.
//
// WHY IT EXISTS. Sub-list changes behave differently on a draft than on a saved worklog, and nothing
// on screen says so. On a SAVED worklog, adding a comment or a photo writes to disk immediately, so
// Escape can only cost a half-typed Title or Description. On a NEW one nothing is written until the
// user clicks "Add worklog" - by design, so a cancelled draft leaves nothing behind - which means
// Escape discards every comment, photo and typed field with it. After adding a comment the way you
// would on a saved worklog, it FEELS saved, and Escape then feels like closing a finished thing.
// Reported as real data loss.
//
// The fix does NOT change when anything is written: a draft still writes nothing until Save. It only
// asks before throwing that work away.
//
// WHAT THESE TESTS COVER. The DECISION - whether the confirmation appears at all - which is the
// whole of the correctness here: it must fire for a draft holding anything and never for a saved
// worklog. The dialog itself asks one fixed question and takes no arguments, so there is nothing in
// it left to test; its click-through cannot be driven headlessly (ShowDialog resolves with no user)
// and is verified by running the app.
[Collection("HeadlessUi")]
public sealed class WorklogEditorDiscardGuardTests
{
    private static Bitmap CreateBitmap()
    {
        var target = new RenderTargetBitmap(new PixelSize(40, 40), new Vector(96, 96));
        return target;
    }

    /// <summary>A shown editor on a NEW worklog - the draft case the guard exists for.</summary>
    private static void WithNewEntryEditor(Action<WorklogEntryEditorWindow> body)
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow { Width = 1000, Height = 700 };
            using var bitmap = CreateBitmap();

            // Workbook 0 does not exist, so nothing is written - which is the point: a draft touches
            // no workbook until Save.
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

    /// <summary>A shown editor on an already-SAVED worklog.</summary>
    private static void WithSavedEntryEditor(WorklogEntryRecord entry, Action<WorklogEntryEditorWindow> body)
    {
        UiTest.Run(() =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();

            var window = new WorklogEntryEditorWindow { Width = 1000, Height = 700 };
            using var bitmap = CreateBitmap();

            window.Initialize(0, entry, bitmap);

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

    private static TextBox Box(WorklogEntryEditorWindow window, string name) =>
        window.FindControl<TextBox>(name)!;

    // Opening a new worklog and immediately backing out must NOT prompt - there is nothing to
    // protect, and a dialog on the most common way to abandon an empty form is pure friction.
    [Fact]
    public void An_untouched_new_worklog_closes_without_asking()
    {
        WithNewEntryEditor(window =>
        {
            Assert.False(window.WouldDiscardEnteredWorkForTests());
        });
    }

    // A typed title is read from the CONTROL, not from the record: the direct fields are only synced
    // into the record on a save, so a user who types a title and hits Escape has it nowhere else.
    // This is the case the guard most needs to catch, and the one a record-only check would miss.
    [Fact]
    public void A_typed_title_alone_is_enough_to_ask()
    {
        WithNewEntryEditor(window =>
        {
            Box(window, "EditorTitleTextBox").Text = "Dead VIC";
            Dispatcher.UIThread.RunJobs();

            Assert.True(window.WouldDiscardEnteredWorkForTests());
        });
    }

    [Fact]
    public void A_typed_description_alone_is_enough_to_ask()
    {
        WithNewEntryEditor(window =>
        {
            Box(window, "EditorDescriptionTextBox").Text = "Traced it to the PLA";
            Dispatcher.UIThread.RunJobs();

            Assert.True(window.WouldDiscardEnteredWorkForTests());
        });
    }

    // Whitespace is not work. SyncDirectFieldsToEntry trims before writing, so a title of spaces
    // would be saved as empty - the guard has to agree with what a save would actually keep.
    [Fact]
    public void Whitespace_in_the_fields_does_not_count_as_work()
    {
        WithNewEntryEditor(window =>
        {
            Box(window, "EditorTitleTextBox").Text = "   ";
            Box(window, "EditorDescriptionTextBox").Text = "\t ";
            Dispatcher.UIThread.RunJobs();

            Assert.False(window.WouldDiscardEnteredWorkForTests());
        });
    }

    // THE REPORTED CASE. A photo (or a comment - same path) added to a new worklog behaves on screen
    // exactly like the instant save it would be on a SAVED worklog, so it feels committed. On a
    // draft it lives only in memory, and Escape discards it along with its bytes.
    //
    // Driven through the attachment seam, which appends to the in-memory list without the file copy
    // a real add would do - which is precisely the state under test: a record the user can see and
    // believes is stored, with nothing on disk behind it.
    [Fact]
    public void A_photo_added_to_a_new_worklog_is_enough_to_ask()
    {
        WithNewEntryEditor(window =>
        {
            Assert.False(window.WouldDiscardEnteredWorkForTests());

            window.AddAttachmentRecordForTests(photos: true, id: 1, fileName: "1_rail.png", comment: "4.8V");

            Assert.True(window.WouldDiscardEnteredWorkForTests());
        });
    }

    // A SAVED worklog never prompts. Escape there is a cheap, well-understood way to back out of
    // half-typed edits - its sub-lists are already on disk - and prompting on every one of those
    // would train the user to dismiss the dialog, which would then be dismissed on the one that
    // mattered.
    [Fact]
    public void A_saved_worklog_never_asks_even_with_edits_pending()
    {
        var saved = new WorklogEntryRecord
        {
            Id = 4,
            SchematicName = "Sheet 1",
            Title = "Bad cap",
            Description = "Leaking",
            Category = "Issue",
            State = "Open"
        };

        WithSavedEntryEditor(saved, window =>
        {
            Assert.False(window.WouldDiscardEnteredWorkForTests());

            Box(window, "EditorTitleTextBox").Text = "Bad cap - half-typed change";
            Dispatcher.UIThread.RunJobs();

            Assert.False(window.WouldDiscardEnteredWorkForTests());
        });
    }

}
