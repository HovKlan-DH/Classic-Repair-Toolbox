using Avalonia;
using Avalonia.Controls;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// A worklog's comments, added and changed in the worklog editor - WorklogEntryEditorWindow.SubLists.cs. A comment is
// MARKUP (RichNote), typed in the "Add comment" window's formatted box; the window stands in through its seam, as a test
// cannot show a dialog.
//
// HeadlessUi: it builds a window. On a NEW worklog, so nothing is written anywhere - a draft touches no workbook until
// Save (WorklogEditorDiscardGuardTests).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class WorklogEditorCommentsTests
{
    private static List<string> CommentTexts(WorklogEntryEditorWindow window) =>
        window.GetControl<ItemsControl>("EditorCommentsList").Items.OfType<WorklogCommentRow>().Select(row => row.Text).ToList();

    // ###########################################################################################
    // *** A COMMENT'S FIRST LINE KEEPS THE INDENT THAT SAYS WHAT IT IS (code review, 2026-11-08). *** The "Add comment"
    // window trims its markup by RichNote.TrimMarkup, and the editor trimmed it once more with Trim(): a first line nested
    // once - "  - swapped U7" - was stored a bullet of the first level, in the list, the export and every later edit.
    // A comment emptied is no change.
    // ###########################################################################################
    [Fact]
    public async Task A_comment_added_or_changed_keeps_its_first_lines_indent()
    {
        await UiTest.RunAsync(async () =>
        {
            using var placementScope = WorklogEntryEditorWindow.SuppressWindowPlacementPersistence();
            using var bitmap = new RenderTargetBitmap(new PixelSize(40, 40), new Vector(96, 96));

            var window = new WorklogEntryEditorWindow { Width = 1000, Height = 700 };
            window.InitializeForNewEntry(0, "Sheet 1", new Rect(10, 20, 30, 40), bitmap);

            var answers = new Queue<string?>(["  - swapped U7\n- U8\n", "\n  - socketed U7  ", "   "]);
            var asked = new List<string?>();
            window.AskCommentOverrideForTests = text =>
            {
                asked.Add(text);
                return Task.FromResult(answers.Dequeue());
            };

            window.Show();

            try
            {
                Dispatcher.UIThread.RunJobs();

                // A new worklog holds its own "Worklog created" already - only the comment added is this test's.
                List<string> before = CommentTexts(window);

                await window.AddCommentAsync();
                Assert.Equal(["  - swapped U7\n- U8"], CommentTexts(window).Except(before));

                int id = window.GetControl<ItemsControl>("EditorCommentsList").Items.OfType<WorklogCommentRow>()
                    .Single(row => row.Text == "  - swapped U7\n- U8").Id;
                await window.EditCommentAsync(id);
                Assert.Equal(["  - socketed U7"], CommentTexts(window).Except(before));

                // Emptied: the comment stays as it was.
                await window.EditCommentAsync(id);
                Assert.Equal(["  - socketed U7"], CommentTexts(window).Except(before));

                Assert.Equal([null, "  - swapped U7\n- U8", "  - socketed U7"], asked);
            }
            finally
            {
                window.Close();
                Dispatcher.UIThread.RunJobs();
            }
        });
    }
}
