using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The delete-WORKLOG confirmation modal's keyboard behaviour - the per-entry twin of
// DeleteWorkbookWindowTests, and it exists for exactly the same reason.
//
// Deleting a worklog removes its row from entries.json AND its whole attachment folder, photos and
// files included, with nothing to undo it. So like the workbook dialog, this is the one shape of
// modal in the app where Enter must NOT confirm: Enter and Escape both CANCEL, and deleting stays a
// deliberate click on the button.
//
// That guarantee is entirely about EVENT ROUTING. Subscribed the ordinary bubbling way
// (`this.KeyDown += ...`) the handler never runs once the Delete button has focus, because
// Avalonia's Button handles Enter itself and marks the event handled - so Enter on a focused Delete
// button would delete the worklog, the exact opposite of the promise. The fix is the Tunnel route,
// and these tests fail against a bubbling version.
[Collection("HeadlessUi")]
public sealed class DeleteWorklogWindowTests
{
    private static DeleteWorklogWindow BuildWindow(string title = "Cracked trace at CN2")
    {
        var window = new DeleteWorklogWindow();
        window.Initialize(new WorklogEntryRecord
        {
            Id = 4,
            SchematicName = "Motherboard",
            Title = title,
            Category = "Issue",
            State = "Open",
        });

        // Show() so the visual tree is built and focus can actually land on a button - without it
        // there is nothing for the key to be routed through.
        window.Show();
        return window;
    }

    private static Button ButtonWithContent(DeleteWorklogWindow window, string content) =>
        window.GetVisualDescendants()
            .OfType<Button>()
            .First(b => (b.Content as string) == content);

    // A REAL keypress through the headless input stack, not a hand-built RaiseEvent - raising
    // KeyDownEvent on the focused button directly skips Avalonia's own key handling for that
    // button, so the button never marks Enter handled and even a broken bubbling version passes.
    private static void PressKey(DeleteWorklogWindow window, Key key, PhysicalKey physicalKey) =>
        window.KeyPress(key, RawInputModifiers.None, physicalKey, keySymbol: null);

    [Fact]
    public void Enter_cancels_even_when_the_delete_button_has_focus()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            var deleteButton = ButtonWithContent(window, "Delete worklog");

            // Watching the button's OWN Click is what makes this test able to fail. Asserting only
            // that the window closed proves nothing: it closes either way, because the cancel
            // handler runs too - the difference is whether OnDeleteClick ALSO ran and confirmed.
            bool deleteConfirmed = false;
            deleteButton.Click += (_, _) => deleteConfirmed = true;

            deleteButton.Focus();
            PressKey(window, Key.Enter, PhysicalKey.Enter);

            Assert.False(deleteConfirmed);
            Assert.False(window.IsVisible);
        });
    }

    [Fact]
    public void Escape_cancels()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            bool deleteConfirmed = false;
            ButtonWithContent(window, "Delete worklog").Click += (_, _) => deleteConfirmed = true;

            PressKey(window, Key.Escape, PhysicalKey.Escape);

            Assert.False(deleteConfirmed);
            Assert.False(window.IsVisible);
        });
    }

    // The confirmation names the worklog by the SAME "#N · Title" the entry's card, its board
    // pill and the exported PDF all show, so the thing named here is recognisably the thing that
    // was clicked - several cards are on screen at once and "are you sure?" alone does not say
    // which one is about to be lost.
    [Fact]
    public void The_confirmation_names_the_worklog_being_deleted_on_its_own_bold_line()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            var nameBlock = window.GetControl<TextBlock>("WorklogNameText");

            Assert.Equal("#4 · Cracked trace at CN2", nameBlock.Text);
            Assert.Equal(FontWeight.Bold, nameBlock.FontWeight);
        });
    }

    // An untitled worklog is still a real record with real attachments, and "#7 · " trailing off
    // into nothing reads as a rendering fault rather than as a worklog nobody named.
    [Fact]
    public void An_untitled_worklog_is_named_as_untitled_rather_than_left_blank()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow(title: "   ");

            Assert.Equal("#4 · (untitled)", window.GetControl<TextBlock>("WorklogNameText").Text);
        });
    }

    // The copy says WORKLOG throughout, not "workbook" or "entry": this dialog is a near-copy of
    // DeleteWorkbookWindow, and a copy that still names the workbook would tell the user they are
    // about to lose the whole job when they are deleting one line of it.
    [Fact]
    public void The_surrounding_text_is_about_the_worklog_and_warns_that_it_cannot_be_undone()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            var textBlocks = window.GetVisualDescendants().OfType<TextBlock>().ToList();

            Assert.Contains(textBlocks, t => t.Text == "This permanently deletes the worklog:");

            var body = Assert.Single(textBlocks, t =>
                t.Text != null && t.Text.Contains("work done, comments, photos and files", StringComparison.Ordinal));
            Assert.Equal(
                "Everything recorded to the worklog will be deleted; e.g. work done, comments, photos and files.",
                body.Text);

            Assert.Contains(textBlocks, t => t.Text == "This cannot be undone!");

            Assert.DoesNotContain(textBlocks, t =>
                t.Text != null && t.Text.Contains("workbook", StringComparison.OrdinalIgnoreCase));
        });
    }
}
