using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The delete-workbook confirmation modal's keyboard behaviour.
//
// This window is the one modal in the app where Enter must NOT confirm: its "submit" is a permanent
// delete of a workbook and everything in it. Enter and Escape both CANCEL, so a reflexive keypress
// aimed at dismissing what looks like just another dialog can never destroy anything - deleting
// stays a deliberate click on the button.
//
// That guarantee is entirely about EVENT ROUTING, which is why it is worth a test: subscribed the
// ordinary bubbling way (`this.KeyDown += ...`), the handler never ran once the Delete button had
// focus, because Avalonia's Button handles Enter itself and marks the event handled - so Enter on a
// focused Delete button deleted the workbook, the exact opposite of the promise. The fix is the
// Tunnel route, and these tests fail against the bubbling version.
[Collection("HeadlessUi")]
public sealed class DeleteWorkbookWindowTests
{
    private static DeleteWorkbookWindow BuildWindow()
    {
        var window = new DeleteWorkbookWindow();
        window.Initialize(new WorkbookRecord
        {
            Id = 3,
            BoardKey = "Commodore 64|250469 (short board)",
            Title = "Black screen",
            Status = "Open",
        });

        // Show() so the visual tree is built and focus can actually land on a button - without it
        // there is nothing for the key to be routed through.
        window.Show();
        return window;
    }

    private static Button ButtonWithContent(DeleteWorkbookWindow window, string content) =>
        window.GetVisualDescendants()
            .OfType<Button>()
            .First(b => (b.Content as string) == content);

    // A REAL keypress through the headless input stack, not a hand-built RaiseEvent.
    //
    // That distinction is the difference between a test that catches this bug and one that does not:
    // raising KeyDownEvent on the focused button directly skips Avalonia's own key handling for that
    // button, so the button never marks Enter handled and even the broken bubbling version passes.
    // Driving the window's input surface routes the key exactly as the running app does - tunnel
    // first, then the button, then bubble - which is what the fix depends on.
    private static void PressKey(DeleteWorkbookWindow window, Key key, PhysicalKey physicalKey) =>
        window.KeyPress(key, RawInputModifiers.None, physicalKey, keySymbol: null);

    [Fact]
    public void Enter_cancels_even_when_the_delete_button_has_focus()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            // THE regression this file exists for. A focused Button consumes Enter on the bubbling
            // route, so before the Tunnel fix this deleted the workbook.
            var deleteButton = ButtonWithContent(window, "Delete workbook");

            // Watching the button's OWN Click is what makes this test able to fail. Asserting only
            // that the window closed proves nothing: it closes either way, because the cancel
            // handler runs too - the difference is whether OnDeleteClick ALSO ran and confirmed the
            // delete. Against the bubbling version this fires; against the Tunnel fix it does not.
            bool deleteConfirmed = false;
            deleteButton.Click += (_, _) => deleteConfirmed = true;

            deleteButton.Focus();
            PressKey(window, Key.Enter, PhysicalKey.Enter);

            Assert.False(deleteConfirmed);
            Assert.False(window.IsVisible);
        });
    }

    [Fact]
    public void Enter_cancels_when_the_cancel_button_has_focus()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            bool deleteConfirmed = false;
            ButtonWithContent(window, "Delete workbook").Click += (_, _) => deleteConfirmed = true;

            ButtonWithContent(window, "Cancel").Focus();
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
            ButtonWithContent(window, "Delete workbook").Click += (_, _) => deleteConfirmed = true;

            PressKey(window, Key.Escape, PhysicalKey.Escape);

            Assert.False(deleteConfirmed);
            Assert.False(window.IsVisible);
        });
    }

    // The confirmation names the workbook on its OWN bold line, since several cards can be on
    // screen at once and "are you sure?" alone does not say which one is about to be lost.
    [Fact]
    public void The_confirmation_names_the_workbook_being_deleted_on_its_own_bold_line()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            var nameBlock = window.GetControl<TextBlock>("WorkbookNameText");

            Assert.Equal("#3 · Black screen", nameBlock.Text);
            Assert.Equal(FontWeight.Bold, nameBlock.FontWeight);
        });
    }

    // The surrounding copy: an intro sentence naming what is about to happen, then the body
    // explaining what is lost, then a final warning that it cannot be undone - each its own
    // TextBlock so the name can sit alone between them rather than buried in one run-on sentence.
    [Fact]
    public void The_surrounding_text_explains_and_warns_before_and_after_the_name()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            var textBlocks = window.GetVisualDescendants().OfType<TextBlock>().ToList();

            Assert.Contains(textBlocks, t => t.Text == "This permanently deletes the workbook:");

            var body = Assert.Single(textBlocks, t =>
                t.Text != null && t.Text.Contains("work done, comments, photos and files", StringComparison.Ordinal));
            Assert.Equal(
                "Everything recorded to the workbook will be deleted; e.g. work done, comments, photos and files.",
                body.Text);

            Assert.Contains(textBlocks, t => t.Text == "This cannot be undone!");
        });
    }
}
