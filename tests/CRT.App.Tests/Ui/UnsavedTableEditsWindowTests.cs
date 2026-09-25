using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The "Unsaved edits" prompt of the Drafts tab's table editor (2026-09-24).
//
// "Discard edits" throws away work that exists nowhere but in memory, so like the discard and
// delete dialogs this is a modal where Enter must NOT confirm: Enter and Escape both CANCEL.
// That is a matter of EVENT ROUTING - a focused Button handles Enter itself on the bubbling route
// - so the Enter test focuses the Discard button and presses a REAL key through the headless
// input stack, and fails against a bubbling handler. See DeleteWorklogWindowTests.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class UnsavedTableEditsWindowTests
{
    private static UnsavedTableEditsWindow BuildWindow(UnsavedTableEditsPrompt prompt)
    {
        var window = new UnsavedTableEditsWindow();
        window.Initialize(prompt);

        // Shown so the visual tree exists and focus can land on a button.
        window.Show();
        return window;
    }

    private static Button ButtonWithContent(Window window, string content) =>
        window.GetVisualDescendants().OfType<Button>().First(button => (button.Content as string) == content);

    [Fact]
    public void Enter_cancels_even_when_the_Discard_button_has_focus()
    {
        UiTest.Run(() =>
        {
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.Leaving);

            Button discard = ButtonWithContent(window, "Discard edits");
            bool discarded = false;
            discard.Click += (_, _) => discarded = true;

            discard.Focus();
            window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);

            Assert.False(discarded);
            Assert.False(window.IsVisible);
        });
    }

    [Fact]
    public void Escape_cancels()
    {
        UiTest.Run(() =>
        {
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.Leaving);

            bool discarded = false;
            ButtonWithContent(window, "Discard edits").Click += (_, _) => discarded = true;

            window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, keySymbol: null);

            Assert.False(discarded);
            Assert.False(window.IsVisible);
        });
    }

    [Fact]
    public void Leaving_the_table_offers_Save_Discard_and_Cancel()
    {
        UiTest.Run(() =>
        {
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.Leaving);

            Assert.True(ButtonWithContent(window, "Save changes").IsVisible);
            Assert.True(ButtonWithContent(window, "Discard edits").IsVisible);
            Assert.True(ButtonWithContent(window, "Cancel").IsVisible);

            window.Close();
        });
    }

    [Fact]
    public void Reloading_offers_NO_Save_since_saving_is_what_was_just_refused()
    {
        UiTest.Run(() =>
        {
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.Reloading);

            Assert.False(ButtonWithContent(window, "Save changes").IsVisible);
            Assert.Contains("Reloading", window.GetControl<TextBlock>("MessageText").Text);

            window.Close();
        });
    }

    [Fact]
    public void Leaving_a_table_whose_draft_changed_on_disk_offers_NO_Save_and_says_why()
    {
        // A save there is refused, so offering one made "Close table" a loop (reported).
        UiTest.Run(() =>
        {
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.DraftChangedOnDisk);

            Assert.False(ButtonWithContent(window, "Save changes").IsVisible);
            Assert.True(ButtonWithContent(window, "Discard edits").IsVisible);
            Assert.True(ButtonWithContent(window, "Cancel").IsVisible);
            Assert.Contains("changed outside this table", window.GetControl<TextBlock>("MessageText").Text);
            Assert.Contains("Save to draft", window.GetControl<TextBlock>("MessageText").Text);

            window.Close();
        });
    }

    [Fact]
    public void Saving_elsewhere_is_a_notice_with_Cancel_alone_that_sends_you_to_the_Drafts_tab()
    {
        // Owner's design (2026-09-24): from another tab you may not remember what you did in
        // the table, so nothing about it is decided here - no Save, no Discard.
        UiTest.Run(() =>
        {
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.SavingElsewhere);

            Assert.False(ButtonWithContent(window, "Save changes").IsVisible);
            Assert.False(ButtonWithContent(window, "Discard edits").IsVisible);
            Assert.True(ButtonWithContent(window, "Cancel").IsVisible);

            Assert.Equal("Unsaved table edits in the Drafts tab", window.Title);
            Assert.Equal("Unsaved table edits in the Drafts tab", window.GetControl<TextBlock>("HeadingText").Text);
            Assert.Contains("Go to the Drafts tab, save or discard the table's edits", window.GetControl<TextBlock>("MessageText").Text);

            window.Close();
        });
    }

    [Fact]
    public void Leaving_a_table_whose_draft_is_held_open_in_Excel_offers_NO_Save_and_says_why()
    {
        // That save would fail, and offering it would loop as the changed-on-disk one once did.
        UiTest.Run(() =>
        {
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.DraftOpenElsewhere);

            Assert.False(ButtonWithContent(window, "Save changes").IsVisible);
            Assert.True(ButtonWithContent(window, "Discard edits").IsVisible);
            Assert.Contains("open in Excel", window.GetControl<TextBlock>("MessageText").Text);

            window.Close();
        });
    }
}
