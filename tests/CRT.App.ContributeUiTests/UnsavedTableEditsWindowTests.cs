using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The "Unsaved edits" prompt of the table editor (2026-09-24) - leaving a Maintainer tab table, and
// since CRT 4.0 step 1 (2026-10-25, a draft saves itself) leaving the Drafts tab's table only after
// its save was refused or failed.
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
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.LeavingSubmission);

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
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.LeavingSubmission);

            bool discarded = false;
            ButtonWithContent(window, "Discard edits").Click += (_, _) => discarded = true;

            window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, keySymbol: null);

            Assert.False(discarded);
            Assert.False(window.IsVisible);
        });
    }

    [Fact]
    public void Leaving_a_submissions_table_offers_Save_Discard_and_Cancel()
    {
        UiTest.Run(() =>
        {
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.LeavingSubmission);

            Assert.True(ButtonWithContent(window, "Save changes").IsVisible);
            Assert.True(ButtonWithContent(window, "Discard edits").IsVisible);
            Assert.True(ButtonWithContent(window, "Cancel").IsVisible);

            window.Close();
        });
    }

    // A board's table on the Maintainer tab's Boards screen (2026-10-03): Save is offered, and the
    // message says what Save does there - a reason is asked for and the change goes STRAIGHT TO BETA
    // (owner decision, the same day), so nobody saves expecting a review first.
    [Fact]
    public void Leaving_a_boards_table_offers_Save_and_says_it_goes_straight_to_BETA()
    {
        UiTest.Run(() =>
        {
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.LeavingBoard);

            Assert.True(ButtonWithContent(window, "Save changes").IsVisible);
            Assert.True(ButtonWithContent(window, "Discard edits").IsVisible);
            Assert.Contains("straight to BETA", window.GetControl<TextBlock>("MessageText").Text);
            Assert.Contains("asked for a reason", window.GetControl<TextBlock>("MessageText").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // A DRAFT'S TABLE IS ASKED ABOUT ONLY WHEN ITS SAVE COULD NOT BE MADE - it saves itself, and the
    // save was just tried - so there is no Save to offer: Discard or Cancel, and why. Offering one
    // made "Close table" a loop (reported, 2026-09-24).
    // ###########################################################################################
    [Fact]
    public void Leaving_a_drafts_table_whose_save_was_refused_offers_NO_Save_and_says_why()
    {
        UiTest.Run(() =>
        {
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.DraftChangedOnDisk);

            Assert.False(ButtonWithContent(window, "Save changes").IsVisible);
            Assert.True(ButtonWithContent(window, "Discard edits").IsVisible);
            Assert.True(ButtonWithContent(window, "Cancel").IsVisible);
            // Said of the draft, not of the table (CRT 4.0 step 2): the component pane asks it too.
            Assert.Equal("Changes not saved", window.GetControl<TextBlock>("HeadingText").Text);
            Assert.Contains("changed somewhere else", window.GetControl<TextBlock>("MessageText").Text);
            Assert.Contains("use the saved version or to keep yours", window.GetControl<TextBlock>("MessageText").Text);

            window.Close();
        });
    }

    [Fact]
    public void Leaving_a_drafts_table_whose_save_failed_offers_NO_Save_and_says_why()
    {
        // That save just failed, and offering it again would loop as the refused one once did. In
        // CRT 4.0 the draft is a board file no spreadsheet program edits, so the words name none.
        UiTest.Run(() =>
        {
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.DraftNotWritten);

            Assert.False(ButtonWithContent(window, "Save changes").IsVisible);
            Assert.True(ButtonWithContent(window, "Discard edits").IsVisible);
            Assert.Contains("could not be written", window.GetControl<TextBlock>("MessageText").Text);
            Assert.Contains("Try again", window.GetControl<TextBlock>("MessageText").Text);
            Assert.DoesNotContain("Excel", window.GetControl<TextBlock>("MessageText").Text);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** UNDER "CRT HAS TO BE UPDATED" (code review, 2026-10-09). *** A Maintainer tab table's Save
    // sends to the server, which turns this CRT away - so leaving one offers Discard or Cancel and
    // says why. (The notices that held another editor's save back - "SavingElsewhere", and its form
    // under a covered Drafts tab - went with CRT 4.0 step 1: every editor changes the one draft held
    // in memory, so nothing is held back.)
    // ###########################################################################################
    [Fact]
    public void Leaving_a_Maintainer_table_while_CRT_has_to_be_updated_offers_NO_Save_and_says_why()
    {
        UiTest.Run(() =>
        {
            UnsavedTableEditsWindow window = BuildWindow(UnsavedTableEditsPrompt.LeavingUpdateRequired);

            Assert.False(ButtonWithContent(window, "Save changes").IsVisible);
            Assert.True(ButtonWithContent(window, "Discard edits").IsVisible);
            Assert.True(ButtonWithContent(window, "Cancel").IsVisible);
            Assert.Contains("cannot be sent until CRT is updated", window.GetControl<TextBlock>("MessageText").Text);

            window.Close();
        });
    }

}
