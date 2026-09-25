using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// The label editor's save-failed and validation-failed dialogs.
//
// These two were written out separately and were IDENTICAL for 48 straight lines - same sizing,
// the same four theme-brush lookups, the same red banner, the same nested layout - differing only
// in their window title, their headline, the text used when the caller supplies no message, and
// whether a trailing hint paragraph exists at all. They are now one builder taking those four as
// arguments.
//
// That merge is exactly why they are worth testing. Four strings that used to be written in place
// are now arguments, and a swapped pair still compiles and still produces a plausible-looking
// dialog - the save dialog wearing the validation dialog's headline is invisible to the compiler
// and to every other test in this suite. So these tests assert the two are still DIFFERENT where
// they were always meant to differ, and identical where the shared builder now guarantees it.
//
// They go through BuildLabelEditor*DialogForTests, which returns the constructed Window without
// showing it. The shipped entry points end in ShowDialog, which blocks on a real window with
// nothing to dismiss it headlessly - a test driving those would hang rather than fail. The seams
// delegate to the same private builders the shipped path calls, so nothing here can pass against
// strings the application does not actually use.
[Collection("HeadlessUi")]
public sealed class LabelEditorErrorDialogTests
{
    private static TabSchematics BuildTab() => new();

    // The dialog is built from nested Borders and StackPanels, so the assertions read it back the
    // way a user sees it: every string it puts on screen, in layout order.
    //
    // This walks the LOGICAL content tree (Border.Child / StackPanel.Children) rather than calling
    // GetVisualDescendants, which returns nothing here: a Window that has never been shown has no
    // visual tree yet, since that is materialised on show. Showing one to read it back would drag
    // in the ShowDialog problem these seams exist to avoid, so the content is walked directly.
    private static string[] TextsOf(Window dialog)
    {
        var texts = new List<string>();
        Walk(dialog.Content as Control, texts);
        return texts.ToArray();
    }

    private static void Walk(Control? control, List<string> texts)
    {
        switch (control)
        {
            case null:
                return;

            case TextBlock block when !string.IsNullOrEmpty(block.Text):
                texts.Add(block.Text!);
                return;

            case Border border:
                Walk(border.Child as Control, texts);
                return;

            case Panel panel:
                foreach (var child in panel.Children)
                {
                    Walk(child as Control, texts);
                }

                return;
        }
    }

    private static Button[] ButtonsOf(Window dialog)
    {
        var buttons = new List<Button>();
        CollectButtons(dialog.Content as Control, buttons);
        return buttons.ToArray();
    }

    private static void CollectButtons(Control? control, List<Button> buttons)
    {
        switch (control)
        {
            case null:
                return;

            case Button button:
                buttons.Add(button);
                return;

            case Border border:
                CollectButtons(border.Child as Control, buttons);
                return;

            case Panel panel:
                foreach (var child in panel.Children)
                {
                    CollectButtons(child as Control, buttons);
                }

                return;
        }
    }

    [Fact]
    public void The_save_dialog_names_the_save_failure_and_carries_the_callers_message()
    {
        UiTest.Run(() =>
        {
            var dialog = BuildTab().BuildLabelEditorSaveFailedDialogForTests("Sheet 'Components' is locked.");

            Assert.Equal("Label editor save failed", dialog.Title);

            var texts = TextsOf(dialog);
            Assert.Contains("Unable to save label editor changes", texts);

            // The caller's own message is shown verbatim rather than replaced by the generic one.
            Assert.Contains("Sheet 'Components' is locked.", texts);
        });
    }

    [Fact]
    public void The_validation_dialog_names_the_validation_failure_and_carries_the_callers_message()
    {
        UiTest.Run(() =>
        {
            var dialog = BuildTab().BuildLabelEditorValidationFailedDialogForTests("Two labels share the name U1.");

            Assert.Equal("Label editor validation failed", dialog.Title);

            var texts = TextsOf(dialog);
            Assert.Contains("Unable to apply component label editor changes", texts);
            Assert.Contains("Two labels share the name U1.", texts);
        });
    }

    // THE regression this file exists for. Both dialogs come out of one builder now, so a swapped
    // argument would give the save dialog the validation dialog's wording (or vice versa) while
    // compiling cleanly and still looking like a working dialog on screen.
    [Fact]
    public void The_two_dialogs_do_not_share_a_title_or_a_headline()
    {
        UiTest.Run(() =>
        {
            var tab = BuildTab();
            var save = tab.BuildLabelEditorSaveFailedDialogForTests("x");
            var validation = tab.BuildLabelEditorValidationFailedDialogForTests("x");

            Assert.NotEqual(save.Title, validation.Title);

            Assert.Contains("Unable to save label editor changes", TextsOf(save));
            Assert.DoesNotContain("Unable to save label editor changes", TextsOf(validation));

            Assert.Contains("Unable to apply component label editor changes", TextsOf(validation));
            Assert.DoesNotContain("Unable to apply component label editor changes", TextsOf(save));
        });
    }

    // A blank message must not produce a dialog whose body is empty - the user would be told
    // something failed with no indication of what. Each dialog substitutes its OWN fallback, which
    // is the third of the four things that differ between them.
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void A_blank_message_falls_back_to_each_dialogs_own_default_text(string blank)
    {
        UiTest.Run(() =>
        {
            var tab = BuildTab();

            Assert.Contains("The label editor changes could not be saved.", TextsOf(tab.BuildLabelEditorSaveFailedDialogForTests(blank)));
            Assert.Contains("The label editor contains invalid data.", TextsOf(tab.BuildLabelEditorValidationFailedDialogForTests(blank)));
        });
    }

    // The fourth difference, and the one the shared builder handles with a null check rather than
    // with a string: the save dialog has a trailing hint about closing Excel, the validation dialog
    // has nothing to add. A null hint must OMIT the paragraph, not render an empty one - the layout
    // is a StackPanel with Spacing = 14, so a blank TextBlock leaves a visible gap above the button.
    [Fact]
    public void Only_the_save_dialog_shows_the_trailing_hint_and_the_other_omits_it_entirely()
    {
        UiTest.Run(() =>
        {
            var tab = BuildTab();

            var saveTexts = TextsOf(tab.BuildLabelEditorSaveFailedDialogForTests("boom"));
            Assert.Contains(saveTexts, text => text.Contains("open in another program"));

            // Counted rather than merely absent: an empty TextBlock would satisfy a "does not
            // contain the hint" check while still occupying a row in the panel.
            //
            // The counts are 4 and 3, not 3 and 2, because the banner's warning glyph is itself a
            // TextBlock: glyph + headline + message (+ hint). So the ONE text of difference between
            // them is the hint, which is the whole point of the assertion.
            var validationTexts = TextsOf(tab.BuildLabelEditorValidationFailedDialogForTests("boom"));
            Assert.DoesNotContain(validationTexts, text => text.Contains("open in another program"));
            Assert.Equal(4, saveTexts.Length);
            Assert.Equal(3, validationTexts.Length);
        });
    }

    // What the shared builder is FOR: everything that is not one of those four things must be
    // identical between the two, which is what stopped being true by hand once and is now
    // structural. The warning glyph, the Close button and the window sizing all come from the one
    // builder, so this fails only if someone reintroduces a second copy.
    [Fact]
    public void Both_dialogs_share_their_sizing_glyph_and_close_button()
    {
        UiTest.Run(() =>
        {
            var tab = BuildTab();
            var save = tab.BuildLabelEditorSaveFailedDialogForTests("a");
            var validation = tab.BuildLabelEditorValidationFailedDialogForTests("b");

            Assert.Equal(save.Width, validation.Width);
            Assert.Equal(save.MinWidth, validation.MinWidth);
            Assert.Equal(save.CanResize, validation.CanResize);
            Assert.Equal(save.SizeToContent, validation.SizeToContent);

            Assert.Contains("⚠", TextsOf(save));
            Assert.Contains("⚠", TextsOf(validation));

            foreach (var dialog in new[] { save, validation })
            {
                var buttons = ButtonsOf(dialog);
                Assert.Single(buttons);
                Assert.Equal("Close", buttons[0].Content as string);
            }
        });
    }
}
