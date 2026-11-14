using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;
using System.Linq;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE FORMATTED NOTE BOX (RichNoteTextBox; owner request, 2026-11-07: "a very simple WYSIWYG ... Bold, Italic,
// Underscore, Strikethrough, Bulletlist (nested)") - in a shown window, driven by its keys and its buttons as a person
// drives it: the note handed in as markup shows as its words in their styles, every edit writes the markup back.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class RichNoteTextBoxTests
{
    private static (Window Window, RichNoteTextBox Box) Shown(string markup)
    {
        var box = new RichNoteTextBox { Markup = markup };
        var window = new Window { Content = box, Width = 500, Height = 300 };
        window.Show();
        Assert.True(box.Focus());
        return (window, box);
    }

    private static void Press(Window window, Key key, RawInputModifiers modifiers = RawInputModifiers.None)
    {
        window.KeyPress(key, modifiers, PhysicalKey.None, keySymbol: null);
        window.KeyRelease(key, modifiers, PhysicalKey.None, keySymbol: null);
    }

    [Fact]
    public void A_note_handed_in_shows_its_words_in_their_styles_and_its_presenter_draws_them()
    {
        UiTest.Run(() =>
        {
            (Window window, RichNoteTextBox box) = Shown("Recapped **the PSU**\n- fuse");

            try
            {
                Assert.Equal("Recapped the PSU\n• fuse", box.Text);
                Assert.Equal(RichNoteStyle.Bold, box.CharacterStylesForTests[9]);
                Assert.Equal(RichNoteStyle.None, box.CharacterStylesForTests[0]);

                RichNoteTextPresenter presenter = box.GetVisualDescendants().OfType<RichNoteTextPresenter>().Single();
                Assert.Equal(box.CharacterStylesForTests, presenter.CharacterStyles);

                // A note typed before formatting existed reads as it was typed.
                box.Markup = "5 * 2 and C64_PSU";
                Assert.Equal("5 * 2 and C64_PSU", box.Text);
                Assert.All(box.CharacterStylesForTests, style => Assert.Equal(RichNoteStyle.None, style));
            }
            finally
            {
                window.Close();
            }
        });
    }

    [Fact]
    public void Ctrl_B_on_a_selection_makes_it_bold_and_again_takes_it_off()
    {
        UiTest.Run(() =>
        {
            (Window window, RichNoteTextBox box) = Shown("the PSU");

            try
            {
                box.SelectionStart = 4;
                box.SelectionEnd = 7;

                Press(window, Key.B, RawInputModifiers.Control);
                Assert.Equal("the **PSU**", box.Markup);

                Press(window, Key.I, RawInputModifiers.Control);
                Press(window, Key.U, RawInputModifiers.Control);
                Press(window, Key.X, RawInputModifiers.Control | RawInputModifiers.Shift);
                Assert.Equal("the ***__~~PSU~~__***", box.Markup);
                Assert.Equal("the PSU", box.Text);

                Press(window, Key.B, RawInputModifiers.Control);
                Assert.Equal("the *__~~PSU~~__*", box.Markup);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // With nothing selected, Ctrl+B is for what is typed next; typing on after a bold word goes on in bold.
    [Fact]
    public void A_style_asked_for_with_nothing_selected_is_what_is_typed_next()
    {
        UiTest.Run(() =>
        {
            (Window window, RichNoteTextBox box) = Shown("Dead ");

            try
            {
                box.CaretIndex = 5;
                Press(window, Key.B, RawInputModifiers.Control);
                window.KeyTextInput("SID");

                Assert.Equal("Dead **SID**", box.Markup);

                Press(window, Key.B, RawInputModifiers.Control);
                window.KeyTextInput(" chip");
                Assert.Equal("Dead **SID** chip", box.Markup);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // *** A CHANGE IS WHERE THE CARET WAS, NEXT TO THE SAME CHARACTER TOO (code review, 2026-11-08). *** The styles
    // followed a change by comparing the text before and after it, which cannot tell where a character went beside one
    // like it: "b" typed after a bold "a" in "ab" read as typed after the plain "b", and came out plain; Ctrl+B there and
    // "b" typed was taken for typing elsewhere, the bold forgotten; Backspace or Delete beside a twin took the twin's
    // style away instead of its own.
    // ###########################################################################################
    [Fact]
    public void Typing_or_deleting_beside_the_same_character_takes_the_style_where_the_caret_is()
    {
        UiTest.Run(() =>
        {
            (Window window, RichNoteTextBox box) = Shown("**a**b");

            try
            {
                box.CaretIndex = 1;
                window.KeyTextInput("b");
                Assert.Equal("**ab**b", box.Markup);

                box.Markup = "ab";
                box.CaretIndex = 1;
                Press(window, Key.B, RawInputModifiers.Control);
                window.KeyTextInput("b");
                Assert.Equal("a**b**b", box.Markup);

                // The bold "b" taken out by Backspace after it, and by Delete before it - the plain one stays.
                box.Markup = "a**b**b";
                box.CaretIndex = 2;
                Press(window, Key.Back);
                Assert.Equal("ab", box.Markup);
                Assert.All(box.CharacterStylesForTests, style => Assert.Equal(RichNoteStyle.None, style));

                box.Markup = "a**b**b";
                box.CaretIndex = 1;
                Press(window, Key.Delete);
                Assert.Equal("ab", box.Markup);
                Assert.All(box.CharacterStylesForTests, style => Assert.Equal(RichNoteStyle.None, style));
            }
            finally
            {
                window.Close();
            }
        });
    }

    [Fact]
    public void The_format_buttons_do_what_their_keys_do()
    {
        UiTest.Run(() =>
        {
            (Window window, RichNoteTextBox box) = Shown("one\ntwo");

            try
            {
                Button bold = box.GetVisualDescendants().OfType<Button>().Single(button => button.Name == "PART_BoldButton");
                Button bullets = box.GetVisualDescendants().OfType<Button>().Single(button => button.Name == "PART_BulletButton");

                box.SelectionStart = 0;
                box.SelectionEnd = 3;
                bold.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
                Assert.Equal("**one**\ntwo", box.Markup);

                box.SelectionStart = 0;
                box.SelectionEnd = box.Text!.Length;
                bullets.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
                Assert.Equal("- **one**\n- two", box.Markup);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // *** A FORMAT BUTTON SHOWS PRESSED WHILE A PRESS WOULD TURN ITS STYLE OFF (owner decision, 2026-11-08). *** The
    // selection all bold: Bold pressed, and a press takes it off; partly bold: not pressed, as a press puts it on; with
    // nothing selected, what the next character typed takes - Ctrl+B there presses it at once. Bullets by the line.
    // ###########################################################################################
    [Fact]
    public void The_format_buttons_show_pressed_what_a_press_would_take_off()
    {
        UiTest.Run(() =>
        {
            (Window window, RichNoteTextBox box) = Shown("the **PSU** and\n- fuse");

            try
            {
                Button Named(string name) => box.GetVisualDescendants().OfType<Button>().Single(button => button.Name == name);
                static bool IsOn(Button button) => button.Classes.Contains("On");
                Button bold = Named("PART_BoldButton");
                Button italic = Named("PART_ItalicButton");
                Button bullets = Named("PART_BulletButton");

                // "PSU", all bold.
                box.SelectionStart = 4;
                box.SelectionEnd = 7;
                Assert.True(IsOn(bold));
                Assert.False(IsOn(italic));
                Assert.False(IsOn(bullets));

                bold.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
                Assert.Equal("the PSU and\n- fuse", box.Markup);
                Assert.False(IsOn(bold));

                // "e PS", partly bold after Ctrl+B on "PS" - not pressed.
                box.SelectionStart = 4;
                box.SelectionEnd = 6;
                Press(window, Key.B, RawInputModifiers.Control);
                box.SelectionStart = 2;
                box.SelectionEnd = 6;
                Assert.False(IsOn(bold));

                // Nothing selected: the caret after the bold "PS" types bold; Ctrl+B asks for plain, and the button follows.
                box.SelectionStart = 6;
                box.SelectionEnd = 6;
                box.CaretIndex = 6;
                Assert.True(IsOn(bold));
                Press(window, Key.B, RawInputModifiers.Control);
                Assert.False(IsOn(bold));

                // On the bullet's line.
                box.CaretIndex = box.Text!.Length;
                Assert.True(IsOn(bullets));
                Assert.False(IsOn(bold));
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A press that changes nothing - Bold on a bullet's start alone, which takes no style - leaves no step to undo: Ctrl+Z
    // takes back the last real change (code review, 2026-11-08).
    [Fact]
    public void A_press_that_changes_nothing_leaves_no_step_to_undo()
    {
        UiTest.Run(() =>
        {
            (Window window, RichNoteTextBox box) = Shown("- fuse");

            try
            {
                box.CaretIndex = box.Text!.Length;
                window.KeyTextInput("s");
                Assert.Equal("- fuses", box.Markup);

                box.SelectionStart = 0;
                box.SelectionEnd = 2;
                Press(window, Key.B, RawInputModifiers.Control);
                Assert.Equal("- fuses", box.Markup);

                Press(window, Key.Z, RawInputModifiers.Control);
                Assert.Equal("- fuse", box.Markup);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // BULLETS BY THE KEYS: Ctrl+Shift+8 on, Enter the next one, Tab nests it, Enter on an empty one ends the list,
    // Backspace at a bullet's words takes it away.
    // ###########################################################################################
    [Fact]
    public void A_nested_list_is_typed_with_Enter_and_Tab()
    {
        UiTest.Run(() =>
        {
            (Window window, RichNoteTextBox box) = Shown(string.Empty);

            try
            {
                Press(window, Key.D8, RawInputModifiers.Control | RawInputModifiers.Shift);
                window.KeyTextInput("PSU");
                Press(window, Key.Enter);
                Press(window, Key.Tab);
                window.KeyTextInput("fuse");
                Press(window, Key.Enter);
                Press(window, Key.Tab, RawInputModifiers.Shift);
                window.KeyTextInput("RAM");
                Press(window, Key.Enter);
                Press(window, Key.Enter);
                window.KeyTextInput("Done");

                Assert.Equal("- PSU\n  - fuse\n- RAM\nDone", box.Markup);
                Assert.Equal("• PSU\n  ◦ fuse\n• RAM\nDone", box.Text);

                box.CaretIndex = box.Text!.IndexOf("RAM", System.StringComparison.Ordinal);
                Press(window, Key.Back);
                Assert.Equal("- PSU\n  - fuse\nRAM\nDone", box.Markup);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // Undo takes back the styles with the text - the TextBox's own would have left them behind - and Redo puts them back.
    [Fact]
    public void Undo_takes_back_a_style_and_the_words_typed_and_Redo_puts_them_back()
    {
        UiTest.Run(() =>
        {
            (Window window, RichNoteTextBox box) = Shown("the PSU");

            try
            {
                box.SelectionStart = 4;
                box.SelectionEnd = 7;
                Press(window, Key.B, RawInputModifiers.Control);

                // The selection let go of, the caret after the bold word (where Ctrl+B left it).
                box.SelectionStart = 7;
                Assert.Equal((7, 7, 7), (box.SelectionStart, box.SelectionEnd, box.CaretIndex));
                window.KeyTextInput(" died");
                Assert.Equal("the **PSU died**", box.Markup);

                Press(window, Key.Z, RawInputModifiers.Control);
                Assert.Equal("the **PSU**", box.Markup);

                Press(window, Key.Z, RawInputModifiers.Control);
                Assert.Equal("the PSU", box.Markup);
                Assert.All(box.CharacterStylesForTests, style => Assert.Equal(RichNoteStyle.None, style));

                Press(window, Key.Y, RawInputModifiers.Control);
                Assert.Equal("the **PSU**", box.Markup);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // A box made again for another row keeps none of the last row's steps (code review, 2026-11-13): the pane's lists build
    // only the rows in view, so a box scrolled away comes back for another row - and with that row's note equal to the last
    // one it wrote (empty, mostly), nothing new was shown and its steps stayed: Ctrl+Z there brought card A's words into
    // card B's note.
    // ###########################################################################################
    [Fact]
    public void A_box_handed_another_row_keeps_none_of_the_last_rows_steps()
    {
        UiTest.Run(() =>
        {
            (Window window, RichNoteTextBox box) = Shown(string.Empty);

            try
            {
                box.DataContext = "card A";
                window.KeyTextInput("x");
                Press(window, Key.Back);
                Assert.Equal(string.Empty, box.Markup);

                box.DataContext = "card B";
                Press(window, Key.Z, RawInputModifiers.Control);

                Assert.Equal(string.Empty, box.Markup);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // *** ENTER STARTS A LINE WITH "\n", ON WINDOWS TOO (code review, 2026-11-07). *** The TextBox inserts its NewLine,
    // Environment.NewLine by default - "\r\n" on Windows - and the note was then stored with a "\" before every literal
    // star of it, as its "\r" never read back as written.
    // ###########################################################################################
    [Fact]
    public void Enter_starts_a_line_with_n_and_the_note_is_stored_as_typed()
    {
        UiTest.Run(() =>
        {
            (Window window, RichNoteTextBox box) = Shown(string.Empty);

            try
            {
                window.KeyTextInput("Sent *IDN? to the scope");
                Press(window, Key.Enter);
                window.KeyTextInput("No answer");

                Assert.Equal("Sent *IDN? to the scope\nNo answer", box.Text);
                Assert.Equal("Sent *IDN? to the scope\nNo answer", box.Markup);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // A note shown only to be read takes none of the box's own editing keys (code review, 2026-11-07): the component
    // popup's picture note took Enter, Backspace and Tab on its bullets, Ctrl+B and its Undo, and showed a note changed.
    // ###########################################################################################
    [Fact]
    public void A_read_only_note_takes_none_of_its_editing_keys()
    {
        UiTest.Run(() =>
        {
            var box = new RichNoteTextBox { Markup = "- one\n- two", IsReadOnly = true };
            var window = new Window { Content = box, Width = 500, Height = 300 };
            window.Show();
            Assert.True(box.Focus());

            try
            {
                box.CaretIndex = box.Text!.IndexOf("one", System.StringComparison.Ordinal) + 3;
                Press(window, Key.Enter);

                box.CaretIndex = box.Text!.IndexOf("two", System.StringComparison.Ordinal);
                Press(window, Key.Back);
                Press(window, Key.Tab);

                box.SelectionStart = box.Text!.IndexOf("one", System.StringComparison.Ordinal);
                box.SelectionEnd = box.SelectionStart + 3;
                Press(window, Key.B, RawInputModifiers.Control);
                Press(window, Key.D8, RawInputModifiers.Control | RawInputModifiers.Shift);
                Press(window, Key.Z, RawInputModifiers.Control);

                Assert.Equal("- one\n- two", box.Markup);
                Assert.Equal("• one\n• two", box.Text);
                Assert.All(box.CharacterStylesForTests, style => Assert.Equal(RichNoteStyle.None, style));
            }
            finally
            {
                window.Close();
            }
        });
    }

    // A note shown only to be read has no format row.
    [Fact]
    public void A_read_only_note_shows_no_format_buttons()
    {
        UiTest.Run(() =>
        {
            var box = new RichNoteTextBox { Markup = "**read** me", IsReadOnly = true };
            var window = new Window { Content = box, Width = 500, Height = 300 };
            window.Show();

            try
            {
                Control formatBar = box.GetVisualDescendants().OfType<Control>().Single(control => control.Name == "PART_FormatBar");
                Assert.False(formatBar.IsEffectivelyVisible);
                Assert.Equal("read me", box.Text);
            }
            finally
            {
                window.Close();
            }
        });
    }
}
