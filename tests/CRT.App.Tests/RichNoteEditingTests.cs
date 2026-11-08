using Handlers.DataHandling;
using System.Collections.Generic;
using System.Linq;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What an edit does to a formatted note (RichNoteEditing; owner request, 2026-11-07: bold, italic, underline,
// strikethrough and nested bullets in the notes of devices, workbooks and worklogs) - every rule the box follows, read
// here without a box.
// ###########################################################################################
public sealed class RichNoteEditingTests
{
    private const RichNoteStyle B = RichNoteStyle.Bold;
    private const RichNoteStyle N = RichNoteStyle.None;

    private static List<RichNoteStyle> Styles(string text, RichNoteStyle style = RichNoteStyle.None) =>
        Enumerable.Repeat(style, text.Length).ToList();

    private static string Written(RichNoteEdit edit) => RichNote.Serialize(new RichNoteText(edit.Text, edit.Styles));

    // ###########################################################################################
    // A CHANGE OF THE TEXT KEEPS EVERY CHARACTER'S STYLE: what was taken out goes with its styles, what came in takes the
    // style of the character before it on its line - so typing on at the end of a bold word goes on in bold - or the
    // style asked for before typing; a line end has none.
    // ###########################################################################################
    [Fact]
    public void Typing_on_after_a_bold_word_is_bold_and_a_new_line_is_not()
    {
        List<RichNoteStyle> styles = [B, B, B, N, N];

        List<RichNoteStyle> typed = RichNoteEditing.FollowTextChange("abc d", styles, "abcX d");
        Assert.Equal([B, B, B, B, N, N], typed);

        List<RichNoteStyle> newLine = RichNoteEditing.FollowTextChange("abc d", styles, "abc\nX d");
        Assert.Equal([B, B, B, N, N, N, N], newLine);

        List<RichNoteStyle> asked = RichNoteEditing.FollowTextChange("abc d", styles, "abc dY", typing: RichNoteStyle.Italic);
        Assert.Equal(RichNoteStyle.Italic, asked[^1]);
    }

    [Fact]
    public void Text_taken_out_takes_its_styles_with_it_and_a_paste_takes_the_style_before_it()
    {
        List<RichNoteStyle> styles = [N, B, B, N];

        Assert.Equal([N, N], RichNoteEditing.FollowTextChange("abcd", styles, "ad"));
        Assert.Equal([N, B, B, B, B, B, N], RichNoteEditing.FollowTextChange("abcd", styles, "abcXYZd"));
        Assert.Equal([N, N, N, B, B, N], RichNoteEditing.FollowTextChange("abcd", styles, "XYabcd"));
        Assert.Equal([N, B, B, N, N, N, N], RichNoteEditing.FollowTextChange("abcd", styles, "abc\nXYd"));
    }

    // ###########################################################################################
    // *** BESIDE THE SAME CHARACTER, THE CHANGE IS WHERE IT BEGAN (code review, 2026-11-08). *** "b" typed into "ab" after
    // its "a" or after its "b" reads the same from the text alone; where the caret was says which, and so whose style the
    // new one takes - and which of two like characters a Backspace or a Delete took out. Without it, the last place.
    // ###########################################################################################
    [Fact]
    public void Beside_the_same_character_the_change_is_where_it_began()
    {
        Assert.Equal((1, 0, 1), RichNoteEditing.ChangeOf("ab", "abb", changeAt: 1));
        Assert.Equal((2, 0, 1), RichNoteEditing.ChangeOf("ab", "abb"));
        Assert.Equal((1, 1, 0), RichNoteEditing.ChangeOf("abb", "ab", changeAt: 1));

        // Where it began is only a choice among the places that say the same: anywhere else is the nearest of them.
        Assert.Equal((1, 0, 1), RichNoteEditing.ChangeOf("ab", "abb", changeAt: 0));
        Assert.Equal((2, 0, 1), RichNoteEditing.ChangeOf("ab", "abb", changeAt: 9));
        Assert.Equal((1, 2, 3), RichNoteEditing.ChangeOf("abcd", "aXYZd", changeAt: 3));

        Assert.Equal([B, B, N], RichNoteEditing.FollowTextChange("ab", [B, N], "abb", changeAt: 1));
        Assert.Equal([B, N, N], RichNoteEditing.FollowTextChange("ab", [B, N], "abb"));
        Assert.Equal([N, N], RichNoteEditing.FollowTextChange("abb", [N, B, N], "ab", changeAt: 1));
        Assert.Equal([N, B], RichNoteEditing.FollowTextChange("abb", [N, B, N], "ab", changeAt: 2));
    }

    // ###########################################################################################
    // *** WHAT THE FORMAT BUTTONS SHOW PRESSED (owner decision, 2026-11-08) - by the rule a press goes by. *** A selection:
    // each style ALL of it has - a bullet's start and a line end, which take none, not counted - and bullets while every
    // line it touches is one. Nothing selected: the style asked for, else the one before the caret on its line.
    // ###########################################################################################
    [Fact]
    public void A_button_is_pressed_when_all_of_the_selection_has_its_style()
    {
        // "• ab\n• cd", the words "ab" and "cd" bold and "cd" italic too.
        const string text = "• ab\n• cd";
        List<RichNoteStyle> styles = [N, N, B, B, N, N, N, B | RichNoteStyle.Italic, B | RichNoteStyle.Italic];

        // Both lines whole, the bullets' starts and the line end in it: all bold, not all italic, every line a bullet.
        Assert.Equal((B, true), RichNoteEditing.Pressed(text, styles, 0, text.Length));

        // "cd" alone: bold and italic.
        Assert.Equal((B | RichNoteStyle.Italic, true), RichNoteEditing.Pressed(text, styles, 7, 9));

        // A selection holding nothing that takes a style - a bullet's start - shows none pressed.
        Assert.Equal((N, true), RichNoteEditing.Pressed(text, styles, 0, 2));

        // A line that is no bullet among the lines touched: the bullet button is not pressed.
        Assert.False(RichNoteEditing.Pressed("• ab\nplain", Styles("• ab\nplain"), 0, 9).Bullets);
    }

    [Fact]
    public void With_nothing_selected_the_buttons_show_what_is_typed_next()
    {
        const string text = "ab cd";
        List<RichNoteStyle> styles = [B, B, N, N, N];

        Assert.Equal(B, RichNoteEditing.Pressed(text, styles, 2, 2).Styles);
        Assert.Equal(N, RichNoteEditing.Pressed(text, styles, 0, 0).Styles);
        Assert.Equal(N, RichNoteEditing.Pressed(text, styles, 4, 4).Styles);

        // What was asked for with nothing selected - Ctrl+B there - is what shows.
        Assert.Equal(RichNoteStyle.Italic, RichNoteEditing.Pressed(text, styles, 2, 2, asked: RichNoteStyle.Italic).Styles);

        // At a line's start the line before is not what is typed.
        Assert.Equal(N, RichNoteEditing.Pressed("ab\ncd", [B, B, N, N, N], 3, 3).Styles);
    }

    // ###########################################################################################
    // A STYLE TOGGLED ON A SELECTION: on where any of it lacks it, off where all has it - never on a bullet's start or a
    // line end.
    // ###########################################################################################
    [Fact]
    public void A_style_goes_on_where_any_lacks_it_and_off_where_all_have_it()
    {
        string text = "the PSU";
        RichNoteEdit bold = RichNoteEditing.Toggle(text, Styles(text), 4, 7, B);

        Assert.Equal("the **PSU**", Written(bold));

        RichNoteEdit wider = RichNoteEditing.Toggle(text, bold.Styles, 0, 7, B);
        Assert.Equal("**the PSU**", Written(wider));

        RichNoteEdit off = RichNoteEditing.Toggle(text, wider.Styles, 7, 0, B);
        Assert.Equal("the PSU", Written(off));
        Assert.Equal((0, 7), (off.SelectionStart, off.SelectionEnd));
    }

    [Fact]
    public void A_bullet_s_start_and_a_line_end_take_no_style()
    {
        string text = "• one\n• two";
        RichNoteEdit bold = RichNoteEditing.Toggle(text, Styles(text), 0, text.Length, B);

        Assert.Equal("- **one**\n- **two**", Written(bold));
        Assert.Equal(N, bold.Styles[0]);
        Assert.Equal(N, bold.Styles[5]);
    }

    // ###########################################################################################
    // BULLETS ON OR OFF FOR THE LINES THE SELECTION TOUCHES - all of them bullets: taken away, else put on every one at the
    // first level; a selection ending at a line's very start does not touch it. The words keep their styles.
    // ###########################################################################################
    [Fact]
    public void Bullets_go_on_the_lines_the_selection_touches_and_off_again()
    {
        string text = "PSU\nRAM\nCIA";

        RichNoteEdit on = RichNoteEditing.ToggleBullets(text, Styles(text), 1, 6);
        Assert.Equal("• PSU\n• RAM\nCIA", on.Text);
        Assert.Equal("- PSU\n- RAM\nCIA", Written(on));

        // From the first line to the second's very start: the first line alone, a bullet already - taken away.
        RichNoteEdit off = RichNoteEditing.ToggleBullets(on.Text, on.Styles, 0, on.Text.IndexOf("• RAM", System.StringComparison.Ordinal));
        Assert.Equal("PSU\n• RAM\nCIA", off.Text);

        RichNoteEdit both = RichNoteEditing.ToggleBullets(on.Text, on.Styles, 0, on.Text.IndexOf("RAM", System.StringComparison.Ordinal));
        Assert.Equal(text, both.Text);

        // A selection ending at the third line's very start leaves it as it is.
        RichNoteEdit two = RichNoteEditing.ToggleBullets(text, Styles(text), 0, 8);
        Assert.Equal("• PSU\n• RAM\nCIA", two.Text);
    }

    [Fact]
    public void A_caret_on_a_line_puts_the_bullet_on_and_stays_in_its_words()
    {
        string text = "PSU";
        RichNoteEdit on = RichNoteEditing.ToggleBullets(text, Styles(text, B), 1, 1);

        Assert.Equal("• PSU", on.Text);
        Assert.Equal((3, 3), (on.SelectionStart, on.SelectionEnd));
        Assert.Equal("- **PSU**", Written(on));
    }

    // Tab and Shift+Tab nest the bullet lines touched deeper or less - and touching none, leave Tab its own meaning.
    [Fact]
    public void Tab_nests_a_bullet_deeper_and_Shift_Tab_less_and_elsewhere_it_is_Tab()
    {
        string text = "• PSU\n• fuse";
        int caret = text.Length;

        RichNoteEdit deeper = RichNoteEditing.Indent(text, Styles(text), caret, caret, 1)!;
        Assert.Equal("• PSU\n  ◦ fuse", deeper.Text);
        Assert.Equal(deeper.Text.Length, deeper.SelectionStart);

        RichNoteEdit back = RichNoteEditing.Indent(deeper.Text, deeper.Styles, deeper.SelectionStart, deeper.SelectionStart, -1)!;
        Assert.Equal(text, back.Text);

        Assert.Null(RichNoteEditing.Indent("plain", Styles("plain"), 2, 2, 1));

        // ###########################################################################################
        // Where no level would change - Shift+Tab at the top level, Tab at the deepest - there is no edit, so Tab keeps
        // its own meaning and moves on (code review, 2026-11-08: an edit changing nothing held the keyboard in the box and
        // left an empty step to undo).
        // ###########################################################################################
        Assert.Null(RichNoteEditing.Indent(text, Styles(text), 0, 0, -1));

        string deepest = RichNote.BulletPrefix(RichNote.MaxLevel) + "fuse";
        Assert.Null(RichNoteEditing.Indent(deepest, Styles(deepest), deepest.Length, deepest.Length, 1));

        // A selection over a top-level bullet and a nested one: the nested one moves, so it is an edit.
        RichNoteEdit outdented = RichNoteEditing.Indent(deeper.Text, deeper.Styles, 0, deeper.Text.Length, -1)!;
        Assert.Equal(text, outdented.Text);
    }

    // Enter in a bullet's words starts the next at its level; on one with no words, the list ends there.
    [Fact]
    public void Enter_on_a_bullet_starts_the_next_and_on_an_empty_one_ends_the_list()
    {
        string text = "  ◦ fuse";
        RichNoteEdit next = RichNoteEditing.EnterOnBullet(text, Styles(text), text.Length)!;

        Assert.Equal("  ◦ fuse\n  ◦ ", next.Text);
        Assert.Equal(next.Text.Length, next.SelectionStart);

        RichNoteEdit ended = RichNoteEditing.EnterOnBullet(next.Text, next.Styles, next.Text.Length)!;
        Assert.Equal("  ◦ fuse\n", ended.Text);
        Assert.Equal(ended.Text.Length, ended.SelectionStart);

        Assert.Null(RichNoteEditing.EnterOnBullet("plain", Styles("plain"), 5));
        Assert.Null(RichNoteEditing.EnterOnBullet(text, Styles(text), 1));
    }

    // Backspace right where a bullet's words start takes the bullet away, its words where they were.
    [Fact]
    public void Backspace_at_a_bullet_s_words_takes_the_bullet_away()
    {
        string text = "PSU\n• fuse";
        int wordsStart = text.IndexOf("fuse", System.StringComparison.Ordinal);

        RichNoteEdit gone = RichNoteEditing.BackspaceAtBullet(text, Styles(text), wordsStart)!;

        Assert.Equal("PSU\nfuse", gone.Text);
        Assert.Equal(4, gone.SelectionStart);
        Assert.Null(RichNoteEditing.BackspaceAtBullet(text, Styles(text), wordsStart + 1));
        Assert.Null(RichNoteEditing.BackspaceAtBullet("plain", Styles("plain"), 0));
    }
}
