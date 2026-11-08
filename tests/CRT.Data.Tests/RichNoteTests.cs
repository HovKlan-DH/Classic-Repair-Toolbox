using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A NOTE WITH A LITTLE FORMATTING, STORED AS TEXT A PERSON CAN READ (RichNote; owner request, 2026-11-07: "Bold, Italic,
// Underscore, Strikethrough, Bulletlist (nested)") - the marks of Markdown, read into plain text and a style per
// character, and written back with the fewest marks. Every note typed before it must read exactly as it did.
// ###########################################################################################
public sealed class RichNoteTests
{
    private const RichNoteStyle B = RichNoteStyle.Bold;
    private const RichNoteStyle I = RichNoteStyle.Italic;
    private const RichNoteStyle U = RichNoteStyle.Underline;
    private const RichNoteStyle S = RichNoteStyle.Strike;
    private const RichNoteStyle N = RichNoteStyle.None;

    // The runs of a note as "text:style" pieces - each stretch of one style, in order - what a test reads at a glance.
    private static string Runs(RichNoteText note)
    {
        var pieces = new List<string>();
        int start = 0;

        for (int index = 1; index <= note.Text.Length; index++)
        {
            if (index == note.Text.Length || note.Styles[index] != note.Styles[start])
            {
                pieces.Add($"{note.Text[start..index]}:{note.Styles[start]}");
                start = index;
            }
        }

        return string.Join("|", pieces);
    }

    private static RichNoteText Note(params (string Text, RichNoteStyle Style)[] pieces) =>
        new(string.Concat(pieces.Select(piece => piece.Text)), pieces.SelectMany(piece => Enumerable.Repeat(piece.Style, piece.Text.Length)).ToList());

    [Theory]
    [InlineData("**Recapped** the PSU", "Recapped:Bold| the PSU:None")]
    [InlineData("an *italic* word", "an :None|italic:Italic| word:None")]
    [InlineData("__under__", "under:Underline")]
    [InlineData("~~gone~~ new", "gone:Strike| new:None")]
    [InlineData("***both***", "both:Bold, Italic")]
    [InlineData("**a *b** c*", "a :Bold|b:Bold, Italic| c:Italic")]
    public void Each_mark_reads_as_its_style(string markup, string expected)
    {
        Assert.Equal(expected, Runs(RichNote.Parse(markup)));
    }

    // ###########################################################################################
    // A NOTE TYPED BEFORE READS AS IT DID: a mark counts only in pairs on one line, next to the words it marks - "5 * 2"
    // is no italic, a lone "**" or "_" is the character it is, nothing crosses a line.
    // ###########################################################################################
    [Theory]
    [InlineData("5 * 2 * 3")]
    [InlineData("a ** b")]
    [InlineData("C64_PSU and R12_A")]
    [InlineData("~50Hz")]
    [InlineData("**not closed")]
    [InlineData("half *a line\nand* the next")]
    [InlineData("")]
    [InlineData("Plain text\nOn two lines")]
    [InlineData("a \\ back slash")]
    public void A_note_without_pairs_of_marks_reads_as_it_was_typed(string markup)
    {
        RichNoteText note = RichNote.Parse(markup);

        Assert.Equal(markup, note.Text);
        Assert.All(note.Styles, style => Assert.Equal(N, style));
    }

    // Bullets: "- " starting a line, two spaces more for each level, shown with their glyphs.
    [Fact]
    public void A_dash_starting_a_line_is_a_bullet_nested_by_its_indent()
    {
        RichNoteText note = RichNote.Parse("Checked:\n- PSU\n  - fuse\n    - **cap** C12\n- RAM");

        Assert.Equal("Checked:\n• PSU\n  ◦ fuse\n    ▪ cap C12\n• RAM", note.Text);
        Assert.True(note.IsFormatted);

        string[] lines = note.Text.Split('\n');
        Assert.False(RichNote.TryReadBullet(lines[0], out _, out _));
        Assert.True(RichNote.TryReadBullet(lines[2], out int level, out int prefix));
        Assert.Equal((1, 4), (level, prefix));
        Assert.Equal(B, note.Styles[note.Text.IndexOf("cap", StringComparison.Ordinal)]);
    }

    // ###########################################################################################
    // *** A DASH AFTER AN INDENT NO BULLET HAS IS THE DASH IT IS (code review, 2026-11-08). *** A bullet's indent is two
    // spaces a level, MaxLevel levels deep at most: " - replaced C14" (one space) read as a bullet, and the first edit
    // wrote it back as "- replaced C14" - the space gone - while twelve spaces came back as ten. A note typed before
    // formatting existed reads, and is written, as it was typed; a backslash there is the character it is.
    // ###########################################################################################
    [Theory]
    [InlineData(" - replaced C14", " - replaced C14")]
    [InlineData("   - three spaces", "   - three spaces")]
    [InlineData("            - twelve spaces", "            - twelve spaces")]
    [InlineData(" \\- one space and a backslash", " \\- one space and a backslash")]
    [InlineData("Checked:\n - PSU\n- RAM", "Checked:\n - PSU\n• RAM")]
    public void A_dash_after_an_indent_no_bullet_has_reads_and_is_written_as_typed(string markup, string shown)
    {
        RichNoteText note = RichNote.Parse(markup);

        Assert.Equal(shown, note.Text);
        Assert.Equal(markup, RichNote.Serialize(note));
    }

    // ###########################################################################################
    // *** A NOTE'S EDGES ARE TRIMMED WITHOUT CHANGING HOW IT READS (code review, 2026-11-08). *** Trim() took a first
    // line's indent, which says whether it is a bullet and how deep: "  - sub" came back a bullet of the first level and
    // "   - x" (no bullet at that indent) became one. Blank lines before the words and white space at the end go; the
    // first line's own leading space goes only where no dash or "\" follows it, as it is nothing but space there.
    // ###########################################################################################
    [Theory]
    [InlineData("  - sub\n- top", "  - sub\n- top")]
    [InlineData("   - not a bullet", "   - not a bullet")]
    [InlineData("\n\n  - sub  \n\n", "  - sub")]
    [InlineData("  \\- a dash, no bullet", "  \\- a dash, no bullet")]
    [InlineData("  A flea market\nCame with its box\n  ", "A flea market\nCame with its box")]
    [InlineData(" \n  **bold** start", "**bold** start")]
    [InlineData("   ", "")]
    [InlineData("", "")]
    [InlineData(null, "")]
    public void A_notes_edges_are_trimmed_without_changing_how_it_reads(string? markup, string expected)
    {
        Assert.Equal(expected, RichNote.TrimMarkup(markup));
    }

    // The first line nested once still reads as a bullet nested once.
    [Fact]
    public void A_nested_first_bullet_still_reads_nested_once_trimmed()
    {
        Assert.Equal("  ◦ sub\n• top", RichNote.Parse(RichNote.TrimMarkup("\n  - sub\n- top\n")).Text);
    }

    // ###########################################################################################
    // *** THE END IS TRIMMED WITHOUT CHANGING HOW THE LAST LINE READS (code review, 2026-11-08). *** Enter on a bullet
    // leaves an empty one, "- ", and a note saved so lost its space: stored "-", it reopened with a line showing a dash. A
    // typed dash, written "\- ", kept a "\" it no longer needed, and showed it. An empty bullet at the end goes as the
    // blank it is; a typed dash keeps only its own backslashes. A dash typed alone is the dash it is.
    // ###########################################################################################
    [Theory]
    [InlineData("- a\n- ", "- a")]
    [InlineData("- a\n  - \n- \n\n", "- a")]
    [InlineData("  - ", "")]
    [InlineData("a\n\\- ", "a\n-")]
    [InlineData("a\n\\\\\\- ", "a\n\\\\-")]
    [InlineData("a\n-", "a\n-")]
    [InlineData("a\n - ", "a\n -")]
    public void An_empty_bullet_or_a_typed_dash_at_the_end_reads_as_it_did(string markup, string expected)
    {
        string trimmed = RichNote.TrimMarkup(markup);

        Assert.Equal(expected, trimmed);
        Assert.Equal(trimmed, RichNote.TrimMarkup(trimmed));
    }

    // The box's own case: a bullet, Enter, and the note saved - it reopens as the one bullet.
    [Fact]
    public void A_note_saved_after_Enter_on_a_bullet_reopens_without_a_dash_line()
    {
        string markup = RichNote.TrimMarkup(RichNote.Serialize(Note(("• a\n• ", N))));

        Assert.Equal("• a", RichNote.Parse(markup).Text);
    }

    // ###########################################################################################
    // A first line starting with a bullet glyph at a bullet's indent reads as a bullet of that level - a glyph pasted in -
    // so trimming keeps that indent (code review, 2026-11-08: "  • x" came back a level up, " • x" became a bullet).
    // ###########################################################################################
    [Theory]
    [InlineData("  • x\n• y", 1)]
    [InlineData(" • x", -1)]
    [InlineData("\n\n    ▪ x", 2)]
    public void A_first_line_starting_with_a_bullet_glyph_keeps_its_level(string markup, int level)
    {
        string trimmed = RichNote.TrimMarkup(markup);
        string firstLine = RichNote.Parse(trimmed).Text.Split('\n')[0];

        Assert.Equal(level, RichNote.TryReadBullet(firstLine, out int read, out _) ? read : -1);
        Assert.Equal(trimmed, RichNote.TrimMarkup(trimmed));
    }

    // ###########################################################################################
    // *** A LINE STARTING WITH BACKSLASHES AND A DASH, AT A BULLET'S INDENT, IS WRITTEN SO IT READS BACK (code review,
    // 2026-11-08). *** A line shown as two backslashes, a dash and a space was written as it stands, read back with one
    // backslash fewer, and lost one more at every save. At a bullet's indent a run of backslashes before "- " is read with
    // its first taken off - the "\" that says "no bullet" - so any such line has a way to be written.
    // ###########################################################################################
    [Theory]
    [InlineData("- x")]
    [InlineData("\\- x")]
    [InlineData("\\\\- x")]
    [InlineData("  \\\\\\- x")]
    [InlineData("Checked:\n\\\\- x")]
    public void A_line_of_backslashes_and_a_dash_at_a_bullets_indent_reads_back_as_shown(string shown)
    {
        var note = new RichNoteText(shown, Enumerable.Repeat(N, shown.Length).ToList());

        Assert.Equal(shown, RichNote.Parse(RichNote.Serialize(note)).Text);
    }

    [Fact]
    public void A_backslash_keeps_a_character_that_would_be_a_mark()
    {
        Assert.Equal("*not italic* and - no bullet", RichNote.Parse("\\*not italic\\* and - no bullet").Text);
        Assert.Equal("- no bullet", RichNote.Parse("\\- no bullet").Text);
        Assert.Equal("__not underlined__", RichNote.Parse("\\__not underlined\\__").Text);
    }

    // ###########################################################################################
    // *** A BACKSLASH WHERE IT CHANGES NOTHING IS THE CHARACTER IT IS (code review, 2026-11-07). *** Every "\" before a
    // "\", "_", "~" or "-" was taken for an escape: "\\NAS\c64" read as "\NAS\c64", "D:\_old" as "D:_old", and the first
    // edit saved them so. A note typed before formatting existed holds paths like these, and must read as typed.
    // ###########################################################################################
    [Theory]
    [InlineData("Dump copied to \\\\NAS\\c64")]
    // windows-path-literal: the words of a note, read as text - never a path any file system is asked about.
    [InlineData("ROMs in D:\\_old")]
    [InlineData("C:\\~tmp and C:\\-backup")]
    [InlineData("a\\\\b")]
    [InlineData("ends in a backslash \\")]
    [InlineData("\\\\\\\\server\\share")]
    public void A_backslash_before_anything_but_a_mark_is_the_character_it_is(string markup)
    {
        RichNoteText note = RichNote.Parse(markup);

        Assert.Equal(markup, note.Text);
        Assert.All(note.Styles, style => Assert.Equal(N, style));

        // Written back, it is the note as it was typed - no backslash added or lost.
        Assert.Equal(markup, RichNote.Serialize(note));
    }

    // ###########################################################################################
    // *** A NOTE WITH WINDOWS LINE ENDS IS WRITTEN WITH NO BACKSLASH IT DOES NOT NEED (code review, 2026-11-07). *** Its
    // "\r\n" never read back as written (Parse reads "\n"), so every literal mark character got a "\" - in every note of
    // more than one line typed or pasted on Windows. Written with "\n", it reads back as itself.
    // ###########################################################################################
    [Theory]
    [InlineData("Sent *IDN? to the scope\r\nNo answer", "Sent *IDN? to the scope\nNo answer")]
    [InlineData("a_b and ~x\rnext line", "a_b and ~x\nnext line")]
    public void A_note_with_Windows_line_ends_is_written_with_n_and_no_backslash(string shown, string expected)
    {
        RichNoteText note = Note((shown, N));

        string written = RichNote.Serialize(note);

        Assert.Equal(expected, written);
        Assert.Equal(expected, RichNote.Parse(written).Text);
    }

    [Fact]
    public void A_style_after_a_Windows_line_end_stays_on_its_words()
    {
        RichNoteText note = Note(("first\r\n", N), ("bold", B), (" end", N));

        Assert.Equal("first\n**bold** end", RichNote.Serialize(note));
    }

    // Before a mark, backslashes count in pairs: "\\" is one, and an odd one left over keeps the mark's character.
    [Fact]
    public void Backslashes_before_a_mark_are_read_in_pairs()
    {
        Assert.Equal("\\:None|a:Italic", Runs(RichNote.Parse("\\\\*a*")));
        Assert.Equal("\\*a*:None", Runs(RichNote.Parse("\\\\\\*a*")));
    }

    // A backslash in front of a mark that IS meant, and one in front of a mark's character kept as itself, are written so
    // they read back - and nowhere else is one added.
    [Fact]
    public void A_backslash_next_to_a_mark_is_written_so_it_reads_back()
    {
        RichNoteText beforeBold = Note(("path\\", N), ("x", B));
        string written = RichNote.Serialize(beforeBold);
        Assert.Equal("path\\\\**x**", written);
        Assert.True(RichNote.Parse(written).SameAs(beforeBold));

        RichNoteText beforeStar = Note(("a\\*b", N));
        written = RichNote.Serialize(beforeStar);
        Assert.Equal("a\\\\\\*b", written);
        Assert.True(RichNote.Parse(written).SameAs(beforeStar));

        // windows-path-literal: the words of a note, read as text - never a path any file system is asked about.
        RichNoteText underscores = Note(("__x__ and D:\\_old", N));
        written = RichNote.Serialize(underscores);
        Assert.Equal("\\__x\\__ and D:\\_old", written);
        Assert.True(RichNote.Parse(written).SameAs(underscores));

        RichNoteText dashLine = Note(("\\- not a bullet", N));
        written = RichNote.Serialize(dashLine);
        Assert.True(RichNote.Parse(written).SameAs(dashLine));
    }

    // ###########################################################################################
    // *** NOTHING INSIDE A WEB ADDRESS IS A MARK (code review, 2026-11-07). *** "__init__" in a link read as underlined
    // "init", so the link opened "...#object.init" - a dead page. The address is its own characters; only a mark right at
    // its end closes one opened before it, so a link can still be bold as a whole.
    // ###########################################################################################
    [Theory]
    [InlineData("https://docs.python.org/3/reference/datamodel.html#object.__init__")]
    [InlineData("see https://host/files/C64__rev__B.pdf for the board")]
    [InlineData("www.example.com/a**b**c and https://x.dk/~~old~~")]
    [InlineData("https://x.dk/a\\*b")]
    public void A_web_address_reads_as_typed(string markup)
    {
        RichNoteText note = RichNote.Parse(markup);

        Assert.Equal(markup, note.Text);
        Assert.All(note.Styles, style => Assert.Equal(N, style));
        Assert.Equal(markup, RichNote.Serialize(note));
    }

    [Fact]
    public void A_link_can_be_styled_as_a_whole()
    {
        Assert.Equal("see :None|https://x.dk/__init__:Bold", Runs(RichNote.Parse("see **https://x.dk/__init__**")));
        Assert.Equal("https://x.dk:Italic|.:None", Runs(RichNote.Parse("*https://x.dk*.")));

        // A style on part of an address cannot be written - it goes, rather than putting marks into the address.
        RichNoteText partly = Note(("https://x.dk/", N), ("ab", B), (" and ", N), ("more", B));
        Assert.Equal("https://x.dk/ab and **more**", RichNote.Serialize(partly));

        RichNoteText whole = Note(("see ", N), ("https://x.dk/__init__", B));
        string written = RichNote.Serialize(whole);
        Assert.Equal("see **https://x.dk/__init__**", written);
        Assert.True(RichNote.Parse(written).SameAs(whole));

        // An address beside text that needs its backslashes keeps none of its own.
        RichNoteText beside = Note(("https://x/__init__ 5*2*3", N));
        Assert.Equal("https://x/__init__ 5\\*2\\*3", RichNote.Serialize(beside));
    }

    // ###########################################################################################
    // WRITTEN WITH THE FEWEST MARKS, AND READ BACK THE SAME: no "\" where nothing could be misread, one before every
    // character that could be taken for a mark otherwise - and whatever is written reads back as what was shown.
    // ###########################################################################################
    [Fact]
    public void A_note_is_written_with_its_marks_and_reads_back_the_same()
    {
        RichNoteText note = Note(("Recapped ", N), ("the PSU", B), (", ", N), ("old", S), (" one ", N), ("gone", I | U));

        string written = RichNote.Serialize(note);

        Assert.Equal("Recapped **the PSU**, ~~old~~ one *__gone__*", written);
        Assert.True(RichNote.Parse(written).SameAs(note));
    }

    [Theory]
    [InlineData("C64_PSU 5 * 2", "C64_PSU 5 * 2")]
    [InlineData("5*2*3", "5\\*2\\*3")]
    [InlineData("**not bold**", "\\*\\*not bold\\*\\*")]
    [InlineData("- a dash, no bullet", "\\- a dash, no bullet")]
    public void A_plain_note_gets_a_backslash_only_where_it_would_be_misread(string plain, string expected)
    {
        RichNoteText note = Note((plain, N));

        string written = RichNote.Serialize(note);

        Assert.Equal(expected, written);
        Assert.Equal(plain, RichNote.Parse(written).Text);
    }

    [Fact]
    public void Bullets_are_written_as_dashes_by_their_level()
    {
        RichNoteText note = Note(("• PSU\n  ◦ ", N), ("fuse", B), ("\n• RAM", N));

        string written = RichNote.Serialize(note);

        Assert.Equal("- PSU\n  - **fuse**\n- RAM", written);
        Assert.True(RichNote.Parse(written).SameAs(note));
    }

    // A style on the spaces at its edges, or on a bullet's start, is let go of - a mark must stand next to a word.
    [Fact]
    public void A_style_on_spaces_at_its_edges_is_let_go_of()
    {
        RichNoteText note = Note(("a", N), (" bold ", B), ("b", N));

        Assert.Equal("a **bold** b", RichNote.Serialize(note));
        Assert.Equal("a :None|bold:Bold| b:None", Runs(RichNote.Parse(RichNote.Serialize(note))));
    }

    // ###########################################################################################
    // *** ONE LINE THAT CANNOT BE WRITTEN AS SHOWN TAKES NO OTHER LINE WITH IT (code review, 2026-11-08). *** A bold line
    // starting "- " - no bullet, its dash typed - lost its start's bold in writing, so the note did not read back and
    // EVERY line was written with a "\" before each star: "5 * 2" came back "5 \* 2" in the stored text. Such a line's
    // start takes no style, and only a line that still would not read back gets the backslashes.
    // ###########################################################################################
    [Fact]
    public void A_bold_line_starting_with_a_dash_leaves_the_other_lines_as_typed()
    {
        RichNoteText note = Note(("Use 5 * 2 of them\n", N), ("- replaced C14", B));

        string written = RichNote.Serialize(note);

        Assert.Equal("Use 5 * 2 of them\n\\- **replaced C14**", written);
        Assert.Equal("Use 5 * 2 of them\n- :None|replaced C14:Bold", Runs(RichNote.Parse(written)));
    }

    // ###########################################################################################
    // An italic address ending in a star cannot be written so it reads back styled (its star and the closing one read as
    // "**"): that line alone goes without its style - its words exactly as shown, never stars where the style was (code
    // review, 2026-11-08) - and the line above keeps its star as typed.
    // ###########################################################################################
    [Fact]
    public void A_line_that_cannot_read_back_styled_keeps_its_words_and_leaves_the_other_lines_alone()
    {
        RichNoteText note = Note(("5 * 2\n", N), ("https://x.dk/a*", I));

        string written = RichNote.Serialize(note);

        Assert.StartsWith("5 * 2\n", written, StringComparison.Ordinal);
        Assert.Equal("5 * 2\nhttps://x.dk/a*", RichNote.Parse(written).Text);
    }

    // The review's own shapes: a style beside or inside an address, where the marks make an address of another length.
    [Theory]
    [InlineData("_*https://x.dk/ax-*", 2, 17, RichNoteStyle.Italic)]
    [InlineData("bhttps://x.dk/a", 13, 1, RichNoteStyle.Bold | RichNoteStyle.Italic | RichNoteStyle.Underline)]
    public void A_style_beside_an_address_never_leaves_marks_showing(string shown, int styledFrom, int styledLength, RichNoteStyle style)
    {
        RichNoteText note = Note((shown[..styledFrom], N), (shown.Substring(styledFrom, styledLength), style), (shown[(styledFrom + styledLength)..], N));

        Assert.Equal(shown, RichNote.Parse(RichNote.Serialize(note)).Text);
    }

    // ###########################################################################################
    // *** READING AND WRITING TAKE AS LONG AS THE NOTE IS LONG (code review, 2026-11-08, Security: a note is anybody's to
    // send, and every CRT opening its component reads it). *** A long run of stars or backslashes was looked through again
    // at every one of them: 40,000 took up to most of a second, and twice as many four times as long. 200,000 of each, read
    // and written, in well under the seconds that took.
    // ###########################################################################################
    [Fact]
    public void A_long_run_of_marks_or_backslashes_is_read_and_written_in_no_time()
    {
        var clock = System.Diagnostics.Stopwatch.StartNew();

        foreach (string run in new[] { new string('*', 200_000), new string('\\', 200_000), new string('_', 200_000), string.Concat(Enumerable.Repeat("\\*", 100_000)) })
        {
            RichNoteText read = RichNote.Parse(run);
            RichNote.Serialize(read);
        }

        Assert.True(clock.Elapsed < TimeSpan.FromSeconds(5), $"took {clock.Elapsed.TotalSeconds:0.0} s");
    }

    // Every note of every shape read back the same - written, read, written again: the same text.
    [Theory]
    [InlineData("**a**")]
    [InlineData("- one\n  - two\n    - three")]
    [InlineData("x ~~y~~ __z__ *w*")]
    [InlineData("\\- dash\n- bullet")]
    [InlineData("Windows line\r\nend")]
    public void What_is_read_and_written_again_stays_as_it_is(string markup)
    {
        string once = RichNote.Serialize(RichNote.Parse(markup));
        string twice = RichNote.Serialize(RichNote.Parse(once));

        Assert.Equal(once, twice);
        Assert.Equal(RichNote.Parse(markup).Text, RichNote.Parse(once).Text);
    }

    [Fact]
    public void The_plain_text_is_what_a_search_reads()
    {
        Assert.Equal("dead SID\n• socketed", RichNote.PlainOf("dead **SID**\n- socketed"));
        Assert.Equal(string.Empty, RichNote.PlainOf(null));
    }

    [Fact]
    public void A_bullet_s_prefix_is_its_indent_and_glyph_and_the_deepest_level_is_kept()
    {
        Assert.Equal("• ", RichNote.BulletPrefix(0));
        Assert.Equal("  ◦ ", RichNote.BulletPrefix(1));
        Assert.Equal("    ▪ ", RichNote.BulletPrefix(2));
        Assert.Equal(new string(' ', 2 * RichNote.MaxLevel) + "▪ ", RichNote.BulletPrefix(99));

        Assert.False(RichNote.TryReadBullet(" • odd indent", out _, out _));
        Assert.False(RichNote.TryReadBullet("•no space", out _, out _));
        Assert.Throws<ArgumentException>(() => new RichNoteText("ab", [N]));
    }
}
