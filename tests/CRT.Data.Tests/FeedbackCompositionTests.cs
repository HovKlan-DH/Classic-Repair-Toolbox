using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// FeedbackComposition - the rules behind pictures in the Feedback tab's text (2026-10-15), shared
// by CRT (which writes) and the server (which reads): how a part splits around a picture and the
// line break the picture takes, how it joins again when the picture goes, what a picture is
// called, what is SENT - plain text with one marker line per picture - and how the server reads a
// marker line back, so the picture is shown in the mail where it sat.
// ###########################################################################################
public sealed class FeedbackCompositionTests
{
    [Fact]
    public void A_pasted_picture_is_named_by_its_number_and_its_marker_is_the_name_in_brackets()
    {
        Assert.Equal("pasted-image-1.png", FeedbackComposition.PastedName(1));
        Assert.Equal("pasted-image-12.png", FeedbackComposition.PastedName(12));
        Assert.Equal("[pasted-image-3.png]", FeedbackComposition.Marker(FeedbackComposition.PastedName(3)));
    }

    // The server reads what CRT writes - through the same two methods, so they cannot drift.
    [Theory]
    [InlineData("pasted-image-1.png")]
    [InlineData("scope-capture (2).png")]
    [InlineData("Shot.JPG")]
    public void A_marker_CRT_writes_is_read_back_to_the_same_name(string name)
    {
        Assert.True(FeedbackComposition.TryParseMarker(FeedbackComposition.Marker(name), out string read));
        Assert.Equal(name, read);
    }

    [Fact]
    public void A_marker_line_is_read_with_whitespace_around_it()
    {
        Assert.True(FeedbackComposition.TryParseMarker("  [shot.png] \r", out string name));
        Assert.Equal("shot.png", name);
    }

    // The sender's own text with brackets in it names nothing - nor does a path.
    [Theory]
    [InlineData("")]
    [InlineData("[]")]
    [InlineData("[ ]")]
    [InlineData("see [shot.png] above")]
    [InlineData("[shot.png] and more")]
    [InlineData("[a[b].png]")]
    [InlineData("[folder/shot.png]")]
    [InlineData("[folder\\shot.png]")]
    [InlineData("shot.png")]
    public void A_line_that_is_not_a_marker_names_nothing(string line)
    {
        Assert.False(FeedbackComposition.TryParseMarker(line, out string name));
        Assert.Equal("", name);
    }

    [Theory]
    [InlineData("one two", 4, "one ", "two")]
    [InlineData("one two", 0, "", "one two")]
    [InlineData("one two", 7, "one two", "")]
    public void The_text_splits_at_the_caret_exactly_as_typed(string text, int caret, string before, string after)
    {
        PictureSplit split = FeedbackComposition.Split(text, caret);

        Assert.Equal(before, split.Before);
        Assert.Equal(after, split.After);
        Assert.Equal("", split.LineBreakTaken);
    }

    // A caret the control reports past the text (or before it) must not throw in the paste.
    [Theory]
    [InlineData(-3, "", "abc")]
    [InlineData(99, "abc", "")]
    public void A_caret_outside_the_text_is_taken_as_its_nearest_end(int caret, string before, string after)
    {
        PictureSplit split = FeedbackComposition.Split("abc", caret);

        Assert.Equal((before, after), (split.Before, split.After));
    }

    // Enter and then paste (owner report, 2026-10-15): the picture takes the line the Enter made,
    // so no empty line sits above it - and the break comes back, with the caret after it, when
    // the picture goes.
    [Theory]
    [InlineData("see:\n", "\n")]
    [InlineData("see:\r\n", "\r\n")]
    [InlineData("see:\r", "\r")]
    public void A_line_break_just_before_the_caret_goes_with_the_picture_and_comes_back_on_removal(string text, string lineBreak)
    {
        PictureSplit split = FeedbackComposition.Split(text, text.Length);

        Assert.Equal("see:", split.Before);
        Assert.Equal("", split.After);
        Assert.Equal(lineBreak, split.LineBreakTaken);
        Assert.True(split.TakenFromBefore);

        Assert.Equal((text, text.Length), FeedbackComposition.Join(split.Before, split.After, split.LineBreakTaken, split.TakenFromBefore));
    }

    // Pasting at the end of a line in the middle of the text: the break after the caret goes with
    // the picture instead, so the text below it starts on its own first line.
    [Fact]
    public void A_line_break_just_after_the_caret_goes_with_the_picture_when_none_is_before_it()
    {
        PictureSplit split = FeedbackComposition.Split("line1\r\nline2", 5);

        Assert.Equal("line1", split.Before);
        Assert.Equal("line2", split.After);
        Assert.Equal("\r\n", split.LineBreakTaken);
        Assert.False(split.TakenFromBefore);

        Assert.Equal(("line1\r\nline2", 5), FeedbackComposition.Join(split.Before, split.After, split.LineBreakTaken, split.TakenFromBefore));
    }

    // Only ONE line break goes with the picture - a blank line the user typed on purpose stays.
    [Fact]
    public void Only_one_line_break_goes_with_the_picture()
    {
        PictureSplit split = FeedbackComposition.Split("see:\n\n", 6);

        Assert.Equal("see:\n", split.Before);
        Assert.Equal("\n", split.LineBreakTaken);
    }

    [Fact]
    public void Joining_puts_the_caret_where_the_picture_was()
    {
        Assert.Equal(("one two", 4), FeedbackComposition.Join("one ", "two"));
        Assert.Equal(("two", 0), FeedbackComposition.Join("", "two"));
        Assert.Equal(("one", 3), FeedbackComposition.Join("one", ""));
    }

    [Fact]
    public void Split_then_join_gives_the_text_and_the_caret_back_wherever_the_caret_was()
    {
        const string text = "The C64 shows\r\na black screen\n\nafter warm-up";

        for (int caret = 0; caret <= text.Length; caret++)
        {
            PictureSplit split = FeedbackComposition.Split(text, caret);
            Assert.Equal((text, caret), FeedbackComposition.Join(split.Before, split.After, split.LineBreakTaken, split.TakenFromBefore));
        }
    }

    [Fact]
    public void Without_pictures_the_text_is_sent_as_it_is()
    {
        Assert.Equal("", FeedbackComposition.ComposeText([""], []));
        Assert.Equal("as typed\n\nwith blank lines\n", FeedbackComposition.ComposeText(["as typed\n\nwith blank lines\n"], []));
    }

    [Fact]
    public void Each_picture_is_a_marker_line_of_its_own_between_the_texts()
    {
        Assert.Equal(
            "see this:\n[pasted-image-1.png]\nand then this",
            FeedbackComposition.ComposeText(["see this:", "and then this"], ["pasted-image-1.png"]));
    }

    // A line break the user typed before pasting stays one line break - no blank line is added.
    [Fact]
    public void A_text_already_ending_in_a_line_break_gets_no_second_one_before_the_marker()
    {
        Assert.Equal(
            "see this:\n[pasted-image-1.png]\nand then this",
            FeedbackComposition.ComposeText(["see this:\n", "and then this"], ["pasted-image-1.png"]));
    }

    [Fact]
    public void An_empty_part_contributes_nothing_so_pictures_at_the_ends_or_back_to_back_are_marker_lines()
    {
        Assert.Equal("[a.png]\ntext", FeedbackComposition.ComposeText(["", "text"], ["a.png"]));
        Assert.Equal("text\n[a.png]", FeedbackComposition.ComposeText(["text", ""], ["a.png"]));
        Assert.Equal("[a.png]\n[b.png]", FeedbackComposition.ComposeText(["", "", ""], ["a.png", "b.png"]));
        Assert.Equal("x\n[a.png]\n[b.png]\ny", FeedbackComposition.ComposeText(["x", "", "y"], ["a.png", "b.png"]));
    }

    // Every marker in what CRT sends is one the server reads back - line by line.
    [Fact]
    public void Every_marker_in_a_composed_text_reads_back_as_its_picture()
    {
        string sent = FeedbackComposition.ComposeText(["see:", "then", ""], ["pasted-image-1.png", "scope (2).png"]);
        var read = new List<string>();

        foreach (string line in sent.Split('\n'))
        {
            if (FeedbackComposition.TryParseMarker(line, out string name))
                read.Add(name);
        }

        Assert.Equal(["pasted-image-1.png", "scope (2).png"], read);
    }

    [Fact]
    public void The_parts_and_pictures_must_interleave()
    {
        Assert.Throws<ArgumentException>(() => FeedbackComposition.ComposeText(["only one"], ["a.png"]));
        Assert.Throws<ArgumentException>(() => FeedbackComposition.ComposeText([], []));
    }

    // ------------------------------------------------------------------ widths in the mail

    [Theory]
    [InlineData(3840, 980)]
    [InlineData(980, 980)]
    [InlineData(640, 640)]
    [InlineData(0, 1)]
    public void A_picture_is_shown_at_its_own_width_up_to_the_mails_column(int natural, int expected)
    {
        Assert.Equal(expected, FeedbackComposition.DefaultPictureWidth(natural));
    }

    [Fact]
    public void The_widths_CRT_writes_are_read_back_by_name_without_regard_to_capitals()
    {
        string text = FeedbackComposition.PictureWidthsText([("pasted-image-1.png", 640), ("Shot (2).PNG", 980)]);

        IReadOnlyDictionary<string, int> read = FeedbackComposition.ReadPictureWidths(text);

        Assert.Equal(2, read.Count);
        Assert.Equal(640, read["pasted-image-1.png"]);
        Assert.Equal(980, read["shot (2).png"]);
    }

    [Fact]
    public void No_pictures_write_no_widths()
    {
        Assert.Equal("", FeedbackComposition.PictureWidthsText([]));
        Assert.Empty(FeedbackComposition.ReadPictureWidths(""));
        Assert.Empty(FeedbackComposition.ReadPictureWidths(null));
    }

    // The server reads what the SENDER wrote: anything but a whole width in range is not there.
    [Theory]
    [InlineData("not json")]
    [InlineData("[640]")]
    [InlineData("{\"a.png\":\"640\"}")]
    [InlineData("{\"a.png\":640.5}")]
    [InlineData("{\"a.png\":0}")]
    [InlineData("{\"a.png\":-5}")]
    [InlineData("{\"a.png\":981}")]
    [InlineData("{\"a.png\":99999999999}")]
    [InlineData("{\"a.png\":")]
    public void A_width_that_is_not_a_whole_number_in_range_is_not_read(string text)
    {
        Assert.Empty(FeedbackComposition.ReadPictureWidths(text));
    }

    [Fact]
    public void A_bad_entry_does_not_lose_the_good_ones()
    {
        IReadOnlyDictionary<string, int> read = FeedbackComposition.ReadPictureWidths("{\"a.png\":\"x\",\"b.png\":300,\"c.png\":5000}");

        Assert.Equal(300, Assert.Single(read).Value);
    }

    [Fact]
    public void A_dropped_file_keeps_its_name_unless_a_picture_has_it()
    {
        Assert.Equal("shot.png", FeedbackComposition.UniqueName("shot.png", []));
        Assert.Equal("shot.png", FeedbackComposition.UniqueName("shot.png", ["other.png"]));
        Assert.Equal("shot (2).png", FeedbackComposition.UniqueName("shot.png", ["shot.png"]));
        Assert.Equal("shot (3).png", FeedbackComposition.UniqueName("shot.png", ["shot.png", "shot (2).png"]));
    }

    // The zip and the server's feedback folder do not tell SHOT.PNG from shot.png.
    [Fact]
    public void Names_are_compared_without_regard_to_capitals()
    {
        Assert.Equal("Shot (2).PNG", FeedbackComposition.UniqueName("Shot.PNG", ["shot.png"]));
    }

    // ------------------------------------------------------------------ code review, 2026-10-16

    // The same rule keeps the zip's entries apart, and a folder's file keeps its folder.
    [Fact]
    public void A_name_with_a_folder_in_front_keeps_the_folder()
    {
        Assert.Equal("logs/run (2).txt", FeedbackComposition.UniqueName("logs/run.txt", ["logs/run.txt"]));
        Assert.Equal("my.logs/run (2)", FeedbackComposition.UniqueName("my.logs/run", ["my.logs/run"]));
        Assert.Equal("README (2)", FeedbackComposition.UniqueName("README", ["readme"]));
    }

    // A set the caller holds is used as it is - copied per call, thousands of clashes cost their
    // square - but only one comparing as file names do: a case-sensitive one is not trusted.
    [Fact]
    public void A_held_set_gives_the_answer_a_copied_one_would()
    {
        var held = new HashSet<string>(["shot.png", "shot (2).png"], StringComparer.OrdinalIgnoreCase);
        var caseSensitive = new HashSet<string>(["shot.png"], StringComparer.Ordinal);

        Assert.Equal("SHOT (3).PNG", FeedbackComposition.UniqueName("SHOT.PNG", held));
        Assert.Equal(2, held.Count);
        Assert.Equal("Shot (2).png", FeedbackComposition.UniqueName("Shot.png", caseSensitive));
    }

    // A dropped file can be called "pasted-image-1.png" too: a paste takes the next number free.
    [Fact]
    public void A_pasted_picture_takes_the_next_number_no_picture_has()
    {
        Assert.Equal(("pasted-image-1.png", 1), FeedbackComposition.NextPastedName(1, []));
        Assert.Equal(("pasted-image-2.png", 2), FeedbackComposition.NextPastedName(1, ["PASTED-IMAGE-1.png", "shot.png"]));
        Assert.Equal(("pasted-image-4.png", 4), FeedbackComposition.NextPastedName(3, ["pasted-image-3.png"]));
        Assert.Equal(("pasted-image-5.png", 5), FeedbackComposition.NextPastedName(5, ["pasted-image-1.png"]));
    }

    // "Screenshot [scope].png" was shown in CRT and never in the mail: a marker cannot hold brackets.
    [Theory]
    [InlineData("Screenshot [scope].png", "Screenshot (scope).png")]
    [InlineData("[1].jpg", "(1).jpg")]
    [InlineData("a\\b.png", "a-b.png")]
    [InlineData("tab\there.png", "tab-here.png")]
    [InlineData("  shot.png ", "shot.png")]
    [InlineData("scope-capture-12.png", "scope-capture-12.png")]
    [InlineData("   ", "picture")]
    public void A_dropped_pictures_name_is_one_a_marker_line_can_carry(string fileName, string expected)
    {
        string name = FeedbackComposition.PictureNameFor(fileName);

        Assert.Equal(expected, name);
        Assert.True(FeedbackComposition.TryParseMarker(FeedbackComposition.Marker(name), out string read));
        Assert.Equal(name, read);
    }

    // A paste over a selection replaces it, as in any text box - in either direction of selecting.
    [Theory]
    [InlineData("see PLACEHOLDER here", 4, 15, "see ", " here")]
    [InlineData("see PLACEHOLDER here", 15, 4, "see ", " here")]
    [InlineData("all of it", 0, 9, "", "")]
    [InlineData("short", 2, 99, "sh", "")]
    [InlineData("one two", 4, 4, "one ", "two")]
    public void A_picture_over_a_selection_replaces_it(string text, int selectionStart, int selectionEnd, string before, string after)
    {
        PictureSplit split = FeedbackComposition.Split(text, selectionStart, selectionEnd);

        Assert.Equal(before, split.Before);
        Assert.Equal(after, split.After);
    }

    // A selected line, break and all: the picture takes the break before it, as after Enter.
    [Fact]
    public void A_selection_ending_a_line_leaves_the_picture_on_that_line()
    {
        PictureSplit split = FeedbackComposition.Split("see:\nPLACEHOLDER\nmore", 5, 16);

        Assert.Equal("see:", split.Before);
        Assert.Equal("\nmore", split.After);
        Assert.Equal("\n", split.LineBreakTaken);
    }

    // ###########################################################################################
    // The ONE reading of the markers, which the server uses both to take pictures out of the zip and
    // to put them into the mail - every kind of line break, each line as written.
    // ###########################################################################################
    [Fact]
    public void The_text_is_read_line_by_line_with_the_picture_each_marker_names()
    {
        IReadOnlyList<FeedbackLine> lines = FeedbackComposition.ReadLines("one\r\n[a.png]\ntwo\r  [b.png] \n");

        Assert.Equal(
            [new FeedbackLine("one", null), new FeedbackLine("[a.png]", "a.png"), new FeedbackLine("two", null), new FeedbackLine("  [b.png] ", "b.png"), new FeedbackLine("", null)],
            lines);
    }

    [Fact]
    public void Nothing_at_all_is_one_empty_line()
    {
        Assert.Equal([new FeedbackLine("", null)], FeedbackComposition.ReadLines(""));
        Assert.Equal([new FeedbackLine("", null)], FeedbackComposition.ReadLines(null));
    }

    [Fact]
    public void The_pictures_a_text_names_are_each_named_once_in_the_order_it_names_them()
    {
        Assert.Equal(["b.png", "a.png"], FeedbackComposition.MarkedNames("[b.png]\nwords [x.png]\n[a.png]\n[B.PNG]\n[a.png]"));
    }

    // ###########################################################################################
    // The widths are the SENDER's text: a name that is no valid text once unescaped (a lone half of
    // a surrogate pair), or a field cut between the two halves of one, threw past the JSON catch
    // and ended the request in a 500 (code review, 2026-10-16). Such an entry is simply not there.
    // ###########################################################################################
    [Fact]
    public void A_width_whose_name_is_no_valid_text_is_not_read_and_the_others_are()
    {
        IReadOnlyDictionary<string, int> read = FeedbackComposition.ReadPictureWidths("{\"\\uDC00.png\":5,\"a.png\":300}");

        Assert.Equal(300, Assert.Single(read).Value);
    }

    [Fact]
    public void Widths_cut_inside_a_character_are_not_read()
    {
        Assert.Empty(FeedbackComposition.ReadPictureWidths("{\"a.png\":300,\"b\uD83D"));
    }

    // ###########################################################################################
    // *** A LINE THE SENDER TYPED AS A MARKER IS NOT TAKEN FOR THE PICTURE (code review,
    // 2026-10-16). *** "[pasted-image-1.png]" written near the top, to point at the screenshot
    // further down, took the picture to the top and left CRT's own marker line as text. Such a line
    // is sent with an invisible mark after it (MarkerEscape), which no marker has.
    // ###########################################################################################
    [Fact]
    public void A_line_the_sender_typed_as_a_marker_is_not_taken_for_the_picture()
    {
        string text = FeedbackComposition.ComposeText(["[a.png]\r\nis below", ""], ["a.png"]);
        IReadOnlyList<FeedbackLine> lines = FeedbackComposition.ReadLines(text);

        Assert.Equal([null, null, "a.png"], lines.Select(line => line.PictureName));
        Assert.Equal("[a.png]" + FeedbackComposition.MarkerEscape, lines[0].Text);
        Assert.Equal(["a.png"], FeedbackComposition.MarkedNames(text));
    }

    // The mark is invisible and no whitespace, so trimming the line keeps it.
    [Fact]
    public void The_mark_is_an_invisible_character_trimming_keeps()
    {
        Assert.Equal('\u200B', FeedbackComposition.MarkerEscape);
        Assert.False(char.IsWhiteSpace(FeedbackComposition.MarkerEscape));
        Assert.False(FeedbackComposition.TryParseMarker(" [a.png]" + FeedbackComposition.MarkerEscape + " ", out _));
    }

    // Only lines that would be read as markers get the mark; the rest is sent as typed.
    [Fact]
    public void Text_that_reads_as_no_marker_is_sent_as_typed()
    {
        Assert.Equal("see [a.png] here\n[]\n[x/y.png]", FeedbackComposition.ComposeText(["see [a.png] here\n[]\n[x/y.png]"], []));
    }

    // What CRT composes is what the server reads back.
    [Fact]
    public void A_composed_text_reads_back_to_its_parts_and_pictures()
    {
        string text = FeedbackComposition.ComposeText(["one\r\ntwo", "", "three"], ["a.png", "b.png"]);

        Assert.Equal(["a.png", "b.png"], FeedbackComposition.MarkedNames(text));
        Assert.Equal(["one", "two", "[a.png]", "[b.png]", "three"], FeedbackComposition.ReadLines(text).Select(line => line.Text));
    }
}
