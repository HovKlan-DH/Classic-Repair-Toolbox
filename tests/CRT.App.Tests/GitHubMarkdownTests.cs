using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The changelog read as GitHub reads it (GitHubMarkdown, drawn on the About tab by
// MarkdownRenderer). Every construct CHANGELOG.md uses, and the edges where a hand-written changelog
// could go wrong: an unclosed "**", the "_" inside snake_case, a link with "%E2%80%90" in it, lists
// nested several levels, an alert followed by a blank line, and Windows line endings.
//
// A block reads as one line here - "H1 (3.1.0)", "Item1 b(Data)" - and a piece of text as its
// styles (b bold, i italic, c code, s struck through), its text in brackets and its link in <>.
// ###########################################################################################
public sealed class GitHubMarkdownTests
{
    private const string Backslash = @"\";

    // -----------------------------------------------------------------------------------------
    // Blocks
    // -----------------------------------------------------------------------------------------

    // The shape of every version in CHANGELOG.md: a heading, the dashes under it (a rule, as GitHub
    // draws it - never a heading, as nothing but a heading stands above them), nested lists, and
    // an alert after a blank line.
    [Fact]
    public void A_changelog_shaped_text_reads_as_headings_rules_nested_items_and_an_alert()
    {
        string markdown = Lines(
            "# 3.1.0, 2026-October-24",
            "------------------------",
            "- **Application**",
            "  - Added \"Bring my draft up to date\" button",
            "    - Exact same functionality",
            "  - `red` only for delete",
            "- **Data**",
            "  - Added `Rigol MHO900` oscilloscope",
            "",
            "",
            "# 3.0.0, 2026-October-11",
            "------------------------",
            "- Fixed it",
            "",
            "> [!CAUTION]",
            "> Due to breaking data format changes, the major version has been increased.",
            "> No more data updates will be given for versions **below 3.0.0**.",
            "",
            "",
            "# 2.5.0, 2026-September-9");

        Assert.Equal(
            new[]
            {
                "H1 (3.1.0, 2026-October-24)",
                "Rule",
                "Item0 b(Application)",
                "Item1 (Added \"Bring my draft up to date\" button)",
                "Item2 (Exact same functionality)",
                "Item1 c(red)|( only for delete)",
                "Item0 b(Data)",
                "Item1 (Added )|c(Rigol MHO900)|( oscilloscope)",
                "H1 (3.0.0, 2026-October-11)",
                "Rule",
                "Item0 (Fixed it)",
                "Quote:Caution0 [P0 (Due to breaking data format changes, the major version has been increased. No more data updates will be given for versions )|b(below 3.0.0)|(.)]",
                "H1 (2.5.0, 2026-September-9)",
            },
            Describe(markdown));
    }

    // CommonMark nests an item under the one above when it starts at that item's text column -
    // two spaces under "- ", as CHANGELOG.md indents, four levels deep.
    [Fact]
    public void Each_two_spaces_of_indentation_nest_an_item_one_level_deeper()
    {
        string markdown = Lines(
            "- a",
            "  - b",
            "    - c",
            "      - d",
            "    - c2",
            "- e");

        Assert.Equal(
            new[] { "Item0 (a)", "Item1 (b)", "Item2 (c)", "Item3 (d)", "Item2 (c2)", "Item0 (e)" },
            Describe(markdown));
    }

    // One space is short of the text column, so it is a sibling, not a child.
    [Fact]
    public void An_item_indented_short_of_the_text_above_is_its_sibling()
    {
        Assert.Equal(new[] { "Item0 (a)", "Item0 (b)" }, Describe(Lines("- a", " - b")));
    }

    // A tab counts to the next multiple of four, so a list indented with tabs nests as with spaces.
    [Fact]
    public void A_tab_in_the_indentation_nests_like_spaces()
    {
        Assert.Equal(new[] { "Item0 (a)", "Item1 (b)" }, Describe(Lines("- a", "\t- b")));
    }

    [Fact]
    public void All_three_bullet_characters_make_items()
    {
        Assert.Equal(new[] { "Item0 (a)", "Item0 (b)", "Item0 (c)" }, Describe(Lines("- a", "* b", "+ c")));
    }

    // An ordered list counts on from its first number, whatever the items say - as GitHub does.
    [Fact]
    public void An_ordered_list_counts_on_from_its_first_number()
    {
        Assert.Equal(new[] { "Item0#1 (a)", "Item0#2 (b)", "Item0#3 (c)" }, Describe(Lines("1. a", "1. b", "1. c")));
        Assert.Equal(new[] { "Item0#3 (a)", "Item0#4 (b)" }, Describe(Lines("3. a", "7. b")));
    }

    [Fact]
    public void Another_list_marker_starts_another_list()
    {
        Assert.Equal(new[] { "Item0#1 (a)", "Item0#1 (b)" }, Describe(Lines("1. a", "1) b")));
        Assert.Equal(new[] { "Item0#1 (a)", "Item0 (b)" }, Describe(Lines("1. a", "- b")));
    }

    // "1990. was a good year" on the second line of a sentence is no list - an ordered list may
    // break into a paragraph only when it starts at 1 (CommonMark).
    [Fact]
    public void A_number_inside_a_paragraph_is_no_list_unless_it_is_one()
    {
        Assert.Equal(new[] { "P0 (The year 1990. was good)" }, Describe(Lines("The year", "1990. was good")));
        Assert.Equal(new[] { "P0 (Steps:)", "Item0#1 (first)" }, Describe(Lines("Steps:", "1. first")));
    }

    // A line that carries an item's sentence on, indented or not, belongs to that item.
    [Fact]
    public void A_line_carrying_an_item_on_joins_it()
    {
        Assert.Equal(new[] { "Item0 (a sentence carried on)" }, Describe(Lines("- a sentence", "carried on")));
        Assert.Equal(new[] { "Item0 (a)", "Item1 (b and more)" }, Describe(Lines("- a", "  - b", "    and more")));
    }

    // After a blank line, text indented under an item is a paragraph of that item; text at the
    // margin ends the list.
    [Fact]
    public void A_blank_line_then_indented_text_stays_in_the_item_and_text_at_the_margin_ends_the_list()
    {
        Assert.Equal(
            new[] { "Item0 (a)", "P1 (more about a)", "Item0 (b)" },
            Describe(Lines("- a", "", "  more about a", "- b")));

        Assert.Equal(
            new[] { "Item0 (a)", "P0 (after the list)" },
            Describe(Lines("- a", "", "after the list")));
    }

    [Fact]
    public void An_item_with_nothing_after_its_bullet_is_still_an_item()
    {
        Assert.Equal(new[] { "Item0 ", "Item0 (b)" }, Describe(Lines("-", "- b")));
    }

    [Fact]
    public void A_heading_ends_the_list_above_it()
    {
        Assert.Equal(new[] { "Item0 (a)", "H2 (Next)", "Item0 (b)" }, Describe(Lines("- a", "## Next", "- b")));
    }

    [Theory]
    [InlineData("# One", "H1 (One)")]
    [InlineData("###### Six", "H6 (Six)")]
    [InlineData("# Title #", "H1 (Title)")]
    [InlineData("## Written in C#", "H2 (Written in C#)")]
    [InlineData("   # Indented three", "H1 (Indented three)")]
    [InlineData("#84 is an issue number", "P0 (#84 is an issue number)")]
    [InlineData("####### Seven is too many", "P0 (####### Seven is too many)")]
    public void A_heading_is_one_to_six_hashes_and_a_space(string line, string expected)
    {
        Assert.Equal(new[] { expected }, Describe(line));
    }

    [Theory]
    [InlineData("---")]
    [InlineData("------------------------")]
    [InlineData("***")]
    [InlineData("- - -")]
    [InlineData("___")]
    [InlineData("  ---")]
    public void A_line_of_three_or_more_dashes_stars_or_underscores_is_a_rule(string line)
    {
        Assert.Equal(new[] { "Rule" }, Describe(line));
    }

    // Under a paragraph of its own a line of = or - makes it a heading, as on GitHub; under a list
    // item's text the dashes are a rule.
    [Fact]
    public void A_line_of_equals_or_dashes_under_a_paragraph_makes_it_a_heading()
    {
        Assert.Equal(new[] { "H1 (Title)" }, Describe(Lines("Title", "=====")));
        Assert.Equal(new[] { "H2 (Sub title)" }, Describe(Lines("Sub title", "---")));
        Assert.Equal(new[] { "Item0 (item)", "Rule" }, Describe(Lines("- item", "---")));
    }

    [Fact]
    public void Paragraph_lines_join_with_a_space_and_a_blank_line_parts_paragraphs()
    {
        Assert.Equal(new[] { "P0 (one two)", "P0 (three)" }, Describe(Lines("one", "   two", "", "three")));
    }

    // -----------------------------------------------------------------------------------------
    // Alerts and quotes
    // -----------------------------------------------------------------------------------------

    [Theory]
    [InlineData("NOTE", "Note")]
    [InlineData("TIP", "Tip")]
    [InlineData("IMPORTANT", "Important")]
    [InlineData("WARNING", "Warning")]
    [InlineData("CAUTION", "Caution")]
    [InlineData("caution", "Caution")]
    public void Each_of_GitHubs_five_alerts_is_read(string marker, string kind)
    {
        Assert.Equal(
            new[] { $"Quote:{kind}0 [P0 (Words.)]" },
            Describe(Lines($"> [!{marker}]", "> Words.")));
    }

    // The blank line ends the alert; what follows is not drawn inside it.
    [Fact]
    public void An_alert_followed_by_a_blank_line_ends_there()
    {
        Assert.Equal(
            new[] { "Quote:Note0 [P0 (Inside.)]", "P0 (Outside.)" },
            Describe(Lines("> [!NOTE]", "> Inside.", "", "Outside.")));

        Assert.Equal(
            new[] { "Quote:Warning0 [P0 (Inside.)]", "H1 (2.0.0)" },
            Describe(Lines("> [!WARNING]", "> Inside.", "", "", "# 2.0.0")));
    }

    // A blank ">" line parts the alert's paragraphs, and a list inside reads as a list.
    [Fact]
    public void An_alert_holds_paragraphs_and_lists_of_its_own()
    {
        Assert.Equal(
            new[] { "Quote:Tip0 [P0 (First.) / P0 (Second:) / Item0 (one) / Item0 (two)]" },
            Describe(Lines("> [!TIP]", "> First.", ">", "> Second:", "> - one", "> - two")));
    }

    // A line without ">" that carries the alert's sentence on still belongs to it (CommonMark's
    // "lazy" line).
    [Fact]
    public void A_line_carrying_the_alerts_sentence_on_belongs_to_it()
    {
        Assert.Equal(
            new[] { "Quote:Note0 [P0 (One sentence carried on.)]" },
            Describe(Lines("> [!NOTE]", "> One sentence", "carried on.")));
    }

    [Fact]
    public void A_quote_without_a_marker_is_a_plain_quote()
    {
        Assert.Equal(new[] { "Quote0 [P0 (Said elsewhere.)]" }, Describe(Lines("> Said elsewhere.")));
    }

    // GitHub needs the marker on a line of its own.
    [Fact]
    public void A_marker_with_text_after_it_is_no_alert()
    {
        Assert.Equal(new[] { "Quote0 [P0 ([!NOTE] text)]" }, Describe(Lines("> [!NOTE] text")));
    }

    // -----------------------------------------------------------------------------------------
    // Code blocks
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void A_fenced_block_is_code_as_typed()
    {
        Assert.Equal(
            new[] { "Code0 dotnet test *all*\n  indented", "P0 (after)" },
            Describe(Lines("```bash", "dotnet test *all*", "  indented", "```", "after")));
    }

    [Fact]
    public void An_unclosed_fence_runs_to_the_end()
    {
        Assert.Equal(new[] { "Code0 a\nb" }, Describe(Lines("~~~", "a", "b")));
    }

    // -----------------------------------------------------------------------------------------
    // Line endings and the empty case
    // -----------------------------------------------------------------------------------------

    // CHANGELOG.md is checked out with Windows line endings on Windows and Unix ones on Linux; both
    // read the same, and so do old Mac ones.
    [Fact]
    public void Windows_and_old_Mac_line_endings_read_as_Unix_ones()
    {
        string[] lines = { "# 1.0.0", "-------", "- a", "  - b", "", "> [!NOTE]", "> c" };
        var unix = Describe(string.Join("\n", lines));

        Assert.Equal(unix, Describe(string.Join("\r\n", lines)));
        Assert.Equal(unix, Describe(string.Join("\r", lines)));
        Assert.Equal(7 - 2, unix.Count);
    }

    [Fact]
    public void A_byte_order_mark_is_not_part_of_the_first_heading()
    {
        Assert.Equal(new[] { "H1 (1.0.0)" }, Describe("\uFEFF# 1.0.0"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("  \n\n   ")]
    public void Nothing_reads_as_no_blocks(string? markdown)
    {
        Assert.Empty(GitHubMarkdown.Parse(markdown));
    }

    // -----------------------------------------------------------------------------------------
    // Bold and italic
    // -----------------------------------------------------------------------------------------

    [Theory]
    [InlineData("**Application**", "b(Application)")]
    [InlineData("Added **Commodore 128** / **310378 (Open128)**", "(Added )|b(Commodore 128)|( / )|b(310378 (Open128))")]
    [InlineData("Added __experimental__ possibility", "(Added )|b(experimental)|( possibility)")]
    [InlineData("will not blindly _execute_ it", "(will not blindly )|i(execute)|( it)")]
    [InlineData("an *italic* word", "(an )|i(italic)|( word)")]
    [InlineData("***both***", "bi(both)")]
    [InlineData("**bold with _italic_ inside**", "b(bold with )|bi(italic)|b( inside)")]
    public void Stars_and_underscores_make_bold_and_italic(string text, string expected)
    {
        Assert.Equal(expected, Inline(text));
    }

    // A slip in a hand-written changelog must not swallow the rest of the line into bold.
    [Theory]
    [InlineData("**not closed", "(**not closed)")]
    [InlineData("a ** b", "(a ** b)")]
    [InlineData("* a *", "(* a *)")]
    [InlineData("closed** only", "(closed** only)")]
    public void A_delimiter_without_a_partner_stays_as_typed(string text, string expected)
    {
        Assert.Equal(expected, Inline(text));
    }

    // GitHub leaves an underscore inside a word alone - names like snake_case are common in a
    // changelog - while a star inside a word does make emphasis.
    [Theory]
    [InlineData("snake_case_name", "(snake_case_name)")]
    [InlineData("the --workbooks_root and data_root options", "(the --workbooks_root and data_root options)")]
    [InlineData("_foo_bar_baz_", "i(foo_bar_baz)")]
    [InlineData("foo*bar*baz", "(foo)|i(bar)|(baz)")]
    public void An_underscore_inside_a_word_is_no_italic(string text, string expected)
    {
        Assert.Equal(expected, Inline(text));
    }

    // CommonMark's own examples for the delimiter rules - punctuation beside a delimiter, nesting,
    // and the "rule of three".
    [Theory]
    [InlineData("foo-_(bar)_", "(foo-)|i((bar))")]
    [InlineData("*foo**bar**baz*", "i(foo)|bi(bar)|i(baz)")]
    [InlineData("*foo**bar*", "i(foo**bar)")]
    [InlineData("**foo \"*bar*\" foo**", "b(foo \")|bi(bar)|b(\" foo)")]
    public void The_delimiter_rules_follow_CommonMark(string text, string expected)
    {
        Assert.Equal(expected, Inline(text));
    }

    // GitHub strikes through with one or two tildes - but a lone "~31%" is just a tilde.
    [Theory]
    [InlineData("~~gone~~ and ~this~", "s(gone)|( and )|s(this)")]
    [InlineData("1074 cases and ~31% coverage", "(1074 cases and ~31% coverage)")]
    [InlineData("~~~three~~~", "(~~~three~~~)")]
    public void Tildes_strike_through(string text, string expected)
    {
        Assert.Equal(expected, Inline(text));
    }

    // -----------------------------------------------------------------------------------------
    // Code
    // -----------------------------------------------------------------------------------------

    [Theory]
    [InlineData("Fixed `ESCAPE` did not close", "(Fixed )|c(ESCAPE)|( did not close)")]
    [InlineData("`CTRL`+`Z`", "c(CTRL)|(+)|c(Z)")]
    [InlineData("``a`b``", "c(a`b)")]
    [InlineData("` padded `", "c(padded)")]
    [InlineData("`**not bold**`", "c(**not bold**)")]
    [InlineData("`unclosed code", "(`unclosed code)")]
    [InlineData("**`7402` now**", "bc(7402)|b( now)")]
    public void Backticks_make_code_and_nothing_inside_is_read(string text, string expected)
    {
        Assert.Equal(expected, Inline(text));
    }

    // -----------------------------------------------------------------------------------------
    // Links
    // -----------------------------------------------------------------------------------------

    // The address is kept exactly as written - "%E2%80%90" is the Wiki's own escape for the
    // hyphen in "how-to", and decoding it would open another page.
    [Fact]
    public void A_link_keeps_its_address_exactly_as_written()
    {
        Assert.Equal(
            "(Read )|(MiniPro programmer documentation)<https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki/MiniPro-programmer-how%E2%80%90to>|( to see)",
            Inline("Read [MiniPro programmer documentation](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki/MiniPro-programmer-how%E2%80%90to) to see"));
    }

    [Theory]
    [InlineData("[How to use traces](https://youtu.be/Y55nC_gJbH4)", "(How to use traces)<https://youtu.be/Y55nC_gJbH4>")]
    [InlineData("[**bold** link](https://x.example)", "b(bold)<https://x.example>|( link)<https://x.example>")]
    [InlineData("[Foo](https://en.wikipedia.org/wiki/Foo_(bar))", "(Foo)<https://en.wikipedia.org/wiki/Foo_(bar)>")]
    [InlineData("[titled](https://x.example \"A title\")", "(titled)<https://x.example>")]
    [InlineData("[angle](<https://x.example/a b>)", "(angle)<https://x.example/a b>")]
    [InlineData("![CRT icon](https://x.example/icon.png)", "(CRT icon)<https://x.example/icon.png>")]
    public void Square_brackets_and_an_address_make_a_link(string text, string expected)
    {
        Assert.Equal(expected, Inline(text));
    }

    [Theory]
    [InlineData("[no address]", "([no address])")]
    [InlineData("[space] (somewhere)", "([space] (somewhere))")]
    [InlineData("[unclosed](somewhere", "([unclosed](somewhere)")]
    public void Brackets_that_complete_no_link_stay_as_typed(string text, string expected)
    {
        Assert.Equal(expected, Inline(text));
    }

    // What is left is read on, as on GitHub - a bare address in it is still a link.
    [Fact]
    public void An_address_in_brackets_that_complete_no_link_is_still_linked_bare()
    {
        Assert.Equal(
            "([unclosed]()|(https://x.example)<https://x.example>",
            Inline("[unclosed](https://x.example"));
    }

    // GitHub links a bare address too, less the sentence's own full stop.
    [Theory]
    [InlineData("see https://example.com.", "(see )|(https://example.com)<https://example.com>|(.)")]
    [InlineData("see www.lemon64.com today", "(see )|(www.lemon64.com)<https://www.lemon64.com>|( today)")]
    [InlineData("**https://example.com/a**", "b(https://example.com/a)<https://example.com/a>")]
    [InlineData("xhttps://example.com", "(xhttps://example.com)")]
    [InlineData("the https protocol", "(the https protocol)")]
    public void A_bare_web_address_is_a_link(string text, string expected)
    {
        Assert.Equal(expected, Inline(text));
    }

    [Theory]
    [InlineData("<https://example.com/a>", "(https://example.com/a)<https://example.com/a>")]
    [InlineData("mail <anna@example.org>", "(mail )|(anna@example.org)<mailto:anna@example.org>")]
    [InlineData("a <b> tag", "(a <b> tag)")]
    public void An_address_in_angle_brackets_is_a_link(string text, string expected)
    {
        Assert.Equal(expected, Inline(text));
    }

    // A link's text holds no second link: the address shown inside it goes where the link goes.
    [Fact]
    public void An_address_inside_a_links_text_is_not_a_second_link()
    {
        Assert.Equal("(https://a.example)<https://b.example>", Inline("[https://a.example](https://b.example)"));
    }

    // -----------------------------------------------------------------------------------------
    // Escapes, entities, line breaks and emoji
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void A_backslash_before_punctuation_shows_the_punctuation_as_typed()
    {
        Assert.Equal("(*not italic*)", Inline(Backslash + "*not italic" + Backslash + "*"));
        Assert.Equal("(C:" + Backslash + "path)", Inline("C:" + Backslash + "path"));
    }

    [Fact]
    public void Html_entities_are_decoded_and_a_lone_ampersand_stays()
    {
        Assert.Equal("(AT&T A < & so &bogus;)", Inline("AT&amp;T &#65; &lt; & so &bogus;"));
    }

    // Two spaces or a backslash at a line's end break the line; otherwise the line goes on.
    [Fact]
    public void A_line_end_is_a_space_unless_it_is_a_hard_break()
    {
        Assert.Equal(new[] { "P0 (a b)" }, Describe(Lines("a", "b")));
        Assert.Equal(new[] { "P0 (a)|(\n)|(b)" }, Describe(Lines("a  ", "b")));
        Assert.Equal(new[] { "P0 (a)|(\n)|(b)" }, Describe(Lines("a" + Backslash, "b")));
    }

    [Fact]
    public void An_emoji_is_kept_as_it_is()
    {
        Assert.Equal("(no longer supported \uD83D\uDE25)", Inline("no longer supported \uD83D\uDE25"));
    }

    // -----------------------------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------------------------

    private static string Lines(params string[] lines) => string.Join("\n", lines);

    private static List<string> Describe(string markdown) =>
        GitHubMarkdown.Parse(markdown).Select(Describe).ToList();

    private static string Inline(string text) => Describe(GitHubMarkdown.ParseInline(text));

    private static string Describe(MarkdownBlock block) => block switch
    {
        MarkdownHeading heading => $"H{heading.Level} {Describe(heading.Spans)}",
        MarkdownRule => "Rule",
        MarkdownListItem item => $"Item{item.Depth}{(item.Number is int number ? "#" + number : string.Empty)} {Describe(item.Spans)}",
        MarkdownParagraph paragraph => $"P{paragraph.Indent} {Describe(paragraph.Spans)}",
        MarkdownQuote quote => $"Quote{(quote.Alert is MarkdownAlertKind kind ? ":" + kind : string.Empty)}{quote.Indent} [{string.Join(" / ", quote.Blocks.Select(Describe))}]",
        MarkdownCode code => $"Code{code.Indent} {code.Text}",
        _ => block.GetType().Name,
    };

    private static string Describe(IReadOnlyList<MarkdownSpan> spans) => string.Join("|", spans.Select(Describe));

    private static string Describe(MarkdownSpan span)
    {
        string styles =
            (span.Style.HasFlag(MarkdownStyle.Bold) ? "b" : string.Empty) +
            (span.Style.HasFlag(MarkdownStyle.Italic) ? "i" : string.Empty) +
            (span.Style.HasFlag(MarkdownStyle.Code) ? "c" : string.Empty) +
            (span.Style.HasFlag(MarkdownStyle.Strikethrough) ? "s" : string.Empty);

        return $"{styles}({span.Text})" + (span.Url != null ? $"<{span.Url}>" : string.Empty);
    }
}
