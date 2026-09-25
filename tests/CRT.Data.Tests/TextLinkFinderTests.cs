using System.Linq;
using Handlers.DataHandling;
using Xunit;

namespace Classic_Repair_Toolbox.Tests
{
    // ###########################################################################################
    // TextLinkFinder turns a user-typed repair note into plain and link runs so the UI can make the
    // links clickable. Two properties matter above all and are asserted repeatedly here:
    //
    //   1. The spans always concatenate back to the ORIGINAL text. A dropped or doubled character
    //      is text loss on screen, in a field the user typed by hand.
    //   2. Ordinary repair prose is NOT turned into links. These fields are full of part numbers
    //      ("74LS08"), measurements ("5.0V") and file names, and a false link is worse than none.
    // ###########################################################################################
    public class TextLinkFinderTests
    {
        // The spans are a partition of the input, so re-joining them must give it back verbatim.
        private static string Rejoin(string text) =>
            string.Concat(TextLinkFinder.FindSpans(text).Select(span => text.Substring(span.Start, span.Length)));

        [Fact]
        public void Blank_text_produces_no_spans_at_all()
        {
            // Nothing to render - a caller must not have to special-case an empty plain span.
            Assert.Empty(TextLinkFinder.FindSpans(null));
            Assert.Empty(TextLinkFinder.FindSpans(string.Empty));
        }

        [Fact]
        public void Text_with_no_link_is_one_plain_span()
        {
            const string text = "Replaced U7 and reflowed the joints on CN2.";

            var spans = TextLinkFinder.FindSpans(text);

            var span = Assert.Single(spans);
            Assert.False(span.IsLink);
            Assert.Null(span.Url);
            Assert.Equal(0, span.Start);
            Assert.Equal(text.Length, span.Length);
            Assert.False(TextLinkFinder.ContainsLink(text));
        }

        [Theory]
        [InlineData("https://example.com")]
        [InlineData("http://example.com")]
        [InlineData("https://example.com/page?a=1&b=2#frag")]
        public void An_explicit_scheme_is_a_link_and_is_opened_exactly_as_typed(string text)
        {
            var span = Assert.Single(TextLinkFinder.FindSpans(text));

            Assert.True(span.IsLink);
            Assert.Equal(text, span.Url);
            Assert.Equal(text, Rejoin(text));
        }

        [Fact]
        public void A_www_link_displays_as_typed_but_opens_with_https_prepended()
        {
            // The user typed no scheme, so the displayed run must stay "www...." while the target
            // handed to the launcher has to be absolute or it cannot be opened at all.
            const string text = "www.classic-repair-toolbox.dk";

            var span = Assert.Single(TextLinkFinder.FindSpans(text));

            Assert.Equal(text.Length, span.Length);
            Assert.Equal("https://www.classic-repair-toolbox.dk", span.Url);
        }

        [Fact]
        public void A_link_inside_a_sentence_splits_into_plain_link_plain()
        {
            const string text = "See https://example.com for the pinout";

            var spans = TextLinkFinder.FindSpans(text);

            Assert.Equal(3, spans.Count);
            Assert.False(spans[0].IsLink);
            Assert.Equal("See ", text.Substring(spans[0].Start, spans[0].Length));
            Assert.True(spans[1].IsLink);
            Assert.Equal("https://example.com", text.Substring(spans[1].Start, spans[1].Length));
            Assert.False(spans[2].IsLink);
            Assert.Equal(" for the pinout", text.Substring(spans[2].Start, spans[2].Length));
            Assert.Equal(text, Rejoin(text));
        }

        [Fact]
        public void Several_links_in_one_text_are_all_found_and_the_text_survives_intact()
        {
            const string text = "Refs: https://a.example and www.b.example plus http://c.example done";

            var spans = TextLinkFinder.FindSpans(text);

            Assert.Equal(3, spans.Count(s => s.IsLink));
            Assert.Equal(text, Rejoin(text));
        }

        [Theory]
        // Sentence punctuation sits against a URL far more often than a URL ends in one.
        [InlineData("Try https://example.com.", "https://example.com")]
        [InlineData("Try https://example.com, then reflow", "https://example.com")]
        [InlineData("Try https://example.com!", "https://example.com")]
        [InlineData("Try https://example.com;", "https://example.com")]
        [InlineData("Try (https://example.com)", "https://example.com")]
        [InlineData("Try \"https://example.com\"", "https://example.com")]
        public void Trailing_sentence_punctuation_is_not_part_of_the_link(string text, string expectedLink)
        {
            var link = Assert.Single(TextLinkFinder.FindSpans(text), s => s.IsLink);

            Assert.Equal(expectedLink, text.Substring(link.Start, link.Length));
            Assert.Equal(text, Rejoin(text));
        }

        [Fact]
        public void A_closing_bracket_that_belongs_to_the_url_is_kept()
        {
            // The Wikipedia-style case: the ")" has its own "(" inside the URL, so stripping it
            // would produce a link that 404s. Balance is what tells this from "(https://x)".
            const string text = "https://en.wikipedia.org/wiki/MOS_Technology_6502_(CPU)";

            var span = Assert.Single(TextLinkFinder.FindSpans(text));

            Assert.Equal(text, span.Url);
        }

        [Theory]
        // The whole reason detection is restricted to explicit prefixes: repair notes are full of
        // dotted tokens that are not links and must never be rendered as one.
        [InlineData("Replaced the 74LS08.pin3 trace")]
        [InlineData("Measured 5.0V on the rail")]
        [InlineData("See notes.txt in the folder")]
        [InlineData("example.com")]
        [InlineData("Board rev 250407.b")]
        public void Ordinary_repair_prose_is_never_turned_into_a_link(string text)
        {
            Assert.False(TextLinkFinder.ContainsLink(text));
            Assert.DoesNotContain(TextLinkFinder.FindSpans(text), span => span.IsLink);
        }

        [Theory]
        // A link must START at a word boundary, or the clickable run would begin mid-word.
        [InlineData("xhttp://example.com")]
        [InlineData("somewww.thing")]
        [InlineData("a_www.example.com")]
        public void A_prefix_in_the_middle_of_a_word_is_not_a_link(string text)
        {
            Assert.False(TextLinkFinder.ContainsLink(text));
        }

        [Theory]
        // A prefix with no host after it has nothing to open.
        [InlineData("https://")]
        [InlineData("Ends with www.")]
        [InlineData("http://")]
        public void A_bare_prefix_with_no_host_is_not_a_link(string text)
        {
            Assert.False(TextLinkFinder.ContainsLink(text));
        }

        [Fact]
        public void Scheme_matching_is_case_insensitive()
        {
            // People type "HTTPS://" and "WWW." - and the displayed run keeps their casing.
            const string text = "HTTPS://Example.COM and WWW.Example.COM";

            var links = TextLinkFinder.FindSpans(text).Where(s => s.IsLink).ToList();

            Assert.Equal(2, links.Count);
            Assert.Equal("HTTPS://Example.COM", text.Substring(links[0].Start, links[0].Length));
            Assert.Equal("HTTPS://Example.COM", links[0].Url);
            Assert.Equal("WWW.Example.COM", text.Substring(links[1].Start, links[1].Length));
            Assert.Equal("https://WWW.Example.COM", links[1].Url);
        }

        [Fact]
        public void A_link_across_a_newline_stops_at_the_line_break()
        {
            // Description is a multi-line field, so a URL at the end of a line must not swallow the
            // next line's first word.
            const string text = "https://example.com\nNext line";

            var spans = TextLinkFinder.FindSpans(text);

            var link = Assert.Single(spans, s => s.IsLink);
            Assert.Equal("https://example.com", text.Substring(link.Start, link.Length));
            Assert.Equal(text, Rejoin(text));
        }

        [Fact]
        public void A_link_at_the_very_start_produces_no_empty_leading_span()
        {
            // An empty Run renders as nothing but still costs a layout pass - and an empty leading
            // span would break the "spans partition the text" reading for anyone debugging it.
            const string text = "https://example.com is the source";

            var spans = TextLinkFinder.FindSpans(text);

            Assert.Equal(2, spans.Count);
            Assert.True(spans[0].IsLink);
            Assert.All(spans, span => Assert.True(span.Length > 0));
        }

        [Fact]
        public void End_reports_the_exclusive_end_of_the_span()
        {
            const string text = "go https://example.com";

            var link = Assert.Single(TextLinkFinder.FindSpans(text), s => s.IsLink);

            Assert.Equal(link.Start + link.Length, link.End);
            Assert.Equal(text.Length, link.End);
        }
    }
}
