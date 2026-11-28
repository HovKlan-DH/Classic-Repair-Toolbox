using Handlers.DataHandling;

namespace CRT.Data.Tests
{
    // How a collection's link is written for where it is pasted (owner decision, 2026-11-28) - one rule for CRT's "Copy link" and
    // the page's "Link" panel.
    public sealed class CollectionLinkFormatsTests
    {
        private const string Url = "https://classic-repair-toolbox.dk/collections/k7mq3x9vbn2p#view.U7-12";

        [Fact]
        public void Each_format_writes_the_link_as_its_kind_of_forum_reads_it()
        {
            Assert.Equal($"[url={Url}]U7, pin 12 - /RAS[/url]", CollectionLinkFormats.Write(CollectionLinkFormat.BbCode, Url, "U7, pin 12 - /RAS"));
            Assert.Equal($"[U7, pin 12 - /RAS]({Url})", CollectionLinkFormats.Write(CollectionLinkFormat.Markdown, Url, "U7, pin 12 - /RAS"));
            Assert.Equal($"U7, pin 12 - /RAS: {Url}", CollectionLinkFormats.Write(CollectionLinkFormat.Plain, Url, "U7, pin 12 - /RAS"));
            Assert.Equal($"<a href=\"{Url}\">U7, pin 12 - /RAS</a>", CollectionLinkFormats.Write(CollectionLinkFormat.Html, Url, "U7, pin 12 - /RAS"));
        }

        // Owner request, 2026-11-28: "add an option of "HTML" ... embed the link in "<a href>"" - what a title holds stays words, so
        // a "<" or a quote typed in it can neither end the link nor open a tag.
        [Fact]
        public void An_HTML_link_encodes_its_words_and_its_address() =>
            Assert.Equal(
                "<a href=\"https://x/collections/k?a=1&amp;b=&quot;2&quot;\">&lt;script&gt; &amp; [C64]</a>",
                CollectionLinkFormats.Write(CollectionLinkFormat.Html, "https://x/collections/k?a=1&b=\"2\"", "<script> & [C64]"));

        // A title typed with "[/url]" in it would end a BBCode link, and a "]" a Markdown link's words.
        [Fact]
        public void Square_brackets_in_the_words_are_written_round()
        {
            Assert.Equal($"[url={Url}]Help (/url) me[/url]", CollectionLinkFormats.Write(CollectionLinkFormat.BbCode, Url, "Help [/url] me"));
            Assert.Equal($"[No picture (C64)]({Url})", CollectionLinkFormats.Write(CollectionLinkFormat.Markdown, Url, "No picture [C64]"));
        }

        [Fact]
        public void A_plain_link_with_no_words_is_the_address_alone() =>
            Assert.Equal(Url, CollectionLinkFormats.Write(CollectionLinkFormat.Plain, Url, string.Empty));

        // Stored by name in the settings - one not known, a number or none reads as BBCode, the most forums' format.
        [Theory]
        [InlineData("Markdown", CollectionLinkFormat.Markdown)]
        [InlineData("plain", CollectionLinkFormat.Plain)]
        [InlineData("Html", CollectionLinkFormat.Html)]
        [InlineData("BbCode", CollectionLinkFormat.BbCode)]
        [InlineData("Reddit", CollectionLinkFormat.BbCode)]
        [InlineData("1", CollectionLinkFormat.BbCode)]
        [InlineData("", CollectionLinkFormat.BbCode)]
        [InlineData(null, CollectionLinkFormat.BbCode)]
        public void A_stored_format_is_read_back_by_its_name(string? stored, CollectionLinkFormat expected) =>
            Assert.Equal(expected, CollectionLinkFormats.Parse(stored));

        [Fact]
        public void Every_format_has_a_name_and_says_where_it_works()
        {
            Assert.Equal(["BBCode", "Markdown", "HTML", "Plain link"], CollectionLinkFormats.All.Select(CollectionLinkFormats.NameOf));
            Assert.All(CollectionLinkFormats.All, format => Assert.False(string.IsNullOrWhiteSpace(CollectionLinkFormats.WhereOf(format))));
            Assert.Contains("Lemon64", CollectionLinkFormats.WhereOf(CollectionLinkFormat.BbCode), StringComparison.Ordinal);
            Assert.Contains("Reddit", CollectionLinkFormats.WhereOf(CollectionLinkFormat.Markdown), StringComparison.Ordinal);
        }
    }
}
