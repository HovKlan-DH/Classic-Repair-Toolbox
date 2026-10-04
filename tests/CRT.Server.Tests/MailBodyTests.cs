using CRT.Server.Handlers.Email;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // MailBody - one mail written once, rendered as HTML and as plain text (owner request,
    // 2026-10-03: "All mails should be sent in HTML format").
    //
    // The two renderings are the whole point: they must say the same thing, and the HTML must never
    // carry markup that came from somebody's typing. Each block kind is pinned in both.
    // ###########################################################################################
    public sealed class MailBodyTests
    {
        [Fact]
        public void A_paragraph_renders_each_style_in_html_and_plain_in_text()
        {
            MailBody body = new MailBody().Paragraph(
                "Plain, ", MailText.Bold("bold"), ", ", MailText.Italic("italic"), ", ", MailText.Named("C64"), ", ",
                MailText.Link("a link", "mailto:x@example.com"));

            Assert.Equal("Plain, bold, italic, [C64], a link", body.ToText());

            Assert.Contains(
                "Plain, <b>bold</b>, <i>italic</i>, [<b>C64</b>], <a href=\"mailto:x@example.com\" style=\"color:#1a5fb4;\">a link</a>",
                body.ToHtml("Subject"),
                StringComparison.Ordinal);
        }

        // A quotation keeps its lines - indented in the text, line breaks and italics in the HTML.
        [Fact]
        public void A_quote_keeps_its_lines()
        {
            MailBody body = new MailBody().Quote("First line\r\nSecond line  ", "(none)");

            Assert.Equal("    First line\r\n    Second line", body.ToText());
            Assert.Contains("<i>First line<br>Second line</i>", body.ToHtml("S"), StringComparison.Ordinal);
        }

        [Fact]
        public void A_blank_quote_says_what_stands_in_for_it()
        {
            Assert.Equal("    (no description given)", new MailBody().Quote("   ", "(no description given)").ToText());
        }

        [Fact]
        public void Steps_are_numbered_in_text_and_an_ordered_list_in_html()
        {
            MailBody body = new MailBody().Steps(["Install CRT."], ["Open the ", MailText.Bold("Maintainer"), " tab."]);

            Assert.Equal("1. Install CRT.\r\n2. Open the Maintainer tab.", body.ToText());

            string html = body.ToHtml("S");
            Assert.Contains("<ol", html, StringComparison.Ordinal);
            Assert.Contains(">Open the <b>Maintainer</b> tab.</li>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void A_code_is_on_its_own_and_fixed_width()
        {
            MailBody body = new MailBody().Code("  ABC-123  ");

            Assert.Equal("    ABC-123", body.ToText());
            Assert.Contains("monospace", body.ToHtml("S"), StringComparison.Ordinal);
            Assert.Contains(">ABC-123</span>", body.ToHtml("S"), StringComparison.Ordinal);
        }

        // Blocks are separated by one blank line, with CRLF throughout - what SMTP wants.
        [Fact]
        public void Blocks_are_a_blank_line_apart_with_CRLF()
        {
            Assert.Equal("One\r\n\r\nTwo", new MailBody().Paragraph("One").Paragraph("Two").ToText());
        }

        // Every run is escaped - a link's address and the subject in the page title too.
        [Fact]
        public void Everything_is_escaped_in_the_html()
        {
            string html = new MailBody()
                .Paragraph("<b>", MailText.Bold("<i>"), MailText.Link("<x>", "mailto:a@b.com\"><script>"))
                .Quote("<script>", string.Empty)
                .ToHtml("<title>");

            Assert.DoesNotContain("<script>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<title><title>", html, StringComparison.Ordinal);
            Assert.Contains("&lt;b&gt;<b>&lt;i&gt;</b>", html, StringComparison.Ordinal);
            Assert.Contains("href=\"mailto:a@b.com&quot;&gt;&lt;script&gt;\"", html, StringComparison.Ordinal);
        }

        [Fact]
        public void The_message_carries_both_bodies()
        {
            EmailMessage message = new MailBody().Paragraph("Hi there,").ToMessage("a@b.com", "Hello");

            Assert.Equal("a@b.com", message.ToAddress);
            Assert.Equal("Hello", message.Subject);
            Assert.Equal("Hi there,", message.Body);
            Assert.Contains("<title>Hello</title>", message.HtmlBody, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // A log in the feedback mail (owner request, 2026-10-03: "logfile should still show as
        // monospace as this is kind of quoted text"): line for line, fixed width, escaped - and in
        // the plain text exactly as written, with no indent added.
        // ###########################################################################################
        [Fact]
        public void Preformatted_text_keeps_its_lines_in_a_fixed_width_font_and_is_never_markup()
        {
            MailBody body = new MailBody(MailBody.WideWidth).Preformatted("12:00 <Started>\r\n    indented & more\n");

            string html = body.ToHtml("Log");

            Assert.Contains("<pre style=", html, StringComparison.Ordinal);
            Assert.Contains("monospace", html, StringComparison.Ordinal);
            Assert.Contains("12:00 &lt;Started&gt;\n    indented &amp; more</pre>", html, StringComparison.Ordinal);
            Assert.Contains($"max-width:{MailBody.WideWidth}px", html, StringComparison.Ordinal);

            Assert.Equal("12:00 <Started>\r\n    indented & more", body.ToText());
        }

        [Fact]
        public void A_reply_to_address_travels_with_the_message()
        {
            EmailMessage message = new MailBody().Paragraph("Hello").ToMessage("to@example.com", "Subject", "reply@example.com");

            Assert.Equal("reply@example.com", message.ReplyToAddress);
            Assert.Null(new MailBody().Paragraph("Hello").ToMessage("to@example.com", "Subject").ReplyToAddress);
        }
    }
}
