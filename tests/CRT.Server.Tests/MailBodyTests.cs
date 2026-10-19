using CRT.Server.Handlers.Email;
using Handlers.DataHandling;
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

        // A quotation with a picture in it (a feedback's pasted screenshot, 2026-10-15, owner report
        // the same day: inside the sender's words, not a box apart): ONE panel holding the
        // words, the picture by its content id, and the words after it; the plain text names the
        // picture in brackets - the marker CRT wrote - on a line of the quotation; the message
        // carries the bytes as a linked image. The name is the sender's (a dropped file's), so it is
        // encoded like every other run.
        [Fact]
        public void A_quote_holds_its_pictures_between_its_words()
        {
            var image = new MailLinkedImage("picture-1@crt-feedback", "shot <1>.png", "image/png", [1, 2, 3]);
            MailBody body = new MailBody().Quote([MailQuotePart.Words("Above"), MailQuotePart.Picture(image), MailQuotePart.Words("Below")], "(none)");

            Assert.Equal("    Above\r\n    [shot <1>.png]\r\n    Below", body.ToText());

            string html = body.ToHtml("S");
            Assert.Contains(
                MailBody.QuotePanelStart + "<i>Above</i><img src=\"cid:picture-1@crt-feedback\" alt=\"shot &lt;1&gt;.png\"",
                html,
                StringComparison.Ordinal);
            Assert.Contains("><i>Below</i>" + MailBody.QuotePanelEnd, html, StringComparison.Ordinal);
            Assert.DoesNotContain("shot <1>.png", html, StringComparison.Ordinal);

            EmailMessage message = body.ToMessage("to@example.com", "S");
            Assert.Same(image, Assert.Single(message.LinkedImages!));
        }

        // Blank words around a picture are dropped, and a quotation left with nothing says what
        // stands in for it - as the one-text quotation always has.
        [Fact]
        public void A_quote_drops_blank_words_and_says_what_stands_in_for_nothing()
        {
            var image = new MailLinkedImage("p@x", "a.png", "image/png", [1]);

            Assert.Equal("    [a.png]", new MailBody().Quote([MailQuotePart.Words(" \n "), MailQuotePart.Picture(image), MailQuotePart.Words("")], "(none)").ToText());
            Assert.Equal("    (none)", new MailBody().Quote([MailQuotePart.Words("  ")], "(none)").ToText());
        }

        // A picture given a width (its sender's, in CRT's field) is shown at it - as an attribute too,
        // which a client ignoring styles still reads - and one without keeps its own.
        [Fact]
        public void A_picture_given_a_width_is_shown_at_it()
        {
            var sized = new MailLinkedImage("a@x", "a.png", "image/png", [1], Width: 480);
            var natural = new MailLinkedImage("b@x", "b.png", "image/png", [1]);

            string html = new MailBody().Quote([MailQuotePart.Picture(sized), MailQuotePart.Picture(natural)], "(none)").ToHtml("S");

            Assert.Contains("<img src=\"cid:a@x\" alt=\"a.png\" width=\"480\" style=\"display:block; width:480px; max-width:100%;", html, StringComparison.Ordinal);
            Assert.Contains("<img src=\"cid:b@x\" alt=\"b.png\" style=\"display:block; max-width:100%;", html, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // CRT's Feedback tab shows a picture no wider than FeedbackComposition.MaxPictureWidth, as
        // the size it will have in the mail (2026-10-15): the mail's quotation must have room for
        // that, or the field would promise a size the mail squeezes. One side moving alone fails here.
        // ###########################################################################################
        [Fact]
        public void The_widest_picture_CRT_offers_fits_the_feedback_mails_quotation()
        {
            Assert.True(
                MailBody.WidestQuotedPicture >= FeedbackComposition.MaxPictureWidth,
                $"The quotation has room for {MailBody.WidestQuotedPicture} px; CRT offers {FeedbackComposition.MaxPictureWidth}.");
            Assert.Contains(
                $"padding:6px {MailBody.QuotePaddingRight}px 6px {MailBody.QuotePaddingLeft}px; border-left:{MailBody.QuoteBarWidth}px solid",
                new MailBody(MailBody.WideWidth).Quote("x", "(none)").ToHtml("S"),
                StringComparison.Ordinal);
            Assert.Contains($"max-width:{MailBody.WideWidth}px", new MailBody(MailBody.WideWidth).Quote("x", "(none)").ToHtml("S"), StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** THE QUOTATION'S PANEL IS DRAWN BY THE MAIL ITSELF (owner report, 2026-10-16: a feedback of
        // pictures alone showed them outside "that blue'ish panel"). *** The panel was never ours:
        // Fastmail paints its own #E8ECF2 behind a <blockquote> with words in it, and not behind one
        // holding only a picture - the HTML had the picture inside all along. So the quotation is a
        // plain block carrying the panel's background itself, the same with words, pictures or both,
        // in any mail program.
        // ###########################################################################################
        [Fact]
        public void A_quotation_is_a_panel_the_mail_draws_itself_with_words_or_pictures_alone()
        {
            var image = new MailLinkedImage("p@x", "a.png", "image/png", [1]);

            foreach (MailBody body in new[]
            {
                new MailBody().Quote("Only words", "(none)"),
                new MailBody().Quote([MailQuotePart.Picture(image)], "(none)"),
            })
            {
                string html = body.ToHtml("S");

                Assert.DoesNotContain("<blockquote", html, StringComparison.Ordinal);
                Assert.Contains("background-color:#e8ecf2;", html, StringComparison.Ordinal);
            }
        }

        // The one-text quotation every other mail uses renders exactly as it did before pictures.
        [Fact]
        public void A_quote_of_words_alone_is_one_italic_run()
        {
            Assert.Contains(
                MailBody.QuotePanelStart + "<i>Only words</i>" + MailBody.QuotePanelEnd,
                new MailBody().Quote("Only words", "(none)").ToHtml("S"),
                StringComparison.Ordinal);
        }

        // A picture shown more than once - in one quote or in two - is carried ONCE (code review,
        // 2026-10-16): every linked image is a base64 copy in the mail, while the HTML may refer to
        // one "cid:" as often as it likes.
        [Fact]
        public void A_picture_shown_more_than_once_is_carried_once()
        {
            var image = new MailLinkedImage("p@x", "a.png", "image/png", [1]);
            var other = new MailLinkedImage("q@x", "b.png", "image/png", [2]);

            MailBody body = new MailBody()
                .Quote([MailQuotePart.Picture(image), MailQuotePart.Words("again"), MailQuotePart.Picture(image)], "(none)")
                .Quote([MailQuotePart.Picture(other), MailQuotePart.Picture(image)], "(none)");

            Assert.Equal([image, other], body.ToMessage("to@example.com", "S").LinkedImages!);
            Assert.Equal(3, body.ToHtml("S").Split("cid:p@x").Length - 1);
        }

        // A mail without pictures carries none - the sender then adds no linked resources.
        [Fact]
        public void A_body_without_pictures_carries_no_linked_images()
        {
            Assert.Null(new MailBody().Paragraph("Words").ToMessage("to@example.com", "S").LinkedImages);
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
