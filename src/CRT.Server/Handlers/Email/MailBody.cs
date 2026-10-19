using System.Net;
using System.Text;

namespace CRT.Server.Handlers.Email
{
    // ###########################################################################################
    // ONE MAIL, WRITTEN ONCE, SENT AS HTML (owner request, 2026-10-03: "All mails should be sent in
    // HTML format, as I expect all email clients will support that").
    //
    // A template builds its mail from a few kinds of block - a paragraph, a quotation (somebody's
    // own words), a code to paste, numbered steps - and this renders the SAME blocks twice: as HTML,
    // which every mail client shows, and as plain text, which goes along as the alternative part.
    // The plain part is not for showing - spam filters score an HTML-only mail as suspicious - and
    // because both come from the one list of blocks they can never say different things.
    //
    // *** EVERYTHING IS ESCAPED HERE, NOT IN THE TEMPLATES. *** A contributor's description, a
    // maintainer's comment and a display name are text somebody typed, and they reach a mail that
    // goes to somebody else - so no template ever writes markup itself, and every run of text is
    // HTML-encoded on its way out. A template that wants bold says so with a style, never with "<b>".
    //
    // Inline styles only: several mail clients drop a <style> block, and none drops a style
    // attribute. No remote resources - a mail that loads nothing cannot be tracked and is not held
    // back by a client that blocks remote content. The one kind of picture (a quote's, since
    // 2026-10-15: a feedback's pasted screenshots) travels INSIDE the mail as a linked resource
    // the HTML refers to by "cid:", which such a client shows.
    // ###########################################################################################
    public sealed class MailBody
    {
        private const string FontStack = "'Segoe UI', Helvetica, Arial, sans-serif";

        private const string MonospaceStack = "Consolas, Menlo, 'Courier New', monospace";

        // A mail's width, as a reading column. The feedback mail is wider: it carries log files, and
        // pictures up to CRT's widest beside the quotation's panel (1020 since 2026-10-16, when the
        // panel got its right padding).
        public const int StandardWidth = 640;
        public const int WideWidth = 1020;

        // ###########################################################################################
        // *** A QUOTATION IS A PANEL THE MAIL DRAWS ITSELF (owner report, 2026-10-16). *** The
        // "blue'ish panel" the owner saw behind the sender's words was Fastmail's own, painted behind
        // a <blockquote> with words in it - and not behind one holding only a picture, so a feedback
        // of pictures alone showed them bare. A quotation is therefore a plain <div> carrying the
        // panel itself: Fastmail's colour (#E8ECF2, measured off the owner's screen), the grey bar on
        // its left, padding all round - the same in every mail program, with words, pictures or both.
        // A <div> also keeps a client from folding it away as "quoted text", which a feedback's own
        // words never are.
        // ###########################################################################################
        public const int QuotePaddingLeft = 14;
        public const int QuotePaddingRight = 12;
        public const int QuoteBarWidth = 3;
        public const string QuoteBackground = "#e8ecf2";

        public static readonly string QuotePanelStart =
            $"<div style=\"margin:0 0 14px 0; padding:6px {MailBody.QuotePaddingRight}px 6px {MailBody.QuotePaddingLeft}px; " +
            $"border-left:{MailBody.QuoteBarWidth}px solid #c8c8c8; background-color:{MailBody.QuoteBackground}; border-radius:4px; color:#444444;\">";

        public const string QuotePanelEnd = "</div>";

        // ###########################################################################################
        // The widest a picture in a wide mail's quotation is shown without being squeezed: the
        // column, less the panel's padding and bar and the picture's own 1px border each side. CRT's
        // Feedback tab offers no wider (CRT.Data's FeedbackComposition.MaxPictureWidth - a test
        // holds the two together).
        // ###########################################################################################
        public const int WidestQuotedPicture =
            MailBody.WideWidth - MailBody.QuotePaddingLeft - MailBody.QuotePaddingRight - MailBody.QuoteBarWidth - 2;

        private readonly List<Block> thisBlocks = [];

        // The pictures the quotes show, handed to the message as its linked images - each ONCE,
        // however often a quote shows it (code review, 2026-10-16): the HTML can refer to one
        // "cid:" as often as it likes, but every linked image is a base64 copy in the mail.
        private readonly List<MailLinkedImage> thisImages = [];
        private readonly HashSet<string> thisImageIds = new(StringComparer.Ordinal);

        private readonly int thisMaxWidth;

        public MailBody(int maxWidth = MailBody.StandardWidth)
        {
            this.thisMaxWidth = maxWidth;
        }

        // A paragraph of runs - plain, bold, italic, a board's name, a link.
        public MailBody Paragraph(params MailText[] parts)
        {
            this.thisBlocks.Add(new Block(BlockKind.Paragraph, parts, string.Empty, []));
            return this;
        }

        // Somebody's own words - a contributor's description, a maintainer's comment - set apart
        // and in italics, so they read as a quotation and not as the service speaking. `whenBlank`
        // stands in for nothing at all.
        public MailBody Quote(string? text, string whenBlank) =>
            this.Quote([MailQuotePart.Words(text)], whenBlank);

        // ###########################################################################################
        // Somebody's own words with their pictures between them, as ONE quotation (owner report,
        // 2026-10-15: a feedback's words and its pasted screenshots were a box each - the pictures
        // belong inside the sender's text). Each run of words is trimmed and a blank one dropped;
        // `whenBlank` stands in when nothing is left. The plain text names each picture in brackets
        // - the marker CRT wrote for it - on a line of the quotation.
        // ###########################################################################################
        public MailBody Quote(IReadOnlyList<MailQuotePart> parts, string whenBlank)
        {
            ArgumentNullException.ThrowIfNull(parts);

            var kept = new List<MailQuotePart>();

            foreach (MailQuotePart part in parts)
            {
                if (part.Image is { } image)
                {
                    kept.Add(part);

                    if (this.thisImageIds.Add(image.ContentId))
                        this.thisImages.Add(image);
                }
                else if (part.Text?.Trim() is { Length: > 0 } words)
                {
                    kept.Add(MailQuotePart.Words(words));
                }
            }

            if (kept.Count == 0)
                kept.Add(MailQuotePart.Words(whenBlank));

            this.thisBlocks.Add(new Block(BlockKind.Quote, [], string.Empty, [], [.. kept]));
            return this;
        }

        // A code to paste into CRT, on its own line in a fixed-width font.
        public MailBody Code(string code)
        {
            this.thisBlocks.Add(new Block(BlockKind.Code, [], code.Trim(), []));
            return this;
        }

        // ###########################################################################################
        // Text as it was written, line for line, in a fixed-width font - a log file, a settings
        // file, a list of paths (owner request, 2026-10-03: "logfile should still show as monospace
        // as this is kind of quoted text"). Long lines wrap rather than widen the mail; nothing in
        // it is ever read as markup. The plain-text part carries it untouched.
        // ###########################################################################################
        public MailBody Preformatted(string text)
        {
            this.thisBlocks.Add(new Block(BlockKind.Preformatted, [], text.Replace("\r\n", "\n").TrimEnd(), []));
            return this;
        }

        // Numbered steps, each a run of parts.
        public MailBody Steps(params MailText[][] steps)
        {
            this.thisBlocks.Add(new Block(BlockKind.Steps, [], string.Empty, steps));
            return this;
        }

        // The mail itself: the plain text (CRLF line ends, as SMTP wants), the HTML, and the
        // pictures the HTML shows.
        public EmailMessage ToMessage(string toAddress, string subject, string? replyToAddress = null) =>
            new(toAddress, subject, this.ToText(), this.ToHtml(subject), replyToAddress,
                LinkedImages: this.thisImages.Count == 0 ? null : this.thisImages.ToList());

        public string ToText()
        {
            var text = new StringBuilder();

            foreach (Block block in this.thisBlocks)
            {
                if (text.Length > 0)
                    text.Append("\r\n\r\n");

                switch (block.Kind)
                {
                    case BlockKind.Paragraph:
                        text.Append(MailBody.RunsAsText(block.Parts));
                        break;

                    case BlockKind.Quote:
                        text.Append(string.Join("\r\n", block.Quote!
                            .SelectMany(part => part.Image is { } image ? [$"[{image.FileName}]"] : MailBody.Lines(part.Text!))
                            .Select(line => "    " + line)));
                        break;

                    case BlockKind.Code:
                        text.Append(string.Join("\r\n", MailBody.Lines(block.Text).Select(line => "    " + line)));
                        break;

                    case BlockKind.Steps:
                        text.Append(string.Join("\r\n", block.Steps.Select((step, index) => $"{index + 1}. {MailBody.RunsAsText(step)}")));
                        break;

                    case BlockKind.Preformatted:
                        text.Append(block.Text.Replace("\n", "\r\n"));
                        break;
                }
            }

            return text.ToString();
        }

        public string ToHtml(string subject)
        {
            var html = new StringBuilder();

            html.Append("<!DOCTYPE html>\r\n<html lang=\"en\">\r\n<head>\r\n")
                .Append("<meta charset=\"utf-8\">\r\n")
                .Append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\r\n")
                .Append("<title>").Append(WebUtility.HtmlEncode(subject)).Append("</title>\r\n")
                .Append("</head>\r\n")
                .Append("<body style=\"margin:0; padding:16px; background-color:#ffffff;\">\r\n")
                .Append($"<div style=\"max-width:{this.thisMaxWidth}px; font-family:{FontStack}; font-size:15px; line-height:1.5; color:#1f1f1f;\">\r\n");

            foreach (Block block in this.thisBlocks)
            {
                switch (block.Kind)
                {
                    case BlockKind.Paragraph:
                        html.Append("<p style=\"margin:0 0 14px 0;\">").Append(MailBody.RunsAsHtml(block.Parts)).Append("</p>\r\n");
                        break;

                    case BlockKind.Quote:
                        html.Append(MailBody.QuotePanelStart);

                        foreach (MailQuotePart part in block.Quote!)
                        {
                            // A picture on a line of its own, at the width given it - as an attribute
                            // too, which a client ignoring styles still reads - and never wider than
                            // the quotation; its name is the alt text, for a client that does not show it.
                            if (part.Image is { } image)
                            {
                                html.Append("<img src=\"cid:")
                                    .Append(WebUtility.HtmlEncode(image.ContentId))
                                    .Append("\" alt=\"")
                                    .Append(WebUtility.HtmlEncode(image.FileName))
                                    .Append('"');

                                if (image.Width is int width)
                                    html.Append(" width=\"").Append(width).Append('"');

                                html.Append(" style=\"display:block; ");

                                if (image.Width is int styled)
                                    html.Append("width:").Append(styled).Append("px; ");

                                html.Append("max-width:100%; height:auto; margin:8px 0; border:1px solid #dddddd;\">");
                            }
                            else
                            {
                                html.Append("<i>").Append(MailBody.LinesAsHtml(part.Text!)).Append("</i>");
                            }
                        }

                        html.Append(MailBody.QuotePanelEnd).Append("\r\n");
                        break;

                    case BlockKind.Code:
                        html.Append($"<p style=\"margin:0 0 14px 0;\"><span style=\"font-family:{MonospaceStack}; font-size:17px; letter-spacing:1px; background-color:#f0f0f0; padding:6px 10px; border-radius:4px;\">")
                            .Append(WebUtility.HtmlEncode(block.Text))
                            .Append("</span></p>\r\n");
                        break;

                    case BlockKind.Steps:
                        html.Append("<ol style=\"margin:0 0 14px 0; padding-left:22px;\">\r\n");
                        foreach (MailText[] step in block.Steps)
                            html.Append("<li style=\"margin:0 0 4px 0;\">").Append(MailBody.RunsAsHtml(step)).Append("</li>\r\n");
                        html.Append("</ol>\r\n");
                        break;

                    case BlockKind.Preformatted:
                        html.Append($"<pre style=\"margin:0 0 14px 0; padding:8px 10px; background-color:#f4f4f4; border:1px solid #dddddd; border-radius:4px; font-family:{MonospaceStack}; font-size:12px; line-height:1.4; color:#1f1f1f; white-space:pre-wrap; word-wrap:break-word;\">")
                            .Append(WebUtility.HtmlEncode(block.Text))
                            .Append("</pre>\r\n");
                        break;
                }
            }

            html.Append("</div>\r\n</body>\r\n</html>\r\n");

            return html.ToString();
        }

        private static string RunsAsText(IEnumerable<MailText> parts) =>
            string.Concat(parts.Select(part => part.Style == MailTextStyle.Named ? $"[{part.Text}]" : part.Text));

        private static string RunsAsHtml(IEnumerable<MailText> parts) =>
            string.Concat(parts.Select(part =>
            {
                string text = WebUtility.HtmlEncode(part.Text);

                return part.Style switch
                {
                    MailTextStyle.Bold => $"<b>{text}</b>",
                    MailTextStyle.Italic => $"<i>{text}</i>",
                    MailTextStyle.Named => $"[<b>{text}</b>]",
                    MailTextStyle.Link => $"<a href=\"{WebUtility.HtmlEncode(part.Href ?? string.Empty)}\" style=\"color:#1a5fb4;\">{text}</a>",
                    _ => text
                };
            }));

        private static IEnumerable<string> Lines(string text) =>
            text.Replace("\r\n", "\n").Split('\n').Select(line => line.TrimEnd());

        private static string LinesAsHtml(string text) =>
            string.Join("<br>", MailBody.Lines(text).Select(WebUtility.HtmlEncode));

        private enum BlockKind
        {
            Paragraph,
            Quote,
            Code,
            Steps,
            Preformatted
        }

        private sealed record Block(BlockKind Kind, MailText[] Parts, string Text, MailText[][] Steps, MailQuotePart[]? Quote = null);
    }

    // One piece of a quotation: a run of somebody's words, or a picture shown between them.
    public sealed record MailQuotePart(string? Text, MailLinkedImage? Image)
    {
        public static MailQuotePart Words(string? text) => new(text ?? string.Empty, null);

        public static MailQuotePart Picture(MailLinkedImage image) =>
            new(null, image ?? throw new ArgumentNullException(nameof(image)));
    }

    public enum MailTextStyle
    {
        Plain,
        Bold,
        Italic,

        // A board's name: "[Commodore/C128/250477]", bold inside the brackets in the HTML (owner
        // wording, 2026-10-03) - CRT's own way of setting off a value in running text.
        Named,

        Link
    }

    // ###########################################################################################
    // One run of text in a mail. A plain string converts to one, so a template reads as the
    // sentence it writes: Paragraph("Your contribution to ", MailText.Named(board), " was ...").
    // ###########################################################################################
    public readonly record struct MailText(string Text, MailTextStyle Style = MailTextStyle.Plain, string? Href = null)
    {
        public static implicit operator MailText(string text) => new(text ?? string.Empty);

        public static MailText Bold(string text) => new(text, MailTextStyle.Bold);

        public static MailText Italic(string text) => new(text, MailTextStyle.Italic);

        public static MailText Named(string text) => new(text, MailTextStyle.Named);

        public static MailText Link(string text, string href) => new(text, MailTextStyle.Link, href);
    }
}
