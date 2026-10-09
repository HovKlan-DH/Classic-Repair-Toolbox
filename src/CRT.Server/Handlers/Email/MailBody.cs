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
    // attribute. No images, no remote resources - a mail that loads nothing cannot be tracked and
    // is not held back by a client that blocks remote content.
    // ###########################################################################################
    public sealed class MailBody
    {
        private const string FontStack = "'Segoe UI', Helvetica, Arial, sans-serif";

        private const string MonospaceStack = "Consolas, Menlo, 'Courier New', monospace";

        // A mail's width, as a reading column. The feedback mail is wider: it carries log files.
        public const int StandardWidth = 640;
        public const int WideWidth = 1000;

        private readonly List<Block> thisBlocks = [];

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
        public MailBody Quote(string? text, string whenBlank)
        {
            string trimmed = text?.Trim() ?? string.Empty;
            this.thisBlocks.Add(new Block(BlockKind.Quote, [], trimmed.Length == 0 ? whenBlank : trimmed, []));
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

        // The mail itself: the plain text (CRLF line ends, as SMTP wants) and the HTML.
        public EmailMessage ToMessage(string toAddress, string subject, string? replyToAddress = null) =>
            new(toAddress, subject, this.ToText(), this.ToHtml(subject), replyToAddress);

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
                        html.Append("<blockquote style=\"margin:0 0 14px 0; padding:6px 0 6px 14px; border-left:3px solid #c8c8c8; color:#444444;\"><i>")
                            .Append(MailBody.LinesAsHtml(block.Text))
                            .Append("</i></blockquote>\r\n");
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

        private sealed record Block(BlockKind Kind, MailText[] Parts, string Text, MailText[][] Steps);
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
