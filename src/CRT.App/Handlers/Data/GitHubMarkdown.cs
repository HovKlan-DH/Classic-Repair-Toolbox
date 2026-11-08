using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE CHANGELOG AS GITHUB SHOWS IT, AS BLOCKS AND STYLED PIECES OF TEXT (owner request,
    // 2026-11-06: "add the entire changelog and then style it alike GitHub, so it gets indented and
    // bold/italic etc."). The About tab draws what this hands back (TabAbout.Changelog.cs); this
    // decides what each line IS, and holds no control, so every rule here is tested on its own
    // (GitHubMarkdownTests).
    //
    // A SUBSET of GitHub's Markdown, read the way GitHub (CommonMark + its extensions) reads it:
    //
    //   - ATX headings ("# 3.1.0, ...") and setext ones (a line of = or - under a paragraph);
    //   - thematic breaks (a line of three or more -, * or _) - the dashes under every version;
    //   - bullet and ordered lists, nested by indentation as CommonMark nests them (an item is a
    //     child when it starts at or past its parent's text column);
    //   - paragraphs, with a soft line break read as a space and a hard one (two trailing spaces or
    //     a trailing backslash) kept;
    //   - block quotes, and GitHub's alerts ("> [!CAUTION]" on a line of its own);
    //   - fenced code blocks;
    //   - inline: **bold** and __bold__, *italic* and _italic_ by CommonMark's flanking rules (so the
    //     "_" in snake_case is no italic), ~~strike~~, `code`, [links](...), <autolinks>, bare
    //     http(s):// and www. addresses (GitHub's autolink extension), backslash escapes and HTML
    //     entities.
    //
    // Not read: raw HTML (shown as typed), tables, reference-style links, task lists, emoji
    // shortcodes, and indented code blocks (four spaces of indentation is a nested list here, not
    // code). The changelog uses none of them.
    //
    //   GitHubMarkdown.cs        - the blocks, and the model handed back (below the class)
    //   GitHubMarkdown.Inline.cs - the pieces of one line of text
    // ###########################################################################################
    internal static partial class GitHubMarkdown
    {
        public static IReadOnlyList<MarkdownBlock> Parse(string? markdown)
        {
            string text = (markdown ?? string.Empty).TrimStart('\uFEFF').Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = text.Split('\n').Select(ExpandLeadingTabs).ToList();

            return new BlockReader(lines).Read();
        }

        // The pieces of one line of text - a heading's, an item's, a paragraph's.
        public static IReadOnlyList<MarkdownSpan> ParseInline(string? text)
        {
            var spans = new List<MarkdownSpan>();
            Flatten(InlineReader.Read(text ?? string.Empty, insideLink: false), MarkdownStyle.None, null, spans);
            return spans;
        }

        // A tab in the indentation counts to the next multiple of four, as CommonMark counts it, so
        // a list nested with tabs nests the same as one nested with spaces.
        private static string ExpandLeadingTabs(string line)
        {
            if (!line.Contains('\t'))
            {
                return line;
            }

            var builder = new StringBuilder();
            int i = 0;

            for (; i < line.Length && (line[i] == ' ' || line[i] == '\t'); i++)
            {
                if (line[i] == ' ')
                {
                    builder.Append(' ');
                }
                else
                {
                    builder.Append(' ', 4 - (builder.Length % 4));
                }
            }

            return builder.Append(line, i, line.Length - i).ToString();
        }

        // ###########################################################################################
        // The block structure, line by line. One reader per level: a quote's inside is read by a
        // reader of its own, so a list or an alert inside a quote reads as it would outside one.
        // ###########################################################################################
        private sealed class BlockReader
        {
            private static readonly Regex AlertMarker = new(@"^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            private static bool IsBlank(string line) => string.IsNullOrWhiteSpace(line);

            private static int LeadingSpaces(string line)
            {
                int count = 0;
                while (count < line.Length && line[count] == ' ')
                {
                    count++;
                }

                return count;
            }

            private readonly IReadOnlyList<string> thisLines;
            private readonly List<MarkdownBlock> thisBlocks = new();

            // The list items open, outermost first; a line belongs to every item whose text column it
            // reaches.
            private readonly List<OpenItem> thisItems = new();

            // The paragraph being collected: a list item's first lines, or a paragraph of its own.
            private readonly List<string> thisParagraph = new();
            private OpenItem? thisParagraphItem;
            private int thisParagraphIndent;

            public BlockReader(IReadOnlyList<string> lines)
            {
                this.thisLines = lines;
            }

            public IReadOnlyList<MarkdownBlock> Read()
            {
                bool afterBlank = false;

                for (int i = 0; i < this.thisLines.Count; i++)
                {
                    string line = this.thisLines[i];

                    if (IsBlank(line))
                    {
                        this.FlushParagraph();
                        afterBlank = true;
                        continue;
                    }

                    int indent = LeadingSpaces(line);
                    string content = line[indent..];
                    int holding = this.ItemsHolding(indent);
                    int relative = indent - (holding == 0 ? 0 : this.thisItems[holding - 1].ContentIndent);
                    bool paragraphOpen = this.thisParagraph.Count > 0;

                    if (relative <= 3 && this.TryReadBlockStart(ref i, content, indent, holding, paragraphOpen))
                    {
                        afterBlank = false;
                        continue;
                    }

                    // A line of text: the paragraph open goes on (a "lazy" line need not be indented),
                    // otherwise a new paragraph starts inside every item it is indented under.
                    if (paragraphOpen && !afterBlank)
                    {
                        this.thisParagraph.Add(content);
                    }
                    else
                    {
                        this.CloseItemsFrom(holding);
                        this.thisParagraphItem = null;
                        this.thisParagraphIndent = holding;
                        this.thisParagraph.Add(content);
                    }

                    afterBlank = false;
                }

                this.FlushParagraph();
                return this.thisBlocks;
            }

            // A heading, a rule, a quote, a code fence or a list item starting on this line - true when
            // it was one, with i moved past whatever lines it took.
            private bool TryReadBlockStart(ref int i, string content, int indent, int holding, bool paragraphOpen)
            {
                // "Title" over "-----" is a heading - but only under a paragraph of its own, never
                // under a list item's text, where the dashes are a rule (CommonMark).
                if (paragraphOpen && this.thisParagraphItem == null && IsSetextUnderline(content, out int setextLevel))
                {
                    string title = string.Join("\n", this.thisParagraph).Trim();
                    this.thisParagraph.Clear();
                    this.thisBlocks.Add(new MarkdownHeading(setextLevel, ParseInline(title)));
                    return true;
                }

                if (IsFenceStart(content, out char fenceChar, out int fenceLength))
                {
                    this.FlushParagraph();
                    this.CloseItemsFrom(holding);
                    i = this.ReadFence(i, indent, fenceChar, fenceLength, holding);
                    return true;
                }

                if (TryReadHeading(content, out int level, out string headingText))
                {
                    this.FlushParagraph();
                    this.thisItems.Clear();
                    this.thisBlocks.Add(new MarkdownHeading(level, ParseInline(headingText)));
                    return true;
                }

                if (IsRule(content))
                {
                    this.FlushParagraph();
                    this.thisItems.Clear();
                    this.thisBlocks.Add(MarkdownRule.Instance);
                    return true;
                }

                if (content.StartsWith('>'))
                {
                    this.FlushParagraph();
                    this.CloseItemsFrom(holding);
                    i = this.ReadQuote(i, holding);
                    return true;
                }

                if (TryReadListMarker(content, out ListMarker marker))
                {
                    OpenItem? sibling = this.thisItems.Count > holding ? this.thisItems[holding] : null;
                    bool continuesList = sibling != null && sibling.Delimiter == marker.Delimiter;

                    // A new list may interrupt a paragraph only with text in its first item, and an
                    // ordered one only when it starts at 1 - "1990. was a good year" in the middle of a
                    // sentence's lines is no list (CommonMark).
                    if (paragraphOpen && !continuesList && (marker.Rest.Length == 0 || (marker.Number.HasValue && marker.Number != 1)))
                    {
                        return false;
                    }

                    this.FlushParagraph();
                    this.CloseItemsFrom(holding);

                    int? number = marker.Number.HasValue && continuesList && sibling!.Number.HasValue
                        ? sibling.Number + 1
                        : marker.Number;

                    var item = new OpenItem(indent + marker.ContentOffset, marker.Delimiter, number, holding);
                    this.thisItems.Add(item);

                    this.thisParagraphItem = item;
                    this.thisParagraph.Add(marker.Rest);
                    return true;
                }

                return false;
            }

            // How many of the open items a line indented this far belongs to - those whose text
            // column it reaches. Their columns grow outwards-in, so it is always the outermost ones.
            private int ItemsHolding(int indent)
            {
                int count = 0;

                while (count < this.thisItems.Count && indent >= this.thisItems[count].ContentIndent)
                {
                    count++;
                }

                return count;
            }

            private void CloseItemsFrom(int depth)
            {
                if (this.thisItems.Count > depth)
                {
                    this.thisItems.RemoveRange(depth, this.thisItems.Count - depth);
                }
            }

            private void FlushParagraph()
            {
                if (this.thisParagraph.Count == 0)
                {
                    return;
                }

                string text = string.Join("\n", this.thisParagraph).Trim();
                this.thisParagraph.Clear();

                if (this.thisParagraphItem != null)
                {
                    // An item is an item even with nothing written after its bullet.
                    this.thisBlocks.Add(new MarkdownListItem(this.thisParagraphItem.Depth, this.thisParagraphItem.Number, ParseInline(text)));
                    this.thisParagraphItem = null;
                    return;
                }

                if (text.Length > 0)
                {
                    this.thisBlocks.Add(new MarkdownParagraph(this.thisParagraphIndent, ParseInline(text)));
                }
            }

            // Everything to the closing fence (or the end) is the code, as typed; the fence's own
            // indentation is taken off every line, as CommonMark does.
            private int ReadFence(int start, int fenceIndent, char fenceChar, int fenceLength, int holding)
            {
                var code = new List<string>();
                int i = start + 1;

                for (; i < this.thisLines.Count; i++)
                {
                    string line = this.thisLines[i];
                    string trimmed = line.TrimStart(' ');

                    if (LeadingSpaces(line) - fenceIndent <= 3 &&
                        trimmed.Length >= fenceLength &&
                        trimmed.TrimEnd().All(c => c == fenceChar) &&
                        trimmed.TrimEnd().Length >= fenceLength)
                    {
                        break;
                    }

                    int remove = Math.Min(fenceIndent, LeadingSpaces(line));
                    code.Add(line[remove..]);
                }

                this.thisBlocks.Add(new MarkdownCode(holding, string.Join("\n", code)));
                return Math.Min(i, this.thisLines.Count - 1);
            }

            // The quote's lines with their ">" taken off, read as blocks of their own. A line without
            // ">" still belongs to it while it carries on the quote's paragraph (a "lazy" line); a
            // blank line ends it.
            private int ReadQuote(int start, int holding)
            {
                int containerIndent = holding == 0 ? 0 : this.thisItems[holding - 1].ContentIndent;
                var inner = new List<string>();
                bool lastWasText = false;
                int i = start;

                for (; i < this.thisLines.Count; i++)
                {
                    string line = this.thisLines[i];

                    if (IsBlank(line))
                    {
                        break;
                    }

                    int indent = LeadingSpaces(line);
                    string content = line[indent..];

                    if (indent <= containerIndent + 3 && content.StartsWith('>'))
                    {
                        string after = content[1..];
                        if (after.StartsWith(' '))
                        {
                            after = after[1..];
                        }

                        inner.Add(after);
                        string afterTrimmed = after.TrimStart(' ');
                        lastWasText = !IsBlank(after) && !StartsBlock(afterTrimmed);
                        continue;
                    }

                    if (lastWasText && !StartsBlock(content))
                    {
                        inner.Add(content);
                        continue;
                    }

                    break;
                }

                MarkdownAlertKind? alert = null;

                if (inner.Count > 0)
                {
                    var match = AlertMarker.Match(inner[0].Trim());
                    if (match.Success)
                    {
                        alert = Enum.Parse<MarkdownAlertKind>(match.Groups[1].Value, ignoreCase: true);
                        inner.RemoveAt(0);
                    }
                }

                this.thisBlocks.Add(new MarkdownQuote(alert, holding, new BlockReader(inner).Read()));
                return i - 1;
            }

            // Whether a line starts something other than paragraph text - what stops a "lazy" line
            // from carrying a quote's paragraph on.
            private static bool StartsBlock(string content) =>
                content.StartsWith('>') ||
                IsRule(content) ||
                IsFenceStart(content, out _, out _) ||
                TryReadHeading(content, out _, out _) ||
                (TryReadListMarker(content, out ListMarker marker) && marker.Rest.Length > 0);

            private static bool IsSetextUnderline(string content, out int level)
            {
                string trimmed = content.TrimEnd();
                level = 0;

                if (trimmed.Length == 0 || (trimmed[0] != '=' && trimmed[0] != '-'))
                {
                    return false;
                }

                char c = trimmed[0];
                if (!trimmed.All(ch => ch == c))
                {
                    return false;
                }

                level = c == '=' ? 1 : 2;
                return true;
            }

            private static bool IsFenceStart(string content, out char fenceChar, out int length)
            {
                fenceChar = content.Length > 0 ? content[0] : '\0';
                length = 0;

                if (fenceChar != '`' && fenceChar != '~')
                {
                    return false;
                }

                while (length < content.Length && content[length] == fenceChar)
                {
                    length++;
                }

                // A backtick fence's info string may hold no backtick - "```a``" is inline code.
                return length >= 3 && (fenceChar == '~' || !content[length..].Contains('`'));
            }

            private static bool TryReadHeading(string content, out int level, out string text)
            {
                level = 0;
                text = string.Empty;

                while (level < content.Length && content[level] == '#')
                {
                    level++;
                }

                if (level == 0 || level > 6 || (level < content.Length && content[level] != ' ' && content[level] != '\t'))
                {
                    return false;
                }

                string rest = content[level..].Trim();

                // An optional closing run of #s, set off by a space: "# Title #" is "Title", while
                // "# C#" keeps its "#".
                if (rest.All(c => c == '#'))
                {
                    rest = string.Empty;
                }
                else
                {
                    int end = rest.Length;
                    while (end > 0 && rest[end - 1] == '#')
                    {
                        end--;
                    }

                    if (end < rest.Length && end > 0 && (rest[end - 1] == ' ' || rest[end - 1] == '\t'))
                    {
                        rest = rest[..end].TrimEnd();
                    }
                }

                text = rest;
                return true;
            }

            private static bool IsRule(string content)
            {
                char kind = '\0';
                int count = 0;

                foreach (char c in content)
                {
                    if (c == ' ' || c == '\t')
                    {
                        continue;
                    }

                    if (c != '-' && c != '*' && c != '_')
                    {
                        return false;
                    }

                    if (kind == '\0')
                    {
                        kind = c;
                    }
                    else if (c != kind)
                    {
                        return false;
                    }

                    count++;
                }

                return count >= 3;
            }

            private static bool TryReadListMarker(string content, out ListMarker marker)
            {
                marker = default;
                int width;
                int? number = null;
                char delimiter;

                if (content.Length > 0 && (content[0] == '-' || content[0] == '+' || content[0] == '*'))
                {
                    width = 1;
                    delimiter = content[0];
                }
                else
                {
                    int digits = 0;
                    while (digits < content.Length && digits < 10 && char.IsAsciiDigit(content[digits]))
                    {
                        digits++;
                    }

                    if (digits == 0 || digits > 9 || digits >= content.Length || (content[digits] != '.' && content[digits] != ')'))
                    {
                        return false;
                    }

                    number = int.Parse(content[..digits], System.Globalization.CultureInfo.InvariantCulture);
                    delimiter = content[digits];
                    width = digits + 1;
                }

                if (width < content.Length && content[width] != ' ')
                {
                    return false;
                }

                int spaces = 0;
                while (width + spaces < content.Length && content[width + spaces] == ' ')
                {
                    spaces++;
                }

                string rest = content[(width + spaces)..];

                // With nothing after the marker, or more than four spaces (code, in CommonMark), the
                // item's text starts one space past it.
                if (rest.Length == 0 || spaces > 4)
                {
                    rest = content[Math.Min(content.Length, width + 1)..].TrimStart(' ');
                    spaces = 1;
                }

                marker = new ListMarker(delimiter, number, width + spaces, rest);
                return true;
            }
        }

        private readonly record struct ListMarker(char Delimiter, int? Number, int ContentOffset, string Rest);

        private sealed record OpenItem(int ContentIndent, char Delimiter, int? Number, int Depth);
    }

    // ###########################################################################################
    // What GitHubMarkdown hands back.
    // ###########################################################################################

    [Flags]
    internal enum MarkdownStyle
    {
        None = 0,
        Bold = 1,
        Italic = 2,
        Code = 4,
        Strikethrough = 8,
    }

    // A piece of text in one look. Url is the link's target as written, null when it is no link;
    // a hard line break is a piece of its own, "\n".
    internal sealed record MarkdownSpan(string Text, MarkdownStyle Style, string? Url);

    // GitHub's five alerts, in its own order.
    internal enum MarkdownAlertKind
    {
        Note,
        Tip,
        Important,
        Warning,
        Caution,
    }

    // Indent is how many list items a block sits inside - how far in it is drawn. 0 for a block of
    // its own; a list item's is its depth.
    internal abstract record MarkdownBlock(int Indent);

    // A heading is bold as a whole, as GitHub draws it - or, BoldAsMarked, plain with only its bold pieces
    // bold: a changelog version's, "Version **3.0.0**, 2026-October-11" (ChangelogVersionHeadings).
    internal sealed record MarkdownHeading(int Level, IReadOnlyList<MarkdownSpan> Spans, bool BoldAsMarked = false) : MarkdownBlock(0);

    // A thematic break - the line of dashes under every version. Every one is the same, so there is
    // one.
    internal sealed record MarkdownRule : MarkdownBlock
    {
        public static MarkdownRule Instance { get; } = new();

        private MarkdownRule()
            : base(0)
        {
        }
    }

    // Depth 0 is a top-level item. Number is an ordered item's number, null for a bullet.
    internal sealed record MarkdownListItem(int Depth, int? Number, IReadOnlyList<MarkdownSpan> Spans) : MarkdownBlock(Depth);

    internal sealed record MarkdownParagraph(int Indent, IReadOnlyList<MarkdownSpan> Spans) : MarkdownBlock(Indent);

    // A block quote, or one of GitHub's alerts when Alert is set; what it holds is blocks of its own.
    internal sealed record MarkdownQuote(MarkdownAlertKind? Alert, int Indent, IReadOnlyList<MarkdownBlock> Blocks) : MarkdownBlock(Indent);

    internal sealed record MarkdownCode(int Indent, string Text) : MarkdownBlock(Indent);
}
