using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // GitHubMarkdown.Inline.cs - THE PIECES OF ONE LINE OF TEXT: bold, italic, struck through, code,
    // links and line breaks, read the way CommonMark and GitHub read them. GitHubMarkdown.cs holds
    // the blocks and what is handed back.
    // ###########################################################################################
    internal static partial class GitHubMarkdown
    {
        // ###########################################################################################
        // The inline pieces. Read in one pass into text, code, links and delimiter runs, then the
        // runs are paired into emphasis by CommonMark's "process emphasis" procedure - which is what
        // makes "**Application**" bold, "__experimental__" bold and the "_" in snake_case nothing.
        // ###########################################################################################
        private enum NodeKind
        {
            Text,
            Delimiter,
            Code,
            HardBreak,
            Link,
            Emphasis,
            Strong,
            Strike,
        }

        private sealed class Node
        {
            public NodeKind Kind;
            public string Text = string.Empty;
            public char DelimiterChar;
            public int OriginalLength;
            public bool CanOpen;
            public bool CanClose;
            public int Order;
            public string? Url;
            public List<Node>? Children;
        }

        private static void Flatten(List<Node> nodes, MarkdownStyle style, string? url, List<MarkdownSpan> spans)
        {
            foreach (var node in nodes)
            {
                switch (node.Kind)
                {
                    case NodeKind.Text:
                    case NodeKind.Delimiter:
                        AddSpan(spans, node.Text, style, url);
                        break;
                    case NodeKind.Code:
                        AddSpan(spans, node.Text, style | MarkdownStyle.Code, url);
                        break;
                    case NodeKind.HardBreak:
                        spans.Add(new MarkdownSpan("\n", style, url));
                        break;
                    case NodeKind.Link:
                        Flatten(node.Children!, style, url ?? node.Url, spans);
                        break;
                    case NodeKind.Emphasis:
                        Flatten(node.Children!, style | MarkdownStyle.Italic, url, spans);
                        break;
                    case NodeKind.Strong:
                        Flatten(node.Children!, style | MarkdownStyle.Bold, url, spans);
                        break;
                    case NodeKind.Strike:
                        Flatten(node.Children!, style | MarkdownStyle.Strikethrough, url, spans);
                        break;
                }
            }
        }

        // Pieces of the same look and link run together as one; two code spans side by side stay
        // two, as they were written.
        private static void AddSpan(List<MarkdownSpan> spans, string text, MarkdownStyle style, string? url)
        {
            if (text.Length == 0)
            {
                return;
            }

            if (spans.Count > 0 && (style & MarkdownStyle.Code) == 0)
            {
                var last = spans[^1];
                if (last.Style == style && last.Url == url && last.Text != "\n")
                {
                    spans[^1] = last with { Text = last.Text + text };
                    return;
                }
            }

            spans.Add(new MarkdownSpan(text, style, url));
        }

        private sealed class InlineReader
        {
            private static readonly Regex Entity = new(@"\G&(?:#[0-9]{1,7}|#[xX][0-9a-fA-F]{1,6}|[A-Za-z][A-Za-z0-9]{1,31});", RegexOptions.CultureInvariant);
            private static readonly Regex Autolink = new(@"\G<([A-Za-z][A-Za-z0-9+.\-]{1,31}:[^\s<>]*)>", RegexOptions.CultureInvariant);
            private static readonly Regex EmailAutolink = new(@"\G<([A-Za-z0-9.!#$%&'*+/=?^_`{|}~\-]+@[A-Za-z0-9](?:[A-Za-z0-9\-]{0,61}[A-Za-z0-9])?(?:\.[A-Za-z0-9](?:[A-Za-z0-9\-]{0,61}[A-Za-z0-9])?)*)>", RegexOptions.CultureInvariant);

            private readonly string thisText;
            private readonly bool thisInsideLink;
            private readonly List<Node> thisNodes = new();
            private readonly List<Node> thisDelimiters = new();
            private readonly StringBuilder thisBuffer = new();

            private InlineReader(string text, bool insideLink)
            {
                this.thisText = text;
                this.thisInsideLink = insideLink;
            }

            public static List<Node> Read(string text, bool insideLink)
            {
                var reader = new InlineReader(text, insideLink);
                reader.Scan();
                ProcessEmphasis(reader.thisNodes, reader.thisDelimiters);
                return reader.thisNodes;
            }

            private void Scan()
            {
                string text = this.thisText;
                int i = 0;

                while (i < text.Length)
                {
                    char c = text[i];

                    switch (c)
                    {
                        case '\\':
                            i = this.ReadBackslash(i);
                            break;
                        case '`':
                            i = this.ReadCode(i);
                            break;
                        case '\n':
                            i = this.ReadLineEnd(i);
                            break;
                        case '*':
                        case '_':
                        case '~':
                            i = this.ReadDelimiterRun(i);
                            break;
                        case '&':
                            i = this.ReadEntity(i);
                            break;
                        case '<':
                            i = this.ReadAngleAutolink(i);
                            break;
                        case '[':
                            i = this.ReadLink(i, i);
                            break;
                        case '!' when i + 1 < text.Length && text[i + 1] == '[':
                            i = this.ReadLink(i, i + 1);
                            break;
                        default:
                            i = this.ReadBareLinkOrChar(i);
                            break;
                    }
                }

                this.FlushText();
            }

            private void FlushText()
            {
                if (this.thisBuffer.Length > 0)
                {
                    this.thisNodes.Add(new Node { Kind = NodeKind.Text, Text = this.thisBuffer.ToString() });
                    this.thisBuffer.Clear();
                }
            }

            private void AddNode(Node node)
            {
                this.FlushText();
                this.thisNodes.Add(node);
            }

            // "\*" is a star, "\" at a line's end a hard break, and any other backslash just itself.
            private int ReadBackslash(int i)
            {
                string text = this.thisText;

                if (i + 1 < text.Length && text[i + 1] == '\n')
                {
                    this.AddNode(new Node { Kind = NodeKind.HardBreak });
                    return SkipSpaces(text, i + 2);
                }

                if (i + 1 < text.Length && IsAsciiPunctuation(text[i + 1]))
                {
                    this.thisBuffer.Append(text[i + 1]);
                    return i + 2;
                }

                this.thisBuffer.Append('\\');
                return i + 1;
            }

            // A run of backticks opens a code span closed by the next run of the SAME length; with
            // none, the backticks are just text.
            private int ReadCode(int i)
            {
                string text = this.thisText;
                int run = RunLength(text, i, '`');
                int close = FindBacktickRun(text, i + run, run);

                if (close < 0)
                {
                    this.thisBuffer.Append('`', run);
                    return i + run;
                }

                this.AddNode(new Node { Kind = NodeKind.Code, Text = CodeContent(text[(i + run)..close]) });
                return close + run;
            }

            // A line's end inside a paragraph is a space - or a hard break after two or more spaces.
            private int ReadLineEnd(int i)
            {
                int spaces = 0;
                while (spaces < this.thisBuffer.Length && this.thisBuffer[^(spaces + 1)] == ' ')
                {
                    spaces++;
                }

                this.thisBuffer.Length -= spaces;

                if (spaces >= 2)
                {
                    this.AddNode(new Node { Kind = NodeKind.HardBreak });
                }
                else
                {
                    this.thisBuffer.Append(' ');
                }

                return SkipSpaces(this.thisText, i + 1);
            }

            // A run of *, _ or ~ that may open or close emphasis, by CommonMark's flanking rules; which
            // ones pair up is decided once the whole text is read (ProcessEmphasis).
            private int ReadDelimiterRun(int i)
            {
                string text = this.thisText;
                char c = text[i];
                int run = RunLength(text, i, c);

                // GitHub strikes through with one or two tildes; three or more are just tildes.
                if (c == '~' && run > 2)
                {
                    this.thisBuffer.Append(c, run);
                    return i + run;
                }

                char before = i == 0 ? '\n' : text[i - 1];
                char after = i + run >= text.Length ? '\n' : text[i + run];

                bool leftFlanking = !char.IsWhiteSpace(after) && (!IsPunctuation(after) || char.IsWhiteSpace(before) || IsPunctuation(before));
                bool rightFlanking = !char.IsWhiteSpace(before) && (!IsPunctuation(before) || char.IsWhiteSpace(after) || IsPunctuation(after));

                bool canOpen;
                bool canClose;

                if (c == '_')
                {
                    // Inside a word an underscore is part of it: snake_case stays as typed.
                    canOpen = leftFlanking && (!rightFlanking || IsPunctuation(before));
                    canClose = rightFlanking && (!leftFlanking || IsPunctuation(after));
                }
                else
                {
                    canOpen = leftFlanking;
                    canClose = rightFlanking;
                }

                var node = new Node
                {
                    Kind = NodeKind.Delimiter,
                    Text = new string(c, run),
                    DelimiterChar = c,
                    OriginalLength = run,
                    CanOpen = canOpen,
                    CanClose = canClose,
                    Order = i,
                };

                this.AddNode(node);

                if (canOpen || canClose)
                {
                    this.thisDelimiters.Add(node);
                }

                return i + run;
            }

            private int ReadEntity(int i)
            {
                var match = Entity.Match(this.thisText, i);

                if (match.Success)
                {
                    string decoded = WebUtility.HtmlDecode(match.Value);
                    if (decoded != match.Value)
                    {
                        this.thisBuffer.Append(decoded);
                        return i + match.Length;
                    }
                }

                this.thisBuffer.Append('&');
                return i + 1;
            }

            private int ReadAngleAutolink(int i)
            {
                var match = Autolink.Match(this.thisText, i);
                string? url = null;
                string shown = string.Empty;

                if (match.Success)
                {
                    shown = match.Groups[1].Value;
                    url = shown;
                }
                else
                {
                    match = EmailAutolink.Match(this.thisText, i);
                    if (match.Success)
                    {
                        shown = match.Groups[1].Value;
                        url = "mailto:" + shown;
                    }
                }

                if (url == null || this.thisInsideLink)
                {
                    this.thisBuffer.Append('<');
                    return i + 1;
                }

                this.AddNode(new Node
                {
                    Kind = NodeKind.Link,
                    Url = url,
                    Children = new List<Node> { new() { Kind = NodeKind.Text, Text = shown } },
                });

                return i + match.Length;
            }

            // "[text](url)", or "![alt](url)" for an image - shown as a link to the picture, by its
            // alternative text. Anything that does not complete one is just text.
            private int ReadLink(int start, int bracket)
            {
                string text = this.thisText;

                if (this.thisInsideLink || !TryReadLink(text, bracket, out int end, out string label, out string url))
                {
                    this.thisBuffer.Append(text[start]);
                    return start + 1;
                }

                var children = Read(label, insideLink: true);
                if (children.Count == 0)
                {
                    children.Add(new Node { Kind = NodeKind.Text, Text = url });
                }

                this.AddNode(new Node { Kind = NodeKind.Link, Url = url, Children = children });
                return end;
            }

            // A bare web address is a link too (GitHub's autolink extension), when it starts a word.
            private int ReadBareLinkOrChar(int i)
            {
                string text = this.thisText;
                char c = text[i];

                if (!this.thisInsideLink &&
                    (c == 'h' || c == 'H' || c == 'w' || c == 'W') &&
                    (i == 0 || char.IsWhiteSpace(text[i - 1]) || "*_~(".Contains(text[i - 1])) &&
                    TryReadBareLink(text, i, out int length, out string? url))
                {
                    this.AddNode(new Node
                    {
                        Kind = NodeKind.Link,
                        Url = url,
                        Children = new List<Node> { new() { Kind = NodeKind.Text, Text = text.Substring(i, length) } },
                    });

                    return i + length;
                }

                this.thisBuffer.Append(c);
                return i + 1;
            }

            // The address runs to the first space or "<", less the punctuation that ends a sentence
            // rather than the address (GitHub's list: ?!.,:*_~ and quotes) - TextLinkFinder then
            // decides the prefix and an unbalanced closing bracket, as it does for the worklog's text.
            private static bool TryReadBareLink(string text, int i, out int length, out string? url)
            {
                length = 0;
                url = null;

                int end = i;
                while (end < text.Length && !char.IsWhiteSpace(text[end]) && text[end] != '<')
                {
                    end++;
                }

                while (end > i && "?!.,:;*_~'\"".Contains(text[end - 1]))
                {
                    end--;
                }

                var spans = TextLinkFinder.FindSpans(text[i..end]);
                if (spans.Count == 0 || !spans[0].IsLink || spans[0].Start != 0)
                {
                    return false;
                }

                length = spans[0].Length;
                url = spans[0].Url;
                return true;
            }

            private static bool TryReadLink(string text, int bracket, out int end, out string label, out string url)
            {
                end = 0;
                label = string.Empty;
                url = string.Empty;

                int close = FindClosingBracket(text, bracket);
                if (close < 0 || close + 1 >= text.Length || text[close + 1] != '(')
                {
                    return false;
                }

                int k = SkipWhitespace(text, close + 2);
                var destination = new StringBuilder();

                if (k < text.Length && text[k] == '<')
                {
                    int gt = k + 1;
                    while (gt < text.Length && text[gt] != '>' && text[gt] != '<' && text[gt] != '\n')
                    {
                        gt++;
                    }

                    if (gt >= text.Length || text[gt] != '>')
                    {
                        return false;
                    }

                    destination.Append(text, k + 1, gt - k - 1);
                    k = gt + 1;
                }
                else
                {
                    int depth = 0;

                    while (k < text.Length && !char.IsWhiteSpace(text[k]) && !char.IsControl(text[k]))
                    {
                        char c = text[k];

                        if (c == '\\' && k + 1 < text.Length && IsAsciiPunctuation(text[k + 1]))
                        {
                            destination.Append(text[k + 1]);
                            k += 2;
                            continue;
                        }

                        if (c == '(')
                        {
                            depth++;
                        }
                        else if (c == ')')
                        {
                            if (depth == 0)
                            {
                                break;
                            }

                            depth--;
                        }

                        destination.Append(c);
                        k++;
                    }
                }

                int beforeTitle = k;
                k = SkipWhitespace(text, k);

                // An optional title, which GitHub shows only as a tooltip - read past and dropped.
                if (k < text.Length && k > beforeTitle && (text[k] == '"' || text[k] == '\'' || text[k] == '('))
                {
                    char closer = text[k] == '(' ? ')' : text[k];
                    int t = k + 1;
                    while (t < text.Length && text[t] != closer)
                    {
                        t += text[t] == '\\' ? 2 : 1;
                    }

                    if (t >= text.Length)
                    {
                        return false;
                    }

                    k = SkipWhitespace(text, t + 1);
                }

                if (k >= text.Length || text[k] != ')')
                {
                    return false;
                }

                end = k + 1;
                label = text[(bracket + 1)..close];
                url = destination.ToString();
                return true;
            }

            // The "]" that closes the "[" at bracket - brackets nest, and neither an escaped one nor
            // one inside a code span counts.
            private static int FindClosingBracket(string text, int bracket)
            {
                int depth = 0;
                int j = bracket + 1;

                while (j < text.Length)
                {
                    char c = text[j];

                    if (c == '\\' && j + 1 < text.Length)
                    {
                        j += 2;
                        continue;
                    }

                    if (c == '`')
                    {
                        int run = RunLength(text, j, '`');
                        int close = FindBacktickRun(text, j + run, run);
                        j = close < 0 ? j + run : close + run;
                        continue;
                    }

                    if (c == '[')
                    {
                        depth++;
                    }
                    else if (c == ']')
                    {
                        if (depth == 0)
                        {
                            return j;
                        }

                        depth--;
                    }

                    j++;
                }

                return -1;
            }

            // ###########################################################################################
            // CommonMark's "process emphasis": each closer looks back for the nearest opener of its
            // kind, and what lies between them becomes emphasis (one character each side), strong
            // (two), or struck through (tildes). What pairs with nothing stays as typed.
            //
            // "Below" is kept by position in the text (Order), so a delimiter taken out of the list
            // still compares correctly against those left.
            // ###########################################################################################
            private static void ProcessEmphasis(List<Node> nodes, List<Node> delimiters)
            {
                var bottoms = new Dictionary<(char, int, bool), int>();
                int ci = 0;

                while (ci < delimiters.Count)
                {
                    Node closer = delimiters[ci];
                    if (!closer.CanClose)
                    {
                        ci++;
                        continue;
                    }

                    var key = (closer.DelimiterChar, closer.OriginalLength % 3, closer.CanOpen);
                    int bottom = bottoms.TryGetValue(key, out int found) ? found : -1;
                    int oi = -1;

                    for (int j = ci - 1; j >= 0 && delimiters[j].Order > bottom; j--)
                    {
                        if (IsOpenerFor(delimiters[j], closer))
                        {
                            oi = j;
                            break;
                        }
                    }

                    if (oi < 0)
                    {
                        bottoms[key] = closer.Order - 1;

                        if (closer.CanOpen)
                        {
                            ci++;
                        }
                        else
                        {
                            delimiters.RemoveAt(ci);
                        }

                        continue;
                    }

                    Node opener = delimiters[oi];
                    int use = closer.DelimiterChar == '~'
                        ? closer.Text.Length
                        : opener.Text.Length >= 2 && closer.Text.Length >= 2 ? 2 : 1;

                    NodeKind kind = closer.DelimiterChar == '~'
                        ? NodeKind.Strike
                        : use == 2 ? NodeKind.Strong : NodeKind.Emphasis;

                    int openerAt = nodes.IndexOf(opener);
                    int closerAt = nodes.IndexOf(closer);
                    var inner = nodes.GetRange(openerAt + 1, closerAt - openerAt - 1);
                    nodes.RemoveRange(openerAt + 1, closerAt - openerAt - 1);
                    nodes.Insert(openerAt + 1, new Node { Kind = kind, Children = inner });

                    delimiters.RemoveRange(oi + 1, ci - oi - 1);
                    ci = oi + 1;

                    opener.Text = opener.Text[..^use];
                    closer.Text = closer.Text[..^use];

                    if (opener.Text.Length == 0)
                    {
                        nodes.Remove(opener);
                        delimiters.RemoveAt(oi);
                        ci--;
                    }

                    if (closer.Text.Length == 0)
                    {
                        nodes.Remove(closer);
                        delimiters.RemoveAt(ci);
                    }
                }
            }

            private static bool IsOpenerFor(Node candidate, Node closer)
            {
                if (candidate.DelimiterChar != closer.DelimiterChar || !candidate.CanOpen)
                {
                    return false;
                }

                if (closer.DelimiterChar == '~')
                {
                    return candidate.Text.Length == closer.Text.Length;
                }

                // The "rule of three": "*foo**bar*" is one emphasis, not a strong one half-matched.
                bool eitherBoth = candidate.CanClose || closer.CanOpen;
                int sum = candidate.OriginalLength + closer.OriginalLength;

                return !(eitherBoth && sum % 3 == 0 && !(candidate.OriginalLength % 3 == 0 && closer.OriginalLength % 3 == 0));
            }

            private static string CodeContent(string raw)
            {
                string content = raw.Replace('\n', ' ');

                if (content.Length >= 2 && content[0] == ' ' && content[^1] == ' ' && content.Any(c => c != ' '))
                {
                    content = content[1..^1];
                }

                return content;
            }

            private static int FindBacktickRun(string text, int from, int length)
            {
                int j = from;

                while (j < text.Length)
                {
                    if (text[j] != '`')
                    {
                        j++;
                        continue;
                    }

                    int run = RunLength(text, j, '`');
                    if (run == length)
                    {
                        return j;
                    }

                    j += run;
                }

                return -1;
            }

            private static int RunLength(string text, int i, char c)
            {
                int run = 0;
                while (i + run < text.Length && text[i + run] == c)
                {
                    run++;
                }

                return run;
            }

            private static int SkipSpaces(string text, int i)
            {
                while (i < text.Length && text[i] == ' ')
                {
                    i++;
                }

                return i;
            }

            private static int SkipWhitespace(string text, int i)
            {
                while (i < text.Length && (text[i] == ' ' || text[i] == '\t' || text[i] == '\n'))
                {
                    i++;
                }

                return i;
            }

            private static bool IsAsciiPunctuation(char c) =>
                c < 128 && (char.IsPunctuation(c) || char.IsSymbol(c));

            private static bool IsPunctuation(char c) =>
                char.IsPunctuation(c) || char.IsSymbol(c);
        }
    }
}
