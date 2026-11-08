using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A NOTE WITH A LITTLE FORMATTING (owner request, 2026-11-07: "For both the device, workbook and worklog - for all
    // textareas in there, I would like the possibility to have a very simple WYSIWYG ... Bold, Italic, Underscore,
    // Strikethrough, Bulletlist (nested). I do not want a full fledged editor ... nor should it feel 'heavy' and slow",
    // and the same for the "Note" of a component's pictures and oscilloscope baselines).
    //
    // *** STORED AS TEXT A PERSON CAN STILL READ *** - the few marks of Markdown, so a note shown somewhere that does not
    // draw them (the Maintainer tab's table, an older CRT, a file opened by hand) reads as it was meant:
    //   **bold**   *italic*   __underlined__   ~~struck through~~   "- " starting a line for a bullet, two spaces more
    //   in front for each level it is nested, and "\" before a character that would otherwise be read as a mark, to
    //   have it as itself (IsEscapeAt).
    // A note typed before this - every note there is - reads as it always did: a mark only counts in pairs on one line,
    // next to the words it marks ("5 * 2" is no italic), a "\" only where it changes the reading ("\\NAS\c64" and
    // "D:\_old" keep every backslash), and NOTHING INSIDE A WEB ADDRESS IS A MARK (LinkParts - "...#object.__init__"
    // stays the address it is). A line a note already starts with "- " shows as a bullet, as it was meant - after no
    // indent, or one a bullet can have: two spaces a level, MaxLevel levels deep at most (IsBulletIndent). A "\" before
    // that "- " says the line is no bullet - the first of a run of them, the rest kept (BackslashesBeforeDash).
    //
    // *** WHAT A NOTE IS WHILE SHOWN OR EDITED: its plain text and each character's style (RichNoteText). *** A bullet
    // line's text starts with its indent and its glyph - "• ", "  ◦ ", "    ▪ " - so it is drawn and edited as the
    // text it is; written, it is "- " again (Serialize). Written with the fewest marks: a "\" only where a character
    // would otherwise be taken for a mark, worked out by reading the result back (Serialize).
    // ###########################################################################################
    [Flags]
    public enum RichNoteStyle
    {
        None = 0,
        Bold = 1,
        Italic = 2,
        Underline = 4,
        Strike = 8,
    }

    // ###########################################################################################
    // A board row's property holding a note with formatting - so it is compared as a note is trimmed (RichNote.TrimMarkup,
    // never Trim(): a first line's indent says whether it is a bullet and how deep). BoardDataDiffer reads it (owner
    // decision, 2026-11-08: an indent-only change of a note "is a change like any other change so it should count").
    // ###########################################################################################
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class RichNoteMarkupAttribute : Attribute
    {
    }

    // A note as shown: its plain text, and the style of each of its characters - the same length.
    public sealed class RichNoteText
    {
        public RichNoteText(string text, IReadOnlyList<RichNoteStyle> styles)
        {
            this.Text = text ?? string.Empty;
            this.Styles = styles ?? throw new ArgumentNullException(nameof(styles));

            if (this.Styles.Count != this.Text.Length)
            {
                throw new ArgumentException($"[{this.Styles.Count}] styles for [{this.Text.Length}] characters", nameof(styles));
            }
        }

        public string Text { get; }

        public IReadOnlyList<RichNoteStyle> Styles { get; }

        // Whether anything in it is styled or a bullet - a note that is none of them is plain text, drawn as such.
        public bool IsFormatted =>
            this.Styles.Any(style => style != RichNoteStyle.None) ||
            this.Text.Split('\n').Any(line => RichNote.TryReadBullet(line, out _, out _));

        public bool SameAs(RichNoteText other) =>
            other is not null && string.Equals(this.Text, other.Text, StringComparison.Ordinal) && this.Styles.SequenceEqual(other.Styles);
    }

    public static class RichNote
    {
        // The bullet glyphs by level - the third for every deeper level.
        public const string BulletGlyphs = "•◦▪";

        // How deep a bullet can be nested.
        public const int MaxLevel = 5;

        // The characters marks are made of.
        private const string MarkCharacters = "*_~";

        public static char GlyphOf(int level) => BulletGlyphs[Math.Min(Math.Max(level, 0), BulletGlyphs.Length - 1)];

        // A bullet line's start as shown, "  ◦ " for level 1.
        public static string BulletPrefix(int level)
        {
            int within = Math.Clamp(level, 0, MaxLevel);
            return new string(' ', within * 2) + GlyphOf(within) + " ";
        }

        // ###########################################################################################
        // Whether a line AS SHOWN is a bullet: an even indent, a bullet glyph, a space. Its level, and how many
        // characters its start takes.
        // ###########################################################################################
        public static bool TryReadBullet(string? line, out int level, out int prefixLength)
        {
            level = 0;
            prefixLength = 0;

            if (string.IsNullOrEmpty(line))
            {
                return false;
            }

            int indent = IndentOf(line);

            if (!IsBulletIndent(indent) || line.Length < indent + 2 || BulletGlyphs.IndexOf(line[indent]) < 0 || line[indent + 1] != ' ')
            {
                return false;
            }

            level = indent / 2;
            prefixLength = indent + 2;
            return true;
        }

        // How many spaces a line starts with - its indent, the one count every rule of a line's start reads.
        private static int IndentOf(string line)
        {
            int indent = 0;
            while (indent < line.Length && line[indent] == ' ')
            {
                indent++;
            }

            return indent;
        }

        // The plain text of a note - what a search looks in, and what is shown where no style can be drawn.
        public static string PlainOf(string? markup) => Parse(markup).Text;

        // ###########################################################################################
        // *** A NOTE AS STORED, ITS EDGES TRIMMED WITHOUT CHANGING HOW IT READS (code review, 2026-11-08). *** Trimmed whole,
        // a note's first line lost its indent, and with it what it is: "  - recapped" (a bullet nested once) came back a
        // bullet of the first level, "   - x" (no bullet at that indent) became one. So: blank lines before the first words
        // go, white space at the end goes as far as the last line still reads the same (TrimEndKeepingReading), and the first
        // line's own leading white space goes only when no dash, "\" or bullet glyph follows it - where it is nothing but
        // space (a bullet glyph pasted in reads as a bullet too, TryReadBullet: "  • x" a level deep). Every note's save goes
        // through this, never Trim().
        // ###########################################################################################
        public static string TrimMarkup(string? markup)
        {
            string text = TrimEndKeepingReading(markup ?? string.Empty);
            int start = 0;

            for (int end = text.IndexOf('\n'); end >= 0 && string.IsNullOrWhiteSpace(text[start..end]); end = text.IndexOf('\n', start))
            {
                start = end + 1;
            }

            int words = start;

            while (words < text.Length && text[words] != '\n' && char.IsWhiteSpace(text[words]))
            {
                words++;
            }

            bool indentReads = words < text.Length && (text[words] is '-' or '\\' || BulletGlyphs.IndexOf(text[words]) >= 0);

            return indentReads ? text[start..] : text[words..];
        }

        // ###########################################################################################
        // A note's markup as it is kept and compared: its line ends "\n", whatever the clipboard brought, and its edges
        // trimmed as TrimMarkup trims them. A picture's or a baseline's note (ComponentPaneEdits), and every [RichNoteMarkup]
        // value BoardDataDiffer and the table compare.
        // ###########################################################################################
        public static string NormalisedMarkup(string? markup) =>
            TrimMarkup((markup ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n'));

        // ###########################################################################################
        // *** WHITE SPACE AT THE END GOES, WITHOUT CHANGING HOW THE LAST LINE READS (code review, 2026-11-08). *** Trimmed
        // whole, a last line "- " - a bullet with nothing typed after it, as Enter on a bullet leaves one - lost its space
        // and was stored as a line "-", shown as a dash; one written "\- " - a typed dash, no bullet - kept a "\" it no longer
        // needed, and showed it. So an empty bullet at the end goes whole, as the blank it is, and a typed dash keeps only
        // the "\"s that were its own (BackslashesBeforeDash: the first says "no bullet", and a dash with no space after it is
        // none anyway).
        // ###########################################################################################
        private static string TrimEndKeepingReading(string text)
        {
            while (true)
            {
                int end = text.TrimEnd().Length;

                // Nothing trimmed, or no space right after the last word: no "- " lost its space.
                if (end == 0 || end == text.Length || text[end] != ' ')
                {
                    return text[..end];
                }

                int lineStart = text.LastIndexOf('\n', end - 1) + 1;
                string line = text[lineStart..end];
                int indent = IndentOf(line);
                int backslashes = BackslashesAt(line, indent);

                if (!IsBulletIndent(indent) || line.Length != indent + backslashes + 1 || line[^1] != '-')
                {
                    return text[..end];
                }

                if (backslashes > 0)
                {
                    return text[..lineStart] + line[..indent] + new string('\\', backslashes - 1) + "-";
                }

                // An empty bullet: the line goes, and what is above it is trimmed in turn.
                text = text[..lineStart];
            }
        }

        // ###########################################################################################
        // A note as written, read: its plain text and styles. Never fails - a mark it cannot pair is the character it
        // is. Line ends are "\n".
        // ###########################################################################################
        public static RichNoteText Parse(string? markup)
        {
            string text = (markup ?? string.Empty).ReplaceLineEndings("\n");
            var plain = new StringBuilder(text.Length);
            var styles = new List<RichNoteStyle>(text.Length);

            string[] lines = text.Split('\n');

            for (int index = 0; index < lines.Length; index++)
            {
                if (index > 0)
                {
                    plain.Append('\n');
                    styles.Add(RichNoteStyle.None);
                }

                ParseLine(lines[index], plain, styles);
            }

            return new RichNoteText(plain.ToString(), styles);
        }

        // ###########################################################################################
        // *** A BULLET'S INDENT IS TWO SPACES A LEVEL, MaxLevel LEVELS DEEP AT MOST (code review, 2026-11-08). *** After any
        // other indent "- " is the dash it is, and so is "\- " - a "\" only where it changes the reading: " - replaced C14"
        // (one space) read as a bullet, and the first edit wrote it back without its space, while twelve spaces came back
        // as ten. Write keeps to the very same rule.
        // ###########################################################################################
        private static bool IsBulletIndent(int indent) => indent % 2 == 0 && indent / 2 <= MaxLevel;

        private static void ParseLine(string line, StringBuilder plain, List<RichNoteStyle> styles)
        {
            int indent = IndentOf(line);
            string content = line;

            // At any other indent no bullet can start, and the line is as typed.
            bool canBeBullet = IsBulletIndent(indent);
            int backslashes = BackslashesBeforeDash(line, indent);

            if (canBeBullet && backslashes == 0)
            {
                Append(plain, styles, BulletPrefix(indent / 2), RichNoteStyle.None);
                content = line[(indent + 2)..];
            }
            else if (canBeBullet && backslashes > 0)
            {
                // ###########################################################################################
                // A run of backslashes before "- ": a line that starts with a dash and is no bullet - the first "\" says
                // so, the rest are themselves. "\- " reads "- ", "\\- " reads "\- ", and so on (code review, 2026-11-08:
                // three or more were read as typed, so a line shown with two had no way to be written and lost one at
                // every save).
                // ###########################################################################################
                Append(plain, styles, line[..indent] + new string('\\', backslashes - 1) + "- ", RichNoteStyle.None);
                content = line[(indent + backslashes + 2)..];
            }

            ParseInline(content, plain, styles);
        }

        // ###########################################################################################
        // How many backslashes stand at `indent` right before a "- " - 0 for "- " itself, -1 when no "- " follows them.
        // ###########################################################################################
        private static int BackslashesBeforeDash(string line, int indent)
        {
            int at = indent + BackslashesAt(line, indent);

            return line.Length >= at + 2 && line[at] == '-' && line[at + 1] == ' ' ? at - indent : -1;
        }

        // How many backslashes stand in a row from `at` on.
        private static int BackslashesAt(string line, int at)
        {
            int count = 0;
            while (at + count < line.Length && line[at + count] == '\\')
            {
                count++;
            }

            return count;
        }

        // `OnlyCloses`: a mark at the end of a web address, which may close one opened before the address and open none.
        private sealed record Token(RichNoteStyle Marker, string Raw, bool OnlyCloses = false)
        {
            public bool IsMarker => this.Marker != RichNoteStyle.None;
        }

        // ###########################################################################################
        // *** A "\" IS AN ESCAPE ONLY WHERE IT CHANGES THE READING (code review, 2026-11-07: "Dump copied to \\NAS\c64"
        // read as "\NAS\c64", "D:\_old" as "D:_old" - and the first edit saved them so, for good). *** Before a star,
        // always - a star alone is a mark; before "_" or "~" only when another of it follows ("__" and "~~" are the marks,
        // one alone is none); before a "\" only when, past the run of backslashes it starts, a mark comes - a "\" before a
        // mark that is meant is written "\\". Anywhere else a "\" is the character it is.
        //
        // So a run of backslashes before a mark is read in pairs, each pair one "\"; an odd one left over takes the mark's
        // first character as itself. A run before anything else is every backslash as typed. Serialize writes by the very
        // same rule (WrittenAfter.NeedsEscapeBefore), so what it writes reads back.
        //
        // `runEnds`: where each run of backslashes ends (BackslashRunEnds) - worked out once for the line, as asked here at
        // every "\" a run of 40,000 took most of a second (code review, 2026-11-08, Security: a note is anybody's to send).
        // ###########################################################################################
        private static bool IsEscapeAt(string text, int at, int[] runEnds) =>
            at < text.Length - 1 && text[at] == '\\' && StartsMark(text, runEnds[at]);

        // For each "\" of the text, the index just past the run of backslashes it stands in.
        private static int[] BackslashRunEnds(string text)
        {
            var ends = new int[text.Length];

            for (int at = text.Length - 1; at >= 0; at--)
            {
                ends[at] = text[at] != '\\' ? at : at + 1 < text.Length && text[at + 1] == '\\' ? ends[at + 1] : at + 1;
            }

            return ends;
        }

        // Whether a mark starts at `at`: a star, or two of "_" or "~".
        private static bool StartsMark(string text, int at) =>
            at < text.Length &&
            (text[at] == '*' || ((text[at] == '_' || text[at] == '~') && at + 1 < text.Length && text[at + 1] == text[at]));

        // ###########################################################################################
        // *** NOTHING INSIDE A WEB ADDRESS IS A MARK (code review, 2026-11-07: "...datamodel.html#object.__init__" read the
        // "__init__" as underlined, the link then opening "...#object.init"). *** The addresses TextLinkFinder finds in the
        // line - the links shown - are their own characters: no mark, no escape. Only a run of mark characters at an
        // address's very end is read as marks, and only to CLOSE one opened before it: "**https://x.dk**" is a bold link,
        // while "...#object.__init__" has nothing open to close, so its "__" stays the address's.
        //
        // Null when the line has no address - the common case, no arrays made.
        // ###########################################################################################
        private static (bool[] Literal, bool[] OnlyCloses)? LinkParts(string text)
        {
            if (!TextLinkFinder.ContainsLink(text))
            {
                return null;
            }

            var literal = new bool[text.Length];
            var onlyCloses = new bool[text.Length];

            foreach (TextLinkSpan span in TextLinkFinder.FindSpans(text))
            {
                if (!span.IsLink)
                {
                    continue;
                }

                int marks = span.End;
                while (marks > span.Start && MarkCharacters.IndexOf(text[marks - 1]) >= 0)
                {
                    marks--;
                }

                for (int at = span.Start; at < span.End; at++)
                {
                    literal[at] = at < marks;
                    onlyCloses[at] = at >= marks;
                }
            }

            return (literal, onlyCloses);
        }

        private static void ParseInline(string text, StringBuilder plain, List<RichNoteStyle> styles)
        {
            var tokens = new List<Token>(text.Length);
            (bool[] Literal, bool[] OnlyCloses)? links = LinkParts(text);
            int[] runEnds = BackslashRunEnds(text);

            for (int at = 0; at < text.Length;)
            {
                char character = text[at];
                bool onlyCloses = links?.OnlyCloses[at] == true;

                if (links?.Literal[at] == true)
                {
                    tokens.Add(new Token(RichNoteStyle.None, character.ToString()));
                    at++;
                }
                else if (IsEscapeAt(text, at, runEnds))
                {
                    tokens.Add(new Token(RichNoteStyle.None, text[at + 1].ToString()));
                    at += 2;
                }
                else if (Starts(text, at, "**"))
                {
                    tokens.Add(new Token(RichNoteStyle.Bold, "**", onlyCloses));
                    at += 2;
                }
                else if (Starts(text, at, "__"))
                {
                    tokens.Add(new Token(RichNoteStyle.Underline, "__", onlyCloses));
                    at += 2;
                }
                else if (Starts(text, at, "~~"))
                {
                    tokens.Add(new Token(RichNoteStyle.Strike, "~~", onlyCloses));
                    at += 2;
                }
                else if (character == '*')
                {
                    tokens.Add(new Token(RichNoteStyle.Italic, "*", onlyCloses));
                    at++;
                }
                else
                {
                    tokens.Add(new Token(RichNoteStyle.None, character.ToString()));
                    at++;
                }
            }

            // Which marks pair: one opens before a word (a character that is no space, past any other marks), the next
            // of its kind closes after one; one left open at the line's end is the character it is.
            var opens = new bool[tokens.Count];
            var closes = new bool[tokens.Count];
            var open = new Dictionary<RichNoteStyle, int>();
            (bool[] wordBefore, bool[] wordAfter) = WordsBesideTokens(tokens);

            for (int index = 0; index < tokens.Count; index++)
            {
                RichNoteStyle marker = tokens[index].Marker;

                if (marker == RichNoteStyle.None)
                {
                    continue;
                }

                if (open.TryGetValue(marker, out int opener) && wordBefore[index])
                {
                    opens[opener] = true;
                    closes[index] = true;
                    open.Remove(marker);
                }
                else if (!open.ContainsKey(marker) && !tokens[index].OnlyCloses && wordAfter[index])
                {
                    open[marker] = index;
                }
            }

            RichNoteStyle current = RichNoteStyle.None;

            for (int index = 0; index < tokens.Count; index++)
            {
                Token token = tokens[index];

                if (opens[index])
                {
                    current |= token.Marker;
                }
                else if (closes[index])
                {
                    current &= ~token.Marker;
                }
                else
                {
                    Append(plain, styles, token.Raw, current);
                }
            }
        }

        private static bool Starts(string text, int at, string mark) => string.CompareOrdinal(text, at, mark, 0, mark.Length) == 0;

        // ###########################################################################################
        // For each token, whether a word stands right before it and right after it, past any other marks: the nearest
        // character that is no mark is no space - none at the line's edge. Worked out once for the line, both ways
        // (code review, 2026-11-08, Security: looked for again at every mark, a line of 40,000 stars took a quarter second).
        // ###########################################################################################
        private static (bool[] Before, bool[] After) WordsBesideTokens(List<Token> tokens)
        {
            var before = new bool[tokens.Count];
            var after = new bool[tokens.Count];
            bool word = false;

            for (int index = 0; index < tokens.Count; index++)
            {
                before[index] = word;
                word = tokens[index].IsMarker ? word : !char.IsWhiteSpace(tokens[index].Raw[0]);
            }

            word = false;

            for (int index = tokens.Count - 1; index >= 0; index--)
            {
                after[index] = word;
                word = tokens[index].IsMarker ? word : !char.IsWhiteSpace(tokens[index].Raw[0]);
            }

            return (before, after);
        }

        private static void Append(StringBuilder plain, List<RichNoteStyle> styles, string text, RichNoteStyle style)
        {
            plain.Append(text);

            for (int index = 0; index < text.Length; index++)
            {
                styles.Add(style);
            }
        }

        // ###########################################################################################
        // A note as shown, written - with the fewest marks: first with no "\" at all, and only a line that would read back
        // as something else with one before every character that could be taken for a mark (code review, 2026-11-08: the
        // whole note went so, one line's trouble putting a "\" before every star of the others - lines are read one by
        // one, so each can be written on its own). A style on spaces at a word's edge is let go of (a mark must stand next
        // to a word), and a bullet's glyph is the one its level has.
        // ###########################################################################################
        public static string Serialize(RichNoteText note)
        {
            ArgumentNullException.ThrowIfNull(note);

            RichNoteText tidy = Tidy(note);
            string written = Write(tidy, escapeLines: null);

            if (Parse(written).SameAs(tidy))
            {
                return written;
            }

            // Write keeps the note's lines - a mark or a "\" is never a line end - so line N written is line N shown.
            List<RichNoteText> shown = LinesOf(tidy);
            bool[] escape = LinesNotReadBack(written, shown, null);
            string escaped = Write(tidy, escape);

            // ###########################################################################################
            // A line that does not read back even so - a style beside a web address, which the written line's marks make
            // an address of another length - goes without its styles, every mark character escaped: its words exactly as
            // shown, never stars or underscores showing where a style was (code review, 2026-11-08). Lines are read one by
            // one, and with no style the address is the shown one's, so that reads back.
            // ###########################################################################################
            bool[] unstyle = LinesNotReadBack(escaped, shown, escape);

            if (!unstyle.Contains(true))
            {
                return escaped;
            }

            var styles = tidy.Styles.ToList();
            int offset = 0;

            for (int line = 0; line < shown.Count; line++)
            {
                if (unstyle[line])
                {
                    styles.RemoveRange(offset, shown[line].Text.Length);
                    styles.InsertRange(offset, Enumerable.Repeat(RichNoteStyle.None, shown[line].Text.Length));
                }

                offset += shown[line].Text.Length + 1;
            }

            return Write(new RichNoteText(tidy.Text, styles), escape);
        }

        // Each line of a note, with its styles.
        private static List<RichNoteText> LinesOf(RichNoteText note)
        {
            var lines = new List<RichNoteText>();
            int offset = 0;

            foreach (string line in note.Text.Split('\n'))
            {
                lines.Add(new RichNoteText(line, note.Styles.Skip(offset).Take(line.Length).ToList()));
                offset += line.Length + 1;
            }

            return lines;
        }

        // Which lines of `written` do not read back as the line shown - of those `only` names, when given.
        private static bool[] LinesNotReadBack(string written, List<RichNoteText> shown, bool[]? only)
        {
            string[] writtenLines = written.Split('\n');
            var notReadBack = new bool[shown.Count];

            for (int line = 0; line < shown.Count; line++)
            {
                notReadBack[line] = (only is null || only[line]) && !Parse(writtenLines[line]).SameAs(shown[line]);
            }

            return notReadBack;
        }

        // ###########################################################################################
        // The note as it will read back: bullet glyphs by level, no style on a bullet's start, and no style on spaces at
        // the edges of what is styled.
        // ###########################################################################################
        internal static RichNoteText Tidy(RichNoteText note)
        {
            note = WithLineEndsOfOneCharacter(note);
            var text = new StringBuilder(note.Text.Length);
            var styles = new List<RichNoteStyle>(note.Text.Length);
            int offset = 0;
            bool first = true;

            foreach (string line in note.Text.Split('\n'))
            {
                if (!first)
                {
                    text.Append('\n');
                    styles.Add(RichNoteStyle.None);
                    offset++;
                }

                first = false;
                int contentStart = 0;
                int unstyledUntil = 0;

                if (TryReadBullet(line, out int level, out int prefixLength))
                {
                    Append(text, styles, BulletPrefix(level), RichNoteStyle.None);
                    contentStart = prefixLength;
                }
                else
                {
                    // A line that starts with "- " and is no bullet is written "\- " (Write) - no mark can stand inside
                    // that, so its start carries no style (code review, 2026-11-08: a bold "- replaced C14" did not read
                    // back, and the whole note was written with a "\" before every star).
                    int indent = IndentOf(line);
                    int backslashes = BackslashesBeforeDash(line, indent);

                    if (IsBulletIndent(indent) && backslashes >= 0)
                    {
                        unstyledUntil = indent + backslashes + 2;
                    }
                }

                for (int index = contentStart; index < line.Length; index++)
                {
                    text.Append(line[index]);
                    styles.Add(index < unstyledUntil ? RichNoteStyle.None : note.Styles[offset + index]);
                }

                offset += line.Length;
            }

            List<RichNoteStyle> trimmed = styles.ToList();
            string plain = text.ToString();

            // A web address is one style throughout, its first character's: no mark can stand inside one (LinkParts), so
            // a style on part of it could not be written - only around the whole.
            foreach (TextLinkSpan span in TextLinkFinder.FindSpans(plain))
            {
                if (span.IsLink)
                {
                    for (int at = span.Start + 1; at < span.End; at++)
                    {
                        trimmed[at] = trimmed[span.Start];
                    }
                }
            }

            foreach (RichNoteStyle flag in new[] { RichNoteStyle.Bold, RichNoteStyle.Italic, RichNoteStyle.Underline, RichNoteStyle.Strike })
            {
                TrimEdges(plain, trimmed, flag);
            }

            return new RichNoteText(plain, trimmed);
        }

        // ###########################################################################################
        // The note with every line end "\n", as Parse reads them (code review, 2026-11-07: a note pasted with Windows'
        // "\r\n" never read back as written, so Serialize put a "\" before every literal mark character in it). "\r\n" loses
        // its "\r"; a "\r" alone is a line end.
        // ###########################################################################################
        private static RichNoteText WithLineEndsOfOneCharacter(RichNoteText note)
        {
            if (!note.Text.Contains('\r', StringComparison.Ordinal))
            {
                return note;
            }

            var text = new StringBuilder(note.Text.Length);
            var styles = new List<RichNoteStyle>(note.Text.Length);

            for (int index = 0; index < note.Text.Length; index++)
            {
                char character = note.Text[index];

                if (character == '\r' && index + 1 < note.Text.Length && note.Text[index + 1] == '\n')
                {
                    continue;
                }

                text.Append(character == '\r' ? '\n' : character);
                styles.Add(character == '\r' ? RichNoteStyle.None : note.Styles[index]);
            }

            return new RichNoteText(text.ToString(), styles);
        }

        // Each run of `flag` let go of on the spaces and line ends at its edges.
        private static void TrimEdges(string text, List<RichNoteStyle> styles, RichNoteStyle flag)
        {
            int index = 0;

            while (index < text.Length)
            {
                if ((styles[index] & flag) == 0 || text[index] == '\n')
                {
                    styles[index] &= ~flag;
                    index++;
                    continue;
                }

                int start = index;
                while (index < text.Length && (styles[index] & flag) != 0 && text[index] != '\n')
                {
                    index++;
                }

                int end = index;

                for (int at = start; at < end && char.IsWhiteSpace(text[at]); at++)
                {
                    styles[at] &= ~flag;
                }

                for (int at = end - 1; at >= start && char.IsWhiteSpace(text[at]); at--)
                {
                    styles[at] &= ~flag;
                }
            }
        }

        private static readonly (RichNoteStyle Flag, string Mark)[] Marks =
        [
            (RichNoteStyle.Bold, "**"),
            (RichNoteStyle.Italic, "*"),
            (RichNoteStyle.Underline, "__"),
            (RichNoteStyle.Strike, "~~"),
        ];

        // ###########################################################################################
        // `escapeLines`, per line, null for none: a "\" before every character of that line that would be read as a mark
        // otherwise - worked out from the end of the line back, as whether a "\" is needed depends on what is written after
        // it (WrittenAfter, IsEscapeAt's rule). A web address's own characters get none: nothing in one is read as a mark or
        // an escape (LinkParts).
        // ###########################################################################################
        private static string Write(RichNoteText note, bool[]? escapeLines)
        {
            var written = new StringBuilder(note.Text.Length + 16);
            int offset = 0;
            int lineNumber = -1;

            foreach (string line in note.Text.Split('\n'))
            {
                lineNumber++;

                if (lineNumber > 0)
                {
                    written.Append('\n');
                    offset++;
                }

                bool escapeAll = escapeLines?[lineNumber] == true;
                int contentStart = 0;

                if (TryReadBullet(line, out int level, out int prefixLength))
                {
                    written.Append(new string(' ', level * 2)).Append("- ");
                    contentStart = prefixLength;
                }
                else
                {
                    int indent = IndentOf(line);

                    // A line that starts with "- " - after any run of backslashes - and is no bullet keeps them and its
                    // dash as themselves: one more "\" in front, ParseLine's rule, at an indent a bullet can have
                    // (IsBulletIndent), where it would read otherwise.
                    int backslashes = BackslashesBeforeDash(line, indent);

                    if (IsBulletIndent(indent) && backslashes >= 0)
                    {
                        written.Append(line[..indent]).Append('\\', backslashes + 1).Append("- ");
                        contentStart = indent + backslashes + 2;
                    }
                }

                bool[]? inLink = escapeAll ? LinkCharacters(line) : null;
                var pieces = new List<(string Text, bool CanEscape)>(line.Length);
                RichNoteStyle current = RichNoteStyle.None;

                for (int index = contentStart; index < line.Length; index++)
                {
                    RichNoteStyle style = note.Styles[offset + index];

                    foreach ((RichNoteStyle flag, string mark) in Marks.Reverse())
                    {
                        if ((current & flag) != 0 && (style & flag) == 0)
                        {
                            pieces.Add((mark, false));
                        }
                    }

                    foreach ((RichNoteStyle flag, string mark) in Marks)
                    {
                        if ((current & flag) == 0 && (style & flag) != 0)
                        {
                            pieces.Add((mark, false));
                        }
                    }

                    current = style;
                    pieces.Add((line[index].ToString(), escapeAll && inLink?[index] != true));
                }

                foreach ((RichNoteStyle flag, string mark) in Marks.Reverse())
                {
                    if ((current & flag) != 0)
                    {
                        pieces.Add((mark, false));
                    }
                }

                // From the end back: `after` holds what is written after the piece, reversed, and `follows` what is asked
                // of it. An escapable piece is one character.
                var after = new StringBuilder(line.Length + 8);
                var follows = new WrittenAfter();

                for (int piece = pieces.Count - 1; piece >= 0; piece--)
                {
                    (string text, bool canEscape) = pieces[piece];
                    bool escape = canEscape && follows.NeedsEscapeBefore(text[0]);

                    for (int at = text.Length - 1; at >= 0; at--)
                    {
                        after.Append(text[at]);
                        follows.Prepend(text[at]);
                    }

                    if (escape)
                    {
                        after.Append('\\');
                        follows.Prepend('\\');
                    }
                }

                for (int at = after.Length - 1; at >= 0; at--)
                {
                    written.Append(after[at]);
                }

                offset += line.Length;
            }

            return written.ToString();
        }

        // ###########################################################################################
        // Every character of the web addresses in a line as shown, its ending mark characters too - none is written with a
        // "\": inside an address one would be read as itself (LinkParts), and the ending marks close only marks written
        // outside it, which get their own "\" where needed. Null with none.
        // ###########################################################################################
        private static bool[]? LinkCharacters(string line) =>
            LinkParts(line) is { } parts ? parts.Literal.Select((literal, at) => literal || parts.OnlyCloses[at]).ToArray() : null;

        // ###########################################################################################
        // What is written after a place in a line, as far as whether a character there needs a "\" (IsEscapeAt's rule): its
        // first character, how many backslashes it starts with, and whether a mark comes right after them. Kept as each
        // character goes in front, so the question costs the same however long a run of backslashes is (code review,
        // 2026-11-08, Security: walked again at every "\", 40,000 of them took most of a second to write).
        // ###########################################################################################
        private struct WrittenAfter
        {
            private char thisFirst;
            private int thisBackslashes;
            private bool thisMarkAfterBackslashes;

            // `character` written in front of what is after.
            public void Prepend(char character)
            {
                if (character == '\\')
                {
                    this.thisBackslashes++;
                }
                else
                {
                    this.thisBackslashes = 0;
                    this.thisMarkAfterBackslashes = character == '*' || ((character == '_' || character == '~') && this.thisFirst == character);
                }

                this.thisFirst = character;
            }

            // Whether `character`, as itself, needs a "\" right in front of what is after: a star always, "_" or "~" before
            // another of it, a "\" when a mark comes past the backslashes after it.
            public readonly bool NeedsEscapeBefore(char character) => character switch
            {
                '*' => true,
                '_' or '~' => this.thisFirst == character,
                '\\' => this.thisMarkAfterBackslashes,
                _ => false,
            };
        }
    }
}
