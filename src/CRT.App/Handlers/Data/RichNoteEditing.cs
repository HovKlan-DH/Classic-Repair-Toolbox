using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT AN EDIT DOES TO A FORMATTED NOTE (RichNoteTextBox; owner request, 2026-11-07: bold, italic, underline,
    // strikethrough and nested bullets, "a very simple WYSIWYG ... not ... 'heavy' and slow"). The box keeps a note as its
    // plain text and each character's style (RichNote); every rule of how a change moves them is here, pure, so a test
    // reads it:
    //   - a change of the text - typing, a paste, a cut, a delete - keeps every character's style, and a character typed
    //     takes the style of the one before it on its line, or the style asked for before typing: Ctrl+B with nothing
    //     selected;
    //   - a style toggled on a selection: on where any of it lacks it, off where all has it - never on a bullet's start
    //     or a line end;
    //   - bullets: on or off for the lines the selection touches, Tab and Shift+Tab nesting them deeper or less, Enter
    //     on one starting the next at its level (on an empty one ending the list), Backspace at its words' start taking
    //     the bullet away;
    //   - the format buttons pressed by that same rule: a pressed one is one a press turns off (Pressed).
    // ###########################################################################################
    public sealed record RichNoteEdit(string Text, IReadOnlyList<RichNoteStyle> Styles, int SelectionStart, int SelectionEnd);

    public static class RichNoteEditing
    {
        // ###########################################################################################
        // Where one change took the text from `before` to `after`: at what index, how many characters it took out there and
        // how many it put in - the fewest of each. Beside a character like the one typed or taken out, several places say
        // the same ("b" typed into "ab" after its "a", or after its "b"); `changeAt`, where the caller knows the change
        // began - the caret, or the selection's start - picks among them (code review, 2026-11-08: worked out from the text
        // alone, "b" typed after a bold "a" took the plain "b"'s style). Without it, the last of them.
        // ###########################################################################################
        public static (int At, int Removed, int Inserted) ChangeOf(string before, string after, int? changeAt = null)
        {
            before ??= string.Empty;
            after ??= string.Empty;

            int shorter = Math.Min(before.Length, after.Length);

            int prefix = 0;
            while (prefix < shorter && before[prefix] == after[prefix])
            {
                prefix++;
            }

            int suffix = 0;
            while (suffix < shorter && before[before.Length - 1 - suffix] == after[after.Length - 1 - suffix])
            {
                suffix++;
            }

            // The most both share around the change, and the places the change can start at for that.
            int shared = Math.Min(prefix + suffix, shorter);
            int earliest = Math.Max(0, shared - suffix);
            int latest = Math.Min(prefix, shared);
            int at = changeAt is { } wanted ? Math.Clamp(wanted, earliest, latest) : latest;

            return (at, before.Length - shared, after.Length - shared);
        }

        // ###########################################################################################
        // The styles after the text went from `before` to `after` in one change (ChangeOf, `changeAt` where it began): the
        // part both share at the start and the end keeps its styles, what was taken out goes, and what came in takes
        // `typing` - or the style of the character before it on its line, none at a line's start; a line end never has one.
        // ###########################################################################################
        public static List<RichNoteStyle> FollowTextChange(string before, IReadOnlyList<RichNoteStyle> styles, string after, RichNoteStyle? typing = null, int? changeAt = null)
        {
            before ??= string.Empty;
            after ??= string.Empty;

            var kept = styles.Take(before.Length).ToList();
            while (kept.Count < before.Length)
            {
                kept.Add(RichNoteStyle.None);
            }

            (int prefix, int removed, int inserted) = ChangeOf(before, after, changeAt);

            kept.RemoveRange(prefix, removed);

            // What stands before the change kept its place - so the style before it is read from what is kept.
            RichNoteStyle inherited = typing ?? StyleTypedAt(before, kept, prefix);

            for (int index = 0; index < inserted; index++)
            {
                bool lineEnd = after[prefix + index] == '\n';
                kept.Insert(prefix + index, lineEnd ? RichNoteStyle.None : inherited);

                // The words after a line end put in have nothing before them on their line.
                if (lineEnd && typing is null)
                {
                    inherited = RichNoteStyle.None;
                }
            }

            return kept;
        }

        // ###########################################################################################
        // `flag` on the characters from `start` to `end`: on all of them where any lacks it, off where all have it. A
        // bullet's start and a line end take none. The selection stays as it was.
        // ###########################################################################################
        public static RichNoteEdit Toggle(string text, IReadOnlyList<RichNoteStyle> styles, int start, int end, RichNoteStyle flag)
        {
            (start, end) = Ordered(text, start, end);
            bool[] styleable = Styleable(text);
            var next = styles.ToList();

            bool allHave = AllHave(next, styleable, start, end, flag);

            for (int index = start; index < end; index++)
            {
                if (styleable[index])
                {
                    next[index] = allHave ? next[index] & ~flag : next[index] | flag;
                }
            }

            return new RichNoteEdit(text, next, start, end);
        }

        // ###########################################################################################
        // Bullets on the lines the selection touches - all of them bullets already: taken away, else put on every one
        // that has none, at the first level. A selection ending at a line's very start does not touch that line.
        // ###########################################################################################
        public static RichNoteEdit ToggleBullets(string text, IReadOnlyList<RichNoteStyle> styles, int start, int end)
        {
            List<Line> lines = LinesOf(text);
            (int first, int last) = Touched(lines, start, end);
            bool allBullets = AllBullets(lines, first, last);

            return Rebuild(text, styles, lines, start, end, (line, index) =>
                index < first || index > last ? Kept(line)
                : allBullets ? -1
                : line.IsBullet ? line.Level : 0);
        }

        // ###########################################################################################
        // *** WHAT THE FORMAT BUTTONS SHOW PRESSED (owner decision, 2026-11-08: "I think that would be nice so do it"). ***
        // By the very rule a press goes by, so a pressed button is one a press turns OFF: with a selection, each style all
        // of it has (AllHave, Toggle's) and bullets while every line it touches is one (AllBullets, ToggleBullets'); with
        // nothing selected, the style the next character typed takes - `asked` (Ctrl+B with nothing selected), else the one
        // before the caret on its line (StyleTypedAt) - and whether the caret's line is a bullet.
        // ###########################################################################################
        public static (RichNoteStyle Styles, bool Bullets) Pressed(string text, IReadOnlyList<RichNoteStyle> styles, int start, int end, RichNoteStyle? asked = null)
        {
            text ??= string.Empty;
            (start, end) = Ordered(text, start, end);

            List<Line> lines = LinesOf(text);
            (int first, int last) = Touched(lines, start, end);
            bool bullets = AllBullets(lines, first, last);

            if (start == end)
            {
                return (asked ?? StyleTypedAt(text, styles, start), bullets);
            }

            bool[] styleable = Styleable(text);
            RichNoteStyle pressed = RichNoteStyle.None;

            foreach (RichNoteStyle flag in new[] { RichNoteStyle.Bold, RichNoteStyle.Italic, RichNoteStyle.Underline, RichNoteStyle.Strike })
            {
                if (AllHave(styles, styleable, start, end, flag))
                {
                    pressed |= flag;
                }
            }

            return (pressed, bullets);
        }

        // ###########################################################################################
        // The style a character typed at `caret` takes without being asked: the one before it on its line - none at a
        // line's start.
        // ###########################################################################################
        public static RichNoteStyle StyleTypedAt(string text, IReadOnlyList<RichNoteStyle> styles, int caret) =>
            caret > 0 && caret <= text.Length && caret <= styles.Count && text[caret - 1] != '\n' ? styles[caret - 1] : RichNoteStyle.None;

        // Whether every character from `start` to `end` that can carry a style has `flag` - no for a range with none.
        private static bool AllHave(IReadOnlyList<RichNoteStyle> styles, bool[] styleable, int start, int end, RichNoteStyle flag) =>
            Enumerable.Range(start, end - start).Where(index => styleable[index]).DefaultIfEmpty(-1)
                .All(index => index >= 0 && (styles[index] & flag) != 0);

        private static bool AllBullets(List<Line> lines, int first, int last) =>
            Enumerable.Range(first, last - first + 1).All(index => lines[index].IsBullet);

        // ###########################################################################################
        // The bullet lines the selection touches nested `by` levels deeper (or less), within 0 to MaxLevel - null when no
        // level would change (none touched, Shift+Tab at the top level, Tab at the deepest), so Tab keeps its own meaning
        // there and no empty step is left to undo (code review, 2026-11-08: Shift+Tab at the top level held the keyboard).
        // ###########################################################################################
        public static RichNoteEdit? Indent(string text, IReadOnlyList<RichNoteStyle> styles, int start, int end, int by)
        {
            List<Line> lines = LinesOf(text);
            (int first, int last) = Touched(lines, start, end);

            if (!Enumerable.Range(first, last - first + 1).Any(index => lines[index].IsBullet && Math.Clamp(lines[index].Level + by, 0, RichNote.MaxLevel) != lines[index].Level))
            {
                return null;
            }

            return Rebuild(text, styles, lines, start, end, (line, index) =>
                index < first || index > last || !line.IsBullet ? Kept(line) : Math.Clamp(line.Level + by, 0, RichNote.MaxLevel));
        }

        // ###########################################################################################
        // Enter on a bullet line, in its words: the next line starts with a bullet of its level - or, on a bullet with no
        // words, the bullet goes and the list ends there. Null anywhere else: Enter is a line end as ever.
        // ###########################################################################################
        public static RichNoteEdit? EnterOnBullet(string text, IReadOnlyList<RichNoteStyle> styles, int caret)
        {
            List<Line> lines = LinesOf(text);
            Line line = lines.First(each => caret >= each.Start && caret <= each.End);

            if (!line.IsBullet || caret < line.ContentStart)
            {
                return null;
            }

            if (text[line.ContentStart..line.End].Trim().Length == 0)
            {
                return Rebuild(text, styles, lines, caret, caret, (each, index) => ReferenceEquals(each, line) ? -1 : Kept(each)) with
                {
                    SelectionStart = line.Start,
                    SelectionEnd = line.Start,
                };
            }

            string prefix = "\n" + RichNote.BulletPrefix(line.Level);
            var next = styles.ToList();
            next.InsertRange(caret, Enumerable.Repeat(RichNoteStyle.None, prefix.Length));

            return new RichNoteEdit(text.Insert(caret, prefix), next, caret + prefix.Length, caret + prefix.Length);
        }

        // ###########################################################################################
        // Backspace right where a bullet's words start: the bullet goes, its words stay where they were in the text. Null
        // anywhere else.
        // ###########################################################################################
        public static RichNoteEdit? BackspaceAtBullet(string text, IReadOnlyList<RichNoteStyle> styles, int caret)
        {
            List<Line> lines = LinesOf(text);
            Line? line = lines.FirstOrDefault(each => each.IsBullet && caret == each.ContentStart);

            if (line is null)
            {
                return null;
            }

            return Rebuild(text, styles, lines, caret, caret, (each, index) => ReferenceEquals(each, line) ? -1 : Kept(each)) with
            {
                SelectionStart = line.Start,
                SelectionEnd = line.Start,
            };
        }

        // A line left as it is: its bullet at its level, or none.
        private static int Kept(Line line) => line.IsBullet ? line.Level : -1;

        // ---- the lines of a note ---------------------------------------------------------------

        private sealed record Line(int Start, int End, bool IsBullet, int Level, int PrefixLength)
        {
            public int ContentStart => this.Start + this.PrefixLength;
        }

        private static List<Line> LinesOf(string text)
        {
            var lines = new List<Line>();
            int start = 0;

            foreach (string line in text.Split('\n'))
            {
                bool bullet = RichNote.TryReadBullet(line, out int level, out int prefix);
                lines.Add(new Line(start, start + line.Length, bullet, bullet ? level : 0, bullet ? prefix : 0));
                start += line.Length + 1;
            }

            return lines;
        }

        private static (int First, int Last) Touched(List<Line> lines, int start, int end)
        {
            int low = Math.Min(start, end);
            int high = Math.Max(start, end);

            int first = lines.FindIndex(line => low <= line.End);
            int last = lines.FindIndex(line => high <= line.End);

            first = first < 0 ? lines.Count - 1 : first;
            last = last < 0 ? lines.Count - 1 : last;

            // A selection ending at a line's very start does not touch it.
            if (last > first && high == lines[last].Start)
            {
                last--;
            }

            return (first, last);
        }

        // ###########################################################################################
        // The note made again with each line's bullet as `levelOf` says (-1 for none), its words and their styles kept,
        // and the selection moved with the words it was in.
        // ###########################################################################################
        private static RichNoteEdit Rebuild(string text, IReadOnlyList<RichNoteStyle> styles, List<Line> lines, int start, int end, Func<Line, int, int> levelOf)
        {
            var built = new StringBuilder(text.Length + 8);
            var next = new List<RichNoteStyle>(styles.Count + 8);
            var moved = new int[text.Length + 1];

            for (int index = 0; index < lines.Count; index++)
            {
                Line line = lines[index];

                if (index > 0)
                {
                    moved[line.Start - 1] = built.Length;
                    built.Append('\n');
                    next.Add(RichNoteStyle.None);
                }

                int level = levelOf(line, index);
                string prefix = level < 0 ? string.Empty : RichNote.BulletPrefix(level);
                int lineStartNow = built.Length;

                built.Append(prefix);
                next.AddRange(Enumerable.Repeat(RichNoteStyle.None, prefix.Length));

                // A position in the old start maps to the new start; one in the words, by its place in them.
                for (int at = line.Start; at < line.ContentStart; at++)
                {
                    moved[at] = lineStartNow + (at == line.Start ? 0 : prefix.Length);
                }

                for (int at = line.ContentStart; at <= line.End; at++)
                {
                    moved[at] = built.Length + (at - line.ContentStart);
                }

                built.Append(text, line.ContentStart, line.End - line.ContentStart);
                next.AddRange(styles.Skip(line.ContentStart).Take(line.End - line.ContentStart));
            }

            int Map(int position) => moved[Math.Clamp(position, 0, text.Length)];

            return new RichNoteEdit(built.ToString(), next, Map(start), Map(end));
        }

        private static (int Start, int End) Ordered(string text, int start, int end) =>
            (Math.Clamp(Math.Min(start, end), 0, text.Length), Math.Clamp(Math.Max(start, end), 0, text.Length));

        // Which characters can carry a style: not a line end, not a bullet's start.
        private static bool[] Styleable(string text)
        {
            var styleable = new bool[text.Length];

            foreach (Line line in LinesOf(text))
            {
                for (int at = line.ContentStart; at < line.End; at++)
                {
                    styleable[at] = true;
                }
            }

            return styleable;
        }
    }
}
