using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** PICTURES IN THE FEEDBACK TEXT (owner request, 2026-10-15). *** The Feedback tab's text is
    // a column of plain text PARTS with a picture between each pair, drawn as one field (CRT.App's
    // FeedbackComposer). These are the rules the control follows, and what travels: the parts
    // joined into one plain text with a marker line "[pasted-image-1.png]" where each picture
    // sat, the picture itself at the zip's root under that name. The SERVER reads the same rule
    // back (TryParseMarker) to show each picture in the mail where its line sits - so the two
    // ends can only agree, both being this class.
    // ###########################################################################################
    public static class FeedbackComposition
    {
        public const string PastedPrefix = "pasted-image-";
        public const string PastedExtension = ".png";

        // A pasted picture's name: numbered in the order pasted, a number never used twice in one
        // feedback (the control keeps the counter), so the markers stay unambiguous after a removal.
        public static string PastedName(int number) => $"{FeedbackComposition.PastedPrefix}{number}{FeedbackComposition.PastedExtension}";

        // ###########################################################################################
        // The next pasted picture's name from number `from` on, skipping a name a picture already
        // has - a dropped file can be called "pasted-image-1.png" too, and two zip entries of one
        // name showed the first picture at both markers (code review, 2026-10-16). Number is the one
        // used; the control counts on from the one after it.
        // ###########################################################################################
        public static (string Name, int Number) NextPastedName(int from, IEnumerable<string> taken)
        {
            ArgumentNullException.ThrowIfNull(taken);

            var set = new HashSet<string>(taken, StringComparer.OrdinalIgnoreCase);
            int number = Math.Max(1, from);

            while (set.Contains(FeedbackComposition.PastedName(number)))
                number++;

            return (FeedbackComposition.PastedName(number), number);
        }

        // The line that stands for a picture in the text that is sent.
        public static string Marker(string pictureName) => $"[{pictureName}]";

        // A zero-width space: put after a line the SENDER typed that would read as a marker, so it
        // never takes a picture (ComposeText). Invisible in the mail, and no whitespace, so the
        // trimming TryParseMarker does keeps it - and the line is no marker.
        public const char MarkerEscape = '​';

        // ###########################################################################################
        // *** HOW WIDE A PICTURE IS SHOWN IN THE MAIL (owner request, 2026-10-15: "show the inserted
        // image with the real size it will be shown in mail, and maybe give the option to resize in
        // the corners"). *** The mail's quotation has room for MaxPictureWidth (CRT.Server's
        // MailBody, held to it by a test there), so a picture is shown at its own width up to that.
        // CRT's field shows it at the width it will have, and its corners make it smaller. The width
        // chosen travels beside the text (FeedbackContract.PictureWidthsField) and becomes the
        // picture's width in the mail; the picture itself goes at full size, to be opened from there.
        // Never wider than its own width - that would only blur it.
        // ###########################################################################################
        public const int MaxPictureWidth = 980;

        public const int MinPictureWidth = 32;

        public static int DefaultPictureWidth(int naturalWidth) => Math.Clamp(naturalWidth, 1, FeedbackComposition.MaxPictureWidth);

        // The widths as they travel: {"pasted-image-1.png":640,...}; empty when there are none.
        public static string PictureWidthsText(IEnumerable<(string Name, int Width)> widths)
        {
            ArgumentNullException.ThrowIfNull(widths);

            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach ((string name, int width) in widths)
                map.TryAdd(name, width);

            return map.Count == 0 ? string.Empty : JsonSerializer.Serialize(map);
        }

        // ###########################################################################################
        // The widths read back - by the server, from text the SENDER wrote, so nothing in it is
        // trusted: not JSON, not an object, a value that is not a whole number, or a width outside
        // 1..MaxPictureWidth, and that entry (or all of it) is simply not there. A picture with no
        // width is shown at its own, no wider than the mail. Names compare as file names do.
        //
        // Not only JsonException (code review, 2026-10-16): a field cut between the two halves of
        // a surrogate pair cannot even be handed to the parser (ArgumentException), and a name that
        // unescapes to half of one cannot be read as text (InvalidOperationException) - each ended
        // the request in a 500.
        // ###########################################################################################
        public static IReadOnlyDictionary<string, int> ReadPictureWidths(string? text)
        {
            var widths = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(text))
                return widths;

            try
            {
                using JsonDocument document = JsonDocument.Parse(text);

                if (document.RootElement.ValueKind != JsonValueKind.Object)
                    return widths;

                foreach (JsonProperty property in document.RootElement.EnumerateObject())
                {
                    if (property.Value.ValueKind == JsonValueKind.Number &&
                        property.Value.TryGetInt32(out int width) &&
                        width is >= 1 and <= FeedbackComposition.MaxPictureWidth &&
                        FeedbackComposition.TryReadName(property, out string name))
                    {
                        widths.TryAdd(name, width);
                    }
                }
            }
            catch (Exception ex) when (ex is JsonException or ArgumentException)
            {
                widths.Clear();
            }

            return widths;
        }

        // A property's name, when it is text at all once unescaped.
        private static bool TryReadName(JsonProperty property, out string name)
        {
            try
            {
                name = property.Name;
                return true;
            }
            catch (InvalidOperationException)
            {
                name = string.Empty;
                return false;
            }
        }

        // ###########################################################################################
        // The picture a line of the text names, when the whole line is a marker - "[shot.png]",
        // whitespace around it allowed. A line with words around the brackets, or brackets inside,
        // or a path rather than a name, is the sender's own text and names nothing.
        // ###########################################################################################
        public static bool TryParseMarker(string? line, out string pictureName)
        {
            pictureName = string.Empty;
            string trimmed = line?.Trim() ?? string.Empty;

            if (trimmed.Length < 3 || trimmed[0] != '[' || trimmed[^1] != ']')
                return false;

            string inner = trimmed[1..^1].Trim();

            if (inner.Length == 0 || inner.IndexOfAny(['[', ']', '/', '\\']) >= 0)
                return false;

            pictureName = inner;
            return true;
        }

        // ###########################################################################################
        // *** THE TEXT'S LINES, EACH WITH THE PICTURE IT NAMES - THE ONE READING OF THE MARKERS
        // (code review, 2026-10-16). *** The server read the lines twice - once for the pictures to
        // take out of the zip, once to put them in the mail - and two readings can drift apart.
        // A line break is "\r\n", "\n" or "\r", as Split has them.
        // ###########################################################################################
        public static IReadOnlyList<FeedbackLine> ReadLines(string? text)
        {
            var lines = new List<FeedbackLine>();
            string all = text ?? string.Empty;
            int start = 0;

            for (int i = 0; i <= all.Length; i++)
            {
                if (i < all.Length && all[i] != '\n' && all[i] != '\r')
                    continue;

                string line = all[start..i];
                lines.Add(new FeedbackLine(line, FeedbackComposition.TryParseMarker(line, out string name) ? name : null));

                if (i + 1 < all.Length && all[i] == '\r' && all[i + 1] == '\n')
                    i++;

                start = i + 1;
            }

            return lines;
        }

        // The pictures the text's marker lines name, in the order it names them, each once -
        // compared as file names are.
        public static IReadOnlyList<string> MarkedNames(string? text)
        {
            var names = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (FeedbackLine line in FeedbackComposition.ReadLines(text))
            {
                if (line.PictureName is { } name && seen.Add(name))
                    names.Add(name);
            }

            return names;
        }

        // ###########################################################################################
        // A dropped picture's name as it travels: its own, with what a marker line cannot carry put
        // right (code review, 2026-10-16: "Screenshot [scope].png" was shown in CRT and never in the
        // mail - TryParseMarker refuses brackets inside a marker). Brackets become parentheses, a
        // backslash (a Linux name may hold one) and a control character a hyphen, and the ends are
        // trimmed, as a marker's name is. A name nothing is left of is "picture".
        // ###########################################################################################
        public static string PictureNameFor(string fileName)
        {
            ArgumentNullException.ThrowIfNull(fileName);

            var name = new StringBuilder(fileName.Length);

            foreach (char character in fileName)
            {
                name.Append(character switch
                {
                    '[' => '(',
                    ']' => ')',
                    '\\' or '/' => '-',
                    _ when char.IsControl(character) => '-',
                    _ => character
                });
            }

            string trimmed = name.ToString().Trim();
            return trimmed.Length == 0 ? "picture" : trimmed;
        }

        // ###########################################################################################
        // A dropped file keeps its own name - the reader of the mail knows "scope-capture-12.png" -
        // unless a picture already has it: then "name (2).ext", "name (3).ext", ... Names are
        // compared without regard to capitals, as the zip and the feedback folder on the server are.
        // The same rule keeps the zip's entries apart (code review, 2026-10-16: an attachment named
        // as a picture was a second entry of that name, and one of them was lost) - so a name with a
        // folder in front keeps the folder: "logs/run.txt", "logs/run (2).txt".
        //
        // A set the caller already holds, comparing as file names do, is used as it is (code review,
        // 2026-10-16): copied on every call, a zip of thousands of clashing names cost its square.
        // ###########################################################################################
        public static string UniqueName(string wanted, IEnumerable<string> taken)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(wanted);
            ArgumentNullException.ThrowIfNull(taken);

            HashSet<string> set = taken is HashSet<string> held && held.Comparer.Equals(StringComparer.OrdinalIgnoreCase)
                ? held
                : new HashSet<string>(taken, StringComparer.OrdinalIgnoreCase);

            if (!set.Contains(wanted))
                return wanted;

            int slash = wanted.LastIndexOfAny(['/', '\\']);
            string folder = wanted[..(slash + 1)];
            string fileName = wanted[(slash + 1)..];
            string extension = Path.GetExtension(fileName);
            string stem = fileName[..^extension.Length];
            int number = 1;
            string candidate;

            do
            {
                number++;
                candidate = $"{folder}{stem} ({number}){extension}";
            }
            while (set.Contains(candidate));

            return candidate;
        }

        // ###########################################################################################
        // The text split at the caret, for a picture to go between the halves - and the ONE line
        // break the picture takes over (owner report, 2026-10-15: Enter and then Ctrl+V showed an
        // empty line above the picture): a break just before the caret goes with the picture, else
        // one just after it does, so the picture sits on the line the user gave it, as in Word.
        // Join gives the break back when the picture goes.
        // ###########################################################################################
        public static PictureSplit Split(string text, int caret)
        {
            ArgumentNullException.ThrowIfNull(text);

            int at = Math.Clamp(caret, 0, text.Length);
            string before = text[..at];
            string after = text[at..];

            foreach (string lineBreak in FeedbackComposition.LineBreaks)
            {
                if (before.EndsWith(lineBreak, StringComparison.Ordinal))
                    return new PictureSplit(before[..^lineBreak.Length], after, lineBreak, TakenFromBefore: true);
            }

            foreach (string lineBreak in FeedbackComposition.LineBreaks)
            {
                if (after.StartsWith(lineBreak, StringComparison.Ordinal))
                    return new PictureSplit(before, after[lineBreak.Length..], lineBreak, TakenFromBefore: false);
            }

            return new PictureSplit(before, after, string.Empty, TakenFromBefore: false);
        }

        // ###########################################################################################
        // A picture pasted over a selection REPLACES it, as any paste into a text box does (code
        // review, 2026-10-16: the selected text stayed beside the picture). The selection goes - its
        // ends in either order, clamped to the text - and the text splits where it began. Removing
        // the picture later does not bring the selection back: that is a removal, not an undo.
        // ###########################################################################################
        public static PictureSplit Split(string text, int selectionStart, int selectionEnd)
        {
            ArgumentNullException.ThrowIfNull(text);

            int start = Math.Clamp(Math.Min(selectionStart, selectionEnd), 0, text.Length);
            int end = Math.Clamp(Math.Max(selectionStart, selectionEnd), 0, text.Length);

            return FeedbackComposition.Split(text.Remove(start, end - start), start);
        }

        // Two parts made one again when the picture between them goes, the line break it took
        // put back; the caret belongs where it was before the picture came.
        public static (string Text, int Caret) Join(string before, string after, string lineBreakTaken = "", bool takenFromBefore = false)
        {
            ArgumentNullException.ThrowIfNull(before);
            ArgumentNullException.ThrowIfNull(after);
            ArgumentNullException.ThrowIfNull(lineBreakTaken);

            return (before + lineBreakTaken + after, takenFromBefore ? before.Length + lineBreakTaken.Length : before.Length);
        }

        // ###########################################################################################
        // What is sent: texts[0], a marker for pictures[0], texts[1], ... texts[n] - every marker on
        // a line of its own, with no blank line added around it. texts has one more entry than
        // pictureNames; an empty text contributes nothing, so two pictures with nothing between
        // them are two marker lines.
        //
        // *** ONLY CRT WRITES A MARKER (code review, 2026-10-16). *** A line the sender typed that
        // would read as one - "[pasted-image-1.png]", to point at the screenshot further down -
        // took the picture to itself and left CRT's own marker as text. It is sent with MarkerEscape
        // after it, which no marker has.
        // ###########################################################################################
        public static string ComposeText(IReadOnlyList<string> texts, IReadOnlyList<string> pictureNames)
        {
            ArgumentNullException.ThrowIfNull(texts);
            ArgumentNullException.ThrowIfNull(pictureNames);

            if (texts.Count != pictureNames.Count + 1)
                throw new ArgumentException($"{texts.Count} texts cannot hold {pictureNames.Count} pictures between them.", nameof(texts));

            var result = new StringBuilder(FeedbackComposition.EscapeMarkerLines(texts[0]));

            for (int i = 0; i < pictureNames.Count; i++)
            {
                if (result.Length > 0 && result[^1] != '\n')
                    result.Append('\n');

                result.Append(FeedbackComposition.Marker(pictureNames[i]));

                string next = FeedbackComposition.EscapeMarkerLines(texts[i + 1]);

                if (next.Length > 0)
                    result.Append('\n').Append(next);
            }

            return result.ToString();
        }

        // The sender's text with MarkerEscape after every line that would read as a marker; every
        // other character, line breaks included, as typed.
        private static string EscapeMarkerLines(string text)
        {
            if (text.IndexOf('[') < 0)
                return text;

            var escaped = new StringBuilder(text.Length + 4);
            int start = 0;

            for (int i = 0; i <= text.Length; i++)
            {
                if (i < text.Length && text[i] != '\n' && text[i] != '\r')
                    continue;

                string line = text[start..i];
                escaped.Append(line);

                if (FeedbackComposition.TryParseMarker(line, out _))
                    escaped.Append(FeedbackComposition.MarkerEscape);

                if (i < text.Length)
                    escaped.Append(text[i]);

                start = i + 1;
            }

            return escaped.ToString();
        }

        // Longest first, so "\r\n" is never read as a "\n" with a stray "\r" left behind.
        private static readonly string[] LineBreaks = ["\r\n", "\n", "\r"];
    }

    // ###########################################################################################
    // The text either side of a picture put in at the caret, and the one line break the picture
    // took over - empty when it took none - with which side it came from, so Join can put the
    // caret back where the picture came in.
    // ###########################################################################################
    public sealed record PictureSplit(string Before, string After, string LineBreakTaken, bool TakenFromBefore);

    // One line of a feedback's text, without its line break, and the picture it names when the
    // whole line is a marker (null otherwise).
    public sealed record FeedbackLine(string Text, string? PictureName);
}
