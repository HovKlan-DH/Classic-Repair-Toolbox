using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // One run inside a user-typed text: either plain prose or a web link the UI should make
    // clickable. Start/Length index into the ORIGINAL text, so a caller can rebuild the whole
    // string from the spans in order without losing or doubling a character.
    //
    // Url is the target to hand ExternalTargetLauncher - it carries an explicit scheme even when
    // the text itself did not ("www.example.com" is displayed as typed but opened as
    // "https://www.example.com"). For a plain span it is null.
    // ###########################################################################################
    public readonly struct TextLinkSpan
    {
        public TextLinkSpan(int start, int length, string? url)
        {
            this.Start = start;
            this.Length = length;
            this.Url = url;
        }

        public int Start { get; }

        public int Length { get; }

        public int End => this.Start + this.Length;

        // Null for a plain-prose span; the absolute http/https URL to open for a link span.
        public string? Url { get; }

        public bool IsLink => this.Url != null;
    }

    // ###########################################################################################
    // Finds the web links inside a free-text field the user typed - a workbook's Note, a worklog's
    // Description, and the Work done / Comment / Photo comment / File comment rows - so the UI can
    // render those runs as clickable links instead of dead text.
    //
    // Deliberately conservative about what counts as a link. These fields are repair notes, not
    // markup: they are full of part numbers, net names, file names and measurements, and turning
    // "74LS08.pin3" or "5.0V" into a link would be worse than not linking at all. So only three
    // shapes are recognised, all of which a user typing a URL actually produces:
    //
    //   http://...   https://...     an explicit scheme, unambiguous
    //   www.<host>...                the other form people type by habit
    //
    // A bare "example.com" is NOT a link - that is the shape ordinary prose collides with (file
    // names, version numbers, "1.5mm"), and there is no way to tell one from the other without
    // guessing. Users who want a link on a bare domain can type the scheme or the "www.".
    //
    // Only http/https are produced. mailto: and local file paths are deliberately out of scope: the
    // request was for a link that "opens the web page", and ExternalTargetLauncher's local-file path
    // has its own containment and extension rules that a substring pulled out of prose has no
    // business feeding.
    //
    // Everything here is pure string work - no Avalonia, no I/O - so it is unit tested directly.
    // ###########################################################################################
    public static class TextLinkFinder
    {
        private const string HttpScheme = "http://";

        private const string HttpsScheme = "https://";

        private const string WwwPrefix = "www.";

        // ###########################################################################################
        // Trailing characters stripped from a detected link, because sentence punctuation sits
        // directly against a URL far more often than a URL genuinely ends in one:
        // "see https://example.com/page." should link the page, not the page-plus-full-stop.
        // ###########################################################################################
        private const string TrailingPunctuation = ".,;:!?'\"";

        // ###########################################################################################
        // Splits text into ordered spans covering it completely: plain runs and link runs.
        //
        // The spans always concatenate back to the input exactly - a caller renders them in order
        // and gets the original string on screen, with only the link runs styled differently. Blank
        // or null input yields no spans at all (there is nothing to render).
        // ###########################################################################################
        public static IReadOnlyList<TextLinkSpan> FindSpans(string? text)
        {
            var spans = new List<TextLinkSpan>();

            if (string.IsNullOrEmpty(text))
            {
                return spans;
            }

            int position = 0;

            while (position < text.Length)
            {
                if (!TryFindNextLink(text, position, out int linkStart, out int linkLength, out string? url))
                {
                    break;
                }

                if (linkStart > position)
                {
                    spans.Add(new TextLinkSpan(position, linkStart - position, null));
                }

                spans.Add(new TextLinkSpan(linkStart, linkLength, url));
                position = linkStart + linkLength;
            }

            if (position < text.Length)
            {
                spans.Add(new TextLinkSpan(position, text.Length - position, null));
            }

            return spans;
        }

        // ###########################################################################################
        // True when the text contains at least one link. Lets a caller skip the whole span-rendering
        // path - which costs an Inlines collection per block - for the overwhelmingly common case of
        // a note with no URL in it.
        // ###########################################################################################
        public static bool ContainsLink(string? text) =>
            TryFindNextLink(text ?? string.Empty, 0, out _, out _, out _);

        // ###########################################################################################
        // Locates the next link at or after startIndex. Candidates are found by scanning for the
        // three recognised prefixes; the match must begin at a word boundary, so the "http://"
        // inside "xhttp://y" is not a link.
        // ###########################################################################################
        private static bool TryFindNextLink(
            string text,
            int startIndex,
            out int linkStart,
            out int linkLength,
            out string? url)
        {
            linkStart = -1;
            linkLength = 0;
            url = null;

            for (int i = startIndex; i < text.Length; i++)
            {
                if (!IsAtWordBoundary(text, i))
                {
                    continue;
                }

                string? scheme = MatchPrefix(text, i);
                if (scheme == null)
                {
                    continue;
                }

                int end = FindLinkEnd(text, i);
                int length = end - i;

                // A bare prefix with nothing after it ("https://" typed alone, or "www." at the end
                // of a sentence) is not a link - there is no host to open.
                if (length <= scheme.Length)
                {
                    continue;
                }

                string matched = text.Substring(i, length);

                linkStart = i;
                linkLength = length;
                url = string.Equals(scheme, WwwPrefix, StringComparison.OrdinalIgnoreCase)
                    ? HttpsScheme + matched
                    : matched;
                return true;
            }

            return false;
        }

        // ###########################################################################################
        // Which of the recognised prefixes starts at index, or null for none. Returns the prefix
        // itself rather than a bool so the caller knows whether a scheme has to be prepended.
        //
        // https:// is tested before http:// deliberately - the shorter is not a prefix of the
        // longer, but keeping the most specific first is the habit that survives someone adding a
        // fourth prefix later.
        // ###########################################################################################
        private static string? MatchPrefix(string text, int index)
        {
            if (StartsWithAt(text, index, HttpsScheme))
            {
                return HttpsScheme;
            }

            if (StartsWithAt(text, index, HttpScheme))
            {
                return HttpScheme;
            }

            if (StartsWithAt(text, index, WwwPrefix))
            {
                return WwwPrefix;
            }

            return null;
        }

        private static bool StartsWithAt(string text, int index, string prefix) =>
            index + prefix.Length <= text.Length &&
            string.Compare(text, index, prefix, 0, prefix.Length, StringComparison.OrdinalIgnoreCase) == 0;

        // ###########################################################################################
        // A link may only START where a word does. Without this, "shttp://x" and the "www." inside
        // "somewww.thing" would both be picked up as links, and the run rendered as clickable would
        // begin mid-word - visibly wrong even before the URL it produced was.
        //
        // Letters and digits block a start; everything else (whitespace, brackets, punctuation, the
        // start of the string) allows one. An underscore blocks too: it reads as part of a word.
        // ###########################################################################################
        private static bool IsAtWordBoundary(string text, int index)
        {
            if (index == 0)
            {
                return true;
            }

            char previous = text[index - 1];
            return !char.IsLetterOrDigit(previous) && previous != '_';
        }

        // ###########################################################################################
        // Where the link ends: at the first whitespace (a URL cannot contain one), then walked back
        // over trailing punctuation that belongs to the sentence rather than to the URL.
        //
        // A closing bracket is only stripped when it is UNBALANCED within the link - the matching
        // opener is what tells the two cases apart. "(https://example.com)" ends at the "m"; but
        // "https://en.wikipedia.org/wiki/Foo_(bar)" keeps its ")", because that bracket has its own
        // opener inside the URL.
        // ###########################################################################################
        private static int FindLinkEnd(string text, int start)
        {
            int end = start;
            while (end < text.Length && !char.IsWhiteSpace(text[end]))
            {
                end++;
            }

            while (end > start)
            {
                char last = text[end - 1];

                if (TrailingPunctuation.IndexOf(last) >= 0)
                {
                    end--;
                    continue;
                }

                if (IsUnbalancedCloser(text, start, end, last))
                {
                    end--;
                    continue;
                }

                break;
            }

            return end;
        }

        // ###########################################################################################
        // True when the closing bracket at the end of the candidate has no matching opener inside
        // it, which means it belongs to the surrounding prose and not to the URL.
        // ###########################################################################################
        private static bool IsUnbalancedCloser(string text, int start, int end, char closer)
        {
            char opener = closer switch
            {
                ')' => '(',
                ']' => '[',
                '}' => '{',
                '>' => '<',
                _ => '\0'
            };

            if (opener == '\0')
            {
                return false;
            }

            int depth = 0;
            for (int i = start; i < end - 1; i++)
            {
                if (text[i] == opener)
                {
                    depth++;
                }
                else if (text[i] == closer)
                {
                    depth--;
                }
            }

            return depth <= 0;
        }
    }
}
