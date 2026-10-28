using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A name CRT already has, as a typed name may be found among them: what is typed is compared with
    // `Says` - the name itself, or another way of writing it (the hardware's folder, "C64" for
    // "Commodore 64") - and `Name` is what is offered.
    // ###########################################################################################
    public sealed record KnownName(string Says, string Name);

    // ###########################################################################################
    // "DID YOU MEAN COMMODORE?" (CRT 4.0 step 3, slice 3e - the plan: "Name: 'Did you mean Commodore?'
    // when a typed name is close to an existing one"). The names of a new board are free text, and a
    // name written another way than the one CRT has makes a SECOND manufacturer or machine beside it
    // in the drop-down lists - "Comodore" beside "Commodore", "C64" beside "Commodore 64" - and, for a
    // manufacturer, a second top folder ("commodore" beside "Commodore" on a file system that tells
    // capitals apart).
    //
    // A typed name is close to a known one when, read as letters, digits and "+" only and without
    // regard to capitals, it is the same ("commodore", "C-64"), or a slip away: one letter wrong,
    // missing, extra or two swapped for a name of 4 to 7 letters, two for a longer one (Damerau-
    // Levenshtein) - never a slip that changes the model, its numbers, "+" signs and the letters after
    // the last digit ("Commodore 65", "C128D" and "Amiga 500+" are other machines than "Commodore 64",
    // "C128" and "Amiga 500"), and never in three letters or fewer ("C16" is not "C64"). A name typed EXACTLY as CRT has it
    // asks nothing; two known names equally close ask nothing either - a guess would be wrong half
    // the time.
    //
    // PURE: the typed name and the known ones in, the name to offer out (null for none).
    // ###########################################################################################
    public static class SimilarNames
    {
        public static string? DidYouMean(string? typed, IEnumerable<KnownName> known)
        {
            ArgumentNullException.ThrowIfNull(known);

            string text = typed?.Trim() ?? string.Empty;
            string letters = SimilarNames.LettersOf(text);

            if (letters.Length == 0)
            {
                return null;
            }

            List<KnownName> names = known
                .Where(name => !string.IsNullOrWhiteSpace(name.Name) && !string.IsNullOrWhiteSpace(name.Says))
                .ToList();

            // Written exactly as CRT has it: nothing to ask.
            if (names.Any(name => string.Equals(name.Name.Trim(), text, StringComparison.Ordinal)))
            {
                return null;
            }

            int allowed = letters.Length <= 3 ? 0 : letters.Length <= 7 ? 1 : 2;
            string model = SimilarNames.ModelOf(letters);

            // A slip never changes the model: "Commodore 65" is another machine than "Commodore 64",
            // "C128D" than "C128", "Amiga 500+" than "Amiga 500".
            List<(string Name, int Distance)> close = names
                .Select(name => (Name: name.Name.Trim(), Says: SimilarNames.LettersOf(name.Says)))
                .Select(name => (name.Name, Distance: name.Says == letters
                    ? 0
                    : SimilarNames.ModelOf(name.Says) == model ? SimilarNames.Distance(letters, name.Says, allowed) : allowed + 1))
                .Where(candidate => candidate.Distance <= allowed)
                .ToList();

            if (close.Count == 0)
            {
                return null;
            }

            int best = close.Min(candidate => candidate.Distance);

            List<string> nearest = close
                .Where(candidate => candidate.Distance == best)
                .Select(candidate => candidate.Name)
                .Distinct(StringComparer.Ordinal)
                .ToList();

            return nearest.Count == 1 ? nearest[0] : null;
        }

        // What "Did you mean" says, and its button.
        public static string Question(string name) => $"Did you mean {name}?";

        public static string UseButton(string name) => $"Use {name}";

        // A name as letters, digits and "+" - "+" is a model, as in "Amiga 500+" (code review, 2026-10-28)
        // - in small letters.
        private static string LettersOf(string? name) =>
            new((name ?? string.Empty).Where(c => char.IsLetterOrDigit(c) || c == '+').Select(char.ToLowerInvariant).ToArray());

        // ###########################################################################################
        // What makes the model: its numbers with their "+" signs, and the letters after the last digit
        // - "C128D" is "128|d", "Spectrum +2A" "+2|a", "Commodore 64" "64|". A slip may change anything
        // else (code review, 2026-10-28: "C128D" was a one-letter slip from "C128").
        // ###########################################################################################
        private static string ModelOf(string letters)
        {
            int lastDigit = letters.LastIndexOfAny("0123456789".ToCharArray());
            string numbers = new(letters.Where(c => char.IsDigit(c) || c == '+').ToArray());

            return lastDigit < 0 ? numbers : $"{numbers}|{letters[(lastDigit + 1)..]}";
        }

        // ###########################################################################################
        // The optimal-string-alignment distance: letters changed, left out, put in, and two side by side
        // swapped. Anything past `limit` is limit + 1 - the exact figure does not matter then.
        // ###########################################################################################
        private static int Distance(string a, string b, int limit)
        {
            if (Math.Abs(a.Length - b.Length) > limit)
            {
                return limit + 1;
            }

            var d = new int[a.Length + 1, b.Length + 1];

            for (int i = 0; i <= a.Length; i++)
            {
                d[i, 0] = i;
            }

            for (int j = 0; j <= b.Length; j++)
            {
                d[0, j] = j;
            }

            for (int i = 1; i <= a.Length; i++)
            {
                for (int j = 1; j <= b.Length; j++)
                {
                    int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);

                    if (i > 1 && j > 1 && a[i - 1] == b[j - 2] && a[i - 2] == b[j - 1])
                    {
                        d[i, j] = Math.Min(d[i, j], d[i - 2, j - 2] + 1);
                    }
                }
            }

            return Math.Min(d[a.Length, b.Length], limit + 1);
        }
    }
}
