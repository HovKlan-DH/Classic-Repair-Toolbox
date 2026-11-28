using System;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // AN ACCOUNT'S DISPLAY NAME - the name others see (owner requests, 2026-11-25, as accounts opened to everybody: "this
    // should be elaborated that this is the display name others will see, and that it needs to be unique for all users", then
    // "allow only normal A-Z + 0-9 and then some common special charatcers like: +-/@.[]() ... In this way we avoid these
    // weird names i foreign languages. Make the max length 30 characters"; and 2026-11-26: spaces and "_" "allow those two
    // also"). Until then a name was 2 to 100 characters of any script, and two accounts could share one.
    //
    // THE ONE RULE, on both sides: the server holds registration, a name changed under "My account" and an accepted invitation
    // to it (AccountRules.ValidateDisplayName), and CRT's account window says the same sentence under the box as it is typed,
    // its boxes taking no more than MaximumLength.
    //
    //   - 2 to 30 characters, counted as kept - trimmed;
    //   - only the letters A-Z (either case), the digits 0-9, spaces, and _ + - / @ . [ ] ( ) - no letters of other
    //     alphabets, so no name can be written in lookalike letters ("Dennis" with a Cyrillic "e");
    //   - a space only between characters, ONE at a time - "Dennis  Hansen" would pass for another "Dennis Hansen";
    //   - UNIQUE, ignoring capitals and the spaces around it: "Dennis" and "dennis" are one name. That half is the server's,
    //     which alone knows the other names, and asks its database (IAccountStore.IsDisplayNameTakenAsync).
    //
    // *** A NAME FROM BEFORE IS KEPT. *** The rule is asked of a name being SET; an account named under the old rule keeps
    // its name until it is changed, and two old accounts may still share one (INSTALLING.md says how to look).
    // ###########################################################################################
    public static class DisplayNames
    {
        public const int MinimumLength = 2;

        public const int MaximumLength = 30;

        // The marks a name may hold besides letters, digits and single spaces - "_" (2026-11-26), then the ones the owner listed
        // first, in his order.
        public const string AllowedMarks = "_+-/@.[]()";

        // ###########################################################################################
        // What a person reads beside the box - in CRT's account window and on the account web pages. Two sentences: what the
        // name is for, and what it may hold.
        // ###########################################################################################
        public const string Explanation = "The name others will see. It must be unique - nobody else can have the same name.";

        public const string AllowedLine = "2 to 30 characters: the letters A-Z, the digits 0-9, single spaces and _ + - / @ . [ ] ( ).";

        public const string Taken = "That display name is taken - choose another.";

        public const string Missing = "Type the display name others will see.";

        // ###########################################################################################
        // What is wrong with `name` as a display name to SET - null when nothing is. One sentence, the first thing wrong:
        // what a person fixes first. Uniqueness is not asked here (see the header). A character outside the rule is followed
        // by the rule itself (AllowedLine) - unless `sayRule` is false, for a place that shows the rule beside the box already
        // (CRT's account window and "My account"), where it would be said twice.
        // ###########################################################################################
        public static string? Problem(string? name, bool sayRule = true)
        {
            string? problem = DisplayNames.ProblemWithRule(name);

            return problem is not null && !sayRule && problem.EndsWith(" " + DisplayNames.AllowedLine, StringComparison.Ordinal)
                ? problem[..^(DisplayNames.AllowedLine.Length + 1)]
                : problem;
        }

        private static string? ProblemWithRule(string? name)
        {
            string kept = name?.Trim() ?? string.Empty;

            if (kept.Length == 0)
            {
                return DisplayNames.Missing;
            }

            // By character as a person counts them (a Rune), so a letter outside the plane of the first 65,536 - an emoji - is
            // said whole, never as half of a pair.
            foreach (System.Text.Rune character in kept.EnumerateRunes())
            {
                if (!character.IsAscii || !DisplayNames.IsAllowed((char)character.Value))
                {
                    return $"A display name cannot hold \"{DisplayNames.Shown(character)}\". " + DisplayNames.AllowedLine;
                }
            }

            if (kept.Contains("  ", StringComparison.Ordinal))
            {
                return "A display name cannot hold two spaces in a row.";
            }

            if (kept.Length < DisplayNames.MinimumLength)
            {
                return $"A display name needs at least {DisplayNames.MinimumLength} characters.";
            }

            if (kept.Length > DisplayNames.MaximumLength)
            {
                return $"A display name can have at most {DisplayNames.MaximumLength} characters - this one has {kept.Length}.";
            }

            return DisplayNames.WordProblem(kept);
        }

        // ###########################################################################################
        // *** THE WORDS WRITTEN BESIDE A NAME ARE NOBODY'S NAME (code review, 2026-11-27, Security). *** With spaces and "()"
        // allowed, "Dennis (Administrator)" read in CRT's collection window exactly as an administrator's comment, and "You" as
        // the owner's own - the window writes a name and its badge as one line. So no name is "You", and none holds the word
        // "Administrator" or "Owner" (whole words, capitals aside - "Ownership" is a word of its own). The comments went on
        // 2026-11-28 (owner decision: accounts for everybody removed), the rule stays: a maintainer's name is still read in lists
        // beside "the administrator".
        // ###########################################################################################
        public const string You = "You";

        private static readonly string[] ReservedWords = ["Administrator", "Owner"];

        private static string? WordProblem(string name)
        {
            if (string.Equals(name, DisplayNames.You, StringComparison.OrdinalIgnoreCase))
                return $"\"{DisplayNames.You}\" is what CRT calls the person reading - choose another display name.";

            foreach (string word in name.Split([' ', '_', '+', '-', '/', '@', '.', '[', ']', '(', ')'], StringSplitOptions.RemoveEmptyEntries))
            {
                if (DisplayNames.ReservedWords.Contains(word, StringComparer.OrdinalIgnoreCase))
                {
                    return $"A display name cannot hold the word \"{word}\" - the pages write it beside the name of whoever it is.";
                }
            }

            return null;
        }

        public static bool IsAllowed(char character) =>
            character is (>= 'A' and <= 'Z') or (>= 'a' and <= 'z') or (>= '0' and <= '9') or ' ' ||
            DisplayNames.AllowedMarks.Contains(character, StringComparison.Ordinal);

        // A character as a sentence can show it - one that cannot be seen (a control character, a zero-width one) by its
        // code, so the sentence never says "cannot hold """.
        private static string Shown(System.Text.Rune character) =>
            System.Text.Rune.IsControl(character) || System.Text.Rune.IsWhiteSpace(character) ||
            System.Text.Rune.GetUnicodeCategory(character) == System.Globalization.UnicodeCategory.Format
                ? $"U+{character.Value:X4}"
                : character.ToString();
    }
}
