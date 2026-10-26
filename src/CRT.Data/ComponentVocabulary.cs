using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE WORDS CRT 4.0's NEW COMPONENT DATA IS WRITTEN IN (step 2, 2026-10-25) - a pin's direction,
    // where a scope baseline comes from, what kind of stand-in a replacement is, and the usual chip
    // cases.
    //
    // Stored as these words, not as codes: a board file is read in a table too, where "Datasheet"
    // says what it is and "2" does not. So the words here ARE the data - rewording one changes
    // what every board already published holds. Each list is what the component pane offers and what
    // the table's checks know (BoardDataChecks: anything else is a WARNING, never a refusal - an
    // empty value is fine everywhere, as unfinished work is). Compared ignoring capitals, as every
    // other value of a board is.
    // ###########################################################################################
    public static class ComponentVocabulary
    {
        // How long a short description may be: one line of the popup (owner decision, 2026-10-24:
        // "about 100 characters"). Longer is a warning, never a refusal (BoardDataChecks).
        public const int ShortDescriptionLimit = 100;

        // A pin's job, as a hobbyist reads a pinout - KiCad's dozen electrical types fold into these.
        public static IReadOnlyList<string> PinDirections { get; } =
        [
            "Input",
            "Output",
            "Input/output",
            "Power",
            "Ground",
            "Not connected",
            "Other",
        ];

        // Where a scope baseline comes from (the plan: "measured on a known-good board, datasheet,
        // service manual, or other").
        public static IReadOnlyList<string> BaselineSources { get; } =
        [
            "Known-good board",
            "Datasheet",
            "Service manual",
            "Other",
        ];

        // A replacement part fits as it is, or only after something is changed (said in Changes).
        public static IReadOnlyList<string> ReplacementKinds { get; } =
        [
            "Drop-in",
            "Needs changes",
        ];

        // ###########################################################################################
        // The cases the pinout drawing knows (the plan: "DIP-8 to DIP-40, wide DIP, SIP, TO-92,
        // TO-220, or 'other'"). Offered, not enforced - a board may name any case, and one CRT cannot
        // draw is shown as its pin list alone. A DIP's pin count is in its name.
        // ###########################################################################################
        public static IReadOnlyList<string> Packages { get; } =
        [
            "DIP-8",
            "DIP-14",
            "DIP-16",
            "DIP-18",
            "DIP-20",
            "DIP-22",
            "DIP-24",
            "DIP-24 wide",
            "DIP-28",
            "DIP-28 wide",
            "DIP-40",
            "DIP-48",
            "DIP-64",
            "SIP",
            "TO-92",
            "TO-220",
            "Other",
        ];

        // ###########################################################################################
        // How many pins a case has, by its name - what the pane needs to number the pins and draw the
        // case: a DIP's count is in its name ("DIP-40", "DIP-24 wide", "dip 28"), TO-92 and TO-220
        // have three. Null for a case whose count is not in its name - SIP, "Other", anything typed.
        // ###########################################################################################
        public static int? PinCountOf(string? package)
        {
            string text = (package ?? string.Empty).Trim();

            if (string.Equals(text, "TO-92", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "TO-220", StringComparison.OrdinalIgnoreCase))
            {
                return 3;
            }

            if (!text.StartsWith("DIP", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            string rest = text[3..].TrimStart('-', ' ');
            string digits = new(rest.TakeWhile(char.IsAsciiDigit).ToArray());

            return int.TryParse(digits, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int count) &&
                   count is > 0 and <= 128 &&
                   count % 2 == 0
                ? count
                : null;
        }

        // Whether a case is a DIP - two rows of pins, drawn from the top with the notch.
        public static bool IsDualInLine(string? package) =>
            (package ?? string.Empty).Trim().StartsWith("DIP", StringComparison.OrdinalIgnoreCase) &&
            ComponentVocabulary.PinCountOf(package) is not null;

        // Whether `value` is one of `words` - or empty, which every one of these allows.
        public static bool IsKnownOrEmpty(IReadOnlyList<string> words, string? value)
        {
            ArgumentNullException.ThrowIfNull(words);

            string trimmed = (value ?? string.Empty).Trim();

            return trimmed.Length == 0 || words.Contains(trimmed, StringComparer.OrdinalIgnoreCase);
        }
    }
}
