using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // ONE VALUE OF A FIELD THE COMPONENTS PICKED DIFFER IN, as the list under the field offers it while the field has the focus
    // (owner requests, 2026-11-09; ComponentPane.Several.cs): the value as stored, as it reads - "(empty)" for none, a note
    // with formatting as its words - and whose it is.
    // ###########################################################################################
    public sealed record ComponentPaneDifferingValue(string Value, string Shown, string Whose)
    {
        // How much of a long value the list shows - a note may run to many lines.
        public const int ShownAtMost = 160;

        // Its words whole, on one line - what typing narrows the list by, past the part shown too (second code review,
        // 2026-11-09: a word after the 160th character of a note found nothing).
        public string Words { get; init; } = string.Empty;

        // ###########################################################################################
        // Each value once, in the order the components shown first have it, with every label having it. `labels` and
        // `values` go together, one value per label (ComponentBulk.ValuesOf); `isMarkup` reads a value as RichNote markup.
        // ###########################################################################################
        public static IReadOnlyList<ComponentPaneDifferingValue> Of(IReadOnlyList<string> labels, IReadOnlyList<string> values, bool isMarkup)
        {
            ArgumentNullException.ThrowIfNull(labels);
            ArgumentNullException.ThrowIfNull(values);

            return labels
                .Zip(values, (label, value) => (Label: label, Value: value ?? string.Empty))
                .GroupBy(each => each.Value, StringComparer.Ordinal)
                .Select(group => new ComponentPaneDifferingValue(
                    group.Key,
                    ComponentPaneDifferingValue.ShownOf(group.Key, isMarkup),
                    string.Join(", ", group.Select(each => each.Label)))
                {
                    Words = ComponentPaneDifferingValue.WordsOf(group.Key, isMarkup),
                })
                .ToList();
        }

        // ###########################################################################################
        // THE LIST NARROWED TO WHAT IS TYPED (owner request, 2026-11-09: "where I can start typing also and then it will filter
        // the list based on that" - as "Board: Links"' "Category" box, FilterMode Contains): the values whose words hold the
        // text typed, capitals aside; nothing typed, all of them. "(empty)" says no value, so typing never finds it.
        // ###########################################################################################
        public static IReadOnlyList<ComponentPaneDifferingValue> Matching(IReadOnlyList<ComponentPaneDifferingValue> values, string? typed)
        {
            ArgumentNullException.ThrowIfNull(values);

            string wanted = (typed ?? string.Empty).Trim();

            return wanted.Length == 0
                ? values
                : values.Where(value => value.Value.Trim().Length > 0 && value.Words.Contains(wanted, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        private static string ShownOf(string value, bool isMarkup)
        {
            string words = ComponentPaneDifferingValue.WordsOf(value, isMarkup);

            if (words.Length == 0)
            {
                return "(empty)";
            }

            return words.Length > ComponentPaneDifferingValue.ShownAtMost
                ? words[..ComponentPaneDifferingValue.ShownAtMost].TrimEnd() + "..."
                : words;
        }

        // A value's words, its lines joined by a space - a note's without its formatting.
        private static string WordsOf(string value, bool isMarkup)
        {
            string words = (isMarkup ? RichNote.PlainOf(value) : value).Trim();

            return string.Join(" ", words.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries).Select(line => line.Trim()));
        }
    }
}
