using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;

namespace Handlers.DataHandling
{
    // One kind of work a draft did, as a credit names it: its category ("what") and its sub-category
    // ("which part of it" - the pictures' names for a schematic image). The credit for a draft's work
    // covers them all (SubmissionCredit.Offer).
    public sealed record CreditSuggestion(string Category, string SubCategory);

    // ###########################################################################################
    // THE CREDITS A DRAFT'S WORK DESERVES (CRT 4.0 step 2, 2026-10-26 - the plan: "After a submission,
    // CRT suggests a credit for what you sent"; owner decision the same day: "Do it now").
    //
    // What the draft changes against the published board, put into the categories the board's credits
    // already use (measured on the live trees: "Schematic image", "Board labelling", "Board data",
    // "Oscilloscope baseline measurement", "Component pinouts"):
    //
    //   - pictures of the board added, renamed or replaced      -> "Schematic image", which ones by name
    //   - components marked on the pictures (the highlights)    -> "Board labelling"
    //   - scope baselines                                       -> "Oscilloscope baseline measurement"
    //   - pins                                                  -> "Component pinouts"
    //   - anything else of the components or of the board's own lists -> "Board data"
    //
    // A change of the credits alone, the KiCad data or a calibration deserves none. A category the
    // board already spells otherwise ("Oscilloscope baseline") is offered in its spelling, so a credit
    // joins its group.
    //
    // PURE: two boards in, words out. SubmissionCredit makes them the credit for the draft's work, a row of
    // "Board: Credits".
    // ###########################################################################################
    public static class CreditSuggestions
    {
        public const string SchematicImage = "Schematic image";
        public const string BoardLabelling = "Board labelling";
        public const string BoardData = "Board data";
        public const string ScopeBaselines = "Oscilloscope baseline measurement";
        public const string ComponentPinouts = "Component pinouts";

        // ###########################################################################################
        // The credits `draft`'s changes against `published` deserve - every category once, in the
        // order above. `published` null: a new board, every part of it the draft's work.
        // ###########################################################################################
        public static List<CreditSuggestion> For(BoardData? published, BoardData draft)
        {
            ArgumentNullException.ThrowIfNull(draft);

            BoardData before = published ?? new BoardData();
            var suggestions = new List<CreditSuggestion>();

            // Pictures: by name and image file - a picture's highlight colours are the board's data.
            List<string> pictures = CreditSuggestions.Added(before.Schematics, draft.Schematics, row => string.Join('\u001F', CreditSuggestions.Key(row.SchematicName), CreditSuggestions.Key(row.SchematicImageFile)))
                .Select(row => row.SchematicName?.Trim() ?? string.Empty)
                .Where(name => name.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            bool picturesRemoved = CreditSuggestions.Differ(before.Schematics, draft.Schematics, "schematic images", row => string.Join('\u001F', CreditSuggestions.Key(row.SchematicName), CreditSuggestions.Key(row.SchematicImageFile)));

            if (pictures.Count > 0 || picturesRemoved)
            {
                suggestions.Add(new CreditSuggestion(CreditSuggestions.Spelled(draft, CreditSuggestions.SchematicImage), string.Join(", ", pictures)));
            }

            if (CreditSuggestions.Differ(before.ComponentHighlights, draft.ComponentHighlights, "highlights", row => string.Join('\u001F', CreditSuggestions.Key(row.SchematicName), CreditSuggestions.Key(row.BoardLabel), row.X?.Trim(), row.Y?.Trim(), row.Width?.Trim(), row.Height?.Trim())))
            {
                suggestions.Add(new CreditSuggestion(CreditSuggestions.Spelled(draft, CreditSuggestions.BoardLabelling), string.Empty));
            }

            if (CreditSuggestions.Differ(before.ComponentImages, draft.ComponentImages, "baselines", CreditSuggestions.ImageKey, ComponentImageKinds.IsScopeBaseline))
            {
                suggestions.Add(new CreditSuggestion(CreditSuggestions.Spelled(draft, CreditSuggestions.ScopeBaselines, "Oscilloscope baseline"), string.Empty));
            }

            if (CreditSuggestions.Differ(before.ComponentPins, draft.ComponentPins, "pins", row => string.Join('\u001F', CreditSuggestions.Key(row.BoardLabel), row.Region?.Trim(), row.Pin?.Trim(), row.Name?.Trim(), row.Direction?.Trim())))
            {
                suggestions.Add(new CreditSuggestion(CreditSuggestions.Spelled(draft, CreditSuggestions.ComponentPinouts), string.Empty));
            }

            if (CreditSuggestions.BoardDataDiffers(before, draft))
            {
                suggestions.Add(new CreditSuggestion(CreditSuggestions.Spelled(draft, CreditSuggestions.BoardData), string.Empty));
            }

            return suggestions;
        }

        // Who a credit names when its contributor wants none (owner request, 2026-11-15: "it should just state "Unknown" for the
        // name") - typed as the name of the credit for the work ("Board: Credits" says so), and never remembered as one.
        public const string UnknownName = "Unknown";

        // Everything "Board data" covers: the components and their pictures, files, links and
        // replacement parts, the board's own files and links, its important signals and its pictures'
        // colours.
        private static bool BoardDataDiffers(BoardData before, BoardData draft) =>
            CreditSuggestions.Differ(before.Components, draft.Components, "components", row => string.Join('\u001F', CreditSuggestions.Key(row.BoardLabel), row.Region?.Trim(), row.FriendlyName?.Trim(), row.TechnicalNameOrValue?.Trim(), row.PartNumber?.Trim(), row.Category?.Trim(), row.Description?.Trim(), row.Package?.Trim(), RichNote.TrimMarkup(row.Purpose))) ||
            CreditSuggestions.Differ(before.ComponentImages, draft.ComponentImages, "pictures", CreditSuggestions.ImageKey, row => !ComponentImageKinds.IsScopeBaseline(row)) ||
            CreditSuggestions.Differ(before.ComponentLocalFiles, draft.ComponentLocalFiles, "component files", row => string.Join('\u001F', CreditSuggestions.Key(row.BoardLabel), row.Region?.Trim(), row.Name?.Trim(), CreditSuggestions.Key(row.File))) ||
            CreditSuggestions.Differ(before.ComponentLinks, draft.ComponentLinks, "component links", row => string.Join('\u001F', CreditSuggestions.Key(row.BoardLabel), row.Region?.Trim(), row.Name?.Trim(), row.Url?.Trim())) ||
            CreditSuggestions.Differ(before.ComponentReplacements, draft.ComponentReplacements, "replacements", row => string.Join('\u001F', CreditSuggestions.Key(row.BoardLabel), row.Region?.Trim(), row.PartNumber?.Trim(), row.Kind?.Trim(), row.Note?.Trim())) ||
            // A board file's and link's note is work too (code review, 2026-11-08: left out, a draft writing only notes was
            // offered no credit) - plain text, kept as ComponentPaneEdits.NormalisedNote keeps it.
            CreditSuggestions.Differ(before.BoardLocalFiles, draft.BoardLocalFiles, "board files", row => string.Join('\u001F', row.Category?.Trim(), row.Region?.Trim(), row.Name?.Trim(), CreditSuggestions.Key(row.File), ComponentPaneEdits.NormalisedNote(row.Note))) ||
            CreditSuggestions.Differ(before.BoardLinks, draft.BoardLinks, "board links", row => string.Join('\u001F', row.Category?.Trim(), row.Region?.Trim(), row.Name?.Trim(), row.Url?.Trim(), ComponentPaneEdits.NormalisedNote(row.Note))) ||
            CreditSuggestions.Differ(before.KiCadImportantSignals, draft.KiCadImportantSignals, "signals", row => string.Join('\u001F', row.DisplayName?.Trim(), row.KiCadNetName?.Trim())) ||
            CreditSuggestions.PictureSettingsDiffer(before.Schematics, draft.Schematics);

        // ###########################################################################################
        // A picture the board had - same name, same image file - with other settings: its CAD name or
        // its highlight colours. A picture added, renamed or replaced is a schematic image's work, not
        // this.
        // ###########################################################################################
        private static bool PictureSettingsDiffer(IEnumerable<BoardSchematicEntry> before, IEnumerable<BoardSchematicEntry> draft)
        {
            static string Picture(BoardSchematicEntry row) => string.Join('\u001F', CreditSuggestions.Key(row.SchematicName), CreditSuggestions.Key(row.SchematicImageFile));

            static string Settings(BoardSchematicEntry row) => string.Join(
                '\u001F',
                row.CadName?.Trim(),
                row.SchematicHighlightColor?.Trim(),
                row.SchematicHighlightOpacity?.Trim(),
                row.OppositeTraceHighlightColor?.Trim(),
                row.ThumbnailHighlightColor?.Trim(),
                row.ThumbnailHighlightOpacity?.Trim(),
                row.TraceHighlightColor?.Trim(),
                row.TraceHighlightOpacity?.Trim(),
                row.OppositeTraceHighlightOpacity?.Trim());

            var had = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (BoardSchematicEntry row in before)
            {
                had.TryAdd(Picture(row), Settings(row));
            }

            return draft.Any(row => had.TryGetValue(Picture(row), out string? settings) && !string.Equals(settings, Settings(row), StringComparison.Ordinal));
        }

        // The note as notes are compared (BoardDataDiffer, [RichNoteMarkup]): a note nested a level deeper is work too.
        private static string ImageKey(ComponentImageEntry row) =>
            string.Join('\u001F', CreditSuggestions.Key(row.BoardLabel), row.Region?.Trim(), row.Pin?.Trim(), row.Name?.Trim(), row.ExpectedOscilloscopeReading?.Trim(), CreditSuggestions.Key(row.File), RichNote.NormalisedMarkup(row.Note), row.TimeDiv?.Trim(), row.VoltsDiv?.Trim(), row.TriggerLevelVolts?.Trim(), row.MachineState?.Trim());

        // A name or label as rows are matched by - capitals are no change.
        private static string Key(string? value) => value?.Trim().ToUpperInvariant() ?? string.Empty;

        // ###########################################################################################
        // Whether two lists hold different rows - their order is no change.
        //
        // *** ITS KEYS WORKED OUT ONCE PER LIST (2026-11-16, performance walkthrough). ***
        // Every change in the component pane asks what the draft's work is (SubmissionCredit.Follow), and each ask built and
        // sorted the keys of every row of both boards - the draft's base, the same at every ask, and the draft, whose sections
        // an edit mostly leaves as the very lists they were. Nothing changes a board it was given (BoardData), so a list's
        // keys are kept while the list lives - `kind` naming the key and `only` the rows it keeps - and the very same list on
        // both sides holds the same rows. A list whose length is not the one its keys were worked out at is worked out again.
        // ###########################################################################################
        private static bool Differ<T>(IReadOnlyList<T> before, IReadOnlyList<T> after, string kind, Func<T, string> key, Func<T, bool>? only = null) =>
            !ReferenceEquals(before, after) &&
            !CreditSuggestions.SortedKeys(before, kind, key, only).AsSpan().SequenceEqual(CreditSuggestions.SortedKeys(after, kind, key, only), StringComparer.Ordinal);

        private static string[] SortedKeys<T>(IReadOnlyList<T> rows, string kind, Func<T, string> key, Func<T, bool>? only)
        {
            ConcurrentDictionary<string, (int Count, string[] Keys)> byKind =
                CreditSuggestions.KeysByList.GetValue(rows, _ => new ConcurrentDictionary<string, (int Count, string[] Keys)>(StringComparer.Ordinal));

            if (byKind.TryGetValue(kind, out (int Count, string[] Keys) known) && known.Count == rows.Count)
            {
                return known.Keys;
            }

            string[] keys = (only is null ? rows : rows.Where(only)).Select(key).Order(StringComparer.Ordinal).ToArray();
            byKind[kind] = (rows.Count, keys);
            Interlocked.Increment(ref CreditSuggestions.KeysWorkedOut);

            return keys;
        }

        private static readonly ConditionalWeakTable<object, ConcurrentDictionary<string, (int Count, string[] Keys)>> KeysByList = new();

        private static int KeysWorkedOut;

        // How many lists' keys were worked out - for tests.
        internal static int KeysWorkedOutForTests => Volatile.Read(ref CreditSuggestions.KeysWorkedOut);

        // The rows of `after` that `before` does not have, in `after`'s order.
        private static IEnumerable<T> Added<T>(IEnumerable<T> before, IEnumerable<T> after, Func<T, string> key)
        {
            var had = before.Select(key).ToHashSet(StringComparer.Ordinal);
            return after.Where(row => !had.Contains(key(row)));
        }

        // `category` as the board's credits already spell it - or one of `alsoKnownAs` it uses - else as given.
        private static string Spelled(BoardData board, string category, params string[] alsoKnownAs) =>
            board.Credits
                .Select(credit => credit.Category?.Trim() ?? string.Empty)
                .FirstOrDefault(used => string.Equals(used, category, StringComparison.OrdinalIgnoreCase) ||
                    alsoKnownAs.Any(other => string.Equals(used, other, StringComparison.OrdinalIgnoreCase)))
            is { Length: > 0 } spelled
                ? spelled
                : category;
    }
}
