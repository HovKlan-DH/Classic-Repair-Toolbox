using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // One credit the Submit dialog offers for what a draft changes: its category ("what") and its
    // sub-category ("which part of it" - the pictures' names for a schematic image).
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
    // PURE: two boards in, words out. ToAdd leaves out a credit the board already gives that name.
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

            bool picturesRemoved = CreditSuggestions.Differ(before.Schematics, draft.Schematics, row => string.Join('\u001F', CreditSuggestions.Key(row.SchematicName), CreditSuggestions.Key(row.SchematicImageFile)));

            if (pictures.Count > 0 || picturesRemoved)
            {
                suggestions.Add(new CreditSuggestion(CreditSuggestions.Spelled(draft, CreditSuggestions.SchematicImage), string.Join(", ", pictures)));
            }

            if (CreditSuggestions.Differ(before.ComponentHighlights, draft.ComponentHighlights, row => string.Join('\u001F', CreditSuggestions.Key(row.SchematicName), CreditSuggestions.Key(row.BoardLabel), row.X?.Trim(), row.Y?.Trim(), row.Width?.Trim(), row.Height?.Trim())))
            {
                suggestions.Add(new CreditSuggestion(CreditSuggestions.Spelled(draft, CreditSuggestions.BoardLabelling), string.Empty));
            }

            if (CreditSuggestions.Differ(
                    before.ComponentImages.Where(ComponentImageKinds.IsScopeBaseline),
                    draft.ComponentImages.Where(ComponentImageKinds.IsScopeBaseline),
                    CreditSuggestions.ImageKey))
            {
                suggestions.Add(new CreditSuggestion(CreditSuggestions.Spelled(draft, CreditSuggestions.ScopeBaselines, "Oscilloscope baseline"), string.Empty));
            }

            if (CreditSuggestions.Differ(before.ComponentPins, draft.ComponentPins, row => string.Join('\u001F', CreditSuggestions.Key(row.BoardLabel), row.Region?.Trim(), row.Pin?.Trim(), row.Name?.Trim(), row.Direction?.Trim())))
            {
                suggestions.Add(new CreditSuggestion(CreditSuggestions.Spelled(draft, CreditSuggestions.ComponentPinouts), string.Empty));
            }

            if (CreditSuggestions.BoardDataDiffers(before, draft))
            {
                suggestions.Add(new CreditSuggestion(CreditSuggestions.Spelled(draft, CreditSuggestions.BoardData), string.Empty));
            }

            return suggestions;
        }

        // ###########################################################################################
        // The credit rows for `picked`, naming `name` with `contact` - each left out when the board
        // already credits that name with that category (capitals and spaces ignored). None for a
        // blank name.
        // ###########################################################################################
        public static List<CreditEntry> ToAdd(BoardData draft, IEnumerable<CreditSuggestion> picked, string? name, string? contact)
        {
            ArgumentNullException.ThrowIfNull(draft);
            ArgumentNullException.ThrowIfNull(picked);

            string wantedName = name?.Trim() ?? string.Empty;

            if (wantedName.Length == 0)
            {
                return [];
            }

            return picked
                .Where(suggestion => !draft.Credits.Any(credit =>
                    string.Equals(credit.Category?.Trim(), suggestion.Category.Trim(), StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(credit.NameOrHandle?.Trim(), wantedName, StringComparison.OrdinalIgnoreCase)))
                .Select(suggestion => new CreditEntry
                {
                    Category = suggestion.Category.Trim(),
                    SubCategory = suggestion.SubCategory.Trim(),
                    NameOrHandle = wantedName,
                    Contact = contact?.Trim() ?? string.Empty,
                })
                .ToList();
        }

        // Everything "Board data" covers: the components and their pictures, files, links and
        // replacement parts, the board's own files and links, its important signals and its pictures'
        // colours.
        private static bool BoardDataDiffers(BoardData before, BoardData draft) =>
            CreditSuggestions.Differ(before.Components, draft.Components, row => string.Join('\u001F', CreditSuggestions.Key(row.BoardLabel), row.Region?.Trim(), row.FriendlyName?.Trim(), row.TechnicalNameOrValue?.Trim(), row.PartNumber?.Trim(), row.Category?.Trim(), row.Description?.Trim(), row.Package?.Trim(), row.RepairNotes?.Trim())) ||
            CreditSuggestions.Differ(before.ComponentImages.Where(row => !ComponentImageKinds.IsScopeBaseline(row)), draft.ComponentImages.Where(row => !ComponentImageKinds.IsScopeBaseline(row)), CreditSuggestions.ImageKey) ||
            CreditSuggestions.Differ(before.ComponentLocalFiles, draft.ComponentLocalFiles, row => string.Join('\u001F', CreditSuggestions.Key(row.BoardLabel), row.Name?.Trim(), CreditSuggestions.Key(row.File))) ||
            CreditSuggestions.Differ(before.ComponentLinks, draft.ComponentLinks, row => string.Join('\u001F', CreditSuggestions.Key(row.BoardLabel), row.Name?.Trim(), row.Url?.Trim())) ||
            CreditSuggestions.Differ(before.ComponentReplacements, draft.ComponentReplacements, row => string.Join('\u001F', CreditSuggestions.Key(row.BoardLabel), row.Region?.Trim(), row.PartNumber?.Trim(), row.Kind?.Trim(), row.Changes?.Trim(), row.Source?.Trim())) ||
            CreditSuggestions.Differ(before.BoardLocalFiles, draft.BoardLocalFiles, row => string.Join('\u001F', row.Category?.Trim(), row.Name?.Trim(), CreditSuggestions.Key(row.File))) ||
            CreditSuggestions.Differ(before.BoardLinks, draft.BoardLinks, row => string.Join('\u001F', row.Category?.Trim(), row.Name?.Trim(), row.Url?.Trim())) ||
            CreditSuggestions.Differ(before.KiCadImportantSignals, draft.KiCadImportantSignals, row => string.Join('\u001F', row.DisplayName?.Trim(), row.KiCadNetName?.Trim())) ||
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
                row.ThumbnailHighlightOpacity?.Trim());

            var had = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (BoardSchematicEntry row in before)
            {
                had.TryAdd(Picture(row), Settings(row));
            }

            return draft.Any(row => had.TryGetValue(Picture(row), out string? settings) && !string.Equals(settings, Settings(row), StringComparison.Ordinal));
        }

        private static string ImageKey(ComponentImageEntry row) =>
            string.Join('\u001F', CreditSuggestions.Key(row.BoardLabel), row.Region?.Trim(), row.Pin?.Trim(), row.Name?.Trim(), row.ExpectedOscilloscopeReading?.Trim(), CreditSuggestions.Key(row.File), row.Note?.Trim(), row.TimeDiv?.Trim(), row.VoltsDiv?.Trim(), row.TriggerLevelVolts?.Trim(), row.MachineState?.Trim(), row.Source?.Trim());

        // A name or label as rows are matched by - capitals are no change.
        private static string Key(string? value) => value?.Trim().ToUpperInvariant() ?? string.Empty;

        // Whether two lists hold different rows - their order is no change.
        private static bool Differ<T>(IEnumerable<T> before, IEnumerable<T> after, Func<T, string> key) =>
            !before.Select(key).Order(StringComparer.Ordinal).SequenceEqual(after.Select(key).Order(StringComparer.Ordinal), StringComparer.Ordinal);

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
