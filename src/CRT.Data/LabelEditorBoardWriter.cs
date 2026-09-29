using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // One rectangle a label editor session is saving.
    //
    // Moved here when LabelEditorDraftWriter was retired (Phase 6, 2026-09-23) - it has lived
    // beside whichever class consumes it since BoardDataWriter first defined it, and this is now
    // the one thing that builds a board row out of it.
    //
    // Coordinates are DOUBLES here and strings on the board row: the editor works in real
    // geometry, and the conversion (invariant, always) happens at the point of writing - see
    // ApplyHighlights.
    // ###########################################################################################
    internal sealed class LabelEditorSaveRow
    {
        public string SchematicName { get; init; } = string.Empty;
        public string BoardLabel { get; init; } = string.Empty;
        public string Category { get; init; } = string.Empty;
        public string Region { get; init; } = string.Empty;
        public double X { get; init; }
        public double Y { get; init; }
        public double Width { get; init; }
        public double Height { get; init; }
    }

    // ###########################################################################################
    // APPLIES A SCHEMATIC LABEL EDITOR SAVE TO A DRAFT'S BOARD
    // (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
    //
    // *** THIS REPLACES LabelEditorDraftWriter, AND IT IS MOSTLY A DELETION. *** That class turned
    // a save into BoardDraft row deltas, which meant it had to reason about three things this one
    // does not:
    //
    //   - whether each row was Added or Modified (BoardDraftApplier drops an Added row whose key
    //     collides with an official one, so getting it wrong silently lost the edit);
    //   - which official highlights needed a Deleted TOMBSTONE to hide them;
    //   - whether a component already existed officially, to avoid a duplicate drafted row.
    //
    // A draft is now the board itself, so the answer to all three is simply "write the rows you
    // want". Deleting a highlight means it is not in the list. Adding a component means adding it.
    // The distinctions existed only to describe an overlay, and there is no overlay.
    //
    // WHAT DID NOT CHANGE, and must not: the save is a SCHEMATIC-SCOPED REPLACE. The editor edits
    // one schematic's whole set of rectangles at a time, so a rect it no longer holds has been
    // deleted even though nothing said "delete" - and highlights on every OTHER schematic are left
    // completely alone. That was true of the original .xlsx writer, true of the delta writer, and
    // is true here.
    //
    // PURE: takes a board, returns a new board, touches no files. The caller persists it through
    // DraftWorkbookStore.Edit, which supplies the board as it is on disk right now.
    // ###########################################################################################
    internal static class LabelEditorBoardWriter
    {
        // ###########################################################################################
        // The board as it should read after saving one schematic's label editor session.
        //
        // `current` is never mutated - a new BoardData is returned, so a failed write leaves the
        // caller's copy untouched. Every section the editor does not touch is carried across by
        // reference, which is safe because nothing here mutates a list it did not create.
        // ###########################################################################################
        public static BoardData ApplyLabelEditorSave(
            BoardData current,
            string schematicName,
            IReadOnlyList<LabelEditorSaveRow> rows,
            string region)
        {
            ArgumentNullException.ThrowIfNull(current);

            IReadOnlyList<LabelEditorSaveRow> saveRows = rows ?? [];
            string normalizedSchematic = schematicName?.Trim() ?? string.Empty;

            return new BoardData
            {
                RevisionDate = current.RevisionDate,
                HardwareName = current.HardwareName,
                BoardName = current.BoardName,
                Schematics = current.Schematics,
                Components = LabelEditorBoardWriter.ApplyComponents(current, saveRows, region),
                ComponentImages = current.ComponentImages,
                ComponentHighlights = LabelEditorBoardWriter.ApplyHighlights(current, normalizedSchematic, saveRows),
                ComponentLocalFiles = current.ComponentLocalFiles,
                ComponentLinks = current.ComponentLinks,
                BoardLocalFiles = current.BoardLocalFiles,
                BoardLinks = current.BoardLinks,
                Credits = current.Credits,
                KiCadImportantSignals = current.KiCadImportantSignals,
            };
        }

        // ###########################################################################################
        // Adds a Components row for every genuinely NEW board label the editor drew.
        //
        // The "already exists" rule is carried over exactly from the previous two writers: case
        // insensitive, and a blank region on EITHER side is a wildcard. Without the wildcard, a
        // label already known under no particular region would gain a duplicate row every time it
        // was drawn on a region-scoped schematic.
        //
        // An EXISTING component is left completely alone. The label editor's job is rectangles; it
        // must never overwrite a component's friendly name, part number or description, all of
        // which it knows nothing about.
        // ###########################################################################################
        private static List<ComponentEntry> ApplyComponents(
            BoardData current,
            IReadOnlyList<LabelEditorSaveRow> rows,
            string region)
        {
            string normalizedRegion = region?.Trim() ?? string.Empty;

            List<string> newLabels = rows
                .Where(row => !string.IsNullOrWhiteSpace(row.BoardLabel))
                .Select(row => row.BoardLabel.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(label => !LabelEditorBoardWriter.ComponentExists(current, label, normalizedRegion))
                .ToList();

            if (newLabels.Count == 0)
            {
                return current.Components;
            }

            var components = new List<ComponentEntry>(current.Components);

            foreach (string label in newLabels)
            {
                string category = rows
                    .First(row => string.Equals(row.BoardLabel?.Trim(), label, StringComparison.OrdinalIgnoreCase))
                    .Category?.Trim() ?? string.Empty;

                // Into its category in label order, not at the bottom (2026-09-24) - the position
                // is where the component appears in the application's own list. ComponentPlacement.
                components.Insert(
                    ComponentPlacement.InsertionIndex(components, category, label),
                    new ComponentEntry
                    {
                        BoardLabel = label,
                        Category = category,
                        Region = normalizedRegion,
                    });
            }

            return components;
        }

        private static bool ComponentExists(BoardData current, string boardLabel, string targetRegion)
        {
            return current.Components.Any(component =>
            {
                if (!string.Equals(component.BoardLabel?.Trim(), boardLabel, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                string existingRegion = component.Region?.Trim() ?? string.Empty;

                return existingRegion.Length == 0
                    || targetRegion.Length == 0
                    || string.Equals(existingRegion, targetRegion, StringComparison.OrdinalIgnoreCase);
            });
        }

        // ###########################################################################################
        // *** SCHEMATIC-SCOPED REPLACE - the one rule in this file that is not obvious. ***
        //
        // Every highlight belonging to THIS schematic is dropped and rebuilt from the editor's
        // rows; highlights on every other schematic are carried across untouched.
        //
        // That is what makes deleting a rectangle work at all. The editor never says "delete this
        // one" - it hands over the set it now holds, and anything missing from that set has been
        // removed. Merging row by row instead would make a deleted rectangle immortal.
        //
        // A row with no board label is skipped: the natural key for a highlight is
        // schematic + label, so a blank one could never be found, replaced or deleted again.
        // ###########################################################################################
        private static List<ComponentHighlightEntry> ApplyHighlights(
            BoardData current,
            string schematicName,
            IReadOnlyList<LabelEditorSaveRow> rows)
        {
            var highlights = current.ComponentHighlights
                .Where(highlight => !string.Equals(
                    highlight.SchematicName?.Trim(),
                    schematicName,
                    StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (LabelEditorSaveRow row in rows)
            {
                if (string.IsNullOrWhiteSpace(row.BoardLabel))
                {
                    continue;
                }

                highlights.Add(new ComponentHighlightEntry
                {
                    SchematicName = schematicName,
                    BoardLabel = row.BoardLabel.Trim(),

                    // *** INVARIANT CULTURE, and this is the third face of a bug documented twice
                    // elsewhere in this codebase. *** These are stored as strings and parsed back
                    // with an invariant parse. On a comma-decimal machine a culture-formatted
                    // "12.5" becomes "12,5", which then reads back as 125 or fails outright - and
                    // the rectangle lands somewhere else entirely, or nowhere.
                    X = row.X.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    Y = row.Y.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    Width = row.Width.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    Height = row.Height.ToString(System.Globalization.CultureInfo.InvariantCulture),
                });
            }

            return highlights;
        }
    }
}
