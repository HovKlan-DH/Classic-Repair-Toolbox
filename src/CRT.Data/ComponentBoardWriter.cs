using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // APPLIES A COMPONENT-CONTRIBUTION EDITING SESSION TO A DRAFT'S BOARD
    // (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
    //
    // *** THIS REPLACES ComponentDraftWriter, AND LIKE THE LABEL EDITOR'S IT IS MOSTLY A DELETION.
    // *** That class turned one editing session into BoardDraft row deltas across six sections,
    // and had to decide Added-vs-Modified per row (against the PURE official board, since a merged
    // view collapses the distinction) and write Deleted tombstones for rows the contributor had
    // removed from the window's collections.
    //
    // A draft is now the board itself, so all of that becomes one rule: the window loads every row
    // belonging to this component before editing starts, so the rows it hands back ARE that
    // component's rows. Anything missing was removed.
    //
    // *** WITH ONE EXCEPTION, AND IT COST 40 FILES BEFORE IT WAS NOTICED (2026-09-23). ***
    // COMPONENT IMAGES are loaded BY REGION, not wholesale - the window's own heading says
    // "Component images relevant for the PAL region". So for that section, and only that section,
    // "anything missing was removed" is FALSE: the other region's rows were never on screen. They
    // are preserved explicitly (see ReplaceForComponentInRegion and `editedRegion`).
    //
    // Before assuming a section can be replaced wholesale, check how LoadComponent filters it.
    // Components, component local files and component links are all loaded unfiltered, so for
    // those the rule holds as written.
    //
    // *** THE SAVE IS COMPONENT-SCOPED, which is this file's equivalent of the label editor's
    // schematic-scoped replace. *** Rows belonging to THIS board label are dropped and rebuilt;
    // every other component's rows are carried across untouched. Getting that boundary wrong would
    // wipe the rest of the board from a window the contributor believes is editing one component.
    //
    // *** AND IT KEEPS POSITIONS (owner request, 2026-09-24). *** An edited component's rows
    // go back where they were; a new component's go into its category in label order
    // (ComponentPlacement). They used to be appended at the bottom on every save, so editing a
    // component moved it to the end of the list the application shows.
    //
    // BOARD-SCOPED SECTIONS (BoardLocalFiles, BoardLinks) belong to the whole board rather than to
    // a component. The window shows all of them, but the board on DISK may have moved since it
    // opened - so they are MERGED against what the window loaded, not replaced (see
    // MergeBoardScoped; code review, 2026-09-25).
    //
    // PURE: takes a board, returns a new board, touches no files.
    // ###########################################################################################
    public static class ComponentBoardWriter
    {
        // ###########################################################################################
        // The board as it should read after saving one component-contribution session.
        //
        // `boardLabel` is the EFFECTIVE label the session belongs to - component-scoped rows with a
        // blank label inherit it, matching the window's own ResolveEffectiveBoardLabel and what
        // ComponentDraftWriter did before.
        // ###########################################################################################
        public static BoardData ApplyComponentSave(
            BoardData current,
            string boardLabel,
            IReadOnlyList<ComponentDraftWriter.ComponentDraftRow> componentRows,
            IReadOnlyList<ComponentDraftWriter.ComponentImageDraftRow> componentImageRows,
            IReadOnlyList<ComponentDraftWriter.LocalFileDraftRow> componentLocalFileRows,
            IReadOnlyList<ComponentDraftWriter.LinkDraftRow> componentLinkRows,
            IReadOnlyList<ComponentDraftWriter.LocalFileDraftRow> boardLocalFileRows,
            IReadOnlyList<ComponentDraftWriter.LinkDraftRow> boardLinkRows,
            string editedRegion,

            // ###########################################################################################
            // The board-scoped rows AS THE WINDOW LOADED THEM, built through the same row builders
            // as the edited rows so the two compare like for like.
            //
            // *** WITHOUT THESE, AN EDIT MADE ELSEWHERE WHILE THE WINDOW WAS OPEN IS REVERTED. ***
            // The window loads every board-level file and link when it opens, and these sections
            // used to be replaced wholesale with its rows on save. Excel is the one writer the app
            // cannot hold back: a board link added there while the window sat open was deleted by
            // the next "Save to draft", even though that save only meant to change one component.
            // Null (a caller that has no baseline) keeps the old wholesale replace.
            // ###########################################################################################
            IReadOnlyList<ComponentDraftWriter.LocalFileDraftRow>? boardLocalFileRowsAtOpen = null,
            IReadOnlyList<ComponentDraftWriter.LinkDraftRow>? boardLinkRowsAtOpen = null)
        {
            ArgumentNullException.ThrowIfNull(current);

            string label = boardLabel?.Trim() ?? string.Empty;
            string region = editedRegion?.Trim() ?? string.Empty;

            return new BoardData
            {
                RevisionDate = current.RevisionDate,
                Schematics = current.Schematics,

                // In place when the component already exists, otherwise in its category in
                // label order - see ComponentPlacement for why the position matters at all.
                Components = ComponentPlacement.ReplaceComponent(
                    current.Components,
                    label,
                    (componentRows ?? []).Select(row => new ComponentEntry
                    {
                        BoardLabel = ComponentBoardWriter.Or(row.BoardLabel, label),
                        FriendlyName = row.FriendlyName?.Trim() ?? string.Empty,
                        TechnicalNameOrValue = row.TechnicalNameOrValue?.Trim() ?? string.Empty,
                        PartNumber = row.PartNumber?.Trim() ?? string.Empty,
                        Category = row.Category?.Trim() ?? string.Empty,
                        Region = row.Region?.Trim() ?? string.Empty,
                        Description = row.Description?.Trim() ?? string.Empty,
                    }).ToList()),

                // ###########################################################################################
                // *** IMAGES ARE REPLACED WITHIN THE EDITED REGION ONLY, and this is the one
                // section where that matters (fixed 2026-09-23). ***
                //
                // The window loads component images filtered BY REGION - its own heading says
                // "Component images relevant for the PAL region" - while every other section it
                // edits is loaded unfiltered. So the rows it hands back are only ever this
                // region's, and replacing the label's rows wholesale deleted the OTHER region's
                // silently.
                //
                // Reported by the project owner: editing U1's short description while viewing PAL
                // removed all 40 of U1's NTSC scope baselines, and the maintainer app correctly showed
                // 40 files being deleted by a submission that was meant to change one line of text.
                //
                // A row with a BLANK region is visible in every region (the window loads it too),
                // so it is inside the replaced set rather than preserved - otherwise saving twice
                // would duplicate it.
                // ###########################################################################################
                ComponentImages = ComponentBoardWriter.ReplaceForComponentInRegion(
                    current.ComponentImages,
                    label,
                    region,
                    existing => existing.BoardLabel,
                    existing => existing.Region,
                    (componentImageRows ?? []).Select(row => new ComponentImageEntry
                    {
                        BoardLabel = ComponentBoardWriter.Or(row.BoardLabel, label),
                        Region = row.Region?.Trim() ?? string.Empty,
                        Pin = row.Pin?.Trim() ?? string.Empty,
                        Name = row.Name?.Trim() ?? string.Empty,
                        ExpectedOscilloscopeReading = row.ExpectedOscilloscopeReading?.Trim() ?? string.Empty,
                        VoltsDiv = row.VoltsDiv?.Trim() ?? string.Empty,
                        TimeDiv = row.TimeDiv?.Trim() ?? string.Empty,
                        TriggerLevelVolts = row.TriggerLevelVolts?.Trim() ?? string.Empty,
                        Note = row.Note?.Trim() ?? string.Empty,
                        File = ComponentBoardWriter.JoinFilePath(row.FileLocation, row.File),
                    })),

                // ComponentHighlights are the LABEL EDITOR's business, never this window's. A
                // component save must not disturb a single rectangle.
                ComponentHighlights = current.ComponentHighlights,

                ComponentLocalFiles = ComponentPlacement.ReplaceRowsForLabel(
                    current.ComponentLocalFiles,
                    label,
                    existing => existing.BoardLabel,
                    (componentLocalFileRows ?? []).Select(row => new ComponentLocalFileEntry
                    {
                        BoardLabel = ComponentBoardWriter.Or(row.BoardLabel, label),
                        Name = row.Name?.Trim() ?? string.Empty,
                        File = ComponentBoardWriter.JoinFilePath(row.FileLocation, row.File),
                    }).ToList()),

                ComponentLinks = ComponentPlacement.ReplaceRowsForLabel(
                    current.ComponentLinks,
                    label,
                    existing => existing.BoardLabel,
                    (componentLinkRows ?? []).Select(row => new ComponentLinkEntry
                    {
                        BoardLabel = ComponentBoardWriter.Or(row.BoardLabel, label),
                        Name = row.Name?.Trim() ?? string.Empty,
                        Url = row.Url?.Trim() ?? string.Empty,
                    }).ToList()),

                // *** BOARD-SCOPED: MERGED against what the window loaded, not replaced. *** These
                // belong to the whole board, and the workbook may have been edited elsewhere while
                // the window was open - see MergeBoardScoped.
                BoardLocalFiles = ComponentBoardWriter.MergeBoardScoped(
                    current.BoardLocalFiles,
                    boardLocalFileRowsAtOpen?.Select(ComponentBoardWriter.ToBoardLocalFile).ToList(),
                    (boardLocalFileRows ?? []).Select(ComponentBoardWriter.ToBoardLocalFile).ToList(),
                    entry => ComponentBoardWriter.Key(entry.Category, entry.Name, ComponentBoardWriter.NormaliseFile(entry.File))),

                BoardLinks = ComponentBoardWriter.MergeBoardScoped(
                    current.BoardLinks,
                    boardLinkRowsAtOpen?.Select(ComponentBoardWriter.ToBoardLink).ToList(),
                    (boardLinkRows ?? []).Select(ComponentBoardWriter.ToBoardLink).ToList(),
                    entry => ComponentBoardWriter.Key(entry.Category, entry.Name, entry.Url)),

                Credits = current.Credits,
                KiCadImportantSignals = current.KiCadImportantSignals,
            };
        }

        // ###########################################################################################
        // "Delete this component" - the window's own toggle.
        //
        // Every row belonging to this board label goes, across all four component-scoped sections
        // INCLUDING the Components row itself - AND its highlights on every schematic.
        //
        // *** THE HIGHLIGHTS GO TOO (owner decision, 2026-09-25): "if a component really is
        // deleted, then it should remove EVERYTHING related to this component." *** They used to be
        // kept, on the reasoning that highlights are the label editor's concern - which left
        // rectangles naming a component the board no longer has, and contradicted this window's
        // own notice ("will be removed ... together with its N schematic highlights"). The board
        // table's delete does the same (BoardTableDocument).
        //
        // BOARD-SCOPED SECTIONS ARE UNTOUCHED: the board's own files and links are not the
        // component's.
        // ###########################################################################################
        public static BoardData ApplyComponentDelete(BoardData current, string boardLabel)
        {
            ArgumentNullException.ThrowIfNull(current);

            string label = boardLabel?.Trim() ?? string.Empty;

            return new BoardData
            {
                RevisionDate = current.RevisionDate,
                Schematics = current.Schematics,

                Components = ComponentBoardWriter.WithoutComponent(
                    current.Components, label, existing => existing.BoardLabel),

                ComponentImages = ComponentBoardWriter.WithoutComponent(
                    current.ComponentImages, label, existing => existing.BoardLabel),

                ComponentHighlights = ComponentBoardWriter.WithoutComponent(
                    current.ComponentHighlights, label, existing => existing.BoardLabel),

                ComponentLocalFiles = ComponentBoardWriter.WithoutComponent(
                    current.ComponentLocalFiles, label, existing => existing.BoardLabel),

                ComponentLinks = ComponentBoardWriter.WithoutComponent(
                    current.ComponentLinks, label, existing => existing.BoardLabel),

                BoardLocalFiles = current.BoardLocalFiles,
                BoardLinks = current.BoardLinks,
                Credits = current.Credits,
                KiCadImportantSignals = current.KiCadImportantSignals,
            };
        }

        // ###########################################################################################
        // A component-scoped replace confined to ONE REGION - the component images' version of
        // ComponentPlacement.ReplaceRowsForLabel.
        //
        // *** WHERE THE REPLACEMENTS GO (2026-09-24). *** Where the first replaced row was; else
        // just after this component's surviving rows from other regions, so one component's images
        // stay together; else among the other components' rows in label order. They used to be
        // appended at the bottom on every save, which reordered the sheet each time a component
        // was edited - see ComponentPlacement's header.
        //
        // A row is replaced when it belongs to this label AND the window could actually see it:
        // its region matches the edited one, or it is blank (blank means "all regions", and the
        // window loads those too). Every other region's rows are carried across untouched.
        //
        // *** THE REGION IS NOT A WILDCARD WHEN IT IS BLANK ON THE EDITING SIDE. *** A window
        // opened with no region at all can only have loaded the blank-region rows, so it must
        // replace only those. Treating a blank EDITED region as "match everything" would restore
        // exactly the bug this method exists to fix.
        // ###########################################################################################
        private static List<T> ReplaceForComponentInRegion<T>(
            List<T> existing,
            string boardLabel,
            string editedRegion,
            Func<T, string> labelOf,
            Func<T, string> regionOf,
            IEnumerable<T> replacements)
        {
            List<T> replacementRows = [.. replacements];

            if (boardLabel.Length == 0)
            {
                List<T> all = [.. existing];
                all.InsertRange(ComponentPlacement.IndexByLabel(all, boardLabel, labelOf), replacementRows);

                return all;
            }

            var kept = new List<T>(existing.Count + replacementRows.Count);
            int firstReplaced = -1;
            int afterLastSurvivor = -1;

            foreach (T row in existing)
            {
                bool sameComponent = string.Equals(
                    labelOf(row)?.Trim(), boardLabel, StringComparison.OrdinalIgnoreCase);

                if (sameComponent)
                {
                    string rowRegion = regionOf(row)?.Trim() ?? string.Empty;

                    // Visible to the editing window, so it is the window's to replace.
                    bool wasEditable = rowRegion.Length == 0
                        || string.Equals(rowRegion, editedRegion, StringComparison.OrdinalIgnoreCase);

                    if (wasEditable)
                    {
                        if (firstReplaced < 0)
                        {
                            firstReplaced = kept.Count;
                        }

                        continue;
                    }
                }

                kept.Add(row);

                if (sameComponent)
                {
                    afterLastSurvivor = kept.Count;
                }
            }

            int insertAt = firstReplaced >= 0
                ? firstReplaced
                : afterLastSurvivor >= 0
                    ? afterLastSurvivor
                    : ComponentPlacement.IndexByLabel(kept, boardLabel, labelOf);

            kept.InsertRange(insertAt, replacementRows);

            return kept;
        }

        private static List<T> WithoutComponent<T>(
            List<T> existing,
            string boardLabel,
            Func<T, string> labelOf)
        {
            if (boardLabel.Length == 0)
            {
                return existing;
            }

            return existing
                .Where(row => !string.Equals(
                    labelOf(row)?.Trim(),
                    boardLabel,
                    StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        // ###########################################################################################
        // The stored File value from the window's FileLocation + File pair.
        //
        // Carried over VERBATIM from ComponentDraftWriter.BuildStoredFile, so a path stored
        // before this change and one stored after it are byte-identical - otherwise every
        // attachment row would read as modified the first time a component was re-saved.
        // ###########################################################################################
        private static BoardLocalFileEntry ToBoardLocalFile(ComponentDraftWriter.LocalFileDraftRow row) => new()
        {
            Category = row.Category?.Trim() ?? string.Empty,
            Name = row.Name?.Trim() ?? string.Empty,
            File = ComponentBoardWriter.JoinFilePath(row.FileLocation, row.File),
        };

        private static BoardLinkEntry ToBoardLink(ComponentDraftWriter.LinkDraftRow row) => new()
        {
            Category = row.Category?.Trim() ?? string.Empty,
            Name = row.Name?.Trim() ?? string.Empty,
            Url = row.Url?.Trim() ?? string.Empty,
        };

        // ###########################################################################################
        // A board-scoped section after a save: what is on disk NOW, with the window's own edits
        // applied to it - a three-way merge of `current` (the workbook as read at save time),
        // `atOpen` (the rows the window loaded) and `edited` (the rows it hands back).
        //
        //   - The window did not change the section (edited == atOpen): the disk wins, whatever
        //     it now says. This is the case that used to lose data - a component edit that never
        //     touched the board links still overwrote a link added in Excel meanwhile.
        //   - Nothing else changed it (current == atOpen): the window's rows win whole, in the
        //     window's order - exactly the old behaviour, and the common case.
        //   - BOTH changed it: rows the window removed are removed from the disk's version, and
        //     rows the window added are inserted after the row they followed in the window (or
        //     appended when that row has gone). An edited row is a removal plus an addition, so
        //     it lands where the original was.
        //
        // Rows are compared on a normalised key of every field (trimmed, "/"-separated paths), so
        // the same row read from the workbook and rebuilt from the window compares equal.
        //
        // atOpen null means "no baseline" and keeps the old wholesale replace.
        // ###########################################################################################
        private static List<T> MergeBoardScoped<T>(
            List<T> current,
            List<T>? atOpen,
            List<T> edited,
            Func<T, string> keyOf)
        {
            if (atOpen is null)
            {
                return edited;
            }

            if (atOpen.Select(keyOf).SequenceEqual(edited.Select(keyOf), StringComparer.Ordinal))
            {
                return [.. current];
            }

            if (current.Select(keyOf).SequenceEqual(atOpen.Select(keyOf), StringComparer.Ordinal))
            {
                return edited;
            }

            // Multisets, so a row present twice is removed or added the right number of times.
            Dictionary<string, int> openCounts = ComponentBoardWriter.CountKeys(atOpen, keyOf);
            Dictionary<string, int> editedCounts = ComponentBoardWriter.CountKeys(edited, keyOf);
            Dictionary<string, int> editedTotals = ComponentBoardWriter.CountKeys(edited, keyOf);

            List<T> result = [.. current];

            // Removed by the window: in atOpen more times than in edited.
            foreach (T row in atOpen)
            {
                string key = keyOf(row);

                if (editedCounts.TryGetValue(key, out int kept) && kept > 0)
                {
                    editedCounts[key] = kept - 1;
                    continue;
                }

                int index = result.FindIndex(existing => string.Equals(keyOf(existing), key, StringComparison.Ordinal));
                if (index >= 0)
                {
                    result.RemoveAt(index);
                }
            }

            // Added by the window: in edited more times than in atOpen.
            for (int i = 0; i < edited.Count; i++)
            {
                string key = keyOf(edited[i]);

                if (openCounts.TryGetValue(key, out int wasThere) && wasThere > 0)
                {
                    openCounts[key] = wasThere - 1;
                    continue;
                }

                // *** IDEMPOTENT: never more copies than the window itself holds. *** Without this,
                // a second save from the same window - whose baseline predates the first - would add
                // the first save's rows a second time.
                int alreadyThere = result.Count(existing => string.Equals(keyOf(existing), key, StringComparison.Ordinal));
                if (alreadyThere >= editedTotals[key])
                {
                    continue;
                }

                int insertAt = i == 0 ? 0 : result.Count;

                for (int previous = i - 1; previous >= 0; previous--)
                {
                    string previousKey = keyOf(edited[previous]);
                    int found = result.FindLastIndex(existing => string.Equals(keyOf(existing), previousKey, StringComparison.Ordinal));

                    if (found >= 0)
                    {
                        insertAt = found + 1;
                        break;
                    }
                }

                result.Insert(insertAt, edited[i]);
            }

            return result;
        }

        private static Dictionary<string, int> CountKeys<T>(IEnumerable<T> rows, Func<T, string> keyOf)
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);

            foreach (T row in rows)
            {
                string key = keyOf(row);
                counts[key] = counts.TryGetValue(key, out int count) ? count + 1 : 1;
            }

            return counts;
        }

        private static string Key(params string?[] parts) =>
            string.Join('\u001F', parts.Select(part => part?.Trim() ?? string.Empty));

        private static string NormaliseFile(string? file) =>
            (file ?? string.Empty).Trim().Replace('\\', '/').Trim('/');

        private static string JoinFilePath(string? fileLocation, string? file)
        {
            // Backslashes are normalised because a contributor can type a Windows-style folder
            // into the window, and the stored path is read by the Linux server too. Only the
            // LOCATION is slash-trimmed: a file name is a name, and trimming it would quietly
            // alter one that legitimately started with a character being stripped.
            string trimmedLocation = fileLocation?.Trim().Replace('\\', '/').Trim('/') ?? string.Empty;
            string trimmedFile = file?.Trim() ?? string.Empty;

            if (trimmedFile.Length == 0)
            {
                return string.Empty;
            }

            return trimmedLocation.Length == 0 ? trimmedFile : $"{trimmedLocation}/{trimmedFile}";
        }

        private static string Or(string? value, string fallback)
        {
            string trimmed = value?.Trim() ?? string.Empty;

            return trimmed.Length == 0 ? fallback : trimmed;
        }
    }
}
