using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE REVIEWER'S TABLE (maintainer request, 2026-09-25): "make the same 'Edit in table
    // format' available in the review app ... The reviewer should be able to also edit whatever, if
    // he chooses to publish it afterwards."
    //
    // What goes over the wire between CRT.Server and CRT.Review for it, and the one rule for what
    // an amendment may change - in CRT.Data so both ends use the same types and the same rule.
    // ###########################################################################################

    // What the review application's table opens on: the published board (null for a new system)
    // and the submission's current rows, plus the amendment version the rows are at - sent back
    // with an amendment, so two reviewers editing at once cannot silently overwrite each other.
    public sealed record ReviewTableData(int Version, SubmissionRows? Published, SubmissionRows Submitted);

    public static class SubmissionRowsBoard
    {
        // ###########################################################################################
        // Rows to a board - every section copied, never aliased. The table and the change summary
        // both work on BoardData.
        // ###########################################################################################
        public static BoardData ToBoard(SubmissionRows? rows)
        {
            rows ??= new SubmissionRows();

            return new BoardData
            {
                RevisionDate = rows.RevisionDate ?? string.Empty,
                Schematics = [.. rows.Schematics],
                Components = [.. rows.Components],
                ComponentImages = [.. rows.ComponentImages],
                ComponentHighlights = [.. rows.ComponentHighlights],
                ComponentLocalFiles = [.. rows.ComponentLocalFiles],
                ComponentLinks = [.. rows.ComponentLinks],
                BoardLocalFiles = [.. rows.BoardLocalFiles],
                BoardLinks = [.. rows.BoardLinks],
                Credits = [.. rows.Credits],
                KiCadImportantSignals = [.. rows.KiCadImportantSignals]
            };
        }

        // A board to rows. Calibrations are not part of a BoardData and stay empty here.
        public static SubmissionRows FromBoard(BoardData board)
        {
            ArgumentNullException.ThrowIfNull(board);

            return new SubmissionRows
            {
                RevisionDate = board.RevisionDate ?? string.Empty,
                Schematics = board.Schematics.ToList(),
                Components = board.Components.ToList(),
                ComponentImages = board.ComponentImages.ToList(),
                ComponentHighlights = board.ComponentHighlights.ToList(),
                ComponentLocalFiles = board.ComponentLocalFiles.ToList(),
                ComponentLinks = board.ComponentLinks.ToList(),
                BoardLocalFiles = board.BoardLocalFiles.ToList(),
                BoardLinks = board.BoardLinks.ToList(),
                Credits = board.Credits.ToList(),
                KiCadImportantSignals = board.KiCadImportantSignals.ToList()
            };
        }

        // ###########################################################################################
        // WHAT AN AMENDMENT MAY CHANGE: the table's nine sheets (BoardWorkbookSchema.AllSheets),
        // and nothing else. The highlights, the KiCad calibrations and the revision date are not
        // in the table, so they are ALWAYS taken from the submission as it stands - whatever a
        // request carries for them. The server applies this, so a hand-made request cannot use the
        // table's route to change what the table does not show.
        //
        // *** EXCEPT THAT THEY FOLLOW THE SCHEMATIC THEY ARE DRAWN ON (code review, 2026-09-25). ***
        // Highlights and calibrations name a schematic, and the table can delete or rename one.
        // Kept unchanged, they named a schematic that no longer existed, the validator refused the
        // save (highlight.unknown_schematic), and nothing in the table could fix it - the reviewer
        // could not correct a schematic at all. See FollowSchematics for the rule.
        //
        // *** AND A DELETED COMPONENT'S HIGHLIGHTS GO WITH IT (maintainer decision, 2026-09-25). ***
        // The table drops them when it saves (BoardTableDocument.ApplyTo), because only the table
        // can tell a component deleted from one renamed. So a highlight the edit no longer carries
        // is dropped - but ONLY when no component in the edit has its label: a request cannot use
        // this to remove the highlights of a component the board still has, nor add or move one.
        // ###########################################################################################
        public static SubmissionRows WithTableSections(SubmissionRows current, SubmissionRows edited)
        {
            ArgumentNullException.ThrowIfNull(current);
            ArgumentNullException.ThrowIfNull(edited);

            IReadOnlyDictionary<string, string?> follow = SubmissionRowsBoard.FollowSchematics(current.Schematics, edited.Schematics);

            var editedLabels = new HashSet<string>(
                edited.Components.Select(component => (component.BoardLabel ?? string.Empty).Trim()),
                StringComparer.OrdinalIgnoreCase);

            var editedHighlights = new HashSet<string>(
                edited.ComponentHighlights.Select(SubmissionRowsBoard.HighlightKey),
                StringComparer.OrdinalIgnoreCase);

            return new SubmissionRows
            {
                RevisionDate = current.RevisionDate,

                ComponentHighlights = current.ComponentHighlights
                    .Where(highlight =>
                        editedLabels.Contains((highlight.BoardLabel ?? string.Empty).Trim()) ||
                        editedHighlights.Contains(SubmissionRowsBoard.HighlightKey(highlight)))
                    .Select(highlight => SubmissionRowsBoard.Follow(follow, highlight.SchematicName) switch
                    {
                        null => null,
                        string name when name == highlight.SchematicName => highlight,
                        string name => new ComponentHighlightEntry
                        {
                            SchematicName = name,
                            BoardLabel = highlight.BoardLabel,
                            X = highlight.X,
                            Y = highlight.Y,
                            Width = highlight.Width,
                            Height = highlight.Height
                        }
                    })
                    .OfType<ComponentHighlightEntry>()
                    .ToList(),

                KiCadCalibrations = current.KiCadCalibrations
                    .Select(calibration => SubmissionRowsBoard.Follow(follow, calibration.SchematicName) switch
                    {
                        null => null,
                        string name when name == calibration.SchematicName => calibration,
                        string name => new KiCadCalibrationEntry
                        {
                            SchematicName = name,
                            CadName = calibration.CadName,
                            OffsetX = calibration.OffsetX,
                            OffsetY = calibration.OffsetY,
                            ScaleX = calibration.ScaleX,
                            ScaleY = calibration.ScaleY,
                            MirrorX = calibration.MirrorX,
                            MirrorY = calibration.MirrorY
                        }
                    })
                    .OfType<KiCadCalibrationEntry>()
                    .ToList(),

                Schematics = edited.Schematics.ToList(),
                Components = edited.Components.ToList(),
                ComponentImages = edited.ComponentImages.ToList(),
                ComponentLocalFiles = edited.ComponentLocalFiles.ToList(),
                ComponentLinks = edited.ComponentLinks.ToList(),
                BoardLocalFiles = edited.BoardLocalFiles.ToList(),
                BoardLinks = edited.BoardLinks.ToList(),
                Credits = edited.Credits.ToList(),
                KiCadImportantSignals = edited.KiCadImportantSignals.ToList()
            };
        }

        // ###########################################################################################
        // Where each schematic the edit REMOVED went: current name -> its new name when it was
        // renamed, or null when it was deleted. Names still present are not listed (unchanged).
        //
        // A RENAME is a removed name and an added name drawn from the SAME image file - the table
        // pairs rows by name, so it sees a rename as a deletion plus an addition, and the image is
        // what says they are one schematic. Only an unambiguous pair counts: when the image is
        // shared by several removed or several added rows, the rows are treated as deleted, since
        // moving a highlight onto the wrong board image is worse than dropping it.
        //
        // Names compare ignoring case, as the validator does (a highlight on "sheet 1" draws on
        // "Sheet 1").
        // ###########################################################################################
        internal static IReadOnlyDictionary<string, string?> FollowSchematics(
            IEnumerable<BoardSchematicEntry> current,
            IEnumerable<BoardSchematicEntry> edited)
        {
            var editedNames = new HashSet<string>(
                edited.Select(schematic => schematic.SchematicName ?? string.Empty),
                StringComparer.OrdinalIgnoreCase);

            var currentNames = new HashSet<string>(
                current.Select(schematic => schematic.SchematicName ?? string.Empty),
                StringComparer.OrdinalIgnoreCase);

            List<BoardSchematicEntry> removed = current
                .Where(schematic => !editedNames.Contains(schematic.SchematicName ?? string.Empty))
                .ToList();

            List<BoardSchematicEntry> added = edited
                .Where(schematic => !currentNames.Contains(schematic.SchematicName ?? string.Empty))
                .ToList();

            var follow = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

            foreach (BoardSchematicEntry gone in removed)
            {
                string image = gone.SchematicImageFile ?? string.Empty;

                List<BoardSchematicEntry> sameImageAdded = image.Length == 0
                    ? []
                    : added.Where(schematic => string.Equals(schematic.SchematicImageFile, image, StringComparison.OrdinalIgnoreCase)).ToList();

                int sameImageRemoved = removed.Count(schematic => string.Equals(schematic.SchematicImageFile, image, StringComparison.OrdinalIgnoreCase));

                follow[gone.SchematicName ?? string.Empty] = sameImageAdded.Count == 1 && sameImageRemoved == 1
                    ? sameImageAdded[0].SchematicName
                    : null;
            }

            return follow;
        }

        // One highlight, whole: where it is drawn, for which label, and its rectangle.
        private static string HighlightKey(ComponentHighlightEntry highlight) =>
            string.Join(
                "|",
                (highlight.SchematicName ?? string.Empty).Trim(),
                (highlight.BoardLabel ?? string.Empty).Trim(),
                (highlight.X ?? string.Empty).Trim(),
                (highlight.Y ?? string.Empty).Trim(),
                (highlight.Width ?? string.Empty).Trim(),
                (highlight.Height ?? string.Empty).Trim());

        // The name a highlight or calibration on `schematicName` carries after the edit: unchanged,
        // renamed, or null when its schematic was deleted. A name the CURRENT rows never had is left
        // as it is - it was not the reviewer's edit that made it unknown.
        private static string? Follow(IReadOnlyDictionary<string, string?> follow, string? schematicName)
        {
            string name = schematicName ?? string.Empty;

            return follow.TryGetValue(name, out string? now) ? now : name;
        }
    }
}
