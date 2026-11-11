using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // DELETES A COMPONENT FROM A DRAFT'S BOARD - every row of its label, in every region.
    //
    // All that is left of what applied the component editor's saves: that window went in CRT 4.0's
    // step 2, and the component pane changes a draft through ComponentPaneEdits.
    //
    // PURE: takes a board, returns a new board, touches no files.
    // ###########################################################################################
    public static class ComponentBoardWriter
    {
        // ###########################################################################################
        // "Delete component" in the component pane (ComponentPane.Whole.cs).
        //
        // Every row belonging to this board label goes, across all four component-scoped sections
        // INCLUDING the Components row itself - AND its highlights on every schematic.
        //
        // *** THE HIGHLIGHTS GO TOO (owner decision, 2026-09-25): "if a component really is
        // deleted, then it should remove EVERYTHING related to this component." *** They used to be
        // kept, on the reasoning that highlights are the label editor's concern - which left
        // rectangles naming a component the board no longer has. The board table's delete does the
        // same (BoardTableDocument).
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
                HardwareName = current.HardwareName,
                BoardName = current.BoardName,
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

                // Its pins and replacement parts are the component's too (CRT 4.0, 2026-10-25).
                ComponentPins = ComponentBoardWriter.WithoutComponent(
                    current.ComponentPins, label, existing => existing.BoardLabel),

                ComponentReplacements = ComponentBoardWriter.WithoutComponent(
                    current.ComponentReplacements, label, existing => existing.BoardLabel),

                Regions = current.Regions,
                RegionColors = current.RegionColors,

                BoardLocalFiles = current.BoardLocalFiles,
                BoardLinks = current.BoardLinks,
                Credits = current.Credits,
                KiCadImportantSignals = current.KiCadImportantSignals,
            };
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
    }
}
