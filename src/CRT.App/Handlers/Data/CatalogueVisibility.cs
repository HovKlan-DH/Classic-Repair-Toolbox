using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Decides whether a hardware, board or schematic is visible in the app UI, based on the
    // Configuration tab's checkbox tree. A key is UNCHECKED if it appears in the caller-supplied
    // set; unlisted keys are checked (visible) by default - see UserSettings.CatalogueUncheckedKeys.
    //
    // Visibility is the AND of an item and every one of its ancestors: a schematic whose own key
    // is still checked is still hidden if its board or hardware is unchecked. That is what lets
    // unchecking a hardware or board hide everything beneath it without touching the descendants'
    // own stored state, so re-checking the parent instantly restores exactly what was visible
    // before - nothing about the children needs to change to make that true.
    //
    // Keys are compared CASE-INSENSITIVELY, matching every other hardware/board/schematic name
    // comparison in the app (the drop-downs, the board-key dictionaries, the tree's own GroupBy).
    // The sets these predicates read are built OrdinalIgnoreCase by UserSettings - see
    // NormalizeCatalogueKeySets there for why that has to be re-applied after deserialization -
    // but the comparison is not left to the caller's set: a recased name arriving in synced data
    // would otherwise silently un-hide an item, with no orphaned entry visible anywhere to clear.
    // ###########################################################################################
    public static class CatalogueVisibility
    {
        private const char KeySeparator = '|';

        public static readonly StringComparer KeyComparer = StringComparer.OrdinalIgnoreCase;

        public static string BuildKey(string hardwareName) => hardwareName;

        public static string BuildKey(string hardwareName, string boardName) =>
            $"{hardwareName}{KeySeparator}{boardName}";

        public static string BuildKey(string hardwareName, string boardName, string schematicName) =>
            $"{hardwareName}{KeySeparator}{boardName}{KeySeparator}{schematicName}";

        // ###########################################################################################
        // A key's depth is its separator count: hardware carries none, a board one, a schematic two.
        // ###########################################################################################
        public static bool IsSchematicKey(string key) => CountSeparators(key) == 2;

        public static bool IsBoardKey(string key) => CountSeparators(key) == 1;

        public static bool IsHardwareKey(string key) => CountSeparators(key) == 0;

        // ###########################################################################################
        // Whether the key belongs to the given hardware/board - true for that board's own key and
        // for any schematic key under it, false for a hardware key (which names no single board) and
        // for anything under a different board. Used to decide whether a checkbox change affects the
        // board currently on screen.
        // ###########################################################################################
        public static bool KeyNamesBoard(string key, (string HardwareName, string BoardName) board)
        {
            if (string.IsNullOrWhiteSpace(board.HardwareName) || string.IsNullOrWhiteSpace(board.BoardName))
            {
                return false;
            }

            string boardKey = BuildKey(board.HardwareName, board.BoardName);

            return KeyComparer.Equals(key, boardKey) ||
                   key.StartsWith(boardKey + KeySeparator, StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // Whether the currently selected hardware/board still survives in freshly recomputed visible
        // name lists - i.e. whether a catalogue checkbox change actually affects what is on screen,
        // as opposed to some other, unrelated hardware/board. Used by Main.ApplyCatalogueVisibility
        // to decide between two very different responses to a checkbox toggle: reselecting (and so
        // reloading the whole board - Excel, thumbnails, KiCad, and closing/reopening the detached
        // thumbnails window if one is open) when the item actually on screen was hidden or shown, or
        // simply refreshing the drop-downs' contents with no reselect at all when it was not.
        //
        // No selection (an empty hardware name) never "survives" - GetCurrentBoardKeyParts documents
        // empty strings for that state, matching the same "no board selected" convention used
        // throughout Main (see KeyNamesBoard and GetCurrentBoardKey). A blank board name (hardware
        // selected, no board yet) is treated as trivially surviving, since there is no board
        // selection for a hidden board to invalidate.
        // ###########################################################################################
        public static bool CurrentSelectionSurvives(
            IEnumerable<string> visibleHardwareNames,
            IEnumerable<string> visibleBoardNamesForSelectedHardware,
            string currentHardwareName,
            string currentBoardName)
        {
            if (string.IsNullOrEmpty(currentHardwareName))
            {
                return false;
            }

            bool hardwareSurvives = Contains(visibleHardwareNames, currentHardwareName);
            if (!hardwareSurvives)
            {
                return false;
            }

            return string.IsNullOrEmpty(currentBoardName) ||
                   Contains(visibleBoardNamesForSelectedHardware, currentBoardName);
        }

        private static bool Contains(IEnumerable<string> names, string name)
        {
            foreach (string candidate in names)
            {
                if (KeyComparer.Equals(candidate, name))
                {
                    return true;
                }
            }

            return false;
        }

        private static int CountSeparators(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                return 0;
            }

            int count = 0;
            foreach (char c in key)
            {
                if (c == KeySeparator)
                {
                    count++;
                }
            }

            return count;
        }

        // ###########################################################################################
        // True when the hardware's own key is checked. Boards/schematics use the AND overloads
        // below, which also cover this case - this one exists so callers naming only a hardware
        // (the Hardware drop-down) do not have to pass null board/schematic arguments through.
        // ###########################################################################################
        public static bool IsHardwareVisible(IReadOnlySet<string> uncheckedKeys, string hardwareName) =>
            !ContainsKey(uncheckedKeys, BuildKey(hardwareName));

        // ###########################################################################################
        // True when the board AND its hardware are both checked.
        // ###########################################################################################
        public static bool IsBoardVisible(IReadOnlySet<string> uncheckedKeys, string hardwareName, string boardName) =>
            IsHardwareVisible(uncheckedKeys, hardwareName) &&
            !ContainsKey(uncheckedKeys, BuildKey(hardwareName, boardName));

        // ###########################################################################################
        // True when the schematic AND its board AND its hardware are all checked.
        // ###########################################################################################
        public static bool IsSchematicVisible(
            IReadOnlySet<string> uncheckedKeys, string hardwareName, string boardName, string schematicName) =>
            IsBoardVisible(uncheckedKeys, hardwareName, boardName) &&
            !ContainsKey(uncheckedKeys, BuildKey(hardwareName, boardName, schematicName));

        // ###########################################################################################
        // Case-insensitive membership regardless of the set's own comparer. Public because the
        // Configuration tab's tree asks the same question of both the unchecked AND the collapsed
        // set, and both have to answer it the same way this class does.
        //
        // A HashSet<string> built OrdinalIgnoreCase answers Contains case-insensitively on its own,
        // and that is the fast path taken here. But these predicates are public and are handed sets
        // from several places (UserSettings' live set, a test's literal, a deserialized one), and an
        // ordinal set silently answering "not unchecked" for a recased key un-hides the item with no
        // visible orphan to clear - so a set with any other comparer is scanned instead.
        // ###########################################################################################
        public static bool IsKeyListed(IReadOnlySet<string> keys, string key) => ContainsKey(keys, key);

        private static bool ContainsKey(IReadOnlySet<string> uncheckedKeys, string key)
        {
            if (uncheckedKeys.Count == 0)
            {
                return false;
            }

            if (uncheckedKeys is HashSet<string> set && ReferenceEquals(set.Comparer, KeyComparer))
            {
                return set.Contains(key);
            }

            foreach (string existing in uncheckedKeys)
            {
                if (KeyComparer.Equals(existing, key))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
