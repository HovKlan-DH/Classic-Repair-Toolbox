using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Which Hardware and Board drop-down entries carry the amber "Draft" chip (owner request,
    // 2026-09-24) - the same chip a drafted component row already carries in the component list.
    //
    // A BOARD carries it when it has a local draft. A HARDWARE carries it when ANY of its boards
    // does, so a draft is findable from the hardware list without opening every board under it.
    //
    // *** BUILT FROM THE BOARDS THE DRAFTS TAB LISTS, and nothing else. *** Main hands this the
    // very list TabDrafts.RefreshDrafts produced, so the chip and the Drafts tab cannot disagree
    // about what "has a draft" means - including for a draft that was started and not yet edited,
    // which the tab lists (a draft exists because its marker does; see DraftManager) and which is
    // therefore badged too.
    //
    // Names compare case-INSENSITIVELY and trimmed, like every other hardware/board comparison in
    // the app (the drop-down selection itself matches with OrdinalIgnoreCase).
    // ###########################################################################################
    internal sealed class DraftBadgeSet
    {
        public static readonly DraftBadgeSet Empty = new(Array.Empty<HardwareBoardEntry>());

        private readonly HashSet<string> thisHardware = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> thisBoards = new(StringComparer.OrdinalIgnoreCase);

        private DraftBadgeSet(IEnumerable<HardwareBoardEntry> draftedBoards)
        {
            foreach (HardwareBoardEntry entry in draftedBoards)
            {
                string hardware = entry.HardwareName?.Trim() ?? string.Empty;
                string board = entry.BoardName?.Trim() ?? string.Empty;

                if (hardware.Length == 0)
                {
                    continue;
                }

                this.thisHardware.Add(hardware);

                if (board.Length > 0)
                {
                    this.thisBoards.Add(DraftBadgeSet.BoardKey(hardware, board));
                }
            }
        }

        public static DraftBadgeSet From(IEnumerable<HardwareBoardEntry>? draftedBoards) =>
            draftedBoards == null ? DraftBadgeSet.Empty : new DraftBadgeSet(draftedBoards.ToList());

        public bool HardwareHasDraft(string? hardwareName) =>
            this.thisHardware.Contains(hardwareName?.Trim() ?? string.Empty);

        // A board name alone is not unique - "250407" can exist under more than one hardware - so
        // the question is always asked about a board OF a hardware.
        public bool BoardHasDraft(string? hardwareName, string? boardName) =>
            this.thisBoards.Contains(DraftBadgeSet.BoardKey(hardwareName?.Trim() ?? string.Empty, boardName?.Trim() ?? string.Empty));

        private static string BoardKey(string hardware, string board) => $"{hardware}|{board}";
    }
}
