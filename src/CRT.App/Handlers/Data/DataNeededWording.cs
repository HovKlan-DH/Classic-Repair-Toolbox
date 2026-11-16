using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // What Main's "data needed" card says (Main.DataNeededNotice.cs; owner request, 2026-11-14): a board checked again under
    // "Visible hardware, boards and schematics" whose data is not on this computer, or behind - the data check leaves an
    // unchecked board out (DataSyncScope). The names in brackets are bold on the card (NoticeEmphasis).
    // ###########################################################################################
    internal static class DataNeededWording
    {
        // The card's button, named in its words.
        public const string CheckButton = "Check for new data";

        public static string Of(IReadOnlyList<HardwareBoardEntry> boards)
        {
            ArgumentNullException.ThrowIfNull(boards);

            string which = boards.Count == 1
                ? $"The data for [{boards[0].HardwareName}] [{boards[0].BoardName}] is"
                : $"The data for [{boards.Count}] boards you checked is";

            return $"{which} not on this computer, or not up to date: data is downloaded only for the hardware and boards " +
                   $"checked on the \"Configuration\" tab. Press \"{DataNeededWording.CheckButton}\" to download it.";
        }
    }
}
