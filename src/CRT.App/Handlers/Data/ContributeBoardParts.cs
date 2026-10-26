using System.Globalization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THE CONTRIBUTE TAB'S LEFT COLUMN SAYS ABOUT A BOARD (CRT 4.0 step 2, 2026-10-25 - the
    // plan's mock: "Pictures 12 ... Important signals 4 ... Files and links 6 ... Credits 3"), and its
    // lines for when there is nothing to list. Pure: a board in, counts and words out.
    // ###########################################################################################
    public sealed record ContributeBoardParts(int Pictures, int ImportantSignals, int FilesAndLinks, int Credits)
    {
        public static ContributeBoardParts Of(BoardData? board) => board is null
            ? new ContributeBoardParts(0, 0, 0, 0)
            : new ContributeBoardParts(
                board.Schematics.Count,
                board.KiCadImportantSignals.Count,
                board.BoardLocalFiles.Count + board.BoardLinks.Count,
                board.Credits.Count);

        // A count as the column shows it - "none" rather than a bare 0, which reads like a fault.
        public static string CountText(int count) =>
            count <= 0 ? "none" : count.ToString(CultureInfo.InvariantCulture);

        // The header of a board with no draft: what changing it does.
        public static string NoDraftLine(bool hasBoard) => hasBoard
            ? "The published data, as everybody sees it. Your first change makes your own draft of it - nothing is shared until you submit."
            : "Pick a hardware and a board in the lists on the left to work on it.";

        // The component list when it lists nothing.
        public static string NoComponentsLine(bool hasBoard, bool searching) =>
            !hasBoard ? "No board is shown."
            : searching ? "No component matches the search."
            : "This board has no components yet. Add the first with \"Add component\".";

        // "My drafts" closed: how many there are - or that there are none yet.
        public static string MyDraftsLine(int drafts) => drafts switch
        {
            <= 0 => "No drafts yet",
            1 => "My drafts (1)",
            _ => $"My drafts ({drafts.ToString(CultureInfo.InvariantCulture)})",
        };
    }
}
