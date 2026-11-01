using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THE CONTRIBUTE TAB'S LEFT COLUMN SAYS ABOUT A BOARD (CRT 4.0 step 2, 2026-10-25 - the
    // plan's mock: "Pictures 12 ... Important signals 4 ... Files and links 6 ... Credits 3"), and its
    // lines for when there is nothing to list. Pure: a board in, counts and words out. Files and links
    // are counted apart since 2026-10-29 (owner request: an entry each in the column); the components
    // are counted by the tab, from the list it shows in the region shown.
    // ###########################################################################################
    public sealed record ContributeBoardParts(int Pictures, int ImportantSignals, int Files, int Links, int Credits, int Categories = 0)
    {
        public static ContributeBoardParts Of(BoardData? board) => board is null
            ? new ContributeBoardParts(0, 0, 0, 0, 0)
            : new ContributeBoardParts(
                board.Schematics.Count,
                board.KiCadImportantSignals.Count,
                board.BoardLocalFiles.Count,
                board.BoardLinks.Count,
                board.Credits.Count,
                ComponentListBuilder.BuildDistinctCategories(board).Count);

        // "Regions" in the column: what the board comes in, or "none" as the counts say it.
        public static string RegionsText(BoardData? board) =>
            BoardRegions.Has(board) ? string.Join(", ", BoardRegions.All) : ContributeBoardParts.CountText(0);

        // A count as the column shows it - "none" rather than a bare 0, which reads like a fault.
        public static string CountText(int count) =>
            count <= 0 ? "none" : count.ToString(CultureInfo.InvariantCulture);

        // ###########################################################################################
        // How many KiCad files a board lists (owner request, 2026-10-30: "KiCad data" in the column
        // with "the amount of files imported/used") - the names under "KiCad data/"
        // (SubmissionKiCadFiles.IsInKiCadFolder, the one rule), each once whatever its spelling.
        // ###########################################################################################
        public static int KiCadFileCount(IEnumerable<string> names) => ContributeBoardParts.KiCadFileNames(names).Count;

        // The board's KiCad files among `names`, each once - what the column counts, and what the KiCad
        // panel's summary says that number is made of (KiCadImportSummary.Files).
        public static List<string> KiCadFileNames(IEnumerable<string> names) =>
            names
                .Where(SubmissionKiCadFiles.IsInKiCadFolder)
                .Select(BoardFileNames.Normalise)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

        // The header of a board with no draft: what changing it does.
        public static string NoDraftLine(bool hasBoard) => hasBoard
            ? "The published data, as everybody sees it. Your first change makes your own draft of it - nothing is shared until you submit."
            : "Pick a hardware and a board in the lists on the left to work on it.";

        // ###########################################################################################
        // THE NOTICE OVER A BOARD WITH A DRAFT (owner request, 2026-10-30: "highlighted text clearly
        // stating that this is a local draft and that it must be submitted for a board maintainer to
        // review it, before others can see it").
        // ###########################################################################################
        public static string LocalDraftNotice(bool isNewBoard) => isNewBoard
            ? "This new board is a local draft, saved only on this computer. Nobody else can see it until you submit it and a maintainer has reviewed and published it."
            : "You are working on a local draft, saved only on this computer. Nobody else can see these changes until you submit them and a maintainer of this board has reviewed and published them.";

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
