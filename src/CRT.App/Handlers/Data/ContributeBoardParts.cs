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

        // "Regions" in the column: what the board comes in, or "0" as the counts say it.
        public static string RegionsText(BoardData? board) =>
            BoardRegions.Has(board) ? string.Join(", ", BoardRegions.Of(board)) : ContributeBoardParts.CountText(0);

        // A count as the column shows it - "0" with none (owner request, 2026-11-04: "none" until then).
        public static string CountText(int count) =>
            Math.Max(count, 0).ToString(CultureInfo.InvariantCulture);

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

        // A component's page with no component picked (owner request, 2026-11-03: "state that the user can
        // use and select a component from the normal left-side component list").
        internal static string NoComponentTitle(ComponentPage page) => $"Components: {ComponentPaneWording.PageName(page)}";

        internal static string NoComponentPicked(ComponentPage page) =>
            $"No component is selected. Select one in the component list on the far left - the \"Category\" list and the filter box above it narrow it down - and its \"{ComponentPaneWording.PageName(page)}\" shows here. \"Add component\" adds a new one.";

        // ... every component selected, as "Mark all" on the other tabs selects them - opened, the next field typed would go
        // to every component of the board (code review, 2026-11-09). "Mark all" is off on this tab.
        internal static string EveryComponentSelected(ComponentPage page) =>
            $"Every component in the list is selected, as \"Mark all\" selects them. Click one to see its \"{ComponentPaneWording.PageName(page)}\" here, or pick several with Ctrl or Shift to edit them together.";

        // "Table" picked, on an entry that is no sheet: why nothing changed.
        public const string NoTableForEntry =
            "This page has no table of its own, so it looks the same with \"UI\" and \"Table\". Pick another entry on the left to see its table.";

        // ###########################################################################################
        // THE TAB'S TOP LINE NAMES THE BOARD SHOWN (owner request, 2026-11-05), its two names in bold
        // (BoldRuns.ShowMarked) - then what a change does, which is why "your first change makes your own
        // draft" is said nowhere else any more.
        // ###########################################################################################
        internal static string TopLine(string? hardware, string? board)
        {
            string hardwareName = hardware?.Trim() ?? string.Empty;
            string boardName = board?.Trim() ?? string.Empty;

            string first = hardwareName.Length > 0 && boardName.Length > 0
                ? $"Improve data for the [<b>{hardwareName}</b>] hardware and [<b>{boardName}</b>] board."
                : "Improve data for a board.";

            return $"{first} Every change is saved at once into your own draft and shows up on the board straight away - nobody else sees it until you submit it.";
        }

        // With no board to show: where to pick one.
        public const string NoBoardLine = "Pick a hardware and a board in the lists on the left to work on it.";

        // ###########################################################################################
        // THE NOTICE OVER A BOARD WITH A DRAFT (owner request, 2026-10-30: "highlighted text clearly
        // stating that this is a local draft and that it must be submitted for a board maintainer to
        // review it, before others can see it").
        // ###########################################################################################
        public static string LocalDraftNotice(bool isNewBoard) => isNewBoard
            ? "This new board is a local draft, saved only on this computer. Nobody else can see it until you submit it and a maintainer has reviewed and published it."
            : "You are working on a local draft, saved only on this computer. Nobody else can see these changes until you submit them and a maintainer of this board has reviewed and published them.";

        // "My drafts" closed: how many there are - or that there are none yet.
        public static string MyDraftsLine(int drafts) => drafts switch
        {
            <= 0 => "No drafts yet",
            1 => "My drafts (1)",
            _ => $"My drafts ({drafts.ToString(CultureInfo.InvariantCulture)})",
        };
    }
}
