using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // What is left to do on a board once its KiCad data is in, as the KiCad import's summary says it
    // (KiCadImportSummary.LeftToDo). NotOnBoard: the components the KiCad data has that the board lacks -
    // the summary's button adds them.
    internal sealed record KiCadWorkLeft(bool Match, bool Align, bool Components, int NotOnBoard, bool Mark, bool Signals);

    // ###########################################################################################
    // What the board's KiCad data says: the labels of its components - the very ones "Add N from
    // KiCad" adds (KiCadComponentTable: its first PCB file's footprints, no placeholder, logo or
    // mounting hole) - its view names, and every view as matching a picture sees it (ToMatch - its
    // side, and its PCB file's name). Not Read: the KiCad files could not be read - no data, never data
    // with nothing in it (code review, 2026-10-26).
    //
    // For "Board status" (2026-11-01): how many components have a footprint to mark them from (Markable -
    // none without a PCB file), the parts only the schematic has (OnlyInSchematic - marked by hand), and
    // the board's nets (Nets).
    // ###########################################################################################
    internal sealed record BoardKiCadFacts(
        IReadOnlyList<string> References,
        IReadOnlyList<string> Views,
        bool Read = true,
        IReadOnlyList<KiCadViewToMatch>? ToMatch = null,
        int Markable = 0,
        IReadOnlyList<string>? OnlyInSchematic = null,
        KiCadNetChoices? Nets = null)
    {
        public static BoardKiCadFacts Of(KiCadProjectRoot? root)
        {
            if (root is null)
            {
                return new BoardKiCadFacts([], [], Read: false);
            }

            List<KiCadViewToMatch> views = root.Project.Views
                .Select(view => new KiCadViewToMatch(
                    view.DisplayName?.Trim() ?? string.Empty,
                    BoardKiCadFacts.SideOf(view),
                    BoardKiCadFacts.SideOf(view) == KiCadViewSide.Sheet ? string.Empty : System.IO.Path.GetFileNameWithoutExtension(view.SourceFile ?? string.Empty)))
                .Where(view => view.Name.Length > 0)
                .ToList();

            List<string> references = KiCadComponentTable.Of(root).Select(component => component.Label).ToList();

            return new BoardKiCadFacts(
                references,
                views.Select(view => view.Name).ToList(),
                ToMatch: views,
                Markable: root.Pcb.Count > 0 ? references.Count : 0,
                OnlyInSchematic: KiCadComponentTable.OnlyInSchematic(root),
                Nets: KiCadNetNames.Choices(root.Pcb.SelectMany(pcb => pcb.Nets.List).Select(net => (net.Name, net.NormalizedName, net.IsNoNet))));
        }

        // Which side of the board a view shows - or a schematic sheet.
        internal static KiCadViewSide SideOf(KiCadProjectView view) =>
            string.Equals(view.Type, "pcb_top", StringComparison.OrdinalIgnoreCase) ? KiCadViewSide.Top
            : string.Equals(view.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase) ? KiCadViewSide.Bottom
            : KiCadViewSide.Sheet;

        // The views of `root` that are a side of the board, by their names - for the Pictures part's
        // cards (2026-11-01), without the component table Of builds.
        public static List<string> BoardViewsOf(KiCadProjectRoot root) =>
            root.Project.Views
                .Where(view => BoardKiCadFacts.SideOf(view) != KiCadViewSide.Sheet)
                .Select(view => view.DisplayName?.Trim() ?? string.Empty)
                .Where(name => name.Length > 0)
                .ToList();
    }

    // ###########################################################################################
    // What an existing board still lacks - components by their labels, in the board's order - the rows
    // of its own lists with a field they need left empty (BoardGap), and the components its KiCad data
    // does not have, whose traces cannot light up.
    // ###########################################################################################
    internal sealed record BoardMissing(
        IReadOnlyList<string> WithoutDescription,
        IReadOnlyList<string> Unmarked,
        IReadOnlyList<string> PinsWithoutBaseline,
        IReadOnlyList<BoardGap>? Gaps = null,
        IReadOnlyList<string>? NotInKiCad = null)
    {
        public IReadOnlyList<BoardGap> BoardGaps => this.Gaps ?? [];

        public IReadOnlyList<string> NotInKiCadLabels => this.NotInKiCad ?? [];
    }

    // Which of the board's own lists has rows with a field they need left empty.
    internal enum BoardGapKind
    {
        IncompleteFiles,
        IncompleteLinks,
        IncompleteCredits,
        IncompleteSignals,
    }

    // One such list, said.
    internal sealed record BoardGap(BoardGapKind Kind, string Text);

    // ###########################################################################################
    // WHAT A BOARD STILL LACKS ("Board status" says it, BoardStatusPage), AND WHAT IS LEFT TO DO ONCE ITS
    // KiCad DATA IS IN (the KiCad import's summary says it, KiCadImportSummary.LeftToDo). Worked out from
    // the board as its draft has it - its rows, its KiCad alignments - and what its KiCad data says.
    //
    // PURE: the board in, the answer out.
    // ###########################################################################################
    internal static class BoardReadiness
    {
        // How many labels a "What's missing" list names before "and N more".
        internal const int ListedAtMost = 8;

        // ###########################################################################################
        // What is left to do once the KiCad data is in, by what the board holds - never a box ticked:
        //
        //   - A PICTURE SHOWS A VIEW once its CAD name is one of the views, and is aligned once it has a
        //     calibration.
        //   - ANY PICTURE MATCHED ENDS MATCHING (code review, 2026-10-26): CRT cannot know whether a
        //     picture shows a view - a photo of the ports shows none - and one such photo left it to do
        //     for ever.
        //   - KiCad FILES THAT COULD NOT BE READ ARE NO KiCad DATA AT ALL (code review, 2026-10-26: they
        //     read as data with nothing in it - "All 0 are on the board", ticked): nothing to match,
        //     align or pick signals from, and no component is missing from the board for them.
        // ###########################################################################################
        public static KiCadWorkLeft WorkLeftAfterKiCad(BoardContent content, BoardKiCadFacts kiCad)
        {
            ArgumentNullException.ThrowIfNull(content);
            ArgumentNullException.ThrowIfNull(kiCad);

            BoardData board = content.Board;

            var views = new HashSet<string>(kiCad.Views.Select(view => view.Trim()), StringComparer.OrdinalIgnoreCase);
            var calibrated = new HashSet<string>(content.Calibrations.Select(row => row.SchematicName.Trim()), StringComparer.OrdinalIgnoreCase);

            List<BoardSchematicEntry> matched = board.Schematics
                .Where(row => !string.IsNullOrWhiteSpace(row.SchematicImageFile) && row.CadName?.Trim() is { Length: > 0 } view && views.Contains(view))
                .ToList();

            int aligned = matched.Count(row => calibrated.Contains(row.SchematicName?.Trim() ?? string.Empty));

            List<string> labels = BoardReadiness.Labels(board);
            int notOnBoard = BoardReadiness.NotOnBoard(board, kiCad.References);

            return new KiCadWorkLeft(
                Match: kiCad.Read && matched.Count == 0,
                Align: kiCad.Read && (matched.Count == 0 || aligned < matched.Count),
                Components: labels.Count == 0 || notOnBoard > 0,
                NotOnBoard: notOnBoard,
                Mark: labels.Count == 0 || BoardReadiness.UnmarkedOf(board, labels).Count > 0,
                Signals: kiCad.Read && board.KiCadImportantSignals.Count == 0);
        }

        // ###########################################################################################
        // What an existing board still lacks: components with no short description, components marked
        // on no picture, and components whose pins are listed but which have no scope baseline - each
        // by its label, once, in the board's order. Which of the board's own lists have rows with a
        // field they need left empty. With KiCad data read, the board's components it does not have -
        // their traces cannot light up.
        // ###########################################################################################
        public static BoardMissing Missing(BoardData board, bool hasKiCadFiles = false, BoardKiCadFacts? kiCad = null)
        {
            ArgumentNullException.ThrowIfNull(board);

            List<string> labels = BoardReadiness.Labels(board);

            // Every label with a short description in any region - one pass over the rows (code review,
            // 2026-10-26: a scan of every row per label, on the UI thread at every board shown).
            var described = new HashSet<string>(
                board.Components.Where(row => !string.IsNullOrWhiteSpace(row.Description)).Select(row => row.BoardLabel?.Trim() ?? string.Empty),
                StringComparer.OrdinalIgnoreCase);

            List<string> withoutDescription = labels.Where(label => !described.Contains(label)).ToList();

            var withPins = new HashSet<string>(board.ComponentPins.Select(row => row.BoardLabel?.Trim() ?? string.Empty), StringComparer.OrdinalIgnoreCase);
            var withBaselines = new HashSet<string>(
                board.ComponentImages.Where(ComponentImageKinds.IsScopeBaseline).Select(row => row.BoardLabel?.Trim() ?? string.Empty),
                StringComparer.OrdinalIgnoreCase);

            List<string> pinsWithoutBaseline = labels.Where(label => withPins.Contains(label) && !withBaselines.Contains(label)).ToList();

            BoardKiCadFacts? facts = hasKiCadFiles && kiCad is { Read: true } ? kiCad : null;

            // ###########################################################################################
            // *** A ROW OF THE BOARD'S OWN LISTS WITH A FIELD IT NEEDS EMPTY IS SAID BY ITS LIST (owner
            // request, 2026-11-01: "These errors must be flagged in 'Board status'"). *** Each is an error
            // of the draft's (BoardRowRequired), its list's panel opening it where the empty boxes are marked.
            // ###########################################################################################
            BoardRowsIncomplete incomplete = BoardRowRequired.Of(board);

            var gaps = new List<BoardGap>();

            BoardReadiness.AddIncomplete(gaps, BoardGapKind.IncompleteFiles, incomplete.Files, "file", "files");
            BoardReadiness.AddIncomplete(gaps, BoardGapKind.IncompleteLinks, incomplete.Links, "link", "links");
            BoardReadiness.AddIncomplete(gaps, BoardGapKind.IncompleteCredits, incomplete.Credits, "credit", "credits");
            BoardReadiness.AddIncomplete(gaps, BoardGapKind.IncompleteSignals, incomplete.Signals, "important KiCad signal", "important KiCad signals");

            List<string> notInKiCad = [];
            List<string> fromKiCad = facts is null ? [] : BoardReadiness.Distinct(facts.References);

            // Only once the KiCad data names components at all - an empty PCB says nothing of them.
            if (fromKiCad.Count > 0)
            {
                var inKiCad = new HashSet<string>(fromKiCad, StringComparer.OrdinalIgnoreCase);
                notInKiCad = labels.Where(label => !inKiCad.Contains(label)).ToList();
            }

            return new BoardMissing(withoutDescription, BoardReadiness.UnmarkedOf(board, labels), pinsWithoutBaseline, gaps, notInKiCad);
        }

        private static void AddIncomplete(List<BoardGap> gaps, BoardGapKind kind, int rows, string what, string whats)
        {
            if (rows > 0)
            {
                gaps.Add(new BoardGap(kind, BoardReadiness.IncompleteGap(rows, what, whats)));
            }
        }

        // ###########################################################################################
        // What adding the KiCad components did, said in "KiCad data" and over Board status's panels: how
        // many came, and how many of them with their pins and a datasheet link.
        // ###########################################################################################
        public static string AddedFromKiCad(KiCadComponentsAdded added)
        {
            ArgumentNullException.ThrowIfNull(added);

            if (added.Components == 0)
            {
                return "Every component in the KiCad data is on the board already.";
            }

            string components = added.Components == 1 ? "1 component" : $"{BoardReadiness.Count(added.Components)} components";
            var with = new List<string>();

            if (added.WithPins > 0)
            {
                with.Add($"{BoardReadiness.Count(added.WithPins)} with {(added.WithPins == 1 ? "its" : "their")} pins");
            }

            if (added.WithDatasheet > 0)
            {
                with.Add($"{BoardReadiness.Count(added.WithDatasheet)} with a datasheet link");
            }

            return with.Count == 0
                ? $"Added {components} from KiCad."
                : $"Added {components} from KiCad - {string.Join(", ", with)}.";
        }

        // ###########################################################################################
        // The matches suggested under Board status's Schematic images (slice 3d): each pair, its "Use",
        // the button for all of them, and what taking them did.
        // ###########################################################################################
        public const string MatchesHeading = "Suggested from the names of the pictures and their files:";

        public const string UseMatch = "Use";

        public static string SuggestedMatch(SchematicCadNameSuggestion match) => $"[{match.SchematicName}] shows [{match.CadName}]";

        public static string UseMatchTip(SchematicCadNameSuggestion match) => $"Give [{match.SchematicName}] the KiCad view [{match.CadName}]";

        public static string UseAllMatches(int count) => $"Use all {BoardReadiness.Count(count)} suggestions";

        public static string MatchesUsed(IReadOnlyList<SchematicCadNameSuggestion> used) =>
            used.Count switch
            {
                0 => "Those pictures, or their KiCad views, have been matched already.",
                1 => $"[{used[0].SchematicName}] now shows [{used[0].CadName}].",
                _ => $"{BoardReadiness.Count(used.Count)} pictures now show their KiCad views.",
            };

        public const string KiCadNotRead = "The board's KiCad data could not be read, so nothing was added.";

        public const string KiCadUnreadable =
            "The KiCad files in your draft could not be read - open \"KiCad: Files\" to see which are there, and import the project again.";

        // ---- "Board status" -------------------------------------------------------------------------

        // The panels' buttons - each opens where the part is seen to; a plain button (CLAUDE.md, button
        // colors), so none says "Add" but the one that adds.
        public const string TableAction = "Show in the table";
        public const string PicturesAction = "Open schematic images";
        public const string KiCadAction = "Open KiCad files";
        public const string SignalsAction = "Open important signals";
        public const string CreditsAction = "Open credits";
        public const string FilesAction = "Open files";
        public const string LinksAction = "Open links";

        // A list with rows missing a field they need (BoardRowRequired) - errors, so said as such.
        public static string IncompleteGap(int rows, string what, string whats) => rows == 1
            ? $"1 {what} has a field it needs left empty - fill it in, or the draft cannot be submitted."
            : $"{BoardReadiness.Count(rows)} {whats} have fields they need left empty - fill them in, or the draft cannot be submitted.";

        public const string NoPicturesGap = "No schematic images yet - photos of both sides and the schematic sheets, which everything else is marked on.";

        public const string NoSignalsGap = "No important KiCad signals yet - the signals worth checking first, each lighting up its KiCad net.";

        public const string NoCreditsGap = "Nobody is credited for the board's data yet.";

        // The headings of "What's missing", counted.
        public static string NotInKiCadHeading(int count) => $"Not in the KiCad data - their traces cannot light up ({BoardReadiness.Count(count)})";

        public static string WithoutDescriptionHeading(int count) => $"No short description ({BoardReadiness.Count(count)})";

        // The warning "component.no_highlight" is, said once for them all: each component is a button.
        public static string UnmarkedHeading(int count) =>
            $"Warning - not marked on any schematic image, so CRT cannot point them out on the board ({BoardReadiness.Count(count)})";

        public static string PinsWithoutBaselineHeading(int count) => $"Pins listed, but no scope baseline ({BoardReadiness.Count(count)})";

        // ###########################################################################################
        // Under a "What's missing" group naming its first ListedAtMost: how many more there are - ""
        // for none. Nothing more: the list on the left shows only the region CRT shows, while a group
        // counts every region (code review, 2026-10-26).
        // ###########################################################################################
        public static string MoreThanListed(int count) =>
            count > BoardReadiness.ListedAtMost ? $"and {BoardReadiness.Count(count - BoardReadiness.ListedAtMost)} more" : string.Empty;

        // ---- the rules ----------------------------------------------------------------------------

        // The board's component labels, trimmed, once each, in its order.
        internal static List<string> Labels(BoardData board) =>
            BoardReadiness.Distinct(board.Components.Select(row => row.BoardLabel));

        // How many of the KiCad data's components the board has none of - capitals ignored.
        internal static int NotOnBoard(BoardData board, IEnumerable<string> references)
        {
            var onBoard = new HashSet<string>(BoardReadiness.Labels(board), StringComparer.OrdinalIgnoreCase);

            return BoardReadiness.Distinct(references).Count(reference => !onBoard.Contains(reference));
        }

        private static List<string> Distinct(IEnumerable<string?> names) =>
            names
                .Select(name => name?.Trim() ?? string.Empty)
                .Where(name => name.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

        // The labels marked on no picture.
        private static List<string> UnmarkedOf(BoardData board, IReadOnlyList<string> labels)
        {
            var marked = new HashSet<string>(board.ComponentHighlights.Select(row => row.BoardLabel?.Trim() ?? string.Empty), StringComparer.OrdinalIgnoreCase);
            return labels.Where(label => !marked.Contains(label)).ToList();
        }

        private static string Count(int count) => count.ToString(CultureInfo.InvariantCulture);
    }
}
