using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // Each panel of "Board status", in the order the page shows them - the entries of the left column.
    // Problems first, and only while there are any.
    internal enum BoardStatusPart
    {
        Problems,
        Name,
        Regions,
        KiCad,
        Signals,
        Pictures,
        Categories,
        Components,
        Files,
        Links,
        Credits,
    }

    internal enum BoardStatusState
    {
        // Nothing more to do here.
        Done,

        // Something is missing.
        ToDo,

        // Something stops the draft being submitted, or cannot be read.
        Error,

        // Worth having, never needed - nothing to say either way yet.
        Optional,
    }

    // One line of a panel's status: `Depth` under the line above it, and marked when something is off.
    internal sealed record BoardStatusLine(string Text, int Depth = 0, BoardStatusState? Tone = null);

    // One panel: its headline, what it is about, its button, how far it is, and what it counts.
    // TableSheet: the sheet of the table that shows its errors and warnings, when it has any.
    internal sealed record BoardStatusPanel(
        BoardStatusPart Part,
        BoardStatusState State,
        string Title,
        string Explanation,
        string Action,
        IReadOnlyList<BoardStatusLine> Lines,
        string? TableSheet = null);

    // ###########################################################################################
    // "BOARD STATUS", A PANEL PER PART OF THE BOARD (owner request, 2026-11-01: "each point in its own
    // panel with a headline, then an explanation what this is about, a button for it and lastly also a
    // status where it will count the data for it, if any. Maybe some coloring and the checkmark icon or
    // a 'missing to do' icon"). The same page for a new board and a published one: the checklist and
    // "What's missing" were two pages saying this in two ways.
    //
    //   - DONE IS WHAT THE BOARD HOLDS, never a box ticked (BoardReadiness' rule): pictures there and
    //     aligned, every component marked, a credit named.
    //   - THE KiCad PROJECT IS OPTIONAL: without it the panel is neither done nor missing, and nothing
    //     else is held against the board for it.
    //   - EVERY ERROR AND WARNING IS SAID ON THE PANEL IT IS ABOUT (owner request, 2026-11-01: "it should
    //     have the errors and warnings directly per section"), by the sheet it is on (PartOf); only what
    //     belongs to no panel keeps a panel of its own above the rest, shown only while there is any.
    //   - "NOT KNOWN YET" IS SAID: the KiCad data is read a moment after the rest (`kiCadFiles` null, or
    //     `kiCad` null with files), and no count is guessed meanwhile.
    //
    // PURE: the board and what is known of its KiCad data in, the panels out.
    // ###########################################################################################
    internal static class BoardStatusPage
    {
        public const string NameTitle = "Board: Name";
        public const string ProblemsTitle = "Errors and warnings";
        public const string ComponentsTitle = "Components";
        public const string FilesTitle = "Board: Files";
        public const string LinksTitle = "Board: Links";
        public const string CreditsTitle = "Board: Credits";
        public const string KiCadTitle = "KiCad: Files";

        public const string NameAction = "Open board name";
        public const string RegionsAction = "Open regions";
        public const string CategoriesAction = "Open categories";
        public const string ComponentsAction = "Open components";

        public const string ReadingKiCad = "Reading the KiCad data...";

        public static List<BoardStatusPanel> Panels(
            BoardData board,
            BoardMissing missing,
            IReadOnlyList<KiCadCalibrationEntry>? calibrations,
            IReadOnlyList<string>? kiCadFiles,
            BoardKiCadFacts? kiCad,
            IReadOnlyList<BoardDataProblem>? problems = null)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(missing);

            bool hasFiles = kiCadFiles is { Count: > 0 };
            bool unreadable = hasFiles && kiCad is { Read: false };
            BoardKiCadFacts? facts = hasFiles && !unreadable ? kiCad : null;

            var panels = new List<BoardStatusPanel>();

            panels.Add(new BoardStatusPanel(
                BoardStatusPart.Name,
                BoardStatusState.Done,
                NameTitle,
                "The manufacturer, hardware and board it is listed under in the lists at the top left.",
                NameAction,
                []));

            panels.Add(BoardStatusPage.RegionsPanel(board));
            panels.Add(BoardStatusPage.KiCadPanel(board, kiCadFiles, facts, unreadable));
            panels.Add(BoardStatusPage.SignalsPanel(board, missing, kiCadFiles, facts, unreadable));
            panels.Add(BoardStatusPage.PicturesPanel(board, calibrations, facts));
            panels.Add(BoardStatusPage.CategoriesPanel(board));
            panels.Add(BoardStatusPage.ComponentsPanel(board, facts));

            panels.Add(BoardStatusPage.ListPanel(
                BoardStatusPart.Files, FilesTitle, "Manuals, notes and other documents about the board.", BoardReadiness.FilesAction,
                board.BoardLocalFiles.Count, "board file", "board files",
                board.BoardLocalFiles.Select(row => row.Category),
                missing, BoardGapKind.IncompleteFiles));

            panels.Add(BoardStatusPage.ListPanel(
                BoardStatusPart.Links, LinksTitle, "Pages about the board - a forum, a wiki, a repair guide, a video.", BoardReadiness.LinksAction,
                board.BoardLinks.Count, "board link", "board links",
                board.BoardLinks.Select(row => row.Category),
                missing, BoardGapKind.IncompleteLinks));

            panels.Add(BoardStatusPage.CreditsPanel(board, missing));

            return BoardStatusPage.WithProblems(panels, problems ?? []);
        }

        // ###########################################################################################
        // Which panel a problem is about: its sheet's - a highlight, which is on no sheet, the images'.
        // What belongs to none is the panel of its own's.
        // ###########################################################################################
        private static BoardStatusPart PartOf(BoardDataProblem problem) => problem.Sheet switch
        {
            _ when problem.Code == "component.no_category" => BoardStatusPart.Categories,
            _ when problem.Code == "regions.unknown" => BoardStatusPart.Regions,
            BoardWorkbookSchema.SheetBoardSchematics => BoardStatusPart.Pictures,
            BoardWorkbookSchema.SheetComponents or
            BoardWorkbookSchema.SheetComponentImages or
            BoardWorkbookSchema.SheetComponentLocalFiles or
            BoardWorkbookSchema.SheetComponentLinks or
            BoardWorkbookSchema.SheetComponentPins or
            BoardWorkbookSchema.SheetComponentReplacements => BoardStatusPart.Components,
            BoardWorkbookSchema.SheetBoardLocalFiles => BoardStatusPart.Files,
            BoardWorkbookSchema.SheetBoardLinks => BoardStatusPart.Links,
            BoardWorkbookSchema.SheetCredits => BoardStatusPart.Credits,
            BoardWorkbookSchema.SheetKiCadImportantSignals => BoardStatusPart.Signals,
            _ when problem.Code.StartsWith("highlight.", StringComparison.Ordinal) => BoardStatusPart.Pictures,
            _ => BoardStatusPart.Problems,
        };

        // Said another way on the page, so not again as a line: a list's "has a field it needs left
        // empty" line - a board file's empty file is one of those fields (BoardRowRequired), under a code
        // of its own - and the buttons of the components marked on no image (the tab's group).
        private static bool IsSaidElsewhere(BoardDataProblem problem) =>
            problem.Code is "row.required_empty" or "component.no_highlight" ||
            (problem.Code == "file.unreferenced" && problem.Sheet == BoardWorkbookSchema.SheetBoardLocalFiles);

        private static List<BoardStatusPanel> WithProblems(List<BoardStatusPanel> panels, IReadOnlyList<BoardDataProblem> problems)
        {
            ILookup<BoardStatusPart, BoardDataProblem> byPart = problems.ToLookup(BoardStatusPage.PartOf);

            List<BoardStatusPanel> shown = panels.Select(panel => BoardStatusPage.WithProblemsOf(panel, byPart[panel.Part])).ToList();

            if (byPart[BoardStatusPart.Problems].Any())
            {
                shown.Insert(0, BoardStatusPage.WithProblemsOf(
                    new BoardStatusPanel(
                        BoardStatusPart.Problems,
                        BoardStatusState.ToDo,
                        ProblemsTitle,
                        "What the draft's checks found that is about none of the parts below. An error stops the draft being submitted; a warning does not.",
                        BoardReadiness.TableAction,
                        []),
                    byPart[BoardStatusPart.Problems]));
            }

            return shown;
        }

        // ###########################################################################################
        // `panel` saying `problems`, errors first - the first ListedAtMost in full, the rest counted, the
        // table a button away. An error makes the panel one; a warning leaves nothing done or optional.
        // ###########################################################################################
        private static BoardStatusPanel WithProblemsOf(BoardStatusPanel panel, IEnumerable<BoardDataProblem> problems)
        {
            List<BoardDataProblem> all = problems.OrderByDescending(problem => problem.Level).ToList();

            if (all.Count == 0)
            {
                return panel;
            }

            List<BoardDataProblem> listed = all.Where(problem => !BoardStatusPage.IsSaidElsewhere(problem)).ToList();
            List<BoardStatusLine> lines = panel.Lines.ToList();

            lines.AddRange(listed
                .Take(BoardReadiness.ListedAtMost)
                .Select(problem => problem.Level == BoardProblemLevel.Error
                    ? new BoardStatusLine($"Error - {problem.Message}", Tone: BoardStatusState.Error)
                    : new BoardStatusLine($"Warning - {problem.Message}", Tone: BoardStatusState.ToDo)));

            string more = BoardReadiness.MoreThanListed(listed.Count);

            if (more.Length > 0)
            {
                lines.Add(new BoardStatusLine(more));
            }

            BoardStatusState state =
                all.Any(problem => problem.Level == BoardProblemLevel.Error) ? BoardStatusState.Error
                : panel.State == BoardStatusState.Error ? BoardStatusState.Error
                : BoardStatusState.ToDo;

            return panel with
            {
                State = state,
                Lines = lines,
                TableSheet = all.Select(problem => problem.Sheet).FirstOrDefault(sheet => sheet is not null),
            };
        }

        // Whether the board comes in PAL and NTSC versions - either is an answer, so nothing to do.
        private static BoardStatusPanel RegionsPanel(BoardData board)
        {
            bool has = BoardRegions.Has(board);

            return new BoardStatusPanel(
                BoardStatusPart.Regions,
                has ? BoardStatusState.Done : BoardStatusState.Optional,
                ComponentPaneWording.Title(PaneBoardPart.Regions),
                "Whether the board came in a PAL and an NTSC version, so what a component has can differ between them.",
                RegionsAction,
                [new BoardStatusLine(has ? ComponentPaneWording.RegionsHasOnBoard : "One version - no regions.")]);
        }

        // What the components are sorted under. The ones with none are its error (PartOf).
        private static BoardStatusPanel CategoriesPanel(BoardData board)
        {
            int categories = BoardCategoryEdits.Of(board).Count(row => row.Name.Length > 0);

            return new BoardStatusPanel(
                BoardStatusPart.Categories,
                categories > 0 ? BoardStatusState.Done : BoardStatusState.Optional,
                ComponentPaneWording.Title(PaneBoardPart.Categories),
                "What the board's components are sorted under.",
                CategoriesAction,
                [new BoardStatusLine(categories == 0 ? "No categories yet." : Counted(categories, "category", "categories"))]);
        }

        private static BoardStatusPanel KiCadPanel(BoardData board, IReadOnlyList<string>? kiCadFiles, BoardKiCadFacts? facts, bool unreadable)
        {
            var lines = new List<BoardStatusLine>();

            if (kiCadFiles is null)
            {
                lines.Add(new BoardStatusLine(ReadingKiCad));
            }
            else if (kiCadFiles.Count == 0)
            {
                lines.Add(new BoardStatusLine("No KiCad data - the board works without it."));
            }
            else
            {
                lines.Add(new BoardStatusLine(Counted(kiCadFiles.Count, "KiCad file imported", "KiCad files imported")));

                (int pcb, int schematic, int project, int other) = KiCadImportSummary.KindsOf(kiCadFiles);

                if (project > 0)
                {
                    lines.Add(new BoardStatusLine(Counted(project, "project file", "project files"), 1));
                }

                if (pcb > 0)
                {
                    // A PCB file is always read as two views, its top and its bottom - whatever a side holds.
                    lines.Add(new BoardStatusLine(
                        pcb == 1
                            ? "1 PCB file - 2 images, its top and its bottom"
                            : $"{Number(pcb)} PCB files - {Number(pcb * 2)} images, the top and the bottom of each",
                        1));
                }

                if (schematic > 0)
                {
                    lines.Add(new BoardStatusLine(Counted(schematic, "schematic file", "schematic files"), 1));
                }

                if (other > 0)
                {
                    lines.Add(new BoardStatusLine(Counted(other, "other file", "other files"), 1));
                }

                if (unreadable)
                {
                    lines.Add(new BoardStatusLine(BoardReadiness.KiCadUnreadable, Tone: BoardStatusState.Error));
                }
                else if (facts is null)
                {
                    lines.Add(new BoardStatusLine(ReadingKiCad));
                }
                else
                {
                    lines.Add(new BoardStatusLine(facts.Markable == 1
                        ? "1 component can be marked automatically"
                        : $"{Number(facts.Markable)} components can be marked automatically"));

                    // What "KiCad data" lists as not on the board yet (owner request, 2026-11-01).
                    int notOnBoard = BoardReadiness.NotOnBoard(board, facts.References);

                    if (notOnBoard > 0)
                    {
                        lines.Add(new BoardStatusLine(
                            notOnBoard == 1 ? "1 component is not on the board yet" : $"{Number(notOnBoard)} components are not on the board yet",
                            Tone: BoardStatusState.ToDo));
                    }

                    int byHand = facts.OnlyInSchematic?.Count ?? 0;

                    if (byHand > 0)
                    {
                        lines.Add(new BoardStatusLine(
                            byHand == 1
                                ? "1 component known by KiCad needs marking by hand - it is on no PCB file"
                                : $"{Number(byHand)} components known by KiCad need marking by hand - they are on no PCB file",
                            Tone: BoardStatusState.ToDo));
                    }

                    lines.Add(new BoardStatusLine($"{Counted(facts.Nets?.Names.Count ?? 0, "net", "nets")} available"));
                    lines.Add(new BoardStatusLine($"{Counted(BoardStatusPage.NetsUsed(board, facts), "net", "nets")} used in \"{ComponentPaneWording.ImportantSignals}\"", 1));
                }
            }

            BoardStatusState state = unreadable
                ? BoardStatusState.Error
                : kiCadFiles is { Count: > 0 } ? BoardStatusState.Done : BoardStatusState.Optional;

            return new BoardStatusPanel(
                BoardStatusPart.KiCad,
                state,
                KiCadTitle,
                "The board's KiCad project, if it has one. With it the traces light up on the pictures, and its components can be added and marked for you.",
                BoardReadiness.KiCadAction,
                lines);
        }

        // Counted in the owner's own words (2026-11-01: "25 images known by KiCad data", "4 images added",
        // "2 KiCad files mapped", "1 KiCad file aligned"), though a picture card says "Shows KiCad view". What
        // is counted is the views, so the two sides of one PCB file are two.
        private static BoardStatusPanel PicturesPanel(BoardData board, IReadOnlyList<KiCadCalibrationEntry>? calibrations, BoardKiCadFacts? facts)
        {
            List<BoardSchematicEntry> pictures = board.Schematics.Where(row => !string.IsNullOrWhiteSpace(row.SchematicImageFile)).ToList();
            var lines = new List<BoardStatusLine>();

            if (facts is not null)
            {
                lines.Add(new BoardStatusLine($"{Counted(facts.Views.Count, "image", "images")} known by KiCad data"));
            }

            int matched = 0;
            int aligned = 0;

            if (pictures.Count == 0)
            {
                lines.Add(new BoardStatusLine(BoardReadiness.NoPicturesGap, Tone: BoardStatusState.ToDo));
            }
            else
            {
                lines.Add(new BoardStatusLine($"{Counted(pictures.Count, "image", "images")} added"));

                if (facts is not null)
                {
                    var views = new HashSet<string>(facts.Views.Select(view => view.Trim()), StringComparer.OrdinalIgnoreCase);
                    var calibrated = new HashSet<string>(calibrations?.Select(row => row.SchematicName.Trim()) ?? [], StringComparer.OrdinalIgnoreCase);

                    List<BoardSchematicEntry> showing = pictures.Where(row => row.CadName?.Trim() is { Length: > 0 } view && views.Contains(view)).ToList();
                    matched = showing.Count;
                    aligned = showing.Count(row => calibrated.Contains(row.SchematicName?.Trim() ?? string.Empty));

                    lines.Add(new BoardStatusLine(
                        $"{Counted(matched, "KiCad file", "KiCad files")} mapped",
                        1,
                        matched == 0 ? BoardStatusState.ToDo : null));

                    if (matched > 0)
                    {
                        lines.Add(new BoardStatusLine($"{Counted(aligned, "KiCad file", "KiCad files")} aligned", 2));

                        if (aligned < matched)
                        {
                            lines.Add(new BoardStatusLine($"{Counted(matched - aligned, "KiCad file", "KiCad files")} not aligned yet", 2, BoardStatusState.ToDo));
                        }
                    }
                }
            }

            List<string> labels = BoardReadiness.Labels(board);
            var marked = new HashSet<string>(board.ComponentHighlights.Select(row => row.BoardLabel?.Trim() ?? string.Empty), StringComparer.OrdinalIgnoreCase);
            int markedCount = labels.Count(marked.Contains);

            if (labels.Count > 0)
            {
                lines.Add(new BoardStatusLine(
                    labels.Count == 1 ? $"{Number(markedCount)} of 1 component marked" : $"{Number(markedCount)} of {Number(labels.Count)} components marked",
                    Tone: markedCount < labels.Count ? BoardStatusState.ToDo : null));
            }

            bool done = pictures.Count > 0 &&
                        (facts is null || facts.Views.Count == 0 || matched > 0) &&
                        aligned == matched &&
                        markedCount == labels.Count;

            return new BoardStatusPanel(
                BoardStatusPart.Pictures,
                done ? BoardStatusState.Done : BoardStatusState.ToDo,
                ComponentPaneWording.SchematicImages,
                "Photos of both sides of the board and its schematic sheets - everything else is marked on them.",
                BoardReadiness.PicturesAction,
                lines);
        }

        private static BoardStatusPanel SignalsPanel(BoardData board, BoardMissing missing, IReadOnlyList<string>? kiCadFiles, BoardKiCadFacts? facts, bool unreadable)
        {
            var lines = new List<BoardStatusLine>();
            int rows = board.KiCadImportantSignals.Count;
            bool noKiCad = kiCadFiles is { Count: 0 };

            if (facts is not null)
            {
                lines.Add(new BoardStatusLine($"{Counted(facts.Nets?.Names.Count ?? 0, "net", "nets")} available"));

                if (rows == 0)
                {
                    lines.Add(new BoardStatusLine(BoardReadiness.NoSignalsGap, Tone: BoardStatusState.ToDo));
                }
                else
                {
                    int used = BoardStatusPage.NetsUsed(board, facts);

                    // The names the Schematics tab's panel lists: only rows it finds a net for.
                    int names = board.KiCadImportantSignals
                        .Where(row => facts.Nets?.NormalizedOf(row.KiCadNetName) is not null)
                        .Select(row => row.DisplayName?.Trim() ?? string.Empty)
                        .Where(name => name.Length > 0)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .Count();

                    lines.Add(new BoardStatusLine(names == used
                        ? $"{Counted(used, "net", "nets")} used"
                        : $"{Counted(used, "net", "nets")} used, shown as {Counted(names, "name", "names")} on the \"Schematics\" tab"));

                    int notFound = board.KiCadImportantSignals.Count(row => facts.Nets?.Finds(row.KiCadNetName) == false);

                    if (notFound > 0)
                    {
                        lines.Add(new BoardStatusLine(
                            notFound == 1
                                ? "1 names a net the KiCad data does not have - it will not light up"
                                : $"{Number(notFound)} name a net the KiCad data does not have - they will not light up",
                            1,
                            BoardStatusState.ToDo));
                    }
                }
            }
            else if (noKiCad)
            {
                lines.Add(new BoardStatusLine(rows == 0 ? "They need the board's KiCad data." : Counted(rows, "important KiCad signal", "important KiCad signals")));
            }
            else if (unreadable)
            {
                lines.Add(new BoardStatusLine("They need KiCad data CRT can read."));
            }
            else
            {
                lines.Add(new BoardStatusLine(ReadingKiCad));
            }

            BoardGap? incomplete = BoardStatusPage.GapOf(missing, BoardGapKind.IncompleteSignals);

            if (incomplete is not null)
            {
                lines.Add(new BoardStatusLine(incomplete.Text, Tone: BoardStatusState.Error));
            }

            BoardStatusState state =
                incomplete is not null ? BoardStatusState.Error
                : facts is null ? BoardStatusState.Optional
                : rows == 0 ? BoardStatusState.ToDo
                : BoardStatusState.Done;

            return new BoardStatusPanel(
                BoardStatusPart.Signals,
                state,
                ComponentPaneWording.ImportantSignals,
                "The signals worth checking first - each lights up its KiCad net on the \"Schematics\" tab.",
                BoardReadiness.SignalsAction,
                lines);
        }

        // ###########################################################################################
        // The owner left this panel open (2026-11-01: "if you have any ideas, based on my proposal here,
        // then please implement it"). How many components, what the KiCad data has that the board does
        // not, and how much more they hold. What each still lacks is under them, a button per component
        // (BoardMissing's groups, drawn by the tab) - never counted a second time here.
        // ###########################################################################################
        private static BoardStatusPanel ComponentsPanel(BoardData board, BoardKiCadFacts? facts)
        {
            List<string> labels = BoardReadiness.Labels(board);
            int notOnBoard = facts is null ? 0 : BoardReadiness.NotOnBoard(board, facts.References);

            var lines = new List<BoardStatusLine>();

            if (labels.Count == 0)
            {
                lines.Add(new BoardStatusLine("No components yet.", Tone: BoardStatusState.ToDo));
            }
            else
            {
                lines.Add(new BoardStatusLine(Counted(labels.Count, "component", "components")));
            }

            int withPins = board.ComponentPins.Select(row => row.BoardLabel?.Trim() ?? string.Empty).Where(label => label.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            int baselines = board.ComponentImages.Count(ComponentImageKinds.IsScopeBaseline);
            int replacements = board.ComponentReplacements.Count;

            if (withPins > 0)
            {
                lines.Add(new BoardStatusLine(withPins == 1 ? "1 with its pins listed" : $"{Number(withPins)} with their pins listed", 1));
            }

            if (baselines > 0)
            {
                lines.Add(new BoardStatusLine(Counted(baselines, "scope baseline", "scope baselines"), 1));
            }

            if (replacements > 0)
            {
                lines.Add(new BoardStatusLine(Counted(replacements, "replacement part", "replacement parts"), 1));
            }

            if (notOnBoard > 0)
            {
                lines.Add(new BoardStatusLine(
                    notOnBoard == 1
                        ? "1 component in the KiCad data is not on the board yet - \"KiCad: Files\" adds it"
                        : $"{Number(notOnBoard)} components in the KiCad data are not on the board yet - \"KiCad: Files\" adds them",
                    Tone: BoardStatusState.ToDo));
            }

            return new BoardStatusPanel(
                BoardStatusPart.Components,
                labels.Count > 0 && notOnBoard == 0 ? BoardStatusState.Done : BoardStatusState.ToDo,
                ComponentsTitle,
                "The chips and parts on the board, each by the label printed on it.",
                ComponentsAction,
                lines);
        }

        // Files and Links: how many, in how many categories - and which have a field they need left empty.
        private static BoardStatusPanel ListPanel(
            BoardStatusPart part,
            string title,
            string explanation,
            string action,
            int rows,
            string one,
            string many,
            IEnumerable<string?> categories,
            BoardMissing missing,
            BoardGapKind incompleteKind)
        {
            var lines = new List<BoardStatusLine>();

            if (rows == 0)
            {
                lines.Add(new BoardStatusLine($"No {many} yet."));
            }
            else
            {
                int categoryCount = categories
                    .Select(category => category?.Trim() ?? string.Empty)
                    .Where(category => category.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count();

                lines.Add(new BoardStatusLine($"{Counted(rows, one, many)} defined"));
                lines.Add(new BoardStatusLine($"{Counted(categoryCount, "category", "categories")} defined"));
            }

            BoardGap? incomplete = BoardStatusPage.GapOf(missing, incompleteKind);

            if (incomplete is not null)
            {
                lines.Add(new BoardStatusLine(incomplete.Text, Tone: BoardStatusState.Error));
            }

            BoardStatusState state =
                incomplete is not null ? BoardStatusState.Error
                : rows == 0 ? BoardStatusState.Optional
                : BoardStatusState.Done;

            return new BoardStatusPanel(part, state, title, explanation, action, lines);
        }

        private static BoardStatusPanel CreditsPanel(BoardData board, BoardMissing missing)
        {
            var lines = new List<BoardStatusLine>();
            BoardGap? incomplete = BoardStatusPage.GapOf(missing, BoardGapKind.IncompleteCredits);

            lines.Add(board.Credits.Count == 0
                ? new BoardStatusLine(BoardReadiness.NoCreditsGap, Tone: BoardStatusState.ToDo)
                : new BoardStatusLine($"{Counted(board.Credits.Count, "credit", "credits")} defined"));

            if (incomplete is not null)
            {
                lines.Add(new BoardStatusLine(incomplete.Text, Tone: BoardStatusState.Error));
            }

            BoardStatusState state =
                incomplete is not null ? BoardStatusState.Error
                : board.Credits.Count == 0 ? BoardStatusState.ToDo
                : BoardStatusState.Done;

            return new BoardStatusPanel(
                BoardStatusPart.Credits,
                state,
                CreditsTitle,
                "Who the board's data is thanks to.",
                BoardReadiness.CreditsAction,
                lines);
        }

        // The board's nets its important signals name, each once, by the Schematics tab's own rule.
        private static int NetsUsed(BoardData board, BoardKiCadFacts facts) =>
            facts.Nets is not { } nets
                ? 0
                : board.KiCadImportantSignals
                    .Select(row => nets.NormalizedOf(row.KiCadNetName))
                    .OfType<string>()
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count();

        private static BoardGap? GapOf(BoardMissing missing, BoardGapKind kind) =>
            missing.BoardGaps.FirstOrDefault(gap => gap.Kind == kind);

        private static string Counted(int count, string one, string many) =>
            count == 1 ? $"1 {one}" : $"{Number(count)} {many}";

        private static string Number(int count) => count.ToString(CultureInfo.InvariantCulture);
    }
}
