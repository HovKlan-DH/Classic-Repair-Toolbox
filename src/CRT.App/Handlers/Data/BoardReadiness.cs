using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // The steps of getting a new board ready, in the plan's order.
    internal enum ReadinessStepKind
    {
        Name,
        Pictures,
        KiCad,
        Match,
        Align,
        Components,
        Mark,
        Signals,
        Submit,
    }

    internal enum ReadinessState
    {
        // Nothing more to do here.
        Done,

        // The first step with something to do - where to go on.
        Next,

        // Something to do, after the next step.
        Later,

        // Worth doing, never needed - the KiCad project.
        Optional,

        // Cannot be done without a step that was left out - matching and aligning without KiCad data.
        Skipped,
    }

    // One row of the checklist: its number, what it is, how far along, and what is left - for the
    // components, how many the KiCad data has that the board does not (its button adds them), and for
    // matching, the pictures and views that seem to go together (slice 3d - offered under the step).
    internal sealed record ReadinessStep(
        ReadinessStepKind Kind,
        int Number,
        ReadinessState State,
        string Title,
        string Detail,
        int FromKiCad = 0,
        IReadOnlyList<SchematicCadNameSuggestion>? Matches = null);

    // ###########################################################################################
    // What the board's KiCad data says: the labels of its components - the very ones "Add N from
    // KiCad" adds (KiCadComponentTable: its first PCB file's footprints, no placeholder, logo or
    // mounting hole) - its view names, which of them are a side of the board, and every view as
    // matching a picture sees it (ToMatch - its side, and its PCB file's name). Not Read: the KiCad
    // files could not be read - no data, never data with nothing in it (code review, 2026-10-26).
    // ###########################################################################################
    internal sealed record BoardKiCadFacts(
        IReadOnlyList<string> References,
        IReadOnlyList<string> Views,
        bool Read = true,
        IReadOnlyList<string>? BoardViews = null,
        IReadOnlyList<KiCadViewToMatch>? ToMatch = null)
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

            return new BoardKiCadFacts(
                KiCadComponentTable.Of(root).Select(component => component.Label).ToList(),
                views.Select(view => view.Name).ToList(),
                BoardViews: views.Where(view => view.Side != KiCadViewSide.Sheet).Select(view => view.Name).ToList(),
                ToMatch: views);
        }

        private static KiCadViewSide SideOf(KiCadProjectView view) =>
            string.Equals(view.Type, "pcb_top", StringComparison.OrdinalIgnoreCase) ? KiCadViewSide.Top
            : string.Equals(view.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase) ? KiCadViewSide.Bottom
            : KiCadViewSide.Sheet;
    }

    // What an existing board still lacks - components by their labels, in the board's order.
    internal sealed record BoardMissing(
        IReadOnlyList<string> WithoutDescription,
        IReadOnlyList<string> Unmarked,
        IReadOnlyList<string> PinsWithoutBaseline)
    {
        public bool IsEmpty => this.WithoutDescription.Count == 0 && this.Unmarked.Count == 0 && this.PinsWithoutBaseline.Count == 0;
    }

    // ###########################################################################################
    // GETTING A BOARD READY (CRT 4.0 step 3, slice 3a - the plan: "Getting this board ready:
    // 1 Name the board, 2 Board pictures, 3 KiCad project (optional), 4 Match pictures to KiCad views,
    // 5 Align KiCad on each picture, 6 Components - 214 found in KiCad [Add], 7 Mark components on the
    // pictures, 8 Important signals, 9 Check and submit ... On an existing board, the same panel shows
    // What's missing: components with no description, unmarked components, pins with no baseline").
    //
    // Worked out from the board as its draft has it - its rows, its KiCad alignments - whether it has
    // KiCad files, and what its KiCad data says (null while that is still being read):
    //
    //   - A STEP IS DONE by what the board holds, never by a box ticked: pictures listed, a picture
    //     given a KiCad view, the components KiCad has on the board, every component marked.
    //   - THE KICAD PROJECT IS OPTIONAL; without it, matching and aligning are skipped and the
    //     components are counted on their own.
    //   - THE FIRST STEP WITH SOMETHING TO DO IS "NEXT"; "Check and submit" is next once nothing else
    //     is, and is never done - submitting is.
    //
    // PURE: the board in, the rows out.
    // ###########################################################################################
    internal static class BoardReadiness
    {
        // How many labels a "What's missing" list names before "and N more".
        internal const int ListedAtMost = 8;

        public static List<ReadinessStep> Steps(string boardName, BoardContent content, bool hasKiCadFiles, BoardKiCadFacts? kiCad)
        {
            ArgumentNullException.ThrowIfNull(content);

            BoardData board = content.Board;

            // KiCad files that could not be read are no KiCad data at all (code review, 2026-10-26: they
            // read as data with nothing in it - "All 0 are on the board", ticked).
            bool unreadable = hasKiCadFiles && kiCad is { Read: false };
            bool usable = hasKiCadFiles && !unreadable;
            BoardKiCadFacts? facts = usable ? kiCad : null;

            List<BoardSchematicEntry> pictures = board.Schematics.Where(row => !string.IsNullOrWhiteSpace(row.SchematicImageFile)).ToList();
            var views = new HashSet<string>(facts?.Views.Select(view => view.Trim()) ?? [], StringComparer.OrdinalIgnoreCase);

            List<BoardSchematicEntry> matched = pictures
                .Where(row => (row.CadName?.Trim() ?? string.Empty).Length > 0 && (facts is null || views.Contains(row.CadName!.Trim())))
                .ToList();

            var calibrated = new HashSet<string>(content.Calibrations.Select(row => row.SchematicName.Trim()), StringComparer.OrdinalIgnoreCase);
            int aligned = matched.Count(row => calibrated.Contains(row.SchematicName?.Trim() ?? string.Empty));

            List<string> labels = BoardReadiness.Labels(board);
            var onBoard = new HashSet<string>(labels, StringComparer.OrdinalIgnoreCase);
            List<string> fromKiCad = facts is null ? [] : BoardReadiness.Distinct(facts.References);
            int notOnBoard = fromKiCad.Count(reference => !onBoard.Contains(reference));
            int unmarked = BoardReadiness.UnmarkedOf(board, labels).Count;
            int signals = board.KiCadImportantSignals.Count;
            string needsKiCad = unreadable ? "Needs KiCad data CRT can read." : "Needs the KiCad project.";

            var steps = new List<(ReadinessStepKind Kind, Status Status, string Title, string Detail)>
            {
                (ReadinessStepKind.Name, Status.Done, "Name the board", boardName.Trim()),

                (ReadinessStepKind.Pictures, pictures.Count > 0 ? Status.Done : Status.ToDo,
                    pictures.Count == 0 ? "Board pictures" : $"Board pictures ({BoardReadiness.Count(pictures.Count)})",
                    pictures.Count == 0
                        ? "Photos of both sides, and the schematic sheets - the board everything else is marked on."
                        : "Add more, name them, or put them in order."),

                (ReadinessStepKind.KiCad, unreadable ? Status.ToDo : hasKiCadFiles ? Status.Done : Status.Optional, "KiCad project (optional)",
                    unreadable
                        ? BoardReadiness.KiCadUnreadable
                        : hasKiCadFiles
                            ? "The board's KiCad data is in your draft."
                            : "With the board's KiCad project, its traces light up and its components can be added in one go."),
            };

            // The pictures and views that seem to go together (slice 3d) - once the views are known.
            List<SchematicCadNameSuggestion> suggested = facts is null ? [] : SchematicCadNameSuggestions.For(board.Schematics, facts.ToMatch ?? []);

            if (!usable)
            {
                steps.Add((ReadinessStepKind.Match, Status.Skipped, "Match pictures to KiCad views", needsKiCad));
                steps.Add((ReadinessStepKind.Align, Status.Skipped, "Align KiCad on each picture", needsKiCad));
            }
            else
            {
                Status match = BoardReadiness.MatchStatus(pictures.Count, matched.Count, views.Count, read: facts is not null);
                string shown = $"{BoardReadiness.Count(matched.Count)} of {BoardReadiness.Count(pictures.Count)} pictures show a KiCad view";

                steps.Add((ReadinessStepKind.Match, match, "Match pictures to KiCad views",
                    facts is null
                        ? "Reading the KiCad data..."
                        : match switch
                        {
                            Status.Done => $"{shown}.",
                            _ when suggested.Count > 0 => $"{shown} - check the matches suggested below.",
                            Status.Optional => $"{shown} - name the view of any other picture that shows one.",
                            _ => $"{shown} - name the view each one shows.",
                        }));

                steps.Add((ReadinessStepKind.Align, matched.Count > 0 && aligned == matched.Count ? Status.Done : Status.ToDo, "Align KiCad on each picture",
                    matched.Count == 0
                        ? "Once a picture shows a KiCad view, line the traces up on it."
                        : $"{BoardReadiness.Count(aligned)} of {BoardReadiness.Count(matched.Count)} pictures aligned."));
            }

            steps.Add(facts is not null
                ? (ReadinessStepKind.Components, labels.Count > 0 && notOnBoard == 0 ? Status.Done : Status.ToDo,
                    $"Components - {BoardReadiness.Count(fromKiCad.Count)} found in KiCad",
                    notOnBoard == 0
                        ? $"All {BoardReadiness.Count(fromKiCad.Count)} are on the board."
                        : $"{BoardReadiness.Count(notOnBoard)} of them are not on the board yet.")
                : (ReadinessStepKind.Components, labels.Count > 0 ? Status.Done : Status.ToDo,
                    labels.Count == 0 ? "Components" : $"Components ({BoardReadiness.Count(labels.Count)})",
                    labels.Count == 0 ? "The chips and parts, each by the label printed on the board." : "Add any the board still lacks."));

            // A picture aligned with a SIDE OF THE BOARD can have them all drawn at once (step 3, slice 3c)
            // - said where there is one, once the KiCad data says which views those are (code review,
            // 2026-10-26: one aligned with a schematic sheet was counted, and the button refused).
            var boardViews = new HashSet<string>(facts?.BoardViews ?? [], StringComparer.OrdinalIgnoreCase);
            bool markable = matched.Any(row => boardViews.Contains(row.CadName!.Trim()) && calibrated.Contains(row.SchematicName?.Trim() ?? string.Empty));
            string automatically = markable ? " A picture aligned with KiCad draws them all with Mark all automatically." : string.Empty;

            steps.Add((ReadinessStepKind.Mark, labels.Count > 0 && unmarked == 0 ? Status.Done : Status.ToDo, "Mark components on the pictures",
                labels.Count == 0
                    ? "Once there are components, draw where each one is on the pictures."
                    : unmarked == 0
                        ? "Every component is marked."
                        : $"{BoardReadiness.Count(unmarked)} of {BoardReadiness.Count(labels.Count)} components are not marked on any picture yet.{automatically}"));

            steps.Add(usable
                ? (ReadinessStepKind.Signals, signals > 0 ? Status.Done : Status.ToDo, "Important signals",
                    signals > 0
                        ? $"{BoardReadiness.Count(signals)} - add any other worth checking first."
                        : "The signals worth checking first, each lighting up its KiCad net.")
                : (ReadinessStepKind.Signals, Status.Skipped, "Important signals", needsKiCad));

            List<ReadinessStep> result = BoardReadiness.WithStates(steps);

            return facts is not null
                ? result
                    .Select(step => step.Kind switch
                    {
                        ReadinessStepKind.Components => step with { FromKiCad = notOnBoard },
                        ReadinessStepKind.Match when suggested.Count > 0 => step with { Matches = suggested },
                        _ => step,
                    })
                    .ToList()
                : result;
        }

        // ###########################################################################################
        // Whether matching is done. Nothing matched is to do. While the KiCad data is read, every
        // picture matched is done. Once it is read: every picture matched, or every view, is done -
        // and anything between is worth finishing, never needed (code review, 2026-10-26): CRT cannot
        // know whether a picture shows a view - a photo of the ports shows none - so a board with such
        // a picture and a view no picture shows left matching "next" for ever.
        // ###########################################################################################
        private static Status MatchStatus(int pictures, int matched, int views, bool read)
        {
            if (matched == 0)
            {
                return Status.ToDo;
            }

            if (!read)
            {
                return matched >= pictures ? Status.Done : Status.ToDo;
            }

            return matched >= pictures || matched >= views ? Status.Done : Status.Optional;
        }

        // ###########################################################################################
        // What an existing board still lacks: components with no short description, components marked
        // on no picture, and components whose pins are listed but which have no scope baseline - each
        // by its label, once, in the board's order.
        // ###########################################################################################
        public static BoardMissing Missing(BoardData board)
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

            return new BoardMissing(withoutDescription, BoardReadiness.UnmarkedOf(board, labels), pinsWithoutBaseline);
        }

        // ###########################################################################################
        // The button beside a step - what it opens - or "" for none: naming is done, and a step that
        // needs the KiCad project has nothing to open without it.
        // ###########################################################################################
        public static string ActionOf(ReadinessStep step)
        {
            ArgumentNullException.ThrowIfNull(step);

            if (step.State == ReadinessState.Skipped)
            {
                return string.Empty;
            }

            // The plan's "Components - 214 found in KiCad [Add]": every one the board lacks, in one go.
            if (step.Kind == ReadinessStepKind.Components && step.FromKiCad > 0)
            {
                return $"Add {BoardReadiness.Count(step.FromKiCad)} from KiCad";
            }

            return step.Kind switch
            {
                ReadinessStepKind.Pictures => "Pictures",
                ReadinessStepKind.KiCad => "KiCad data",
                ReadinessStepKind.Match => "Match pictures",
                ReadinessStepKind.Align => "Align KiCad",
                ReadinessStepKind.Components => "Add component",
                ReadinessStepKind.Mark => "Mark components",
                ReadinessStepKind.Signals => "Important signals",
                ReadinessStepKind.Submit => "Submit",
                _ => string.Empty,
            };
        }

        // ###########################################################################################
        // What "Add N from KiCad" did, said under the checklist: how many came, and how many of them
        // with their pins and a datasheet link.
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
        // The matches suggested under "Match pictures to KiCad views" (slice 3d): each pair, its "Use",
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
            "The KiCad files in your draft could not be read - open KiCad data to see which are there, and import the project again.";

        // The headings of "What's missing", counted.
        public static string WithoutDescriptionHeading(int count) => $"No short description ({BoardReadiness.Count(count)})";

        public static string UnmarkedHeading(int count) => $"Not marked on any picture ({BoardReadiness.Count(count)})";

        public static string PinsWithoutBaselineHeading(int count) => $"Pins listed, but no scope baseline ({BoardReadiness.Count(count)})";

        public const string NothingMissing =
            "Nothing - every component has a short description and is marked on a picture, and every one with its pins listed has a scope baseline.";

        // ###########################################################################################
        // Under a "What's missing" group naming its first ListedAtMost: how many more there are - ""
        // for none. Nothing more: the list on the left shows only the region CRT shows, while a group
        // counts every region (code review, 2026-10-26).
        // ###########################################################################################
        public static string MoreThanListed(int count) =>
            count > BoardReadiness.ListedAtMost ? $"and {BoardReadiness.Count(count - BoardReadiness.ListedAtMost)} more" : string.Empty;

        // ---- the rules ----------------------------------------------------------------------------

        // How far along a step is, before which step is next is known.
        private enum Status
        {
            Done,
            ToDo,
            Optional,
            Skipped,
        }

        // Next is the first step with something to do; "Check and submit" is next once nothing is.
        private static List<ReadinessStep> WithStates(List<(ReadinessStepKind Kind, Status Status, string Title, string Detail)> steps)
        {
            var result = new List<ReadinessStep>();
            bool nextGiven = false;

            for (int index = 0; index < steps.Count; index++)
            {
                (ReadinessStepKind kind, Status status, string title, string detail) = steps[index];

                ReadinessState state = status switch
                {
                    Status.Done => ReadinessState.Done,
                    Status.Optional => ReadinessState.Optional,
                    Status.Skipped => ReadinessState.Skipped,
                    _ when !nextGiven => ReadinessState.Next,
                    _ => ReadinessState.Later,
                };

                nextGiven |= state == ReadinessState.Next;
                result.Add(new ReadinessStep(kind, index + 1, state, title, detail));
            }

            result.Add(new ReadinessStep(
                ReadinessStepKind.Submit,
                steps.Count + 1,
                nextGiven ? ReadinessState.Later : ReadinessState.Next,
                "Check and submit",
                nextGiven
                    ? "Send it whenever you like - what is left can come in a later submission."
                    : "Everything is in place - send the board for review."));

            return result;
        }

        // The board's component labels, trimmed, once each, in its order.
        private static List<string> Labels(BoardData board) =>
            BoardReadiness.Distinct(board.Components.Select(row => row.BoardLabel));

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
