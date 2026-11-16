using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A ROW OF THE BOARD'S OWN LISTS MUST HAVE FILLED IN (owner requests, 2026-10-30: "'What' and
    // 'Name or handle' must be set ... In 'Links' then all 3 fields must be filled ... Same for 'Files'
    // ... the 'Important signal name' must not be empty"; and 2026-11-01: "These errors must be flagged
    // in 'Board status' and then it should not be possible to submit the draft board"):
    //
    //   - a BOARD FILE: its category, its name and its file;
    //   - a BOARD LINK: its category, its name and its address;
    //   - a CREDIT: what it is for (its category) and the name or handle - which part of it and the
    //     contact may stay empty, as on many of the live boards' credits - and its main contribution,
    //     on a credit the board has not published as it stands (2026-11-14; PublishedCredits.Judged -
    //     each published credit standing for one row, 2026-11-15);
    //   - an IMPORTANT SIGNAL: its KiCad net and its name.
    //
    // *** ONE RULE, ASKED EVERYWHERE. *** BoardDataChecks makes each empty one an ERROR on its cell -
    // the table marks it, the draft's row counts it, Board status says it, Submit sends nothing while it
    // is there, and the server refuses it (an error in the table is a refusal at submit); the component
    // pane marks the box and puts the row into the draft as it is (BoardPartRequired names the fields
    // by its headings). Until 2026-11-01 the pane held such a row back, so nothing but the pane knew.
    //
    // Columns are named as the board's sheets name them (BoardWorkbookSchema), in the order a row shows
    // them. A blank value is an empty one.
    // ###########################################################################################
    public static class BoardRowRequired
    {
        public static IReadOnlyList<string> EmptyColumns(BoardLocalFileEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return BoardRowRequired.Empty(
                (row.Category, BoardWorkbookSchema.ColCategory),
                (row.Name, BoardWorkbookSchema.ColName),
                (row.File, BoardWorkbookSchema.ColFile));
        }

        public static IReadOnlyList<string> EmptyColumns(BoardLinkEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return BoardRowRequired.Empty(
                (row.Category, BoardWorkbookSchema.ColCategory),
                (row.Name, BoardWorkbookSchema.ColName),
                (row.Url, BoardWorkbookSchema.ColUrl));
        }

        // ###########################################################################################
        // A CREDIT: what it is for and the name or handle - and its MAIN CONTRIBUTION, unless the board has it published
        // as it stands (owner request, 2026-11-14: "This is a mandatory field"; owner decision: "New credits only" - none
        // of the credits published before it has one, and a contributor fixing a typo elsewhere must not be held to
        // writing them). `needsContribution`: the row does not stand as published - which only its list can say, each
        // published credit standing for one row (PublishedCredits.Judged; owner report, 2026-11-15).
        // ###########################################################################################
        public static IReadOnlyList<string> EmptyColumns(CreditEntry row, bool needsContribution)
        {
            ArgumentNullException.ThrowIfNull(row);

            return BoardRowRequired.Empty(
                (row.Category, BoardWorkbookSchema.ColCategory),
                (row.NameOrHandle, BoardWorkbookSchema.ColNameOrHandle),
                (needsContribution ? RichNote.NormalisedMarkup(row.MainContribution) : "-", BoardWorkbookSchema.ColMainContribution));
        }

        public static IReadOnlyList<string> EmptyColumns(KiCadImportantSignalEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return BoardRowRequired.Empty(
                (row.KiCadNetName, BoardWorkbookSchema.ColKiCadNetName),
                (row.DisplayName, BoardWorkbookSchema.ColDisplayName));
        }

        // ###########################################################################################
        // AN OSCILLOSCOPE BASELINE'S SIGNAL NAME AND IMAGE (owner request, 2026-11-02: "the Pin + Signal
        // name + image are mandatory"). Its pin is what makes it one (ComponentImageKinds); a picture needs
        // neither. BoardDataChecks says these under a code of its own, so they are no part of Of below.
        // ###########################################################################################
        public static IReadOnlyList<string> EmptyColumns(ComponentImageEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            return ComponentImageKinds.IsScopeBaseline(row)
                ? BoardRowRequired.Empty(
                    (row.Name, BoardWorkbookSchema.ColName),
                    (row.File, BoardWorkbookSchema.ColFile))
                : [];
        }

        // What a sentence calls such a column.
        public static string BaselineFieldName(string column) => column == BoardWorkbookSchema.ColName ? "signal name" : "image";

        // ###########################################################################################
        // How many rows of each list have a column they need empty, and how many such columns there
        // are in all - each one ERROR of BoardDataChecks (a board file's empty file is its
        // "file.unreferenced", the rest "row.required_empty"), so Board status can tell them apart from
        // the board's other errors.
        // ###########################################################################################
        // `published`: the board's published credits (EmptyColumns' rule) - null leaves a credit's main contribution unjudged.
        public static BoardRowsIncomplete Of(BoardData board, PublishedCredits? published)
        {
            ArgumentNullException.ThrowIfNull(board);

            static (int Rows, int Columns) Count<T>(IEnumerable<T>? rows, Func<T, IReadOnlyList<string>> emptyColumns)
            {
                List<int> empty = (rows ?? []).Select(row => emptyColumns(row).Count).Where(count => count > 0).ToList();
                return (empty.Count, empty.Sum());
            }

            IReadOnlySet<CreditEntry> judged = PublishedCredits.Judged(board.Credits, published);

            (int files, int fileColumns) = Count(board.BoardLocalFiles, BoardRowRequired.EmptyColumns);
            (int links, int linkColumns) = Count(board.BoardLinks, BoardRowRequired.EmptyColumns);
            (int credits, int creditColumns) = Count(board.Credits, row => BoardRowRequired.EmptyColumns(row, judged.Contains(row)));
            (int signals, int signalColumns) = Count(board.KiCadImportantSignals, BoardRowRequired.EmptyColumns);

            return new BoardRowsIncomplete(files, links, credits, signals, fileColumns + linkColumns + creditColumns + signalColumns);
        }

        private static List<string> Empty(params (string? Value, string Column)[] columns) =>
            columns.Where(column => string.IsNullOrWhiteSpace(column.Value)).Select(column => column.Column).ToList();
    }

    // The rows of each of the board's own lists with a column they need empty, and the errors they make.
    public sealed record BoardRowsIncomplete(int Files, int Links, int Credits, int Signals, int Errors)
    {
        public bool Any => this.Errors > 0;
    }

    // ###########################################################################################
    // THE CREDITS A BOARD HAS PUBLISHED, which need no "Main contribution" while they stand AS PUBLISHED (owner decision,
    // 2026-11-14: "New credits only" - BoardRowRequired). A credit stands as published while the published board has one
    // with the same category, sub-category, name or handle, contact and main contribution, trimmed and in the same capitals;
    // its date, the server's, is no part of it. Anything else - a credit added, or one changed - needs its main contribution.
    // None is a board nothing is published of: every credit needs one.
    //
    // *** EACH PUBLISHED CREDIT STANDS FOR ONE ROW (owner report, 2026-11-15). *** The credit made for a draft's work came out
    // the same as one the board had published - "Board data", the same name and contact, no main contribution - and, asked by
    // its values alone, stood "as published" too: it was submitted with its main contribution empty. So a list is judged as a
    // whole (Judged): each published credit answers for the first row with its values, in the board's order, and another like
    // it is new. BETA's and stable's credits together count a credit both have once (Union).
    //
    // *** A SCHEMATIC IMAGE'S CREDIT FOLLOWING ITS IMAGE'S RENAME STANDS AS PUBLISHED (code review, 2026-11-16). *** A picture
    // renamed on its card takes the credits for it along (SchematicRename - owner request, 2026-11-04): "Which part of it"
    // follows the new name. Judged by its values alone, the author's published credit then read as changed - a main
    // contribution demanded of somebody else's credit, and Submit stopped. So a "Schematic image" credit unmatched by its very
    // values is matched, second, by every value BUT "Which part of it" (PartlessKeyOf) - as CreditOrder matches it too.
    // ###########################################################################################
    public sealed class PublishedCredits
    {
        // Each published credit's values, and how many published credits have them.
        private readonly Dictionary<string, int> thisCounts;

        // The same of the "Schematic image" credits by every value but which part of it (above).
        private readonly Dictionary<string, int> thisPartlessCounts;

        private PublishedCredits(Dictionary<string, int> counts, Dictionary<string, int> partlessCounts)
        {
            this.thisCounts = counts;
            this.thisPartlessCounts = partlessCounts;
        }

        public static PublishedCredits None { get; } = new(new Dictionary<string, int>(StringComparer.Ordinal), new Dictionary<string, int>(StringComparer.Ordinal));

        public static PublishedCredits Of(IEnumerable<CreditEntry>? credits)
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            var partlessCounts = new Dictionary<string, int>(StringComparer.Ordinal);

            foreach (CreditEntry credit in credits ?? [])
            {
                string key = PublishedCredits.KeyOf(credit);
                counts[key] = counts.GetValueOrDefault(key) + 1;

                if (PublishedCredits.PartlessKeyOf(credit) is { } partless)
                {
                    partlessCounts[partless] = partlessCounts.GetValueOrDefault(partless) + 1;
                }
            }

            return new PublishedCredits(counts, partlessCounts);
        }

        // Two boards' published credits together - BETA's and stable's, as a draft may start from either. A credit both have is
        // counted once: the more of the two, never their sum.
        public static PublishedCredits Union(PublishedCredits? first, PublishedCredits? second) =>
            new(
                PublishedCredits.UnionOf(first?.thisCounts, second?.thisCounts),
                PublishedCredits.UnionOf(first?.thisPartlessCounts, second?.thisPartlessCounts));

        private static Dictionary<string, int> UnionOf(Dictionary<string, int>? first, Dictionary<string, int>? second)
        {
            var counts = new Dictionary<string, int>(first ?? new Dictionary<string, int>(), StringComparer.Ordinal);

            foreach ((string key, int count) in second ?? new Dictionary<string, int>())
            {
                counts[key] = Math.Max(counts.GetValueOrDefault(key), count);
            }

            return counts;
        }

        // Whether the board has published a credit with these very values. Which ROW of a list stands for it is Judged's to say.
        public bool Has(CreditEntry row) => this.thisCounts.ContainsKey(PublishedCredits.KeyOf(row));

        // ###########################################################################################
        // Whether a credit the board has published stands as published as it is - by its very values, or a "Schematic image"
        // credit by every value but which part of it, its image renamed (above): Judged's two passes, asked of one row (code
        // review, 2026-11-16: the pane asked Has alone, and marked "Main contribution" required on the author's credit that
        // the checks and the server took as it stands). Which ROW stands for which published credit is still Judged's to say.
        // ###########################################################################################
        public bool StandsAsPublished(CreditEntry row) =>
            this.Has(row) || (PublishedCredits.PartlessKeyOf(row) is { } partless && this.thisPartlessCounts.ContainsKey(partless));

        // ###########################################################################################
        // The credits of `credits` that do NOT stand as published - those held to their main contribution and their lengths
        // (BoardRowRequired, CreditLimits). Each published credit answers for ONE row with its very values, the first in the
        // board's order (above); then a "Schematic image" credit left over by its values but which part of it - its image
        // renamed (above). By reference. With `published` null, none: the finalise judges nothing again.
        // ###########################################################################################
        public static IReadOnlySet<CreditEntry> Judged(IReadOnlyList<CreditEntry>? credits, PublishedCredits? published)
        {
            var judged = new HashSet<CreditEntry>(ReferenceEqualityComparer.Instance);

            if (published is null || credits is null)
            {
                return judged;
            }

            var left = new Dictionary<string, int>(published.thisCounts, StringComparer.Ordinal);
            var partlessLeft = new Dictionary<string, int>(published.thisPartlessCounts, StringComparer.Ordinal);
            var unmatched = new List<CreditEntry>();

            foreach (CreditEntry credit in credits)
            {
                string key = PublishedCredits.KeyOf(credit);

                if (left.TryGetValue(key, out int count) && count > 0)
                {
                    left[key] = count - 1;

                    // The published credit this one stands for answers for no other.
                    if (PublishedCredits.PartlessKeyOf(credit) is { } partless && partlessLeft.TryGetValue(partless, out int partlessCount))
                    {
                        partlessLeft[partless] = partlessCount - 1;
                    }
                }
                else
                {
                    unmatched.Add(credit);
                }
            }

            foreach (CreditEntry credit in unmatched)
            {
                if (PublishedCredits.PartlessKeyOf(credit) is { } partless &&
                    partlessLeft.TryGetValue(partless, out int count) && count > 0)
                {
                    partlessLeft[partless] = count - 1;
                }
                else
                {
                    judged.Add(credit);
                }
            }

            return judged;
        }

        private static string KeyOf(CreditEntry row) => string.Join(
            '\u001F',
            row.Category?.Trim(),
            row.SubCategory?.Trim(),
            row.NameOrHandle?.Trim(),
            row.Contact?.Trim(),
            RichNote.NormalisedMarkup(row.MainContribution));

        // ###########################################################################################
        // A "Schematic image" credit by every value but which part of it - the one its image's rename changes (above); null for
        // any other credit, or one naming no part. CreditOrder matches by the same, so the two never disagree on it.
        // ###########################################################################################
        internal static string? PartlessKeyOf(CreditEntry row) =>
            string.Equals(row.Category?.Trim(), CreditSuggestions.SchematicImage, StringComparison.OrdinalIgnoreCase) &&
            !string.IsNullOrWhiteSpace(row.SubCategory)
                ? string.Join(
                    '\u001F',
                    row.Category?.Trim(),
                    row.NameOrHandle?.Trim(),
                    row.Contact?.Trim(),
                    RichNote.NormalisedMarkup(row.MainContribution))
                : null;
    }
}
