using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHICH REGIONS A BOARD COMES IN (owner decisions, 2026-11-01: "PAL and NTSC, or none", stored and
    // published; 2026-11-09: "do the region steps 1-9, so we are prepared for other regions"). THE ONE RULE
    // for which region buttons a board shows - under the component list, on the component popup, in the
    // Contribute tab - and what "Applies to" offers on its rows.
    //
    // A board says so in the board file's "regions" (BoardData.Regions), in its own order - PAL and NTSC, or
    // names of its own for hardware that is not told apart by its video standard ("EU", "Japan", "230V").
    // One without it is worked out from its rows, as CRT always did: any row saying PAL or NTSC, and it comes
    // in those two. That rule asked only the COMPONENT rows until 2026-11-01, so a board whose scope baselines
    // were per region and whose components were not hid the buttons and showed one region's baselines for good.
    // Every published board has a regional component, so nothing there needs converting - and every one of
    // them reads exactly as it did, two regions or none.
    //
    // A name is compared ignoring capitals, as every region is, and shown as the board spells it.
    // ###########################################################################################
    public static class BoardRegions
    {
        public const string Pal = "PAL";

        public const string Ntsc = "NTSC";

        // The video standards a computer of the time came in - what "Board: Regions"' tick adds, and what a board
        // without a "regions" list comes in when its rows say either.
        public static IReadOnlyList<string> PalAndNtsc { get; } = [BoardRegions.Pal, BoardRegions.Ntsc];

        // What a row says to apply to every region - never a region's name.
        public const string EveryRegion = "All regions";

        // How long a region's name may be, and how many a board may come in - the buttons stand three to a row in
        // the narrow panel under the component list.
        public const int LongestName = 20;

        public const int MostRegions = 6;

        // How few a board with regions comes in: "multiple regions" is two at least, and a board of one region is no board
        // told apart by region at all - the tick takes them all away instead.
        public const int FewestRegions = 2;

        // ###########################################################################################
        // The board's regions, in its order: its "regions" list (each name once, blanks left out), else PAL and
        // NTSC when a row says either, else none.
        // ###########################################################################################
        public static IReadOnlyList<string> Of(BoardData? board)
        {
            if (board is null)
            {
                return [];
            }

            List<string> listed = BoardRegions.Cleaned(board.Regions);

            // At most MostRegions (code review, 2026-11-09): a board file listing more - placed by hand, refused at Submit - had
            // a button with no place in the fixed grids, laid over another.
            if (listed.Count > 0)
            {
                return listed.Count > BoardRegions.MostRegions ? listed.Take(BoardRegions.MostRegions).ToList() : listed;
            }

            return BoardRegions.RegionsOfRows(board).Any(region => BoardRegions.SaysOneOf(region, BoardRegions.PalAndNtsc)) ? BoardRegions.PalAndNtsc : [];
        }

        // ###########################################################################################
        // The region each row of the board says, row by row - every list a row of which can be of one region. The ONE list of
        // them (code review, 2026-11-09: two copies, counted and asked, would drift as a list with a region is added).
        // ###########################################################################################
        private static IEnumerable<string?> RegionsOfRows(BoardData board) =>
            board.Components.Select(row => row.Region)
                .Concat(board.ComponentImages.Select(row => row.Region))
                .Concat(board.ComponentPins.Select(row => row.Region))
                .Concat(board.ComponentReplacements.Select(row => row.Region))
                .Concat(board.ComponentLocalFiles.Select(row => row.Region))
                .Concat(board.ComponentLinks.Select(row => row.Region))
                .Concat(board.BoardLocalFiles.Select(row => row.Region))
                .Concat(board.BoardLinks.Select(row => row.Region));

        private static bool SaysOneOf(string? region, IReadOnlyCollection<string> regions) =>
            regions.Contains(region?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase);

        public static bool Has(BoardData? board) => BoardRegions.Of(board).Count > 0;

        // `region` as the board spells it - null for none of its regions.
        public static string? SpellingOf(BoardData? board, string? region)
        {
            string wanted = region?.Trim() ?? string.Empty;

            return wanted.Length == 0
                ? null
                : BoardRegions.Of(board).FirstOrDefault(each => string.Equals(each, wanted, StringComparison.OrdinalIgnoreCase));
        }

        // A list of names as a board keeps it: trimmed, each once (the first spelling), blanks left out.
        public static List<string> Cleaned(IEnumerable<string?>? names) =>
            (names ?? [])
                .Select(name => name?.Trim() ?? string.Empty)
                .Where(name => name.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

        // ###########################################################################################
        // How many rows of the board say one of its regions - what stops the setting being taken away: a row only
        // one region shows would go on being there with no button to reach it. A row saying PAL or NTSC counts on a board of
        // regions of its own too (code review, 2026-11-09): a board that lists none is worked out from its rows, so with the
        // list taken away that row brought PAL and NTSC back - the tick stayed ticked, and the board "now came in PAL and NTSC".
        // ###########################################################################################
        public static int RegionalRows(BoardData board) =>
            BoardRegions.RowsSay(board, BoardRegions.Of(board).Concat(BoardRegions.PalAndNtsc).Distinct(StringComparer.OrdinalIgnoreCase).ToList());

        // How many rows say one of `regions`.
        public static int RowsSay(BoardData board, IReadOnlyCollection<string> regions)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(regions);

            return BoardRegions.RegionsOfRows(board).Count(region => BoardRegions.SaysOneOf(region, regions));
        }

        // Whether two boards differ in what they say about their regions - what a draft has to submit. The order is
        // the buttons' order, so it counts.
        public static bool Differ(BoardData? published, BoardData? draft) =>
            !BoardRegions.Of(published).SequenceEqual(BoardRegions.Of(draft), StringComparer.Ordinal);

        // ###########################################################################################
        // "BOARD: REGIONS"' TICK, "This board comes with multiple regions" (owner request, 2026-11-09: "have the checkbox
        // state "This board comes with multiple regions" and then list all regions and its colors" - "This board comes in
        // PAL and NTSC versions" until then). Ticked on a board with none, it comes in PAL and NTSC - what nearly every board
        // here is told apart by - to be renamed, added to, deleted and ordered from there; one worked out from its rows gets
        // them written down. Unticked, it comes in none - the same board while a row still says one of its regions, which
        // would go on being there with no button to reach it.
        // ###########################################################################################
        public static BoardData WithRegions(BoardData board, bool has)
        {
            ArgumentNullException.ThrowIfNull(board);

            IReadOnlyList<string> now = BoardRegions.Of(board);

            if (has)
            {
                return now.Count == 0 ? board.WithRegions([.. BoardRegions.PalAndNtsc])
                    : board.Regions.Count == 0 ? board.WithRegions([.. now])
                    : board;
            }

            // Every region's colors go with them (RegionColors, 2026-11-11).
            return now.Count == 0 || BoardRegions.RegionalRows(board) > 0
                ? board
                : RegionColors.Tidied(board.WithRegions([]));
        }

        // ###########################################################################################
        // The board's regions in `order` - the buttons' order (owner request, 2026-11-09: "It should be possible to reorder
        // the list of regions up/down"). The same board when that is the order already, or when `order` is not exactly its
        // regions, capitals aside; written as the board spells them.
        // ###########################################################################################
        public static BoardData InOrder(BoardData board, IReadOnlyList<string> order)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(order);

            IReadOnlyList<string> now = BoardRegions.Of(board);
            List<string?> spelt = order.Select(region => BoardRegions.SpellingOf(board, region)).ToList();

            if (spelt.Count != now.Count || spelt.Any(region => region is null) ||
                spelt.Distinct(StringComparer.OrdinalIgnoreCase).Count() != now.Count ||
                spelt.SequenceEqual(now, StringComparer.Ordinal))
            {
                return board;
            }

            // Their colors in the same order (RegionColors, 2026-11-11).
            return RegionColors.Tidied(board.WithRegions([.. spelt.Select(region => region!), .. BoardRegions.BeyondTheMost(board)]));
        }

        // ###########################################################################################
        // The names a board file lists past MostRegions - not read (Of), and refused at Submit, but kept when the list is written
        // back (code review, 2026-11-09: a region dragged wrote the six read, and a seventh placed by hand went without a word).
        // ###########################################################################################
        private static List<string> BeyondTheMost(BoardData board) => BoardRegions.Cleaned(board.Regions).Skip(BoardRegions.MostRegions).ToList();

        // Whether `region` may be deleted from the board: one of its regions, held by no row (RowsHolding), and more than the
        // fewest left.
        public static bool CanDelete(BoardData board, string? region)
        {
            ArgumentNullException.ThrowIfNull(board);

            return BoardRegions.SpellingOf(board, region) is { } spelt &&
                   BoardRegions.HoldingOf(board).First(each => string.Equals(each.Region, spelt, StringComparison.Ordinal)).CanDelete;
        }

        // ###########################################################################################
        // *** HOW MANY ROWS HOLD A REGION ON THE BOARD (code review, 2026-11-09). *** A row saying it that would be left in no
        // region without it: a component of that region alone, a picture, a baseline, a part, a file or a link with it
        // applied. NOT a row of a component told apart in EVERY region, nor a pin of a component in every region whose pins
        // are told apart per region: those go WITH the region, the component staying in every other - the very rows a region
        // added gives such a component (WithRowsInAddedRegion), which kept a region added by mistake for good. 0 for a
        // region the board does not have.
        // ###########################################################################################
        public static int RowsHolding(BoardData board, string? region)
        {
            ArgumentNullException.ThrowIfNull(board);

            return BoardRegions.SpellingOf(board, region) is { } spelt
                ? BoardRegions.HoldingOf(board).First(each => string.Equals(each.Region, spelt, StringComparison.Ordinal)).Rows
                : 0;
        }

        // A region of the board, how many rows hold it (RowsHolding), whether it may be deleted (CanDelete), and how many rows
        // would go WITH it - its rows of components told apart in every region, and their pins (said under it, code review
        // 2026-11-09: they went, with what was typed in them, without a word).
        //
        // AND THE COMPONENTS IT IS APPLIED ON (owner request, 2026-11-11: "it should be possible to see which component has the
        // region applied, so I can find these and change them"): every component with a row saying it - those told apart in
        // every region too, whose rows go WITH the region rather than hold it (owner report the same day: on a PAL and NTSC
        // board that is nearly every one, and none was listed) - each label once in the board's order of components; and how
        // many of the board's own files and links say it.
        public sealed record RegionHolding(string Region, int Rows, bool CanDelete, int Going = 0)
        {
            public IReadOnlyList<string> Components { get; init; } = [];

            public int BoardFiles { get; init; }

            public int BoardLinks { get; init; }
        }

        // ###########################################################################################
        // Every region of the board, in its order, with what RowsHolding and CanDelete say of it - worked out in ONE go
        // (second code review, 2026-11-09: "Board: Regions" worked the board's regions and the components told apart out
        // again for every region, and twice over for whether it could go).
        // ###########################################################################################
        public static List<RegionHolding> HoldingOf(BoardData board)
        {
            ArgumentNullException.ThrowIfNull(board);

            IReadOnlyList<string> regions = BoardRegions.Of(board);
            Spread spread = BoardRegions.SpreadOf(board, regions);

            return regions
                .Select(region =>
                {
                    int going = BoardRegions.RowsGoingWith(board, region, spread).Count;
                    int rows = BoardRegions.RowsSay(board, [region]) - going;

                    return new RegionHolding(region, rows, rows == 0 && regions.Count > BoardRegions.FewestRegions, going)
                    {
                        Components = BoardRegions.ComponentsSaying(board, region),
                        BoardFiles = board.BoardLocalFiles.Count(row => BoardRegions.Says(row.Region, region)),
                        BoardLinks = board.BoardLinks.Count(row => BoardRegions.Says(row.Region, region)),
                    };
                })
                .ToList();
        }

        // The components with a row saying `region` - in every list a component's row can say a region in - each label once, in
        // the board's order of components (a label no component row has, last).
        private static List<string> ComponentsSaying(BoardData board, string region)
        {
            var labels = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void Take(string? label, string? rowRegion)
            {
                string text = BoardRegions.LabelOf(label);

                if (text.Length > 0 && BoardRegions.Says(rowRegion, region) && seen.Add(text))
                {
                    labels.Add(text);
                }
            }

            foreach (ComponentEntry row in board.Components) Take(row.BoardLabel, row.Region);
            foreach (ComponentImageEntry row in board.ComponentImages) Take(row.BoardLabel, row.Region);
            foreach (ComponentPinEntry row in board.ComponentPins) Take(row.BoardLabel, row.Region);
            foreach (ComponentReplacementEntry row in board.ComponentReplacements) Take(row.BoardLabel, row.Region);
            foreach (ComponentLocalFileEntry row in board.ComponentLocalFiles) Take(row.BoardLabel, row.Region);
            foreach (ComponentLinkEntry row in board.ComponentLinks) Take(row.BoardLabel, row.Region);

            var order = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (ComponentEntry row in board.Components)
            {
                order.TryAdd(BoardRegions.LabelOf(row.BoardLabel), order.Count);
            }

            return labels.OrderBy(label => order.TryGetValue(label, out int at) ? at : int.MaxValue).ToList();
        }

        // Which components are told apart in each of the board's regions, and which are in every region - that way or with a
        // row for every region.
        private sealed record Spread(HashSet<string> ToldApartInEach, HashSet<string> InEveryRegion);

        private static string LabelOf(string? label) => label?.Trim() ?? string.Empty;

        private static bool Says(string? rowRegion, string wanted) => string.Equals(rowRegion?.Trim() ?? string.Empty, wanted, StringComparison.OrdinalIgnoreCase);

        private static Spread SpreadOf(BoardData board, IReadOnlyList<string> regions)
        {
            var toldApart = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var inEveryRegion = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (IGrouping<string, ComponentEntry> rows in board.Components
                         .Where(row => BoardRegions.LabelOf(row.BoardLabel).Length > 0)
                         .GroupBy(row => BoardRegions.LabelOf(row.BoardLabel), StringComparer.OrdinalIgnoreCase))
            {
                if (rows.All(row => !string.IsNullOrWhiteSpace(row.Region)) && regions.All(each => rows.Any(row => BoardRegions.Says(row.Region, each))))
                {
                    toldApart.Add(rows.Key);
                    inEveryRegion.Add(rows.Key);
                }
                else if (rows.Any(row => string.IsNullOrWhiteSpace(row.Region)))
                {
                    inEveryRegion.Add(rows.Key);
                }
            }

            return new Spread(toldApart, inEveryRegion);
        }

        // The component rows and pins that go with `region` as it is deleted (RowsHolding), by reference.
        private static HashSet<object> RowsGoingWith(BoardData board, string region, Spread spread)
        {
            var going = new HashSet<object>(ReferenceEqualityComparer.Instance);

            going.UnionWith(board.Components.Where(row => spread.ToldApartInEach.Contains(BoardRegions.LabelOf(row.BoardLabel)) && BoardRegions.Says(row.Region, region)));
            going.UnionWith(board.ComponentPins.Where(pin => spread.InEveryRegion.Contains(BoardRegions.LabelOf(pin.BoardLabel)) && BoardRegions.Says(pin.Region, region)));

            return going;
        }

        // ###########################################################################################
        // The board with a region of its own added at the end - the same board for a name NameProblem refuses.
        //
        // *** A COMPONENT TOLD APART PER REGION GETS A ROW IN IT (code review, 2026-11-09). *** A component with a row of its
        // own in every region the board had is in every region - so it must be in the new one too, or it is gone from its
        // list, its popup and the Schematics tab there, and only unticking "This data is different per region" (which keeps
        // one region's row and drops the rest) could bring it back. It gets a copy of its row in the board's FIRST region,
        // right after its own rows - as ticking that box starts each region with the one row it had. A component in every
        // region whose PINS are told apart per region gets the first region's pins the same way. A component of only some
        // regions is left out of the new one, as it is of the others; anything else per row ("Applies to") is the row's own.
        // ###########################################################################################
        public static BoardData WithRegion(BoardData board, string? name)
        {
            ArgumentNullException.ThrowIfNull(board);

            IReadOnlyList<string> now = BoardRegions.Of(board);

            if (BoardRegions.NameProblem(name, now) is not null)
            {
                return board;
            }

            string added = name!.Trim();

            return BoardRegions.WithRowsInAddedRegion(board.WithRegions([.. now, added]), now, added);
        }

        // Every component told apart in each of `before` given a row in `added`, and its per-region pins too.
        private static BoardData WithRowsInAddedRegion(BoardData board, IReadOnlyList<string> before, string added)
        {
            if (before.Count == 0)
            {
                return board;
            }

            Dictionary<string, List<ComponentEntry>> rowsOf = board.Components
                .GroupBy(row => LabelOf(row.BoardLabel), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);

            bool ToldApartInEach(List<ComponentEntry> rows) =>
                rows.All(row => !string.IsNullOrWhiteSpace(row.Region)) &&
                before.All(region => rows.Any(row => Says(row.Region, region))) &&
                !rows.Any(row => Says(row.Region, added));

            // The new rows, by the label they follow: a copy of the first region's row.
            Dictionary<string, List<ComponentEntry>> newRows = rowsOf
                .Where(pair => pair.Key.Length > 0 && ToldApartInEach(pair.Value))
                .ToDictionary(
                    pair => pair.Key,
                    pair => new List<ComponentEntry> { pair.Value.First(row => Says(row.Region, before[0])).WithRegion(added) },
                    StringComparer.OrdinalIgnoreCase);

            // In every region: a row for every region, or one told apart in each. Its pins told apart per region - one with a
            // region - get the pins of the first region that has any.
            bool InEveryRegion(string label) =>
                newRows.ContainsKey(label) || (rowsOf.TryGetValue(label, out List<ComponentEntry>? rows) && rows.Any(row => string.IsNullOrWhiteSpace(row.Region)));

            Dictionary<string, List<ComponentPinEntry>> newPins = board.ComponentPins
                .GroupBy(pin => LabelOf(pin.BoardLabel), StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Key.Length > 0 &&
                                InEveryRegion(group.Key) &&
                                group.Any(pin => !string.IsNullOrWhiteSpace(pin.Region)) &&
                                !group.Any(pin => Says(pin.Region, added)))
                .Select(group => (group.Key, Pins: before
                    .Select(region => group.Where(pin => Says(pin.Region, region)).ToList())
                    .FirstOrDefault(pins => pins.Count > 0) ?? []))
                .Where(each => each.Pins.Count > 0)
                .ToDictionary(each => each.Key, each => each.Pins.Select(pin => pin.WithRegion(added)).ToList(), StringComparer.OrdinalIgnoreCase);

            if (newRows.Count == 0 && newPins.Count == 0)
            {
                return board;
            }

            return board.WithComponentSections(
                components: BoardRegions.InsertedAfterTheirOwn(board.Components, row => LabelOf(row.BoardLabel), newRows),
                componentPins: BoardRegions.InsertedAfterTheirOwn(board.ComponentPins, pin => LabelOf(pin.BoardLabel), newPins));
        }

        // `rows` with each label's new rows right after its last row.
        private static List<T> InsertedAfterTheirOwn<T>(IReadOnlyList<T> rows, Func<T, string> labelOf, Dictionary<string, List<T>> added)
        {
            if (added.Count == 0)
            {
                return rows.ToList();
            }

            var lastOf = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int index = 0; index < rows.Count; index++)
            {
                lastOf[labelOf(rows[index])] = index;
            }

            var result = new List<T>(rows.Count + added.Values.Sum(list => list.Count));

            for (int index = 0; index < rows.Count; index++)
            {
                result.Add(rows[index]);

                string label = labelOf(rows[index]);

                if (lastOf[label] == index && added.TryGetValue(label, out List<T>? rowsToAdd))
                {
                    result.AddRange(rowsToAdd);
                }
            }

            return result;
        }

        // The board without `region` - the same board while CanDelete says no: a row still holds it, it has no such region, or
        // the fewest would not be left. The rows of a component told apart in every region, and the pins of one in every
        // region told apart per region, go with it (RowsHolding) - one step, which "Undo" takes back.
        public static BoardData WithoutRegion(BoardData board, string? region)
        {
            ArgumentNullException.ThrowIfNull(board);

            if (!BoardRegions.CanDelete(board, region) || BoardRegions.SpellingOf(board, region) is not { } spelt)
            {
                return board;
            }

            IReadOnlyList<string> regions = BoardRegions.Of(board);
            List<string> left = [.. regions.Where(each => !string.Equals(each, spelt, StringComparison.Ordinal)), .. BoardRegions.BeyondTheMost(board)];
            HashSet<object> going = BoardRegions.RowsGoingWith(board, spelt, BoardRegions.SpreadOf(board, regions));

            BoardData without = going.Count == 0
                ? board
                : board.WithComponentSections(
                    components: board.Components.Where(row => !going.Contains(row)).ToList(),
                    componentPins: board.ComponentPins.Where(pin => !going.Contains(pin)).ToList());

            // Its colors go with it (RegionColors, 2026-11-11).
            return RegionColors.Tidied(without.WithRegions(left));
        }

        public static bool IsPalOrNtsc(string? region)
        {
            string text = region?.Trim() ?? string.Empty;

            return string.Equals(text, BoardRegions.Pal, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, BoardRegions.Ntsc, StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // *** WHAT A REGION MAY BE CALLED - the rule the table, the pane and the server share (Security, code review
        // 2026-11-01: a submission's "regions" reaches the publish as sent). *** A name of 1 to LongestName letters,
        // digits, spaces and "- + / ( ) .", never "All regions", never one the board has already; at most MostRegions.
        // Null when `name` may be added to `existing`.
        // ###########################################################################################
        public static string? NameProblem(string? name, IReadOnlyList<string> existing)
        {
            ArgumentNullException.ThrowIfNull(existing);

            string text = name?.Trim() ?? string.Empty;

            if (text.Length == 0)
            {
                return "A region needs a name.";
            }

            if (text.Length > BoardRegions.LongestName)
            {
                return $"A region's name is at most {BoardRegions.LongestName} characters.";
            }

            if (!text.All(BoardRegions.MayBeInName))
            {
                return "A region's name is letters, digits, spaces and \"- + / ( ) .\" only.";
            }

            if (string.Equals(text, BoardRegions.EveryRegion, StringComparison.OrdinalIgnoreCase))
            {
                return $"\"{BoardRegions.EveryRegion}\" is what a row says for every region, so it cannot be a region's name.";
            }

            if (existing.Contains(text, StringComparer.OrdinalIgnoreCase))
            {
                return $"The board comes in [{text}] already.";
            }

            if (existing.Count >= BoardRegions.MostRegions)
            {
                return $"A board comes in at most {BoardRegions.MostRegions} regions.";
            }

            return null;
        }

        private static bool MayBeInName(char character) =>
            char.IsLetterOrDigit(character) || character is ' ' or '-' or '+' or '/' or '(' or ')' or '.';

        // ---- what a change is called - "Review changes", the review and the board's history ----------------

        // "PAL and NTSC", "PAL, NTSC and Japan" - or "PAL or NTSC" with `last` "or".
        public static string Joined(IReadOnlyList<string> regions, string last = "and")
        {
            ArgumentNullException.ThrowIfNull(regions);

            return regions.Count switch
            {
                0 => string.Empty,
                1 => regions[0],
                _ => $"{string.Join(", ", regions.Take(regions.Count - 1))} {last} {regions[^1]}",
            };
        }

        // ###########################################################################################
        // A change of what the board comes in, named by its regions before and after. Null names - a record from before
        // regions had names of their own - are PAL and NTSC, all there was.
        // ###########################################################################################
        public static string ChangedLine(IReadOnlyList<string>? before, IReadOnlyList<string>? now)
        {
            IReadOnlyList<string> after = now ?? BoardRegions.PalAndNtsc;

            // The same regions in another order - the buttons' order, a change a draft submits - is said as that (code review,
            // 2026-11-09: it read "This board now comes in NTSC and PAL versions", a new set of regions to a maintainer).
            if (before is { Count: > 0 } && after.Count == before.Count &&
                new HashSet<string>(before, StringComparer.OrdinalIgnoreCase).SetEquals(after) &&
                !before.SequenceEqual(after, StringComparer.OrdinalIgnoreCase))
            {
                return $"This board's regions are now in the order {string.Join(", ", after)}";
            }

            return after.Count > 0
                ? $"This board now comes in {BoardRegions.Joined(after)} versions"
                : $"This board no longer comes in {BoardRegions.Joined(before is { Count: > 0 } ? before : BoardRegions.PalAndNtsc)} versions";
        }

        // The same, from whether it has regions now alone - all an older record says.
        public static string ChangedLine(bool hasNow) =>
            BoardRegions.ChangedLine(BoardRegions.PalAndNtsc, hasNow ? BoardRegions.PalAndNtsc : []);

        // The same of a review's or history's record: by the names when it has them, else by whether it has regions now.
        public static string ChangedLine(bool hasNow, IReadOnlyList<string>? before, IReadOnlyList<string>? now) =>
            now is null ? BoardRegions.ChangedLine(hasNow) : BoardRegions.ChangedLine(before, now);

        // The same of a new board that has them - the review and the board's history (code review, 2026-11-02).
        public static string NewBoardLine(IReadOnlyList<string>? regions) =>
            $"This new board comes in {BoardRegions.Joined(regions is { Count: > 0 } ? regions : BoardRegions.PalAndNtsc)} versions";
    }
}
