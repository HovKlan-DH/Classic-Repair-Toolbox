using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // What can be copied of a component from another board.
    public enum ComponentCopyPart
    {
        Package,
        Description,
        Pins,
        Replacements,
        Links,
    }

    // ###########################################################################################
    // One component of another board with the same part number - where it is, and what it has.
    // `BoardName` is how the board is shown ("Commodore 64 - 250425"); `RegionShown` the region it is
    // looked at in.
    // ###########################################################################################
    public sealed record ComponentCopySource(string BoardName, BoardData Board, ComponentEntry Row, string BoardIdentity = "", string RegionShown = "")
    {
        // ###########################################################################################
        // The region this component's pins, parts and links are read in: its row's - else, for a component
        // in every region, the one shown (code review, 2026-11-02). Asked of a blank region, the board's rows
        // of one region were never seen, on the usual shape of a regional board: a row for every region,
        // and a link or a pin of PAL.
        // ###########################################################################################
        public string RowsRegion => string.IsNullOrWhiteSpace(this.Row.Region) ? this.RegionShown : this.Row.Region;

        // "Commodore 64 - 250425: U7 - PLA | 906114-01"
        public string DisplayText
        {
            get
            {
                string names = string.Join(" | ", new[] { this.Row.FriendlyName, this.Row.TechnicalNameOrValue }
                    .Select(name => name?.Trim() ?? string.Empty)
                    .Where(name => name.Length > 0));

                return names.Length == 0
                    ? $"{this.BoardName}: {this.Row.BoardLabel?.Trim()}"
                    : $"{this.BoardName}: {this.Row.BoardLabel?.Trim()} - {names}";
            }
        }

        public override string ToString() => this.DisplayText;
    }

    // ###########################################################################################
    // COPY FROM ANOTHER BOARD (CRT 4.0 step 2, slice 5 - the plan: "Copy from another board pulls data
    // across for a part with the same part number").
    //
    // A chip is the same chip on every board it sits on: its case, its pins and what can replace it. So
    // the component pane offers the components of the other boards on this machine with the SAME PART
    // NUMBER (trimmed, capitals ignored - a component without one matches nothing), and copies what is
    // picked into the component shown, as ONE step:
    //
    //   - the package and the short description - each a field set;
    //   - the pins, the replacement parts and the links - the component's own in the region shown
    //     replaced by the other board's, filed under this component's label and region. On a component
    //     in every region a row of one region stays of one region, the one shown - where the board has
    //     regions at all, as "PAL" would turn its PAL/NTSC buttons on.
    //   - ROWS OF ONE REGION REPLACE ONLY THAT REGION'S (code review, 2026-11-02): the component's rows for
    //     every region are the other region's too, so they stay, filed for it - it goes on showing what it
    //     showed. Rows for every region copied replace them in every region: the same chip, in both.
    //
    // Files are not copied (their bytes are another board's); nor is anything board-specific - the
    // label, the category, the friendly name, the highlights, the scope baselines.
    //
    // PURE: boards in, a board out.
    // ###########################################################################################
    public static class ComponentCopy
    {
        // ###########################################################################################
        // Every component of `others` - boards named, the draft's own board left out by the caller -
        // with the part number `partNumber`, the boards in the order given and each board's rows in its
        // own order. One row per component: a component of two regions offers the one of `region`.
        // ###########################################################################################
        public static List<ComponentCopySource> Sources(IEnumerable<(string BoardName, BoardData Board)> others, string? partNumber, string? region)
        {
            ArgumentNullException.ThrowIfNull(others);

            return ComponentCopy.Sources(others.Select(other => (other.BoardName, string.Empty, other.Board)), partNumber, region);
        }

        // The same, each source knowing the board it is on (its identity, the board file) - so a change
        // can be put into that board's own draft ("Also on other boards").
        public static List<ComponentCopySource> Sources(IEnumerable<(string BoardName, string BoardIdentity, BoardData Board)> others, string? partNumber, string? region)
        {
            ArgumentNullException.ThrowIfNull(others);

            string wanted = partNumber?.Trim() ?? string.Empty;

            if (wanted.Length == 0)
            {
                return [];
            }

            var sources = new List<ComponentCopySource>();

            foreach ((string boardName, string identity, BoardData board) in others)
            {
                foreach (IGrouping<string, ComponentEntry> component in board.Components
                    .Where(row => string.Equals(row.PartNumber?.Trim(), wanted, StringComparison.OrdinalIgnoreCase))
                    .GroupBy(row => row.BoardLabel?.Trim().ToUpperInvariant() ?? string.Empty))
                {
                    // Among the rows that MATCHED: the region's, else the first (code review, 2026-10-26 -
                    // the region's row whatever its part number was another region's chip).
                    ComponentEntry row = component.FirstOrDefault(match => ComponentPaneEdits.InRegion(match.Region, region)) ?? component.First();
                    sources.Add(new ComponentCopySource(boardName, board, row, identity, region ?? string.Empty));
                }
            }

            return sources;
        }

        // ###########################################################################################
        // What the draft changed of the component `label` as shown in `region`, against `published` -
        // the parts "Also on other boards" offers to put on the same part elsewhere. Every part it holds
        // for a component the published board does not have. Only parts that can be copied count.
        // ###########################################################################################
        public static List<ComponentCopyPart> Changed(BoardData draft, BoardData? published, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(draft);

            if (ComponentPaneEdits.RowShown(draft, label, region) is not ComponentEntry row)
            {
                return [];
            }

            var here = new ComponentCopySource(string.Empty, draft, row, RegionShown: region ?? string.Empty);
            List<ComponentCopyPart> available = ComponentCopy.Available(here);

            if (published is null || ComponentPaneEdits.RowShown(published, label, region) is not ComponentEntry before)
            {
                return available;
            }

            string ownRegion = here.RowsRegion;
            string beforeRegion = (here with { Row = before }).RowsRegion;

            bool Differs(ComponentCopyPart part) => part switch
            {
                ComponentCopyPart.Package => !ComponentCopy.Same(row.Package, before.Package),
                ComponentCopyPart.Description => !ComponentCopy.Same(row.Description, before.Description),
                ComponentCopyPart.Pins => !ComponentCopy.SameKeys(
                    ComponentPaneEdits.PinsOf(draft, label, ownRegion).Select(pin => $"{pin.Pin?.Trim()}|{pin.Name?.Trim()}|{pin.Direction?.Trim()}"),
                    ComponentPaneEdits.PinsOf(published, label, beforeRegion).Select(pin => $"{pin.Pin?.Trim()}|{pin.Name?.Trim()}|{pin.Direction?.Trim()}")),
                ComponentCopyPart.Replacements => !ComponentCopy.SameKeys(
                    ComponentPaneEdits.ReplacementsOf(draft, label, ownRegion).Select(part => $"{part.PartNumber?.Trim()}|{part.Kind?.Trim()}|{part.Note?.Trim()}"),
                    ComponentPaneEdits.ReplacementsOf(published, label, beforeRegion).Select(part => $"{part.PartNumber?.Trim()}|{part.Kind?.Trim()}|{part.Note?.Trim()}")),
                _ => !ComponentCopy.SameKeys(
                    ComponentPaneEdits.LinksOf(draft, label, ownRegion).Select(link => $"{link.Name?.Trim()}|{link.Url?.Trim()}"),
                    ComponentPaneEdits.LinksOf(published, label, beforeRegion).Select(link => $"{link.Name?.Trim()}|{link.Url?.Trim()}")),
            };

            return available.Where(Differs).ToList();
        }

        private static bool Same(string? first, string? second) =>
            string.Equals(first?.Trim() ?? string.Empty, second?.Trim() ?? string.Empty, StringComparison.Ordinal);

        private static bool SameKeys(IEnumerable<string> first, IEnumerable<string> second) =>
            first.Order(StringComparer.Ordinal).SequenceEqual(second.Order(StringComparer.Ordinal), StringComparer.Ordinal);

        // What `source` has to copy - each part only when it holds something.
        public static List<ComponentCopyPart> Available(ComponentCopySource source)
        {
            ArgumentNullException.ThrowIfNull(source);

            var parts = new List<ComponentCopyPart>();
            string region = source.RowsRegion;

            if (!string.IsNullOrWhiteSpace(source.Row.Package))
                parts.Add(ComponentCopyPart.Package);
            if (!string.IsNullOrWhiteSpace(source.Row.Description))
                parts.Add(ComponentCopyPart.Description);
            if (ComponentPaneEdits.PinsOf(source.Board, source.Row.BoardLabel, region).Count > 0)
                parts.Add(ComponentCopyPart.Pins);
            if (ComponentPaneEdits.ReplacementsOf(source.Board, source.Row.BoardLabel, region).Count > 0)
                parts.Add(ComponentCopyPart.Replacements);
            if (ComponentPaneEdits.LinksOf(source.Board, source.Row.BoardLabel, region).Count > 0)
                parts.Add(ComponentCopyPart.Links);

            return parts;
        }

        // How a part is named beside its tick box - with what it holds: "Pins (40)", "Package (DIP-40)".
        public static string Describe(ComponentCopySource source, ComponentCopyPart part)
        {
            ArgumentNullException.ThrowIfNull(source);

            string region = source.RowsRegion;

            return part switch
            {
                ComponentCopyPart.Package => $"Package ({source.Row.Package?.Trim()})",
                ComponentCopyPart.Description => "Short description",
                ComponentCopyPart.Pins => ComponentCopy.Counted("Pins", ComponentPaneEdits.PinsOf(source.Board, source.Row.BoardLabel, region).Count),
                ComponentCopyPart.Replacements => ComponentCopy.Counted("Replacement parts", ComponentPaneEdits.ReplacementsOf(source.Board, source.Row.BoardLabel, region).Count),
                _ => ComponentCopy.Counted("Links", ComponentPaneEdits.LinksOf(source.Board, source.Row.BoardLabel, region).Count),
            };
        }

        // ###########################################################################################
        // `board` with `parts` of `source` copied into the component `label` as shown in `region` - the
        // same board when that copies nothing new (or the board has no such component).
        // ###########################################################################################
        public static BoardData Apply(BoardData board, string label, string? region, ComponentCopySource source, IReadOnlyCollection<ComponentCopyPart> parts)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(source);
            ArgumentNullException.ThrowIfNull(parts);

            if (ComponentPaneEdits.RowShown(board, label, region) is not ComponentEntry target)
            {
                return board;
            }

            string targetRegion = target.Region?.Trim() ?? string.Empty;
            string sourceRegion = source.RowsRegion;
            string shown = region?.Trim() ?? string.Empty;
            BoardData result = board;

            string FiledRegion(string? rowRegion) =>
                targetRegion.Length == 0 && !string.IsNullOrWhiteSpace(rowRegion) && BoardRegions.Has(board)
                    ? shown
                    : targetRegion;

            // The regions besides the one shown that have this component - they show its rows for every
            // region too. Never one it is not in: "PAL" alone has no NTSC to keep anything for.
            List<string> others = shown.Length > 0 && BoardRegions.Has(board)
                ? BoardRegions.All
                    .Where(other => !string.Equals(other, shown, StringComparison.OrdinalIgnoreCase) &&
                                    board.Components.Any(row => ComponentPaneEdits.InScope(row.BoardLabel, row.Region, label, other)))
                    .ToList()
                : [];

            if (parts.Contains(ComponentCopyPart.Package))
                result = ComponentPaneEdits.WithField(result, label, region, ComponentField.Package, source.Row.Package);
            if (parts.Contains(ComponentCopyPart.Description))
                result = ComponentPaneEdits.WithField(result, label, region, ComponentField.Description, source.Row.Description);

            if (parts.Contains(ComponentCopyPart.Pins))
            {
                static ComponentPinEntry Filed(ComponentPinEntry pin, string filed) =>
                    new() { Region = filed, Pin = pin.Pin, Name = pin.Name, Direction = pin.Direction };

                List<ComponentPinEntry> pins = ComponentPaneEdits.PinsOf(source.Board, source.Row.BoardLabel, sourceRegion)
                    .Select(pin => Filed(pin, FiledRegion(pin.Region)))
                    .ToList();

                result = ComponentPaneEdits.WithPins(
                    result,
                    label,
                    region,
                    ComponentCopy.KeepingOtherRegions(pins, ComponentPaneEdits.PinsOf(result, label, region), pin => pin.Region, Filed, shown, others));
            }

            if (parts.Contains(ComponentCopyPart.Replacements))
            {
                static ComponentReplacementEntry Filed(ComponentReplacementEntry row, string filed) =>
                    new() { Region = filed, PartNumber = row.PartNumber, Kind = row.Kind, Note = row.Note };

                List<ComponentReplacementEntry> replacements = ComponentPaneEdits.ReplacementsOf(source.Board, source.Row.BoardLabel, sourceRegion)
                    .Select(row => Filed(row, FiledRegion(row.Region)))
                    .ToList();

                result = ComponentPaneEdits.WithReplacements(
                    result,
                    label,
                    region,
                    ComponentCopy.KeepingOtherRegions(replacements, ComponentPaneEdits.ReplacementsOf(result, label, region), row => row.Region, Filed, shown, others));
            }

            if (parts.Contains(ComponentCopyPart.Links))
            {
                static ComponentLinkEntry Filed(ComponentLinkEntry row, string filed) =>
                    new() { Region = filed, Name = row.Name, Url = row.Url };

                List<ComponentLinkEntry> links = ComponentPaneEdits.LinksOf(source.Board, source.Row.BoardLabel, sourceRegion)
                    .Select(row => Filed(row, FiledRegion(row.Region)))
                    .ToList();

                result = ComponentPaneEdits.WithLinks(
                    result,
                    label,
                    region,
                    ComponentCopy.KeepingOtherRegions(links, ComponentPaneEdits.LinksOf(result, label, region), row => row.Region, Filed, shown, others));
            }

            return result;
        }

        // ###########################################################################################
        // What goes in place of the rows the region shown lists (`before`), when any copied is of one
        // region: the copied rows, all of the region shown - one for every region would stand beside what
        // the other region keeps - and the rows for every region filed for each other region, so it goes
        // on showing them. Just the copied rows when nothing is left behind by them.
        // ###########################################################################################
        private static List<T> KeepingOtherRegions<T>(
            List<T> copied,
            List<T> before,
            Func<T, string?> regionOf,
            Func<T, string, T> filedIn,
            string shown,
            IReadOnlyList<string> others)
        {
            bool ForEveryRegion(T row) => string.IsNullOrWhiteSpace(regionOf(row));

            List<T> kept = before.Where(ForEveryRegion).ToList();

            if (others.Count == 0 || kept.Count == 0 || copied.All(ForEveryRegion))
            {
                return copied;
            }

            return
            [
                .. copied.Select(row => ForEveryRegion(row) ? filedIn(row, shown) : row),
                .. others.SelectMany(other => kept.Select(row => filedIn(row, other))),
            ];
        }

        private static string Counted(string name, int count) => $"{name} ({count.ToString(CultureInfo.InvariantCulture)})";
    }
}
