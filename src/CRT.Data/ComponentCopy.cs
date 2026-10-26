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
        RepairNotes,
        Pins,
        Replacements,
        Links,
    }

    // ###########################################################################################
    // One component of another board with the same part number - where it is, and what it has.
    // `BoardName` is how the board is shown ("Commodore 64 - 250425").
    // ###########################################################################################
    public sealed record ComponentCopySource(string BoardName, BoardData Board, ComponentEntry Row, string BoardIdentity = "")
    {
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
    // A chip is the same chip on every board it sits on: its case, its pins, what can replace it and
    // what usually goes wrong with it. So the component pane offers the components of the other boards
    // on this machine with the SAME PART NUMBER (trimmed, capitals ignored - a component without one
    // matches nothing), and copies what is picked into the component shown, as ONE step:
    //
    //   - the package, the short description and the repair notes - each a field set;
    //   - the pins, the replacement parts and the links - the component's own (in the region shown, for
    //     pins and parts) replaced by the other board's, filed under this component's label and region.
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
                    sources.Add(new ComponentCopySource(boardName, board, row, identity));
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

            var here = new ComponentCopySource(string.Empty, draft, row);
            List<ComponentCopyPart> available = ComponentCopy.Available(here);

            if (published is null || ComponentPaneEdits.RowShown(published, label, region) is not ComponentEntry before)
            {
                return available;
            }

            string ownRegion = row.Region ?? string.Empty;
            string beforeRegion = before.Region ?? string.Empty;

            bool Differs(ComponentCopyPart part) => part switch
            {
                ComponentCopyPart.Package => !ComponentCopy.Same(row.Package, before.Package),
                ComponentCopyPart.Description => !ComponentCopy.Same(row.Description, before.Description),
                ComponentCopyPart.RepairNotes => !ComponentCopy.Same(row.RepairNotes, before.RepairNotes),
                ComponentCopyPart.Pins => !ComponentCopy.SameKeys(
                    ComponentPaneEdits.PinsOf(draft, label, ownRegion).Select(pin => $"{pin.Pin?.Trim()}|{pin.Name?.Trim()}|{pin.Direction?.Trim()}"),
                    ComponentPaneEdits.PinsOf(published, label, beforeRegion).Select(pin => $"{pin.Pin?.Trim()}|{pin.Name?.Trim()}|{pin.Direction?.Trim()}")),
                ComponentCopyPart.Replacements => !ComponentCopy.SameKeys(
                    ComponentPaneEdits.ReplacementsOf(draft, label, ownRegion).Select(part => $"{part.PartNumber?.Trim()}|{part.Kind?.Trim()}|{part.Changes?.Trim()}|{part.Source?.Trim()}"),
                    ComponentPaneEdits.ReplacementsOf(published, label, beforeRegion).Select(part => $"{part.PartNumber?.Trim()}|{part.Kind?.Trim()}|{part.Changes?.Trim()}|{part.Source?.Trim()}")),
                _ => !ComponentCopy.SameKeys(
                    ComponentPaneEdits.LinksOf(draft, label).Select(link => $"{link.Name?.Trim()}|{link.Url?.Trim()}"),
                    ComponentPaneEdits.LinksOf(published, label).Select(link => $"{link.Name?.Trim()}|{link.Url?.Trim()}")),
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
            string region = source.Row.Region ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(source.Row.Package))
                parts.Add(ComponentCopyPart.Package);
            if (!string.IsNullOrWhiteSpace(source.Row.Description))
                parts.Add(ComponentCopyPart.Description);
            if (!string.IsNullOrWhiteSpace(source.Row.RepairNotes))
                parts.Add(ComponentCopyPart.RepairNotes);
            if (ComponentPaneEdits.PinsOf(source.Board, source.Row.BoardLabel, region).Count > 0)
                parts.Add(ComponentCopyPart.Pins);
            if (ComponentPaneEdits.ReplacementsOf(source.Board, source.Row.BoardLabel, region).Count > 0)
                parts.Add(ComponentCopyPart.Replacements);
            if (ComponentPaneEdits.LinksOf(source.Board, source.Row.BoardLabel).Count > 0)
                parts.Add(ComponentCopyPart.Links);

            return parts;
        }

        // How a part is named beside its tick box - with what it holds: "Pins (40)", "Package (DIP-40)".
        public static string Describe(ComponentCopySource source, ComponentCopyPart part)
        {
            ArgumentNullException.ThrowIfNull(source);

            string region = source.Row.Region ?? string.Empty;

            return part switch
            {
                ComponentCopyPart.Package => $"Package ({source.Row.Package?.Trim()})",
                ComponentCopyPart.Description => "Short description",
                ComponentCopyPart.RepairNotes => "Repair notes",
                ComponentCopyPart.Pins => ComponentCopy.Counted("Pins", ComponentPaneEdits.PinsOf(source.Board, source.Row.BoardLabel, region).Count),
                ComponentCopyPart.Replacements => ComponentCopy.Counted("Replacement parts", ComponentPaneEdits.ReplacementsOf(source.Board, source.Row.BoardLabel, region).Count),
                _ => ComponentCopy.Counted("Links", ComponentPaneEdits.LinksOf(source.Board, source.Row.BoardLabel).Count),
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
            string sourceRegion = source.Row.Region ?? string.Empty;
            BoardData result = board;

            if (parts.Contains(ComponentCopyPart.Package))
                result = ComponentPaneEdits.WithField(result, label, region, ComponentField.Package, source.Row.Package);
            if (parts.Contains(ComponentCopyPart.Description))
                result = ComponentPaneEdits.WithField(result, label, region, ComponentField.Description, source.Row.Description);
            if (parts.Contains(ComponentCopyPart.RepairNotes))
                result = ComponentPaneEdits.WithField(result, label, region, ComponentField.RepairNotes, source.Row.RepairNotes);

            if (parts.Contains(ComponentCopyPart.Pins))
            {
                List<ComponentPinEntry> pins = ComponentPaneEdits.PinsOf(source.Board, source.Row.BoardLabel, sourceRegion)
                    .Select(pin => new ComponentPinEntry { Region = targetRegion, Pin = pin.Pin, Name = pin.Name, Direction = pin.Direction })
                    .ToList();

                result = ComponentPaneEdits.WithPins(result, label, region, pins);
            }

            if (parts.Contains(ComponentCopyPart.Replacements))
            {
                List<ComponentReplacementEntry> replacements = ComponentPaneEdits.ReplacementsOf(source.Board, source.Row.BoardLabel, sourceRegion)
                    .Select(row => new ComponentReplacementEntry { Region = targetRegion, PartNumber = row.PartNumber, Kind = row.Kind, Changes = row.Changes, Source = row.Source })
                    .ToList();

                result = ComponentPaneEdits.WithReplacements(result, label, region, replacements);
            }

            if (parts.Contains(ComponentCopyPart.Links))
            {
                List<ComponentLinkEntry> links = ComponentPaneEdits.LinksOf(source.Board, source.Row.BoardLabel)
                    .Select(row => new ComponentLinkEntry { Name = row.Name, Url = row.Url })
                    .ToList();

                result = ComponentPaneEdits.WithLinks(result, label, links);
            }

            return result;
        }

        private static string Counted(string name, int count) => $"{name} ({count.ToString(CultureInfo.InvariantCulture)})";
    }
}
