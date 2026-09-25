using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Following one net across a SHEET BOUNDARY, so selecting a component lights its wires on the
    // sheet the user is actually looking at (owner report, 2026-09-24).
    //
    // *** THE BUG THIS EXISTS FOR. *** A hierarchical KiCad design renames a net at every boundary.
    // On the C128 Open128 board, CN9 pad 4 is the userport's CNT1 line:
    //
    //     the Userport sheet draws it labelled ...... CNT1
    //     the PCB calls that copper ................. /I{slash}O/Serial Bus/CNT
    //
    // because on the PARENT "I/O" sheet, the Userport sheet symbol's "CNT1" pin and the Serial Bus
    // sheet symbol's "CNT" pin sit on one wire. KiCad names the merged net after the Serial Bus
    // side, so the name the PCB carries is a name the Userport sheet has never heard of.
    //
    // Selecting a component reads net names off the PCB pads and looks them up in a per-sheet index
    // keyed by that sheet's own label text. For CNT1/SP1 the lookup missed, so nothing highlighted -
    // while hovering the same wire still worked, because hover hit-tests geometry and never does
    // this lookup. That asymmetry is what the project owner noticed.
    //
    // *** WHY THIS IS AN ALIAS TABLE RATHER THAN A RE-KEYING. *** The obvious fix is to key the
    // index by the full hierarchical path instead of the leaf. That fails for the sheet's OWN
    // labels: a sheet draws "CNT1" as plain text and nothing in the sheet file says which full path
    // it becomes - that is decided by the parent, and for a sheet instantiated twice there is no
    // single answer. So the sheet keeps its own honest names and gains the other names the same
    // wire answers to elsewhere. Nothing that works today changes key.
    //
    // *** THE AMBIGUITY RULE IS THE IMPORTANT PART. *** Highlighting the WRONG net is worse than
    // highlighting none: someone at a bench would trust it and probe the wrong pin. So an alias is
    // only ever recorded when it is unambiguous - see BuildAliasesForSheet.
    // ###########################################################################################
    public static class KiCadSheetNetAliases
    {
        // How close a parent wire endpoint must come to a sheet pin to count as touching it. The
        // same tolerance the index's own label-to-wire seeding uses, and for the same reason: KiCad
        // writes millimetre coordinates that are exact in practice but not bit-exact after rotation.
        private const double PinTouchTolerance = 0.35;

        // ###########################################################################################
        // Every extra name each sheet's nets answer to, keyed by schematic index.
        //
        // The result maps a sheet index to (alias name -> the sheet's OWN name for that net), so a
        // caller can register the sheet's existing resolved wires under the alias without recomputing
        // any geometry.
        // ###########################################################################################
        public static Dictionary<int, Dictionary<string, string>> Build(
            IReadOnlyList<KiCadSchematic> schematics)
        {
            var result = new Dictionary<int, Dictionary<string, string>>();

            if (schematics == null || schematics.Count == 0)
            {
                return result;
            }

            // Which schematic index each file name corresponds to, so a parent's Sheetfile can be
            // resolved to the child it names. File names only - a KiCad project keeps its sheets in
            // one folder tree and Sheetfile is relative to the parent.
            var indexByFileName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < schematics.Count; i++)
            {
                string file = System.IO.Path.GetFileName(schematics[i].Filename ?? string.Empty);

                if (!string.IsNullOrWhiteSpace(file) && !indexByFileName.ContainsKey(file))
                {
                    indexByFileName[file] = i;
                }
            }

            // Walk every PARENT sheet: its wires are what join one child's pin to another's.
            for (int parentIndex = 0; parentIndex < schematics.Count; parentIndex++)
            {
                var parent = schematics[parentIndex];

                if (parent.ChildSheets.Count == 0)
                {
                    continue;
                }

                KiCadSheetNetAliases.BuildAliasesForSheet(
                    parent, indexByFileName, result);
            }

            return result;
        }

        // ###########################################################################################
        // The alias rule for one parent sheet.
        //
        // *** THE CONNECTION IS PIN-TO-PIN THROUGH THE WIRES, NOT VIA A LABEL. *** An earlier
        // version keyed this on the parent label for the wire, which fails on exactly the reported
        // case: on the C128 I/O sheet the Userport "CNT1" pin is joined to the Serial Bus "CNT" pin
        // by a bare run of ten wire segments carrying no label at all. Flooding the wires from one
        // pin and collecting every other sheet pin reached is what actually finds it, and it needs
        // no label to exist.
        //
        // Any label sitting on that same run is collected too, since the parent own name for the net
        // is a legitimate alias as well - it is simply not required for the match.
        //
        // *** AN ALIAS IS ONLY RECORDED WHEN IT IS UNAMBIGUOUS. *** Two guards, both load-bearing:
        //
        //   - a run reaching two DIFFERENTLY named pins of the SAME child gives that child two
        //     candidate names for one wire, so neither is recorded;
        //   - a later disagreeing alias REMOVES the entry rather than overwriting it.
        //
        // Both are silent: an ambiguous hierarchy is a legitimate way to draw a board, and the cost
        // is only that those nets keep behaving exactly as they do today.
        // ###########################################################################################
        private static void BuildAliasesForSheet(
            KiCadSchematic parent,
            IReadOnlyDictionary<string, int> indexByFileName,
            Dictionary<int, Dictionary<string, string>> result)
        {
            var wires = parent.Wires
                .Concat(parent.Polylines)
                .Where(path => path.Points.Count >= 2)
                .ToList();

            if (wires.Count == 0)
            {
                return;
            }

            // Every sheet pin on this parent, with the child it belongs to.
            var pins = new List<(int ChildIndex, string Name, double X, double Y)>();

            foreach (var child in parent.ChildSheets)
            {
                string file = System.IO.Path.GetFileName(child.FileName ?? string.Empty);

                if (string.IsNullOrWhiteSpace(file) ||
                    !indexByFileName.TryGetValue(file, out int childIndex))
                {
                    continue;
                }

                foreach (var pin in child.Pins)
                {
                    string pinName = pin.NormalizedName?.Trim() ?? string.Empty;

                    if (!string.IsNullOrWhiteSpace(pinName) && pin.At != null)
                    {
                        pins.Add((childIndex, pinName, pin.At.X, pin.At.Y));
                    }
                }
            }

            // ONE pin is enough: a lone child pin wired to a LABEL on the parent is the cassette
            // shape, and requiring two here skipped it before the run was ever examined.
            if (pins.Count == 0)
            {
                return;
            }

            // Group the pins into connected runs. Each run is one net under several names.
            var assigned = new bool[pins.Count];

            for (int seed = 0; seed < pins.Count; seed++)
            {
                if (assigned[seed])
                {
                    continue;
                }

                var run = KiCadSheetNetAliases.FloodFromPoint(wires, pins[seed].X, pins[seed].Y);

                var onThisRun = new List<int>();

                for (int i = 0; i < pins.Count; i++)
                {
                    if (assigned[i])
                    {
                        continue;
                    }

                    if (KiCadSheetNetAliases.RunTouches(run, pins[i].X, pins[i].Y))
                    {
                        onThisRun.Add(i);
                        assigned[i] = true;
                    }
                }

                if (onThisRun.Count == 0)
                {
                    continue;
                }

                // ONE pin is enough when a LABEL on the parent names the net, which is the cassette
                // case: the child sheet WRITE pin is wired to a parent label "P3", and the PCB names
                // that copper P3. Requiring two pins missed every net of that shape.

                // Every name this one net is known by: the pin names, plus any label on the run.
                var names = new HashSet<string>(
                    onThisRun.Select(i => pins[i].Name), StringComparer.OrdinalIgnoreCase);

                foreach (var label in KiCadSheetNetAliases.EnumerateLabels(parent))
                {
                    string text = label.NormalizedText?.Trim() ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(text) || label.At == null)
                    {
                        continue;
                    }

                    if (KiCadSheetNetAliases.RunTouches(run, label.At.X, label.At.Y))
                    {
                        names.Add(text);
                    }
                }

                if (names.Count < 2)
                {
                    continue;
                }

                foreach (var group in onThisRun.GroupBy(i => pins[i].ChildIndex))
                {
                    var ownNames = group
                        .Select(i => pins[i].Name)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    // AMBIGUOUS: this child reaches the run through two differently named pins.
                    if (ownNames.Count != 1)
                    {
                        continue;
                    }

                    string ownName = ownNames[0];

                    if (!result.TryGetValue(group.Key, out var aliases))
                    {
                        aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        result[group.Key] = aliases;
                    }

                    foreach (string alias in names)
                    {
                        if (string.Equals(alias, ownName, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        if (aliases.TryGetValue(alias, out string? existing))
                        {
                            // Two different nets claiming one alias on this sheet is exactly the
                            // wrong-net highlight this class exists to avoid, so neither wins.
                            if (!string.Equals(existing, ownName, StringComparison.OrdinalIgnoreCase))
                            {
                                aliases[alias] = string.Empty;
                            }

                            continue;
                        }

                        aliases[alias] = ownName;
                    }
                }
            }
        }

        // ###########################################################################################
        // The wires reachable from one point by walking a sheet connected run.
        //
        // *** IT RETURNS THE WIRES, NOT JUST THEIR ENDPOINTS, and that distinction is the whole
        // second half of this bug. *** A KiCad label attaches anywhere ALONG a wire, not only at its
        // ends: on the C128 I/O sheet the cassette WRITE pin sits at x=210.82, its wire runs to
        // x=200.66, and the "P3" label that names that net sits at x=205.74 - in the middle. An
        // endpoint-only walk reached the wire and still saw no label, so the net stayed unnamed.
        //
        // Growth is still endpoint-to-endpoint (a KiCad run is a chain of segments meeting at their
        // ends), but what comes back is the segments themselves so a caller can test a point against
        // any position along them - the same rule the schematic index own label seeding uses.
        // ###########################################################################################
        private static List<KiCadSchematicPathItem> FloodFromPoint(
            IReadOnlyList<KiCadSchematicPathItem> wires,
            double x,
            double y)
        {
            var reachedPoints = new List<(double X, double Y)> { (x, y) };
            var reachedWires = new List<KiCadSchematicPathItem>();
            var used = new HashSet<int>();

            bool grew = true;

            while (grew)
            {
                grew = false;

                for (int i = 0; i < wires.Count; i++)
                {
                    if (used.Contains(i))
                    {
                        continue;
                    }

                    var first = wires[i].Points[0];
                    var last = wires[i].Points[^1];

                    bool touches =
                        reachedPoints.Any(point => KiCadSheetNetAliases.IsNear(point.X, point.Y, first.X, first.Y)) ||
                        reachedPoints.Any(point => KiCadSheetNetAliases.IsNear(point.X, point.Y, last.X, last.Y));

                    if (!touches)
                    {
                        continue;
                    }

                    used.Add(i);
                    reachedWires.Add(wires[i]);
                    reachedPoints.Add((first.X, first.Y));
                    reachedPoints.Add((last.X, last.Y));
                    grew = true;
                }
            }

            return reachedWires;
        }

        // ###########################################################################################
        // True when a point lies on any segment of any wire in a run, within tolerance.
        //
        // ANYWHERE ALONG THE SEGMENT, not just at its ends - see FloodFromPoint for why.
        // ###########################################################################################
        private static bool RunTouches(
            IReadOnlyList<KiCadSchematicPathItem> run,
            double x,
            double y)
        {
            foreach (var wire in run)
            {
                for (int i = 0; i + 1 < wire.Points.Count; i++)
                {
                    if (KiCadSheetNetAliases.IsPointOnSegment(
                        wire.Points[i].X,
                        wire.Points[i].Y,
                        wire.Points[i + 1].X,
                        wire.Points[i + 1].Y,
                        x,
                        y))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        // Distance from a point to a line SEGMENT (not the infinite line), within tolerance.
        private static bool IsPointOnSegment(
            double ax, double ay, double bx, double by, double px, double py)
        {
            double dx = bx - ax;
            double dy = by - ay;
            double lengthSquared = (dx * dx) + (dy * dy);

            if (lengthSquared <= double.Epsilon)
            {
                return KiCadSheetNetAliases.IsNear(ax, ay, px, py);
            }

            // How far along the segment the perpendicular foot falls, clamped to its ends so a point
            // beyond either end measures to that end rather than to the infinite line.
            double t = (((px - ax) * dx) + ((py - ay) * dy)) / lengthSquared;
            t = Math.Clamp(t, 0.0, 1.0);

            double closestX = ax + (t * dx);
            double closestY = ay + (t * dy);

            double offsetX = px - closestX;
            double offsetY = py - closestY;

            return ((offsetX * offsetX) + (offsetY * offsetY)) <=
                   KiCadSheetNetAliases.PinTouchTolerance * KiCadSheetNetAliases.PinTouchTolerance;
        }

        private static IEnumerable<KiCadSchematicLabel> EnumerateLabels(KiCadSchematic schematic)
        {
            foreach (var label in schematic.Labels.Local)
            {
                yield return label;
            }

            foreach (var label in schematic.Labels.Global)
            {
                yield return label;
            }

            foreach (var label in schematic.Labels.Hierarchical)
            {
                yield return label;
            }
        }

        private static bool IsNear(double ax, double ay, double bx, double by) =>
            Math.Abs(ax - bx) <= KiCadSheetNetAliases.PinTouchTolerance &&
            Math.Abs(ay - by) <= KiCadSheetNetAliases.PinTouchTolerance;
    }
}
