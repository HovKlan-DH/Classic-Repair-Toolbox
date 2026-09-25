using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Loads and caches raw KiCad files, then builds a schematic net-path index so the Schematics tab
    // can highlight both PCB copper and schematic wire geometry quickly.
    // ###########################################################################################
    internal static class KiCadProjectLoader
    {
        private static readonly Dictionary<string, KiCadProjectBundle?> thisRawCache =
            new(StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // Loads modern raw KiCad files directly from disk and converts them into a normalized project bundle.
        // ###########################################################################################
        public static async Task<KiCadProjectBundle?> LoadRawAsync(
            IReadOnlyList<string> rawPaths,
            string hardwareName = "",
            string boardName = "")
        {
            var existingPaths = rawPaths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(Path.GetFullPath)
                .Where(File.Exists)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (existingPaths.Count == 0)
            {
                return null;
            }

            string cacheKey = string.Join(
                "\u001E",
                existingPaths.Select(path =>
                    $"{path}|{File.GetLastWriteTimeUtc(path).Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture)}"));

            if (KiCadProjectLoader.thisRawCache.TryGetValue(cacheKey, out var cached))
            {
                return cached;
            }

            try
            {
                var root = await KiCadRawProjectLoader.LoadAsync(existingPaths, hardwareName, boardName).ConfigureAwait(false);

                if (root == null || (root.Pcb.Count == 0 && root.Schematics.Count == 0))
                {
                    Logger.Warning($"Raw KiCad files could not be loaded: [{string.Join("], [", existingPaths)}]");
                    KiCadProjectLoader.thisRawCache[cacheKey] = null;
                    return null;
                }

                var bundle = new KiCadProjectBundle
                {
                    Root = root,
                    SchematicNetPathIndexBySchematicIndex =
                        KiCadProjectLoader.BuildSchematicNetPathIndex(root.Schematics)
                };

                KiCadProjectLoader.thisRawCache[cacheKey] = bundle;

                return bundle;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load raw KiCad files [{string.Join("], [", existingPaths)}] - [{ex}]");
                KiCadProjectLoader.thisRawCache[cacheKey] = null;
                return null;
            }
        }

        // ###########################################################################################
        // Builds a per-schematic lookup of normalized net names to connected wire/polyline paths.
        // The MVP logic seeds paths from labels touching a path and then floods through endpoint-
        // connected paths. This is sufficient for most KiCad sheets that use standard labels.
        // ###########################################################################################
        private static Dictionary<int, Dictionary<string, List<KiCadResolvedPath>>> BuildSchematicNetPathIndex(
            IReadOnlyList<KiCadSchematic> schematics)
        {
            var result = new Dictionary<int, Dictionary<string, List<KiCadResolvedPath>>>();

            for (int schematicIndex = 0; schematicIndex < schematics.Count; schematicIndex++)
            {
                var schematic = schematics[schematicIndex];

                var sourcePaths = schematic.Wires
                    .Concat(schematic.Polylines)
                    .Where(path => path.Points.Count >= 2)
                    .Select(path => new IndexedPath
                    {
                        Points = path.Points
                            .Select(point => new KiCadPoint2D
                            {
                                X = point.X,
                                Y = point.Y
                            })
                            .ToList()
                    })
                    .ToList();

                if (sourcePaths.Count == 0)
                {
                    result[schematicIndex] =
                        new Dictionary<string, List<KiCadResolvedPath>>(StringComparer.OrdinalIgnoreCase);
                    continue;
                }

                var adjacency = KiCadProjectLoader.BuildEndpointAdjacency(sourcePaths);
                var indexByNet = new Dictionary<string, List<KiCadResolvedPath>>(StringComparer.OrdinalIgnoreCase);

                foreach (var label in KiCadProjectLoader.EnumerateNetLabels(schematic))
                {
                    string normalizedText = label.NormalizedText?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(normalizedText) || label.At == null)
                    {
                        continue;
                    }

                    var seedPathIndexes = new List<int>();

                    for (int i = 0; i < sourcePaths.Count; i++)
                    {
                        if (KiCadProjectLoader.PathContainsPoint(
                            sourcePaths[i].Points,
                            label.At.X,
                            label.At.Y,
                            0.35))
                        {
                            seedPathIndexes.Add(i);
                        }
                    }

                    if (seedPathIndexes.Count == 0)
                    {
                        continue;
                    }

                    if (!indexByNet.TryGetValue(normalizedText, out var targetList))
                    {
                        targetList = new List<KiCadResolvedPath>();
                        indexByNet[normalizedText] = targetList;
                    }

                    var visited = new HashSet<int>();
                    var queue = new Queue<int>(seedPathIndexes);

                    while (queue.Count > 0)
                    {
                        int current = queue.Dequeue();
                        if (!visited.Add(current))
                        {
                            continue;
                        }

                        targetList.Add(new KiCadResolvedPath
                        {
                            Points = sourcePaths[current].Points
                                .Select(point => new KiCadPoint2D
                                {
                                    X = point.X,
                                    Y = point.Y
                                })
                                .ToList()
                        });

                        foreach (int adjacent in adjacency[current])
                        {
                            if (!visited.Contains(adjacent))
                            {
                                queue.Enqueue(adjacent);
                            }
                        }
                    }
                }

                result[schematicIndex] = indexByNet;
            }

            KiCadProjectLoader.ApplyCrossSheetAliases(schematics, result);

            return result;
        }

        // ###########################################################################################
        // Registers each sheet's already-resolved wires under the OTHER names the same net answers
        // to elsewhere in the hierarchy (maintainer report, 2026-09-24).
        //
        // *** THE REPORTED BUG. *** Selecting a component reads net names off the PCB pads, but a
        // hierarchical design renames a net at every sheet boundary - the C128's userport sheet
        // draws "CNT1" while the PCB calls that copper "/I{slash}O/Serial Bus/CNT". The lookup
        // missed, so selecting CN9 highlighted nothing for those two pins while hovering them still
        // worked (hover hit-tests geometry and never does this lookup).
        //
        // The alias table is built by KiCadSheetNetAliases, which is where the rules and the
        // ambiguity guards live. This method only applies it, and applies it to the SAME
        // KiCadResolvedPath instances the sheet already holds - the alias is another way to ask for
        // one net, not a second copy of its geometry.
        //
        // AN ALIAS NEVER REPLACES A REAL LABEL. A sheet that genuinely draws the alias name keeps
        // its own wires for it; borrowing a name the sheet already uses would move a working
        // highlight onto the wrong wire, which is worse than the bug being fixed.
        // ###########################################################################################
        private static void ApplyCrossSheetAliases(
            IReadOnlyList<KiCadSchematic> schematics,
            Dictionary<int, Dictionary<string, List<KiCadResolvedPath>>> result)
        {
            Dictionary<int, Dictionary<string, string>> aliasesBySheet;

            try
            {
                aliasesBySheet = Handlers.Geometry.KiCadSheetNetAliases.Build(schematics);
            }
            catch (Exception ex)
            {
                // A malformed hierarchy must not cost the whole overlay - without aliases the app
                // behaves exactly as it did before this was added.
                Logger.Warning($"Could not build KiCad cross-sheet net aliases - [{ex.Message}]");
                return;
            }

            foreach (var (schematicIndex, aliases) in aliasesBySheet)
            {
                if (!result.TryGetValue(schematicIndex, out var indexByNet))
                {
                    continue;
                }

                foreach (var (alias, ownName) in aliases)
                {
                    // Emptied by the builder when two nets claimed one alias on this sheet.
                    if (string.IsNullOrWhiteSpace(ownName))
                    {
                        continue;
                    }

                    if (indexByNet.ContainsKey(alias))
                    {
                        continue;
                    }

                    if (indexByNet.TryGetValue(ownName, out var paths) && paths.Count > 0)
                    {
                        indexByNet[alias] = paths;
                    }
                }
            }
        }

        // ###########################################################################################
        // Connects path indexes that share exact or near-exact endpoints. This is the fast backbone
        // used by the schematic net flood-fill.
        // ###########################################################################################
        private static List<HashSet<int>> BuildEndpointAdjacency(IReadOnlyList<IndexedPath> paths)
        {
            const double endpointTolerance = 0.02;

            var adjacency = new List<HashSet<int>>(paths.Count);
            for (int i = 0; i < paths.Count; i++)
            {
                adjacency.Add(new HashSet<int>());
            }

            var endpointBuckets = new Dictionary<string, List<int>>(StringComparer.Ordinal);

            for (int i = 0; i < paths.Count; i++)
            {
                var first = paths[i].Points.First();
                var last = paths[i].Points.Last();

                KiCadProjectLoader.AddEndpointBucket(
                    endpointBuckets,
                    KiCadProjectLoader.MakePointKey(first.X, first.Y, endpointTolerance),
                    i);

                KiCadProjectLoader.AddEndpointBucket(
                    endpointBuckets,
                    KiCadProjectLoader.MakePointKey(last.X, last.Y, endpointTolerance),
                    i);
            }

            foreach (var bucket in endpointBuckets.Values)
            {
                if (bucket.Count < 2)
                {
                    continue;
                }

                for (int i = 0; i < bucket.Count; i++)
                {
                    for (int j = i + 1; j < bucket.Count; j++)
                    {
                        adjacency[bucket[i]].Add(bucket[j]);
                        adjacency[bucket[j]].Add(bucket[i]);
                    }
                }
            }

            return adjacency;
        }

        // ###########################################################################################
        // Returns all electrical label types that can identify a net on a schematic page.
        // ###########################################################################################
        private static IEnumerable<KiCadSchematicLabel> EnumerateNetLabels(KiCadSchematic schematic)
        {
            foreach (var item in schematic.Labels.Local)
            {
                yield return item;
            }

            foreach (var item in schematic.Labels.Global)
            {
                yield return item;
            }

            foreach (var item in schematic.Labels.Hierarchical)
            {
                yield return item;
            }
        }

        // ###########################################################################################
        // Returns true when the supplied point lies on any segment inside the given path.
        // ###########################################################################################
        private static bool PathContainsPoint(
            IReadOnlyList<KiCadPoint2D> points,
            double x,
            double y,
            double tolerance)
        {
            if (points.Count < 2)
            {
                return false;
            }

            for (int i = 0; i < points.Count - 1; i++)
            {
                double distance = KiCadProjectLoader.DistancePointToSegment(
                    x,
                    y,
                    points[i].X,
                    points[i].Y,
                    points[i + 1].X,
                    points[i + 1].Y);

                if (distance <= tolerance)
                {
                    return true;
                }
            }

            return false;
        }

        // ###########################################################################################
        // Returns the shortest distance between one point and one line segment.
        // ###########################################################################################
        private static double DistancePointToSegment(
            double px,
            double py,
            double x1,
            double y1,
            double x2,
            double y2)
        {
            double dx = x2 - x1;
            double dy = y2 - y1;

            if (Math.Abs(dx) < 1e-12 && Math.Abs(dy) < 1e-12)
            {
                double onlyDx = px - x1;
                double onlyDy = py - y1;
                return Math.Sqrt((onlyDx * onlyDx) + (onlyDy * onlyDy));
            }

            double t = (((px - x1) * dx) + ((py - y1) * dy)) / ((dx * dx) + (dy * dy));
            t = Math.Clamp(t, 0.0, 1.0);

            double cx = x1 + (t * dx);
            double cy = y1 + (t * dy);

            double ddx = px - cx;
            double ddy = py - cy;

            return Math.Sqrt((ddx * ddx) + (ddy * ddy));
        }

        // ###########################################################################################
        // Quantizes one point into a stable string key so near-identical coordinates collapse into
        // the same endpoint bucket.
        // ###########################################################################################
        private static string MakePointKey(double x, double y, double tolerance)
        {
            long qx = (long)Math.Round(x / tolerance, MidpointRounding.AwayFromZero);
            long qy = (long)Math.Round(y / tolerance, MidpointRounding.AwayFromZero);
            return $"{qx}:{qy}";
        }

        // ###########################################################################################
        // Adds one path index into the specified endpoint bucket.
        // ###########################################################################################
        private static void AddEndpointBucket(
            Dictionary<string, List<int>> buckets,
            string key,
            int pathIndex)
        {
            if (!buckets.TryGetValue(key, out var bucket))
            {
                bucket = new List<int>();
                buckets[key] = bucket;
            }

            bucket.Add(pathIndex);
        }

        private sealed class IndexedPath
        {
            public List<KiCadPoint2D> Points { get; init; } = new();
        }
    }
}