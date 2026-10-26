using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Loads modern raw KiCad files directly and converts them to the normalized KiCadProjectRoot
    // used by the schematics overlay.
    // ###########################################################################################
    internal static class KiCadRawProjectLoader
    {

        // ###########################################################################################
        // Loads all supported modern KiCad raw files and builds a normalized project root.
        // ###########################################################################################
        // `cancellationToken` stops it between files (OperationCanceledException) - the next board
        // switch does, so a board left behind does not go on parsing a 44 MB PCB.
        public static async Task<KiCadProjectRoot?> LoadAsync(
            IReadOnlyList<string> rawPaths,
            string hardwareName = "",
            string boardName = "",
            CancellationToken cancellationToken = default)
        {
            var expandedPaths = ExpandRawKiCadPaths(rawPaths);

            var pcbs = new List<KiCadPcb>();
            var schematics = new List<KiCadSchematic>();
            var schematicPaths = new List<string>();

            foreach (string path in expandedPaths)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string extension = Path.GetExtension(path);

                if (string.Equals(extension, ".kicad_pcb", StringComparison.OrdinalIgnoreCase))
                {
                    string content = await File.ReadAllTextAsync(path, cancellationToken).ConfigureAwait(false);
                    cancellationToken.ThrowIfCancellationRequested();
                    var forms = SExpressionParser.Parse(content);
                    var root = forms.FirstOrDefault(node =>
                        string.Equals(Head(node), "kicad_pcb", StringComparison.OrdinalIgnoreCase));

                    if (root != null)
                    {
                        pcbs.Add(ParsePcb(root, Path.GetFileName(path)));
                    }

                    continue;
                }

                if (string.Equals(extension, ".kicad_sch", StringComparison.OrdinalIgnoreCase))
                {
                    schematicPaths.Add(path);
                }
            }

            await LoadSchematicsAsync(schematicPaths, schematics, cancellationToken).ConfigureAwait(false);

            if (pcbs.Count == 0 && schematics.Count == 0)
            {
                return null;
            }

            var views = BuildProjectViews(pcbs, schematics);
            LogKiCadInformation(hardwareName, boardName, pcbs, schematics, views);

            return new KiCadProjectRoot
            {
                Ok = true,
                Project = new KiCadProjectInfo
                {
                    Views = views
                },
                Pcb = pcbs,
                Schematics = schematics
            };
        }

        // ###########################################################################################
        // The view names LoadAsync would give these files, in the same order - WITHOUT PARSING A
        // PCB (2026-10-10). The "Schematic images" window offers them as CAD names, and a board's
        // .kicad_pcb runs to tens of megabytes that window has no other use for.
        //
        // A PCB's two views are named after its FILE, so only its opening is read, to keep a file
        // that is no KiCad board out as LoadAsync keeps it out. The schematic sheets ARE read as
        // LoadAsync reads them, because a child sheet's name is the one its parent gives it.
        //
        // Never throws: an unreadable file gives fewer names, not a broken window.
        // ###########################################################################################
        public static async Task<List<string>> LoadViewDisplayNamesAsync(IReadOnlyList<string> rawPaths)
        {
            try
            {
                var pcbs = new List<KiCadPcb>();
                var schematicPaths = new List<string>();

                foreach (string path in ExpandRawKiCadPaths(rawPaths).Where(File.Exists))
                {
                    string extension = Path.GetExtension(path);

                    if (string.Equals(extension, ".kicad_pcb", StringComparison.OrdinalIgnoreCase))
                    {
                        if (await StartsAsKiCadPcbAsync(path).ConfigureAwait(false))
                        {
                            pcbs.Add(new KiCadPcb { Filename = Path.GetFileName(path) });
                        }

                        continue;
                    }

                    if (string.Equals(extension, ".kicad_sch", StringComparison.OrdinalIgnoreCase))
                    {
                        schematicPaths.Add(path);
                    }
                }

                var schematics = new List<KiCadSchematic>();
                await LoadSchematicsAsync(schematicPaths, schematics).ConfigureAwait(false);

                return BuildProjectViews(pcbs, schematics)
                    .Select(view => view.DisplayName?.Trim() ?? string.Empty)
                    .Where(name => name.Length > 0)
                    .ToList();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not read the KiCad view names - [{ex.Message}]");
                return new List<string>();
            }
        }

        // True when the file opens with a "(kicad_pcb" form - LoadAsync's test for a PCB, made on
        // the first few hundred characters instead of the whole parsed file.
        private static async Task<bool> StartsAsKiCadPcbAsync(string path)
        {
            using var reader = new StreamReader(path, detectEncodingFromByteOrderMarks: true);

            char[] buffer = new char[256];
            int read = await reader.ReadBlockAsync(buffer, 0, buffer.Length).ConfigureAwait(false);

            string opening = new string(buffer, 0, read).TrimStart();

            if (!opening.StartsWith('('))
            {
                return false;
            }

            string head = opening.Substring(1).TrimStart();

            return head.StartsWith("kicad_pcb", StringComparison.OrdinalIgnoreCase) &&
                   (head.Length == "kicad_pcb".Length || char.IsWhiteSpace(head["kicad_pcb".Length]) || head["kicad_pcb".Length] == ')');
        }

        // ###########################################################################################
        // Logs a compact KiCad project summary grouped by source file and matching display names.
        // ###########################################################################################
        private static void LogKiCadInformation(
            string hardwareName,
            string boardName,
            IReadOnlyList<KiCadPcb> pcbs,
            IReadOnlyList<KiCadSchematic> schematics,
            IReadOnlyList<KiCadProjectView> views)
        {
            string hardware = string.IsNullOrWhiteSpace(hardwareName) ? "unknown hardware" : hardwareName.Trim();
            string board = string.IsNullOrWhiteSpace(boardName) ? "unknown board" : boardName.Trim();

            Logger.Info($"KiCad information for [{hardware}] [{board}]:");

            for (int i = 0; i < pcbs.Count; i++)
            {
                var pcb = pcbs[i];
                int padCount = pcb.Footprints.Sum(footprint => footprint.Pads.Count);

                Logger.Info(
                    $"  File [{pcb.Filename}]; nets [{pcb.Nets.List.Count}], footprints [{pcb.Footprints.Count}], pads [{padCount}], segments [{pcb.Routing.Segments.Count}], vias [{pcb.Routing.Vias.Count}], arcs [{pcb.Routing.Arcs.Count}], zones [{pcb.Routing.Zones.Count}]");

                foreach (var view in views.Where(view =>
                             string.Equals(view.SourceKind, "pcb", StringComparison.OrdinalIgnoreCase) &&
                             view.SourceIndex == i))
                {
                    Logger.Info(
                        $"    Display name [{view.DisplayName}]; type [{view.Type}], source_kind [{view.SourceKind}], id [{view.Id}]");
                }
            }

            for (int i = 0; i < schematics.Count; i++)
            {
                var schematic = schematics[i];
                int labelCount =
                    schematic.Labels.Local.Count +
                    schematic.Labels.Global.Count +
                    schematic.Labels.Hierarchical.Count;

                Logger.Info(
                    $"  File [{schematic.Filename}]; wires [{schematic.Wires.Count}], polylines [{schematic.Polylines.Count}], labels [{labelCount}], symbols [{schematic.Symbols.Count}]");

                foreach (var view in views.Where(view =>
                             string.Equals(view.SourceKind, "schematic", StringComparison.OrdinalIgnoreCase) &&
                             view.SourceIndex == i))
                {
                    Logger.Info(
                        $"    Display name [{view.DisplayName}]; type [{view.Type}], source_kind [{view.SourceKind}], id [{view.Id}]");
                }
            }
        }

        // ###########################################################################################
        // Expands project-file references and keeps only supported modern KiCad raw files.
        // ###########################################################################################
        private static List<string> ExpandRawKiCadPaths(IReadOnlyList<string> rawPaths)
        {
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AddPath(string path)
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    return;
                }

                string fullPath = Path.GetFullPath(path);

                if (!File.Exists(fullPath))
                {
                    Logger.Warning($"KiCad raw file not found: [{fullPath}]");
                    return;
                }

                string extension = Path.GetExtension(fullPath);
                if (!IsSupportedModernKiCadFile(extension))
                {
                    return;
                }

                if (seen.Add(fullPath))
                {
                    result.Add(fullPath);
                }
            }

            foreach (string rawPath in rawPaths)
            {
                AddPath(rawPath);

                if (!string.Equals(Path.GetExtension(rawPath), ".kicad_pro", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string directory = Path.GetDirectoryName(Path.GetFullPath(rawPath)) ?? string.Empty;
                string baseName = Path.GetFileNameWithoutExtension(rawPath);

                AddPath(Path.Combine(directory, $"{baseName}.kicad_pcb"));
                AddPath(Path.Combine(directory, $"{baseName}.kicad_sch"));
            }

            return result;
        }

        // ###########################################################################################
        // Returns true when the extension is a supported modern KiCad file type.
        // ###########################################################################################
        private static bool IsSupportedModernKiCadFile(string extension)
        {
            return string.Equals(extension, ".kicad_pcb", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".kicad_pro", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".kicad_sch", StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // Parses one KiCad PCB S-expression tree into the normalized DTO.
        // Supports modern footprint nodes and legacy module nodes inside .kicad_pcb files.
        // ###########################################################################################
        private static KiCadPcb ParsePcb(SExpressionNode pcb, string filename)
        {
            var nets = ExtractPcbNets(pcb);
            var netMap = nets
                .Where(net => !string.IsNullOrWhiteSpace(net.Id))
                .ToDictionary(net => net.Id!, net => net.Name ?? net.Id!, StringComparer.OrdinalIgnoreCase);

            // KiCad's "no net" (net 0 "") - a pad on it is on no net (KiCadNetRef.IsNoNet).
            var noNetIds = new HashSet<string>(
                nets.Where(net => net.IsNoNet && !string.IsNullOrWhiteSpace(net.Id)).Select(net => net.Id!),
                StringComparer.OrdinalIgnoreCase);

            var footprints = ExtractFootprints(pcb, netMap, noNetIds);
            var routing = ExtractRouting(pcb, netMap);

            var result = new KiCadPcb
            {
                Filename = filename,
                Nets = new KiCadPcbNets
                {
                    List = nets
                },
                Footprints = footprints,
                Routing = routing
            };

            var normalized = new KiCadPcb
            {
                Filename = result.Filename,
                Nets = result.Nets,
                Footprints = result.Footprints,
                Routing = result.Routing,
                HighlightIndex = BuildPcbHighlightIndex(result)
            };

            return normalized;
        }

        // ###########################################################################################
        // Returns both modern footprint nodes and legacy module nodes from a KiCad PCB file.
        // Older KiCad PCB files can use module even though the board file extension is .kicad_pcb.
        // ###########################################################################################
        private static IEnumerable<SExpressionNode> PcbFootprintNodes(SExpressionNode pcb)
        {
            foreach (var footprint in Children(pcb, "footprint"))
            {
                yield return footprint;
            }

            foreach (var module in Children(pcb, "module"))
            {
                yield return module;
            }
        }

        // ###########################################################################################
        // Extracts all PCB net declarations and inline net references from a PCB.
        // Supports both modern footprint nodes and legacy module nodes.
        // ###########################################################################################
        private static List<KiCadNetRef> ExtractPcbNets(SExpressionNode pcb)
        {
            var netsById = new Dictionary<string, KiCadNetRef>(StringComparer.OrdinalIgnoreCase);

            void AddOrUpdate(string? rawId, string? rawName = null)
            {
                string id = rawId?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(id))
                {
                    return;
                }

                string name = string.IsNullOrWhiteSpace(rawName) ? id : rawName.Trim();

                // KiCad's "no net", declared (net 0 "") - a name GIVEN, and blank. A KiCad 9 net
                // named "0" is (net "0"), with no second value (KiCadNetRef.IsNoNet).
                bool declaredNoNet = id == "0" && rawName is not null && rawName.Trim().Length == 0;

                if (!netsById.ContainsKey(id))
                {
                    netsById[id] = new KiCadNetRef
                    {
                        Id = id,
                        Name = name,
                        NormalizedName = NormalizeNetName(name),
                        IsNoNet = declaredNoNet
                    };

                    return;
                }

                var existing = netsById[id];
                if (string.IsNullOrWhiteSpace(existing.Name) || string.Equals(existing.Name, id, StringComparison.OrdinalIgnoreCase))
                {
                    netsById[id] = new KiCadNetRef
                    {
                        Id = existing.Id,
                        Name = name,
                        NormalizedName = NormalizeNetName(name),

                        // A later mention by number alone - a track on net 0 - keeps what the net
                        // is; a real name given later makes it a net like any other.
                        IsNoNet = string.IsNullOrWhiteSpace(rawName) && (existing.IsNoNet || declaredNoNet)
                    };
                }
            }

            foreach (var netNode in Children(pcb, "net"))
            {
                AddOrUpdate(Arg(netNode, 0), Arg(netNode, 1));
            }

            foreach (var footprint in PcbFootprintNodes(pcb))
            {
                foreach (var pad in Children(footprint, "pad"))
                {
                    var netNode = Child(pad, "net");
                    if (netNode != null)
                    {
                        AddOrUpdate(Arg(netNode, 0), Arg(netNode, 1));
                    }
                }
            }

            foreach (var segment in Children(pcb, "segment"))
            {
                AddOrUpdate(ChildValue(segment, "net"));
            }

            foreach (var via in Children(pcb, "via"))
            {
                AddOrUpdate(ChildValue(via, "net"));
            }

            foreach (var arc in Children(pcb, "arc"))
            {
                AddOrUpdate(ChildValue(arc, "net"));
            }

            foreach (var zone in Children(pcb, "zone"))
            {
                AddOrUpdate(ChildValue(zone, "net"), ChildValue(zone, "net_name"));
            }

            return netsById.Values
                .OrderBy(net => net.Id, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ###########################################################################################
        // Extracts KiCad footprint/module and pad data, including absolute pad centers.
        // Supports modern footprint nodes and legacy module nodes inside .kicad_pcb files.
        // ###########################################################################################
        private static List<KiCadPcbFootprint> ExtractFootprints(
            SExpressionNode pcb,
            IReadOnlyDictionary<string, string> netMap,
            IReadOnlySet<string> noNetIds)
        {
            var footprints = new List<KiCadPcbFootprint>();

            foreach (var footprintNode in PcbFootprintNodes(pcb))
            {
                var footprintAt = ExtractAt(Child(footprintNode, "at"));
                string footprintLayer = ChildValue(footprintNode, "layer") ?? string.Empty;
                string reference = ExtractFootprintTextValue(footprintNode, "reference") ?? string.Empty;

                double footprintAngle = footprintAt?.Angle ?? 0.0;

                var pads = new List<KiCadPcbPad>();

                foreach (var padNode in Children(footprintNode, "pad"))
                {
                    // A pad's (at x y angle) mixes two frames: x/y are footprint-local and
                    // unrotated, while the angle is already absolute - KiCad writes the parent
                    // footprint's rotation into it. So the position needs rotating here and the
                    // angle does not. Back-side footprints need no mirroring either: KiCad bakes
                    // the flip into the stored local coordinates when the footprint is flipped.
                    var padAt = ExtractAt(Child(padNode, "at")) ?? new KiCadPoint2DAngle();
                    var rotated = RotatePoint(padAt.X, padAt.Y, footprintAngle);

                    var netNode = Child(padNode, "net");
                    string? netId = Arg(netNode, 0);
                    string? inlineNetName = Arg(netNode, 1);
                    string? netName = !string.IsNullOrWhiteSpace(inlineNetName)
                        ? inlineNetName
                        : !string.IsNullOrWhiteSpace(netId) && netMap.TryGetValue(netId, out string? mappedName)
                            ? mappedName
                            : netId;

                    KiCadPoint2D? absoluteCenter = footprintAt == null
                        ? null
                        : new KiCadPoint2D
                        {
                            X = footprintAt.X + rotated.X,
                            Y = footprintAt.Y + rotated.Y
                        };

                    pads.Add(new KiCadPcbPad
                    {
                        Number = Arg(padNode, 0),
                        Shape = Arg(padNode, 2),
                        AbsoluteCenter = absoluteCenter,
                        Size = ExtractSize(Child(padNode, "size")),
                        RotationDegrees = padAt.Angle ?? 0.0,
                        Layers = Args(Child(padNode, "layers")).ToList(),
                        Net = string.IsNullOrWhiteSpace(netId) && string.IsNullOrWhiteSpace(netName)
                            ? null
                            : new KiCadNetRef
                            {
                                Id = netId,
                                Name = netName,
                                NormalizedName = NormalizeNetName(netName),

                                // An older file's unconnected pad, (net 0 "") - its name reads "0"
                                // (code review, 2026-10-26: the pin showed net "0").
                                IsNoNet = !string.IsNullOrWhiteSpace(netId) && noNetIds.Contains(netId) && string.IsNullOrWhiteSpace(inlineNetName)
                            },
                        PinFunction = ChildValue(padNode, "pinfunction"),
                        PinType = ChildValue(padNode, "pintype")
                    });
                }

                footprints.Add(new KiCadPcbFootprint
                {
                    Reference = reference,
                    Layer = footprintLayer,
                    FootprintName = Arg(footprintNode, 0),
                    Value = ExtractFootprintTextValue(footprintNode, "value"),
                    Datasheet = ExtractFootprintTextValue(footprintNode, "Datasheet"),
                    Pads = pads
                });
            }

            return footprints;
        }

        // ###########################################################################################
        // Extracts PCB tracks, vias, arcs, and copper zones from a modern KiCad PCB.
        // ###########################################################################################
        private static KiCadPcbRouting ExtractRouting(
            SExpressionNode pcb,
            IReadOnlyDictionary<string, string> netMap)
        {
            var segments = new List<KiCadPcbSegment>();
            var vias = new List<KiCadPcbVia>();
            var arcs = new List<KiCadPcbArc>();
            var zones = new List<KiCadPcbZone>();

            KiCadNetRef? BuildNetRef(string? netId, string? explicitName = null)
            {
                if (string.IsNullOrWhiteSpace(netId) && string.IsNullOrWhiteSpace(explicitName))
                {
                    return null;
                }

                string? trimmedNetId = string.IsNullOrWhiteSpace(netId) ? null : netId.Trim();
                string name = !string.IsNullOrWhiteSpace(explicitName)
                    ? explicitName.Trim()
                    : !string.IsNullOrWhiteSpace(trimmedNetId) &&
                      netMap.TryGetValue(trimmedNetId, out string? mappedName) &&
                      !string.IsNullOrWhiteSpace(mappedName)
                        ? mappedName
                        : trimmedNetId ?? string.Empty;

                return new KiCadNetRef
                {
                    Id = trimmedNetId,
                    Name = name,
                    NormalizedName = NormalizeNetName(name)
                };
            }

            foreach (var segmentNode in Children(pcb, "segment"))
            {
                string? netId = ChildValue(segmentNode, "net");

                segments.Add(new KiCadPcbSegment
                {
                    Start = ExtractPoint(Child(segmentNode, "start")),
                    End = ExtractPoint(Child(segmentNode, "end")),
                    Width = ToDoubleOrNull(ChildValue(segmentNode, "width")),
                    Layer = ChildValue(segmentNode, "layer"),
                    Net = BuildNetRef(netId)
                });
            }

            foreach (var viaNode in Children(pcb, "via"))
            {
                string? netId = ChildValue(viaNode, "net");

                vias.Add(new KiCadPcbVia
                {
                    At = ExtractAt(Child(viaNode, "at")),
                    Size = ToDoubleOrNull(ChildValue(viaNode, "size")),
                    Layers = Args(Child(viaNode, "layers")).ToList(),
                    Net = BuildNetRef(netId)
                });
            }

            foreach (var arcNode in Children(pcb, "arc"))
            {
                string? netId = ChildValue(arcNode, "net");

                arcs.Add(new KiCadPcbArc
                {
                    Start = ExtractPoint(Child(arcNode, "start")),
                    Mid = ExtractPoint(Child(arcNode, "mid")),
                    End = ExtractPoint(Child(arcNode, "end")),
                    Width = ToDoubleOrNull(ChildValue(arcNode, "width")),
                    Layer = ChildValue(arcNode, "layer"),
                    Net = BuildNetRef(netId)
                });
            }

            foreach (var zoneNode in Children(pcb, "zone"))
            {
                string? netId = ChildValue(zoneNode, "net");
                string? netName = ChildValue(zoneNode, "net_name");

                var outlinePolygons = ExtractZonePolygons(zoneNode, "polygon");
                var filledPolygons = ExtractZonePolygons(zoneNode, "filled_polygon");

                if (outlinePolygons.Count == 0 && filledPolygons.Count == 0)
                {
                    continue;
                }

                zones.Add(new KiCadPcbZone
                {
                    Layers = ExtractZoneLayers(zoneNode),
                    Net = BuildNetRef(netId, netName),
                    OutlinePolygons = outlinePolygons,
                    FilledPolygons = filledPolygons
                });
            }

            return new KiCadPcbRouting
            {
                Segments = segments,
                Vias = vias,
                Arcs = arcs,
                Zones = zones
            };
        }

        // ###########################################################################################
        // Builds the PCB highlight index used by render and hover logic.
        // ###########################################################################################
        private static Dictionary<string, KiCadPcbHighlightBucket> BuildPcbHighlightIndex(KiCadPcb pcb)
        {
            var index = new Dictionary<string, KiCadPcbHighlightBucket>(StringComparer.OrdinalIgnoreCase);

            KiCadPcbHighlightBucket GetBucket(string? netId)
            {
                string key = netId?.Trim() ?? string.Empty;

                if (!index.TryGetValue(key, out var bucket))
                {
                    bucket = new KiCadPcbHighlightBucket();
                    index[key] = bucket;
                }

                return bucket;
            }

            for (int i = 0; i < pcb.Routing.Segments.Count; i++)
            {
                string? netId = pcb.Routing.Segments[i].Net?.Id;
                if (!string.IsNullOrWhiteSpace(netId))
                {
                    GetBucket(netId).Segments.Add(i);
                }
            }

            for (int i = 0; i < pcb.Routing.Vias.Count; i++)
            {
                string? netId = pcb.Routing.Vias[i].Net?.Id;
                if (!string.IsNullOrWhiteSpace(netId))
                {
                    GetBucket(netId).Vias.Add(i);
                }
            }

            for (int i = 0; i < pcb.Routing.Arcs.Count; i++)
            {
                string? netId = pcb.Routing.Arcs[i].Net?.Id;
                if (!string.IsNullOrWhiteSpace(netId))
                {
                    GetBucket(netId).Arcs.Add(i);
                }
            }

            for (int i = 0; i < pcb.Routing.Zones.Count; i++)
            {
                string? netId = pcb.Routing.Zones[i].Net?.Id;
                if (!string.IsNullOrWhiteSpace(netId))
                {
                    GetBucket(netId).Zones.Add(i);
                }
            }

            for (int footprintIndex = 0; footprintIndex < pcb.Footprints.Count; footprintIndex++)
            {
                var footprint = pcb.Footprints[footprintIndex];

                for (int padIndex = 0; padIndex < footprint.Pads.Count; padIndex++)
                {
                    var pad = footprint.Pads[padIndex];
                    string? netId = pad.Net?.Id;

                    if (string.IsNullOrWhiteSpace(netId))
                    {
                        continue;
                    }

                    GetBucket(netId).Pads.Add(new KiCadPcbHighlightPadRef
                    {
                        FootprintIndex = footprintIndex,
                        PadIndex = padIndex,
                        Reference = footprint.Reference,
                        PadNumber = pad.Number
                    });
                }
            }

            return index
                .Where(pair => !string.IsNullOrWhiteSpace(pair.Key))
                .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // Extracts the copper layers used by one PCB zone.
        // Supports both single-layer and multi-layer syntax.
        // ###########################################################################################
        private static List<string> ExtractZoneLayers(SExpressionNode zoneNode)
        {
            var layers = Args(Child(zoneNode, "layers"))
                .Where(layer => !string.IsNullOrWhiteSpace(layer))
                .Select(layer => layer.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (layers.Count > 0)
            {
                return layers;
            }

            string singleLayer = ChildValue(zoneNode, "layer")?.Trim() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(singleLayer))
            {
                layers.Add(singleLayer);
            }

            return layers;
        }

        // ###########################################################################################
        // Extracts one group of zone polygons from a zone node.
        // Uses the filled polygons when present and falls back to outline polygons otherwise.
        // ###########################################################################################
        private static List<KiCadPcbZonePolygon> ExtractZonePolygons(SExpressionNode zoneNode, string polygonNodeName)
        {
            return Children(zoneNode, polygonNodeName)
                .Select(polygonNode => new KiCadPcbZonePolygon
                {
                    Points = ExtractPts(polygonNode)
                })
                .Where(polygon => polygon.Points.Count >= 3)
                .ToList();
        }

        // ###########################################################################################
        // Parses one modern KiCad schematic S-expression tree into the normalized DTO.
        // ###########################################################################################
        private static KiCadSchematic ParseSchematic(SExpressionNode schematic, string filename, string? displayName)
        {
            string fallbackDisplayName = Path.GetFileNameWithoutExtension(filename);

            return new KiCadSchematic
            {
                Filename = filename,
                DisplayName = string.IsNullOrWhiteSpace(displayName) ? fallbackDisplayName : displayName,
                Wires = ExtractSchematicPathGroup(schematic, "wire"),
                Polylines = ExtractSchematicPathGroup(schematic, "polyline"),
                Labels = new KiCadSchematicLabels
                {
                    Local = ExtractSchematicLabelGroup(schematic, "label"),
                    Global = ExtractSchematicLabelGroup(schematic, "global_label"),
                    Hierarchical = ExtractSchematicLabelGroup(schematic, "hierarchical_label")
                },
                Symbols = ExtractSchematicSymbols(schematic),
                ChildSheets = ExtractSchematicChildSheets(schematic)
            };
        }

        // ###########################################################################################
        // Extracts the child-sheet symbols drawn on this sheet, each with its boundary pins.
        //
        // *** WITHOUT THIS A HIERARCHICAL NET CANNOT BE FOLLOWED ACROSS A SHEET, *** which is the
        // whole reason it exists - see KiCadSchematicChildSheet. A sheet pin's NAME is the label the
        // CHILD uses, and its POSITION is where the PARENT's wire lands, so the pair is what ties
        // the two sheets' vocabularies together.
        //
        // A KiCad "(sheet ...)" node carries both the Sheetname/Sheetfile properties (already read
        // by ExtractSchematicSheetRefs for the load traversal) and its "(pin ...)" children. The two
        // readers stay separate on purpose: the traversal runs before any sheet is parsed, and needs
        // only the file name.
        // ###########################################################################################
        private static List<KiCadSchematicChildSheet> ExtractSchematicChildSheets(SExpressionNode schematic)
        {
            var result = new List<KiCadSchematicChildSheet>();

            foreach (var sheetNode in Children(schematic, "sheet"))
            {
                string? sheetName = null;
                string? fileName = null;

                foreach (var propertyNode in Children(sheetNode, "property"))
                {
                    string? propertyName = Arg(propertyNode, 0);
                    string? propertyValue = Arg(propertyNode, 1);

                    if (string.Equals(propertyName, "Sheetname", StringComparison.OrdinalIgnoreCase))
                    {
                        sheetName = propertyValue;
                    }
                    else if (string.Equals(propertyName, "Sheetfile", StringComparison.OrdinalIgnoreCase))
                    {
                        fileName = propertyValue;
                    }
                }

                var pins = new List<KiCadSchematicSheetPin>();

                foreach (var pinNode in Children(sheetNode, "pin"))
                {
                    string? name = Arg(pinNode, 0);

                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    pins.Add(new KiCadSchematicSheetPin
                    {
                        Name = name,
                        NormalizedName = NormalizeNetName(name),
                        At = ExtractAt(Child(pinNode, "at"))
                    });
                }

                result.Add(new KiCadSchematicChildSheet
                {
                    SheetName = sheetName,
                    FileName = fileName,
                    Pins = pins
                });
            }

            return result;
        }

        // ###########################################################################################
        // Extracts schematic wire/polyline items.
        // ###########################################################################################
        private static List<KiCadSchematicPathItem> ExtractSchematicPathGroup(SExpressionNode schematic, string type)
        {
            return Children(schematic, type)
                .Select(node => new KiCadSchematicPathItem
                {
                    Type = type,
                    Points = ExtractPts(node)
                })
                .ToList();
        }

        // ###########################################################################################
        // Extracts schematic labels of the requested KiCad label type.
        // ###########################################################################################
        private static List<KiCadSchematicLabel> ExtractSchematicLabelGroup(SExpressionNode schematic, string type)
        {
            return Children(schematic, type)
                .Select(node =>
                {
                    string? text = Arg(node, 0);

                    return new KiCadSchematicLabel
                    {
                        Type = type,
                        Text = text,
                        NormalizedText = NormalizeNetName(text),
                        At = ExtractAt(Child(node, "at"))
                    };
                })
                .ToList();
        }

        // ###########################################################################################
        // Extracts schematic symbols and their detailed properties.
        // ###########################################################################################
        private static List<KiCadSchematicSymbol> ExtractSchematicSymbols(SExpressionNode schematic)
        {
            var result = new List<KiCadSchematicSymbol>();
            Dictionary<string, List<KiCadSchematicPin>> libraryPins = ExtractLibrarySymbolPins(schematic);

            foreach (var symbolNode in Children(schematic, "symbol"))
            {
                var properties = Children(symbolNode, "property")
                    .Select(propertyNode => new KiCadSchematicSymbolPropertyDetailed
                    {
                        Name = Arg(propertyNode, 0),
                        Value = Arg(propertyNode, 1),
                        At = ExtractAt(Child(propertyNode, "at")),
                        Effects = ExtractTextEffects(Child(propertyNode, "effects"))
                    })
                    .ToList();

                string? reference = properties
                    .FirstOrDefault(property => string.Equals(property.Name, "Reference", StringComparison.OrdinalIgnoreCase))
                    ?.Value;

                string? value = properties
                    .FirstOrDefault(property => string.Equals(property.Name, "Value", StringComparison.OrdinalIgnoreCase))
                    ?.Value;

                // An instance names its library symbol by "lib_name" when the sheet keeps an edited
                // copy of it, else by "lib_id".
                string libraryName = ChildValue(symbolNode, "lib_name") ?? ChildValue(symbolNode, "lib_id") ?? string.Empty;

                result.Add(new KiCadSchematicSymbol
                {
                    Reference = reference,
                    Value = value,
                    At = ExtractAt(Child(symbolNode, "at")),
                    PropertiesDetailed = properties,
                    Pins = libraryPins.TryGetValue(libraryName, out List<KiCadSchematicPin>? pins) ? pins : new List<KiCadSchematicPin>()
                });
            }

            return result;
        }

        // ###########################################################################################
        // Every library symbol the sheet keeps a copy of ("lib_symbols"), by name, with the pins of
        // ALL its units - a 74LS00's four gates and its power pins are one chip's pins - each number
        // once. A De Morgan alternative ("<name>_<unit>_2") draws the same pins again and is passed
        // over. A symbol derived from another ("extends") with no pins of its own takes its parent's.
        // ###########################################################################################
        private static Dictionary<string, List<KiCadSchematicPin>> ExtractLibrarySymbolPins(SExpressionNode schematic)
        {
            var result = new Dictionary<string, List<KiCadSchematicPin>>(StringComparer.Ordinal);

            if (Child(schematic, "lib_symbols") is not { } library)
            {
                return result;
            }

            var parents = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (var symbolNode in Children(library, "symbol"))
            {
                string? name = Arg(symbolNode, 0);

                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                var pins = new List<KiCadSchematicPin>();
                var numbers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Each number once - the first unit to draw it names it.
                void Take(SExpressionNode node) =>
                    pins.AddRange(Children(node, "pin")
                        .Select(pinNode => new KiCadSchematicPin
                        {
                            Number = Arg(Child(pinNode, "number"), 0)?.Trim() ?? string.Empty,
                            Name = Arg(Child(pinNode, "name"), 0) ?? string.Empty,
                            Type = Arg(pinNode, 0) ?? string.Empty
                        })
                        .Where(pin => pin.Number.Length > 0 && numbers.Add(pin.Number)));

                Take(symbolNode);

                foreach (var unitNode in Children(symbolNode, "symbol"))
                {
                    if (!IsAlternateBodyStyle(Arg(unitNode, 0)))
                    {
                        Take(unitNode);
                    }
                }

                result[name] = pins;

                if (ChildValue(symbolNode, "extends") is { Length: > 0 } parent)
                {
                    parents[name] = parent;
                }
            }

            foreach ((string name, string parent) in parents)
            {
                // The parent is named without its library: "74xx:74LS00" extends "74LS01".
                int colon = name.IndexOf(':');
                string sameLibrary = colon >= 0 ? name[..(colon + 1)] + parent : parent;

                if (result[name].Count == 0 &&
                    (result.TryGetValue(sameLibrary, out List<KiCadSchematicPin>? parentPins) || result.TryGetValue(parent, out parentPins)))
                {
                    result[name] = parentPins;
                }
            }

            return result;
        }

        // "74LS00_1_2" - unit 1 drawn in body style 2, KiCad's De Morgan alternative.
        private static bool IsAlternateBodyStyle(string? unitName)
        {
            string[] parts = (unitName ?? string.Empty).Split('_');

            return parts.Length >= 3 && parts[^1] == "2" && int.TryParse(parts[^2], NumberStyles.Integer, CultureInfo.InvariantCulture, out _);
        }

        // ###########################################################################################
        // Extracts schematic text effects relevant for label/reference visibility and anchors.
        // ###########################################################################################
        private static KiCadSchematicTextEffects? ExtractTextEffects(SExpressionNode? effectsNode)
        {
            if (effectsNode == null)
            {
                return null;
            }

            return new KiCadSchematicTextEffects
            {
                Hide = Child(effectsNode, "hide") != null,
                Justify = Args(Child(effectsNode, "justify")).ToList()
            };
        }

        // ###########################################################################################
        // Builds top/bottom PCB views and schematic views for the normalized project.
        // ###########################################################################################
        private static List<KiCadProjectView> BuildProjectViews(
            IReadOnlyList<KiCadPcb> pcbs,
            IReadOnlyList<KiCadSchematic> schematics)
        {
            var views = new List<KiCadProjectView>();

            for (int i = 0; i < pcbs.Count; i++)
            {
                string baseName = Path.GetFileNameWithoutExtension(pcbs[i].Filename);

                views.Add(new KiCadProjectView
                {
                    Id = $"pcb:{i}:top",
                    Type = "pcb_top",
                    DisplayName = $"{baseName} - PCB Top",
                    SourceKind = "pcb",
                    SourceIndex = i,
                    SourceFile = pcbs[i].Filename,
                    BoardSide = "top",
                    LayerHint = "F.Cu"
                });

                views.Add(new KiCadProjectView
                {
                    Id = $"pcb:{i}:bottom",
                    Type = "pcb_bottom",
                    DisplayName = $"{baseName} - PCB Bottom",
                    SourceKind = "pcb",
                    SourceIndex = i,
                    SourceFile = pcbs[i].Filename,
                    BoardSide = "bottom",
                    LayerHint = "B.Cu"
                });
            }

            for (int i = 0; i < schematics.Count; i++)
            {
                string displayName = string.IsNullOrWhiteSpace(schematics[i].DisplayName)
                    ? Path.GetFileNameWithoutExtension(schematics[i].Filename)
                    : schematics[i].DisplayName;

                views.Add(new KiCadProjectView
                {
                    Id = $"schematic:{i}",
                    Type = "schematic",
                    DisplayName = displayName,
                    SourceKind = "schematic",
                    SourceIndex = i,
                    SourceFile = schematics[i].Filename,
                    SheetRole = i == 0 ? "top_level" : "child_sheet"
                });
            }

            return views;
        }

        // ###########################################################################################
        // Loads every schematic in a board folder exactly once, starting from its root sheets.
        //
        // A hierarchical KiCad design names its child sheets inside the parent, and the board folder
        // holds those child files too - so treating every .kicad_sch on disk as a root sheet loads
        // each child twice. The two copies then disagree about their own name: reached through its
        // parent a sheet is named by that parent's "Sheetname" property, loaded standalone it falls
        // back to its filename. That name is what a contributor copies into the board workbook's
        // "CAD name" column, so one sheet offering two names is a trap rather than merely wasteful.
        //
        // Root sheets - files no other file references as a child - are therefore identified first,
        // and the walk runs from those through one shared visited set. Anything still unvisited
        // afterwards (a reference cycle, or a child whose parent is not in the folder) is then loaded
        // as its own root, so no sheet is ever silently dropped.
        // ###########################################################################################
        private static async Task LoadSchematicsAsync(
            IReadOnlyList<string> schematicPaths,
            List<KiCadSchematic> schematics,
            CancellationToken cancellationToken = default)
        {
            var parsedRoots = new Dictionary<string, SExpressionNode>(StringComparer.OrdinalIgnoreCase);

            foreach (string path in schematicPaths)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string fullPath = Path.GetFullPath(path);

                if (parsedRoots.ContainsKey(fullPath))
                {
                    continue;
                }

                var parsedRoot = await ReadSchematicRootAsync(fullPath).ConfigureAwait(false);

                if (parsedRoot != null)
                {
                    parsedRoots[fullPath] = parsedRoot;
                }
            }

            var childPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var entry in parsedRoots)
            {
                foreach (var sheetRef in ExtractSchematicSheetRefs(entry.Value))
                {
                    string? childPath = ResolveChildSheetPath(entry.Key, sheetRef.FileName, parsedRoots);

                    if (childPath != null)
                    {
                        childPaths.Add(childPath);
                    }
                }
            }

            var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Root sheets first, so a child is always reached through its parent and keeps the name
            // that parent gives it - regardless of where the child sorts in the folder listing.
            foreach (string path in schematicPaths)
            {
                string fullPath = Path.GetFullPath(path);

                if (parsedRoots.ContainsKey(fullPath) && !childPaths.Contains(fullPath))
                {
                    await LoadSchematicAndChildSheetsAsync(fullPath, null, schematics, visited, parsedRoots)
                        .ConfigureAwait(false);
                }
            }

            foreach (string path in schematicPaths)
            {
                string fullPath = Path.GetFullPath(path);

                if (parsedRoots.ContainsKey(fullPath))
                {
                    await LoadSchematicAndChildSheetsAsync(fullPath, null, schematics, visited, parsedRoots)
                        .ConfigureAwait(false);
                }
            }
        }

        // ###########################################################################################
        // Reads and parses one KiCad schematic file, returning its root node or null with a warning.
        // ###########################################################################################
        private static async Task<SExpressionNode?> ReadSchematicRootAsync(string fullPath)
        {
            if (!File.Exists(fullPath))
            {
                Logger.Warning($"KiCad schematic file not found: [{fullPath}]");
                return null;
            }

            string content = await File.ReadAllTextAsync(fullPath).ConfigureAwait(false);
            var forms = SExpressionParser.Parse(content);
            var root = forms.FirstOrDefault(node =>
                string.Equals(Head(node), "kicad_sch", StringComparison.OrdinalIgnoreCase));

            if (root == null)
            {
                Logger.Warning($"KiCad schematic root not found in file: [{fullPath}]");
            }

            return root;
        }

        // ###########################################################################################
        // Loads a KiCad schematic file and recursively loads any child sheet files referenced by it.
        // Files already parsed by the caller are reused; a child sheet that was not in the board
        // folder listing is read from disk here - only from inside its parent's own folder
        // (ResolveChildSheetPath).
        // ###########################################################################################
        private static async Task LoadSchematicAndChildSheetsAsync(
            string path,
            string? displayName,
            List<KiCadSchematic> schematics,
            HashSet<string> visited,
            IReadOnlyDictionary<string, SExpressionNode> parsedRoots)
        {
            string fullPath = Path.GetFullPath(path);

            if (!visited.Add(fullPath))
            {
                return;
            }

            if (!parsedRoots.TryGetValue(fullPath, out var root))
            {
                root = await ReadSchematicRootAsync(fullPath).ConfigureAwait(false);
            }

            if (root == null)
            {
                return;
            }

            var schematic = ParseSchematic(root, Path.GetFileName(fullPath), displayName);
            schematics.Add(schematic);

            foreach (var sheetRef in ExtractSchematicSheetRefs(root))
            {
                string? childPath = ResolveChildSheetPath(fullPath, sheetRef.FileName, parsedRoots);

                if (childPath == null)
                {
                    if (!string.IsNullOrWhiteSpace(sheetRef.FileName))
                    {
                        Logger.Warning(
                            $"KiCad sheet [{sheetRef.SheetName}] in [{fullPath}] names a file outside the board's KiCad files, which is not read: [{sheetRef.FileName}]");
                    }

                    continue;
                }

                await LoadSchematicAndChildSheetsAsync(
                        childPath,
                        sheetRef.SheetName,
                        schematics,
                        visited,
                        parsedRoots)
                    .ConfigureAwait(false);
            }
        }

        // ###########################################################################################
        // Where a child sheet a parent names may be read from - or null, and it is not read at all.
        //
        // *** ONLY FROM THE BOARD'S OWN KiCad FILES (security review, 2026-10-11). *** The path is
        // written INSIDE the parent file, and a contributed file can write anything there. Combined
        // with the parent's folder as it was, a "Sheetfile" of "//server/share/x.kicad_sch" became
        // a network path, and the File.Exists on it made every Windows CRT that showed the board
        // open an SMB connection to that server - handing it the user's NTLM hash - with no click
        // beyond picking the board. So a child is followed only when it is
        //
        //   - one of the files the loader was handed (`parsedRoots` - the board's "KiCad data"
        //     folders, scanned), however the parent spells the way to it; or
        //   - a .kicad_sch inside the parent's OWN folder or below it, read from disk - the
        //     project folder a .kicad_pro was expanded from (ExpandRawKiCadPaths).
        //
        // Nothing here touches the disk: the full path is worked out as text, and whatever is
        // refused is never looked at. The server refuses such a file too (CRT.Data's
        // KiCadSheetReferences), for the CRTs installed before this.
        // ###########################################################################################
        private static string? ResolveChildSheetPath(
            string parentFullPath,
            string? sheetFile,
            IReadOnlyDictionary<string, SExpressionNode> parsedRoots)
        {
            if (string.IsNullOrWhiteSpace(sheetFile))
            {
                return null;
            }

            string directory = Path.GetDirectoryName(parentFullPath) ?? string.Empty;
            string childPath;

            try
            {
                childPath = Path.GetFullPath(Path.Combine(directory, sheetFile));
            }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
            {
                return null;
            }

            if (parsedRoots.ContainsKey(childPath))
            {
                return childPath;
            }

            // Case folded on Windows only - the comparison ExternalTargetLauncher holds its
            // data-root containment to.
            StringComparison comparison = OperatingSystem.IsWindows()
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;

            string folder = Path.EndsInDirectorySeparator(directory)
                ? directory
                : directory + Path.DirectorySeparatorChar;

            return childPath.StartsWith(folder, comparison) &&
                   string.Equals(Path.GetExtension(childPath), ".kicad_sch", StringComparison.OrdinalIgnoreCase)
                ? childPath
                : null;
        }

        // ###########################################################################################
        // Extracts child sheet references from a KiCad schematic hierarchy sheet.
        // ###########################################################################################
        private static List<KiCadSchematicSheetRef> ExtractSchematicSheetRefs(SExpressionNode schematic)
        {
            var result = new List<KiCadSchematicSheetRef>();

            foreach (var sheetNode in Children(schematic, "sheet"))
            {
                string? sheetName = null;
                string? fileName = null;
                string? uuid = ChildValue(sheetNode, "uuid");

                foreach (var propertyNode in Children(sheetNode, "property"))
                {
                    string? propertyName = Arg(propertyNode, 0);
                    string? propertyValue = Arg(propertyNode, 1);

                    if (string.Equals(propertyName, "Sheetname", StringComparison.OrdinalIgnoreCase))
                    {
                        sheetName = propertyValue;
                        continue;
                    }

                    if (string.Equals(propertyName, KiCadSheetReferences.SheetFileProperty, StringComparison.OrdinalIgnoreCase))
                    {
                        fileName = propertyValue;
                    }
                }

                result.Add(new KiCadSchematicSheetRef
                {
                    SheetName = sheetName,
                    FileName = fileName,
                    Uuid = uuid
                });
            }

            return result;
        }

        private sealed class KiCadSchematicSheetRef
        {
            public string? SheetName { get; init; }

            public string? FileName { get; init; }

            public string? Uuid { get; init; }
        }

        // ###########################################################################################
        // Returns the first text value from a footprint/module text or property entry of the requested
        // kind, supporting both fp_text reference/value and property "Reference"/"Value" forms.
        // ###########################################################################################
        private static string? ExtractFootprintTextValue(SExpressionNode footprintNode, string kind)
        {
            foreach (var textNode in Children(footprintNode, "fp_text"))
            {
                if (string.Equals(Arg(textNode, 0), kind, StringComparison.OrdinalIgnoreCase))
                {
                    return Arg(textNode, 1);
                }
            }

            foreach (var propertyNode in Children(footprintNode, "property"))
            {
                if (string.Equals(Arg(propertyNode, 0), kind, StringComparison.OrdinalIgnoreCase))
                {
                    return Arg(propertyNode, 1);
                }
            }

            return null;
        }

        // ###########################################################################################
        // Extracts one KiCad point node.
        // ###########################################################################################
        private static KiCadPoint2D? ExtractPoint(SExpressionNode? node)
        {
            if (node == null)
            {
                return null;
            }

            double? x = ToDoubleOrNull(Arg(node, 0));
            double? y = ToDoubleOrNull(Arg(node, 1));

            if (!x.HasValue || !y.HasValue)
            {
                return null;
            }

            return new KiCadPoint2D
            {
                X = x.Value,
                Y = y.Value
            };
        }

        // ###########################################################################################
        // Extracts one KiCad at node with optional angle.
        // ###########################################################################################
        private static KiCadPoint2DAngle? ExtractAt(SExpressionNode? node)
        {
            if (node == null)
            {
                return null;
            }

            double? x = ToDoubleOrNull(Arg(node, 0));
            double? y = ToDoubleOrNull(Arg(node, 1));

            if (!x.HasValue || !y.HasValue)
            {
                return null;
            }

            return new KiCadPoint2DAngle
            {
                X = x.Value,
                Y = y.Value,
                Angle = ToDoubleOrNull(Arg(node, 2)) ?? 0.0
            };
        }

        // ###########################################################################################
        // Extracts one KiCad size node.
        // ###########################################################################################
        private static KiCadSize2D? ExtractSize(SExpressionNode? node)
        {
            if (node == null)
            {
                return null;
            }

            return new KiCadSize2D
            {
                X = ToDoubleOrNull(Arg(node, 0)) ?? 0.0,
                Y = ToDoubleOrNull(Arg(node, 1)) ?? 0.0
            };
        }

        // ###########################################################################################
        // Extracts point lists from a KiCad pts node.
        // ###########################################################################################
        private static List<KiCadPoint2D> ExtractPts(SExpressionNode node)
        {
            var result = new List<KiCadPoint2D>();

            foreach (var ptsNode in Children(node, "pts"))
            {
                foreach (var xyNode in Children(ptsNode, "xy"))
                {
                    var point = ExtractPoint(xyNode);
                    if (point != null)
                    {
                        result.Add(point);
                    }
                }
            }

            return result;
        }

        // ###########################################################################################
        // Normalizes hierarchical KiCad net names to their leaf net name.
        // ###########################################################################################
        private static string? NormalizeNetName(string? name)
        {
            if (name == null)
            {
                return null;
            }

            string trimmed = name.Trim();
            if (trimmed.Length == 0)
            {
                return string.Empty;
            }

            int slashIndex = trimmed.LastIndexOf('/');
            return slashIndex >= 0
                ? trimmed[(slashIndex + 1)..]
                : trimmed;
        }

        // ###########################################################################################
        // Rotates one KiCad-local point using KiCad's Y-down angle convention.
        // ###########################################################################################
        private static (double X, double Y) RotatePoint(double x, double y, double angleDegrees)
        {
            double radians = Math.PI * -angleDegrees / 180.0;
            double cos = Math.Cos(radians);
            double sin = Math.Sin(radians);

            return (
                (x * cos) - (y * sin),
                (x * sin) + (y * cos));
        }

        // ###########################################################################################
        // Returns the head atom of one list node.
        // ###########################################################################################
        private static string? Head(SExpressionNode? node)
        {
            return node?.Items.Count > 0 ? node.Items[0].Atom : null;
        }

        // ###########################################################################################
        // Returns child list nodes matching the requested head atom.
        // ###########################################################################################
        private static IEnumerable<SExpressionNode> Children(SExpressionNode node, string head)
        {
            return node.Items.Where(item =>
                item.Items.Count > 0 &&
                string.Equals(Head(item), head, StringComparison.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // Returns the first child list matching the requested head atom.
        // ###########################################################################################
        private static SExpressionNode? Child(SExpressionNode? node, string head)
        {
            return node == null ? null : Children(node, head).FirstOrDefault();
        }

        // ###########################################################################################
        // Returns one positional argument from a list node, ignoring the head atom.
        // ###########################################################################################
        private static string? Arg(SExpressionNode? node, int index)
        {
            if (node == null)
            {
                return null;
            }

            int itemIndex = index + 1;
            return itemIndex >= 0 && itemIndex < node.Items.Count
                ? node.Items[itemIndex].Atom
                : null;
        }

        // ###########################################################################################
        // Returns all positional arguments from a list node, ignoring the head atom.
        // ###########################################################################################
        private static IEnumerable<string> Args(SExpressionNode? node)
        {
            if (node == null)
            {
                yield break;
            }

            for (int i = 1; i < node.Items.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(node.Items[i].Atom))
                {
                    yield return node.Items[i].Atom!;
                }
            }
        }

        // ###########################################################################################
        // Returns the first positional value from a named child node.
        // ###########################################################################################
        private static string? ChildValue(SExpressionNode node, string head)
        {
            return Arg(Child(node, head), 0);
        }

        // ###########################################################################################
        // Parses invariant-culture floating point text.
        // ###########################################################################################
        private static double? ToDoubleOrNull(string? text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                ? value
                : null;
        }

        private sealed class SExpressionNode
        {
            public string? Atom { get; init; }
            public List<SExpressionNode> Items { get; init; } = new();
        }

        private static class SExpressionParser
        {
            // ###########################################################################################
            // Parses KiCad S-expression content into root expression nodes. The tokens are CRT.Data's
            // KiCadSExpressionTokens, which the server's KiCad sheet rule reads too - so the two
            // cannot disagree about what a file says.
            // ###########################################################################################
            public static List<SExpressionNode> Parse(string content)
            {
                var tokens = KiCadSExpressionTokens.Tokenize(content);
                int index = 0;
                var result = new List<SExpressionNode>();

                while (index < tokens.Count)
                {
                    result.Add(ParseNode(tokens, ref index));
                }

                return result;
            }

            // ###########################################################################################
            // Parses one S-expression node from the token stream.
            // ###########################################################################################
            private static SExpressionNode ParseNode(IReadOnlyList<string> tokens, ref int index)
            {
                if (tokens[index] == "(")
                {
                    index++;
                    var items = new List<SExpressionNode>();

                    while (index < tokens.Count && tokens[index] != ")")
                    {
                        items.Add(ParseNode(tokens, ref index));
                    }

                    if (index < tokens.Count && tokens[index] == ")")
                    {
                        index++;
                    }

                    return new SExpressionNode
                    {
                        Items = items
                    };
                }

                string atom = tokens[index];
                index++;

                return new SExpressionNode
                {
                    Atom = atom
                };
            }
        }
    }
}