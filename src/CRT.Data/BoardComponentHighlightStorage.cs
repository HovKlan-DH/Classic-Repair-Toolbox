using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    internal sealed class BoardComponentHighlightJsonRect
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }

    internal sealed class BoardComponentHighlightJsonRoot
    {
        [JsonPropertyName("Component highlights")]
        public Dictionary<string, Dictionary<string, List<BoardComponentHighlightJsonRect>>> ComponentHighlights { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    }

    internal static class BoardComponentHighlightStorage
    {
        private static readonly JsonSerializerOptions JsonSerializerOptions = new()
        {
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };

        // ###########################################################################################
        // Resolves the board highlight JSON path from the board Excel path by replacing the file
        // extension with .json while keeping the same directory and base filename.
        // ###########################################################################################
        public static string GetJsonPath(string excelPath)
        {
            if (string.IsNullOrWhiteSpace(excelPath))
            {
                return string.Empty;
            }

            return Path.ChangeExtension(excelPath, ".json");
        }

        // ###########################################################################################
        // Loads component highlights from the board JSON file and flattens the nested JSON shape
        // into the existing runtime list model used elsewhere in the application.
        // ###########################################################################################
        public static List<ComponentHighlightEntry> LoadComponentHighlights(string excelPath)
        {
            string jsonPath = GetJsonPath(excelPath);
            if (string.IsNullOrWhiteSpace(jsonPath) || !File.Exists(jsonPath))
            {
                CrtLog.Warning($"Board highlight JSON file not found: [{jsonPath}]");
                return new List<ComponentHighlightEntry>();
            }

            try
            {
                string json = File.ReadAllText(jsonPath);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return new List<ComponentHighlightEntry>();
                }

                var root = JsonSerializer.Deserialize<BoardComponentHighlightJsonRoot>(json, JsonSerializerOptions);
                if (root?.ComponentHighlights == null)
                {
                    return new List<ComponentHighlightEntry>();
                }

                var result = new List<ComponentHighlightEntry>();

                foreach (var schematicEntry in root.ComponentHighlights
                             .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
                {
                    string schematicName = schematicEntry.Key?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(schematicName) || schematicEntry.Value == null)
                    {
                        continue;
                    }

                    foreach (var boardLabelEntry in schematicEntry.Value
                                 .OrderBy(item => item.Key, NaturalLabelComparer.Instance))
                    {
                        string boardLabel = boardLabelEntry.Key?.Trim() ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(boardLabel) || boardLabelEntry.Value == null)
                        {
                            continue;
                        }

                        foreach (var rect in boardLabelEntry.Value)
                        {
                            result.Add(new ComponentHighlightEntry
                            {
                                SchematicName = schematicName,
                                BoardLabel = boardLabel,
                                X = rect.X.ToString(CultureInfo.InvariantCulture),
                                Y = rect.Y.ToString(CultureInfo.InvariantCulture),
                                Width = rect.Width.ToString(CultureInfo.InvariantCulture),
                                Height = rect.Height.ToString(CultureInfo.InvariantCulture)
                            });
                        }
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                CrtLog.Warning($"Failed to load board highlight JSON file [{jsonPath}] - [{ex.Message}]");
                return new List<ComponentHighlightEntry>();
            }
        }

        // ###########################################################################################
        // Saves one schematic's highlight rows, keeping every other schematic's highlights, the KiCad
        // calibrations and the schematic renames already in the file.
        //
        // *** THROUGH BoardSidecarWriter, LIKE EVERY OTHER WRITE (code review, 2026-10-13). *** This
        // kept its own file shape (roots sorted, unknown roots carried forward) and wrote in place;
        // the draft writes and the server's publish build the file through BoardSidecarWriter, so one
        // content had three byte forms. It keeps its own row rules: coordinates rounded, labels in
        // natural order, a label's rectangles top to bottom, and the schematic matched ignoring
        // capitals (its rows under "top" replace those under "Top" rather than sitting beside them).
        // ###########################################################################################
        public static void SaveComponentHighlights(
            string excelPath,
            string schematicName,
            IReadOnlyList<LabelEditorSaveRow> rows)
        {
            string jsonPath = GetJsonPath(excelPath);
            if (string.IsNullOrWhiteSpace(jsonPath))
            {
                throw new InvalidOperationException("Could not resolve board highlight JSON path.");
            }

            string normalizedSchematicName = schematicName?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(normalizedSchematicName))
            {
                throw new InvalidOperationException("No schematic name was provided for highlight JSON save.");
            }

            BoardComponentHighlightStorage.EnsureReadableInFull(jsonPath);

            List<ComponentHighlightEntry> highlights = LoadComponentHighlights(excelPath)
                .Where(highlight => !string.Equals(highlight.SchematicName?.Trim(), normalizedSchematicName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            highlights.AddRange(rows
                .Where(item => string.Equals(item.SchematicName?.Trim(), normalizedSchematicName, StringComparison.OrdinalIgnoreCase))
                .Where(item => !string.IsNullOrWhiteSpace(item.BoardLabel))
                .GroupBy(item => item.BoardLabel.Trim(), StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, NaturalLabelComparer.Instance)
                .SelectMany(group => group
                    .OrderBy(item => item.Y)
                    .ThenBy(item => item.X)
                    .Select(item => new ComponentHighlightEntry
                    {
                        SchematicName = normalizedSchematicName,
                        BoardLabel = group.Key,
                        X = RoundToInt(item.X).ToString(CultureInfo.InvariantCulture),
                        Y = RoundToInt(item.Y).ToString(CultureInfo.InvariantCulture),
                        Width = RoundToInt(item.Width).ToString(CultureInfo.InvariantCulture),
                        Height = RoundToInt(item.Height).ToString(CultureInfo.InvariantCulture),
                    })));

            BoardSidecarWriter.Write(
                excelPath,
                highlights,
                BoardComponentHighlightStorage.LoadAllKiCadCalibrations(excelPath),
                LoadSchematicRenames(excelPath));
        }

        // Every KiCad calibration in the sidecar, whatever schematic it names, in name order.
        internal static List<KiCadCalibrationEntry> LoadAllKiCadCalibrations(string excelPath)
        {
            string jsonPath = GetJsonPath(excelPath);

            if (string.IsNullOrWhiteSpace(jsonPath) || !File.Exists(jsonPath))
                return [];

            List<string> names = LoadJsonRootObject(jsonPath)["KiCad calibration points"] is JsonObject root
                ? root.Select(pair => pair.Key).ToList()
                : [];

            return LoadKiCadCalibrations(excelPath, names)
                .OrderBy(entry => entry.SchematicName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ###########################################################################################
        // Rounds highlight coordinates and sizes to integer values using midpoint rounding away
        // from zero so saved JSON matches the old compact Excel persistence behavior.
        // ###########################################################################################
        private static int RoundToInt(double value)
        {
            return (int)Math.Round(value, 0, MidpointRounding.AwayFromZero);
        }

        // ###########################################################################################
        // Loads one saved KiCad box calibration entry for the requested schematic from the board JSON.
        // Returns false when the JSON file, root, or schematic entry does not exist.
        // ###########################################################################################
        public static bool TryLoadKiCadCalibration(
            string excelPath,
            string schematicName,
            out string cadName,
            out double offsetX,
            out double offsetY,
            out double scaleX,
            out double scaleY,
            out bool mirrorX,
            out bool mirrorY)
        {
            cadName = string.Empty;
            offsetX = 0.0;
            offsetY = 0.0;
            scaleX = 1.0;
            scaleY = 1.0;
            mirrorX = false;
            mirrorY = false;

            string jsonPath = GetJsonPath(excelPath);
            if (string.IsNullOrWhiteSpace(jsonPath) || !File.Exists(jsonPath))
            {
                return false;
            }

            string normalizedSchematicName = schematicName?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(normalizedSchematicName))
            {
                return false;
            }

            try
            {
                string json = File.ReadAllText(jsonPath);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return false;
                }

                JsonObject? calibrationRoot = (JsonNode.Parse(json) as JsonObject)?["KiCad calibration points"] as JsonObject;

                if (BoardComponentHighlightStorage.ReadCalibration(calibrationRoot, normalizedSchematicName) is not KiCadCalibrationEntry entry)
                {
                    return false;
                }

                cadName = entry.CadName;
                offsetX = entry.OffsetX;
                offsetY = entry.OffsetY;
                scaleX = entry.ScaleX;
                scaleY = entry.ScaleY;
                mirrorX = entry.MirrorX;
                mirrorY = entry.MirrorY;

                return true;
            }
            catch (Exception ex)
            {
                CrtLog.Warning($"Failed to load KiCad calibration JSON entry for schematic [{schematicName}] from [{jsonPath}] - [{ex.Message}]");
                return false;
            }
        }

        // ###########################################################################################
        // The KiCad calibrations filed for these schematic names, in the order asked, each named as
        // asked (trimmed) - with the sidecar read and parsed ONCE (code review, 2026-10-10).
        //
        // TryLoadKiCadCalibration parses the whole file for one schematic. Collecting a board's
        // calibrations through it parsed the file once per schematic, and every draft write collects
        // them twice (the names before the edit and after it) - 50 parses of a large highlight file
        // per save on a 25-image board. A name with no calibration, a blank one, or an entry that
        // cannot be read is skipped, as TryLoadKiCadCalibration skips it.
        // ###########################################################################################
        public static List<KiCadCalibrationEntry> LoadKiCadCalibrations(string excelPath, IEnumerable<string?> schematicNames)
        {
            var result = new List<KiCadCalibrationEntry>();

            string jsonPath = GetJsonPath(excelPath);
            if (string.IsNullOrWhiteSpace(jsonPath) || !File.Exists(jsonPath))
            {
                return result;
            }

            JsonObject? calibrationRoot;

            try
            {
                string json = File.ReadAllText(jsonPath);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return result;
                }

                calibrationRoot = (JsonNode.Parse(json) as JsonObject)?["KiCad calibration points"] as JsonObject;
            }
            catch (Exception ex)
            {
                CrtLog.Warning($"Failed to load the KiCad calibrations from [{jsonPath}] - [{ex.Message}]");
                return result;
            }

            foreach (string schematicName in schematicNames.Select(name => name?.Trim() ?? string.Empty))
            {
                if (schematicName.Length == 0)
                {
                    continue;
                }

                try
                {
                    if (BoardComponentHighlightStorage.ReadCalibration(calibrationRoot, schematicName) is KiCadCalibrationEntry entry)
                    {
                        result.Add(entry);
                    }
                }
                catch (Exception ex)
                {
                    CrtLog.Warning($"Failed to load KiCad calibration JSON entry for schematic [{schematicName}] from [{jsonPath}] - [{ex.Message}]");
                }
            }

            return result;
        }

        // One schematic's calibration out of the parsed "KiCad calibration points" root, named as
        // asked - or null when it has none. Throws on an entry whose values are not of their type.
        private static KiCadCalibrationEntry? ReadCalibration(JsonObject? calibrationRoot, string schematicName)
        {
            if (BoardComponentHighlightStorage.CalibrationKeyOf(calibrationRoot, schematicName) is not string key ||
                calibrationRoot![key] is not JsonObject schematicObject)
            {
                return null;
            }

            return new KiCadCalibrationEntry
            {
                SchematicName = schematicName,
                CadName = schematicObject["CadName"]?.GetValue<string>()?.Trim() ?? string.Empty,
                OffsetX = schematicObject["OffsetX"]?.GetValue<double>() ?? 0.0,
                OffsetY = schematicObject["OffsetY"]?.GetValue<double>() ?? 0.0,
                ScaleX = schematicObject["ScaleX"]?.GetValue<double>() ?? 1.0,
                ScaleY = schematicObject["ScaleY"]?.GetValue<double>() ?? 1.0,
                MirrorX = schematicObject["MirrorX"]?.GetValue<bool>() ?? false,
                MirrorY = schematicObject["MirrorY"]?.GetValue<bool>() ?? false,
            };
        }

        // ###########################################################################################
        // The board's renamed schematic images, oldest first - the "Schematic renames" root
        // (SchematicRenameLog; owner decision, 2026-10-10). Empty when the file, the root or the
        // list is not there, or does not parse: a board nobody renamed has none, and an unreadable
        // list must not stop a board from loading. An entry without both names is skipped.
        // ###########################################################################################
        public static List<SchematicRenameRecord> LoadSchematicRenames(string excelPath)
        {
            var result = new List<SchematicRenameRecord>();

            string jsonPath = GetJsonPath(excelPath);
            if (string.IsNullOrWhiteSpace(jsonPath) || !File.Exists(jsonPath))
            {
                return result;
            }

            try
            {
                if (JsonNode.Parse(File.ReadAllText(jsonPath)) is not JsonObject root ||
                    root[SchematicRenameLog.Root] is not JsonArray list)
                {
                    return result;
                }

                foreach (JsonObject entry in list.OfType<JsonObject>())
                {
                    string from = BoardComponentHighlightStorage.TextOf(entry["From"]);
                    string to = BoardComponentHighlightStorage.TextOf(entry["To"]);

                    if (from.Length > 0 && to.Length > 0)
                    {
                        result.Add(new SchematicRenameRecord(
                            from,
                            to,
                            BoardComponentHighlightStorage.TextOf(entry["File"]),
                            BoardComponentHighlightStorage.TextOf(entry["Date"])));
                    }
                }
            }
            catch (Exception ex)
            {
                CrtLog.Warning($"Failed to read the schematic renames from [{jsonPath}] - [{ex.Message}]");
                return new List<SchematicRenameRecord>();
            }

            return result;
        }

        // A JSON string value, trimmed - or empty for anything else.
        private static string TextOf(JsonNode? node) =>
            node is JsonValue value && value.TryGetValue(out string? text) ? text.Trim() : string.Empty;

        // ###########################################################################################
        // The key a schematic's calibration is filed under: its exact name, or else the name in other
        // capitals (owner request, 2026-10-10 - "And fix all"). Every other schematic name in CRT
        // and on the server compares without regard to capitals - the highlights, the checks, the
        // Schematics tab - and this exact lookup was the one that did not: a schematic renamed from
        // "top" to "Top" (its Name box, a table cell, Excel) kept its highlights and silently lost
        // its KiCad calibration. Null when there is none.
        // ###########################################################################################
        private static string? CalibrationKeyOf(JsonObject? calibrationRoot, string schematicName)
        {
            if (calibrationRoot is null)
            {
                return null;
            }

            if (calibrationRoot.ContainsKey(schematicName))
            {
                return schematicName;
            }

            return calibrationRoot
                .Select(pair => pair.Key)
                .FirstOrDefault(key => string.Equals(key.Trim(), schematicName, StringComparison.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // Saves one KiCad box calibration for the requested schematic, keeping the highlights, the
        // other calibrations and the schematic renames already in the file.
        //
        // *** THROUGH BoardSidecarWriter (code review, 2026-10-13). *** This wrote its own file shape
        // in place - roots sorted, so a calibrated draft's sidecar had other bytes than the same
        // content written by any other draft write, and the draft's fingerprint (which hashes the
        // sidecar's bytes) brought Submit back for a draft that held what was sent. One writer, one
        // shape, written through a temporary file. The same schematic under other capitals is this
        // one (CalibrationKeyOf) - replaced, not kept beside it.
        // ###########################################################################################
        public static void SaveKiCadCalibration(
            string excelPath,
            string schematicName,
            string cadName,
            double offsetX,
            double offsetY,
            double scaleX,
            double scaleY,
            bool mirrorX,
            bool mirrorY)
        {
            string jsonPath = GetJsonPath(excelPath);
            if (string.IsNullOrWhiteSpace(jsonPath))
            {
                throw new InvalidOperationException("Could not resolve board JSON path for KiCad calibration save.");
            }

            string normalizedSchematicName = schematicName?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(normalizedSchematicName))
            {
                throw new InvalidOperationException("No schematic name was provided for KiCad calibration save.");
            }

            BoardComponentHighlightStorage.EnsureReadableInFull(jsonPath);

            List<KiCadCalibrationEntry> calibrations = LoadAllKiCadCalibrations(excelPath)
                .Where(entry => !string.Equals(entry.SchematicName?.Trim(), normalizedSchematicName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            calibrations.Add(new KiCadCalibrationEntry
            {
                SchematicName = normalizedSchematicName,
                CadName = cadName?.Trim() ?? string.Empty,
                OffsetX = offsetX,
                OffsetY = offsetY,
                ScaleX = scaleX,
                ScaleY = scaleY,
                MirrorX = mirrorX,
                MirrorY = mirrorY,
            });

            BoardSidecarWriter.Write(
                excelPath,
                LoadComponentHighlights(excelPath),
                calibrations.OrderBy(entry => entry.SchematicName, StringComparer.OrdinalIgnoreCase).ToList(),
                LoadSchematicRenames(excelPath));
        }

        // ###########################################################################################
        // *** A SIDECAR THAT CANNOT BE READ IN FULL IS NOT WRITTEN OVER (code review, 2026-10-13). ***
        // Both savers above rebuild the WHOLE file from the typed readers (BoardSidecarWriter), and
        // those readers answer "nothing" for what they cannot parse - so one hand-edited rectangle
        // ("X": 10.5) made a calibration save write the board's highlights as none, and a calibration
        // entry that did not read was dropped. The file is read strictly first: a highlight that is
        // not whole numbers, a calibration entry that is not an object of numbers, a renames root
        // that is not a list - any of these refuses the save with an InvalidOperationException, and
        // the file is left exactly as it is.
        // ###########################################################################################
        internal static void EnsureReadableInFull(string jsonPath)
        {
            if (string.IsNullOrWhiteSpace(jsonPath) || !File.Exists(jsonPath))
            {
                return;
            }

            string json = File.ReadAllText(jsonPath);

            if (string.IsNullOrWhiteSpace(json))
            {
                return;
            }

            try
            {
                _ = JsonSerializer.Deserialize<BoardComponentHighlightJsonRoot>(json, JsonSerializerOptions);

                if (JsonNode.Parse(json) is not JsonObject root)
                {
                    throw new JsonException("The file does not hold a JSON object.");
                }

                if (root["KiCad calibration points"] is JsonNode calibrations)
                {
                    if (calibrations is not JsonObject calibrationRoot)
                    {
                        throw new JsonException("\"KiCad calibration points\" is not an object.");
                    }

                    foreach (KeyValuePair<string, JsonNode?> entry in calibrationRoot)
                    {
                        if (entry.Value is not JsonObject)
                        {
                            throw new JsonException($"The KiCad calibration of [{entry.Key}] is not an object.");
                        }

                        _ = BoardComponentHighlightStorage.ReadCalibration(calibrationRoot, entry.Key);
                    }
                }

                if (root[SchematicRenameLog.Root] is not (null or JsonArray))
                {
                    throw new JsonException($"\"{SchematicRenameLog.Root}\" is not a list.");
                }
            }
            catch (Exception ex) when (ex is JsonException or InvalidOperationException or FormatException or NotSupportedException)
            {
                throw new InvalidOperationException(
                    $"The board's highlight file [{Path.GetFileName(jsonPath)}] could not be read in full, so nothing was saved over it " +
                    $"- [{ex.Message}]. Correct the file, or remove the part that does not read, and save again.",
                    ex);
            }
        }

        // ###########################################################################################
        // Loads the complete board JSON as a mutable JsonObject so independent feature roots can be
        // added and updated without overwriting each other.
        // ###########################################################################################
        private static JsonObject LoadJsonRootObject(string jsonPath)
        {
            try
            {
                if (!File.Exists(jsonPath))
                {
                    return new JsonObject();
                }

                string json = File.ReadAllText(jsonPath);
                if (string.IsNullOrWhiteSpace(json))
                {
                    return new JsonObject();
                }

                return JsonNode.Parse(json) as JsonObject ?? new JsonObject();
            }
            catch (Exception ex)
            {
                CrtLog.Warning($"Failed to parse board JSON file [{jsonPath}] - [{ex.Message}]");
                return new JsonObject();
            }
        }






    }
}