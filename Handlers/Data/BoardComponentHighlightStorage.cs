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

        private static readonly IComparer<string> BoardLabelNaturalComparer =
            Comparer<string>.Create(CompareBoardLabelsNaturally);

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
                Logger.Warning($"Board highlight JSON file not found: [{jsonPath}]");
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
                                 .OrderBy(item => item.Key, BoardLabelNaturalComparer))
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
                Logger.Warning($"Failed to load board highlight JSON file [{jsonPath}] - [{ex.Message}]");
                return new List<ComponentHighlightEntry>();
            }
        }

        // ###########################################################################################
        // Saves the current schematic highlight rows back into the board JSON file while preserving
        // highlight data for all other schematics already present in the same JSON file.
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

            // Load full JSON so we do not destroy other roots (e.g. "KiCad calibration points").
            JsonObject rootObject = LoadJsonRootObject(jsonPath);

            JsonObject highlightsRoot;
            if (rootObject["Component highlights"] is JsonObject existingHighlightsRoot)
            {
                highlightsRoot = existingHighlightsRoot;
            }
            else
            {
                highlightsRoot = new JsonObject();
                rootObject["Component highlights"] = highlightsRoot;
            }

            var schematicRows = rows
                .Where(item => string.Equals(item.SchematicName, normalizedSchematicName, StringComparison.OrdinalIgnoreCase))
                .Where(item => !string.IsNullOrWhiteSpace(item.BoardLabel))
                .GroupBy(item => item.BoardLabel.Trim(), StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, BoardLabelNaturalComparer)
                .ToList();

            var schematicObject = new JsonObject();

            foreach (var group in schematicRows)
            {
                var rectArray = new JsonArray();

                foreach (var item in group
                             .OrderBy(r => r.Y)
                             .ThenBy(r => r.X))
                {
                    rectArray.Add(new JsonObject
                    {
                        ["X"] = RoundToInt(item.X),
                        ["Y"] = RoundToInt(item.Y),
                        ["Width"] = RoundToInt(item.Width),
                        ["Height"] = RoundToInt(item.Height)
                    });
                }

                schematicObject[group.Key] = rectArray;
            }

            highlightsRoot[normalizedSchematicName] = schematicObject;

            // Keep stable ordering at the root level (matches KiCad calibration save behavior).
            JsonObject orderedRootObject = new();

            foreach (var property in rootObject.OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
            {
                orderedRootObject[property.Key] = property.Value?.DeepClone();
            }

            string directory = Path.GetDirectoryName(jsonPath) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string json = orderedRootObject.ToJsonString(new JsonSerializerOptions
            {
                WriteIndented = true
            });

            File.WriteAllText(jsonPath, json);
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
        // Compares board labels using natural ordering so labels like C2, C10, and C105 sort as
        // humans expect instead of using pure lexicographic string ordering.
        // ###########################################################################################
        private static int CompareBoardLabelsNaturally(string? left, string? right)
        {
            string leftValue = left?.Trim() ?? string.Empty;
            string rightValue = right?.Trim() ?? string.Empty;

            int leftIndex = 0;
            int rightIndex = 0;

            while (leftIndex < leftValue.Length && rightIndex < rightValue.Length)
            {
                bool leftIsDigit = char.IsDigit(leftValue[leftIndex]);
                bool rightIsDigit = char.IsDigit(rightValue[rightIndex]);

                if (leftIsDigit && rightIsDigit)
                {
                    int leftDigitStart = leftIndex;
                    int rightDigitStart = rightIndex;

                    while (leftIndex < leftValue.Length && char.IsDigit(leftValue[leftIndex]))
                    {
                        leftIndex++;
                    }

                    while (rightIndex < rightValue.Length && char.IsDigit(rightValue[rightIndex]))
                    {
                        rightIndex++;
                    }

                    string leftDigits = leftValue[leftDigitStart..leftIndex];
                    string rightDigits = rightValue[rightDigitStart..rightIndex];

                    string leftTrimmedDigits = leftDigits.TrimStart('0');
                    string rightTrimmedDigits = rightDigits.TrimStart('0');

                    if (leftTrimmedDigits.Length == 0)
                    {
                        leftTrimmedDigits = "0";
                    }

                    if (rightTrimmedDigits.Length == 0)
                    {
                        rightTrimmedDigits = "0";
                    }

                    if (leftTrimmedDigits.Length != rightTrimmedDigits.Length)
                    {
                        return leftTrimmedDigits.Length.CompareTo(rightTrimmedDigits.Length);
                    }

                    int digitCompare = string.Compare(leftTrimmedDigits, rightTrimmedDigits, StringComparison.Ordinal);
                    if (digitCompare != 0)
                    {
                        return digitCompare;
                    }

                    int originalDigitLengthCompare = leftDigits.Length.CompareTo(rightDigits.Length);
                    if (originalDigitLengthCompare != 0)
                    {
                        return originalDigitLengthCompare;
                    }

                    continue;
                }

                if (leftIsDigit != rightIsDigit)
                {
                    return leftIsDigit ? -1 : 1;
                }

                int leftTextStart = leftIndex;
                int rightTextStart = rightIndex;

                while (leftIndex < leftValue.Length && !char.IsDigit(leftValue[leftIndex]))
                {
                    leftIndex++;
                }

                while (rightIndex < rightValue.Length && !char.IsDigit(rightValue[rightIndex]))
                {
                    rightIndex++;
                }

                string leftText = leftValue[leftTextStart..leftIndex];
                string rightText = rightValue[rightTextStart..rightIndex];

                int textCompare = string.Compare(leftText, rightText, StringComparison.OrdinalIgnoreCase);
                if (textCompare != 0)
                {
                    return textCompare;
                }
            }

            return leftValue.Length.CompareTo(rightValue.Length);
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

                JsonNode? rootNode = JsonNode.Parse(json);
                JsonObject? rootObject = rootNode as JsonObject;
                JsonObject? calibrationRoot = rootObject?["KiCad calibration points"] as JsonObject;
                JsonObject? schematicObject = calibrationRoot?[normalizedSchematicName] as JsonObject;

                if (schematicObject == null)
                {
                    return false;
                }

                cadName = schematicObject["CadName"]?.GetValue<string>()?.Trim() ?? string.Empty;
                offsetX = schematicObject["OffsetX"]?.GetValue<double>() ?? 0.0;
                offsetY = schematicObject["OffsetY"]?.GetValue<double>() ?? 0.0;
                scaleX = schematicObject["ScaleX"]?.GetValue<double>() ?? 1.0;
                scaleY = schematicObject["ScaleY"]?.GetValue<double>() ?? 1.0;
                mirrorX = schematicObject["MirrorX"]?.GetValue<bool>() ?? false;
                mirrorY = schematicObject["MirrorY"]?.GetValue<bool>() ?? false;

                return true;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load KiCad calibration JSON entry for schematic [{schematicName}] from [{jsonPath}] - [{ex.Message}]");
                return false;
            }
        }

        // ###########################################################################################
        // Saves one KiCad box calibration entry for the requested schematic into the board JSON while
        // preserving all other JSON roots already stored in the same file.
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

            JsonObject rootObject = BoardComponentHighlightStorage.LoadJsonRootObject(jsonPath);

            JsonObject calibrationRoot;
            if (rootObject["KiCad calibration points"] is JsonObject existingCalibrationRoot)
            {
                calibrationRoot = existingCalibrationRoot;
            }
            else
            {
                calibrationRoot = new JsonObject();
                rootObject["KiCad calibration points"] = calibrationRoot;
            }

            calibrationRoot[normalizedSchematicName] = new JsonObject
            {
                ["CadName"] = cadName?.Trim() ?? string.Empty,
                ["OffsetX"] = offsetX,
                ["OffsetY"] = offsetY,
                ["ScaleX"] = scaleX,
                ["ScaleY"] = scaleY,
                ["MirrorX"] = mirrorX,
                ["MirrorY"] = mirrorY
            };

            JsonObject orderedRootObject = new();

            foreach (var property in rootObject.OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
            {
                orderedRootObject[property.Key] = property.Value?.DeepClone();
            }

            string directory = Path.GetDirectoryName(jsonPath) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string json = orderedRootObject.ToJsonString(new JsonSerializerOptions
            {
                WriteIndented = true
            });

            File.WriteAllText(jsonPath, json);
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
                Logger.Warning($"Failed to parse board JSON file [{jsonPath}] - [{ex.Message}]");
                return new JsonObject();
            }
        }






    }
}