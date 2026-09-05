using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Root cached KiCad project bundle used by the Schematics tab.
    // ###########################################################################################
    public sealed class KiCadProjectBundle
    {
        public KiCadProjectRoot Root { get; init; } = new();
        public Dictionary<int, Dictionary<string, List<KiCadResolvedPath>>> SchematicNetPathIndexBySchematicIndex { get; init; }
            = new();
    }

    // ###########################################################################################
    // Root JSON DTO for the combined KiCad export.
    // ###########################################################################################
    public sealed class KiCadProjectRoot
    {
        [JsonPropertyName("ok")]
        public bool Ok { get; init; }

        [JsonPropertyName("project")]
        public KiCadProjectInfo Project { get; init; } = new();

        [JsonPropertyName("pcb")]
        public List<KiCadPcb> Pcb { get; init; } = new();

        [JsonPropertyName("schematics")]
        public List<KiCadSchematic> Schematics { get; init; } = new();
    }

    public sealed class KiCadProjectInfo
    {
        [JsonPropertyName("views")]
        public List<KiCadProjectView> Views { get; init; } = new();
    }

    public sealed class KiCadProjectView
    {
        [JsonPropertyName("id")]
        public string Id { get; init; } = string.Empty;

        [JsonPropertyName("type")]
        public string Type { get; init; } = string.Empty;

        [JsonPropertyName("display_name")]
        public string DisplayName { get; init; } = string.Empty;

        [JsonPropertyName("source_kind")]
        public string SourceKind { get; init; } = string.Empty;

        [JsonPropertyName("source_index")]
        public int SourceIndex { get; init; }

        [JsonPropertyName("source_file")]
        public string SourceFile { get; init; } = string.Empty;

        [JsonPropertyName("board_side")]
        public string BoardSide { get; init; } = string.Empty;

        [JsonPropertyName("layer_hint")]
        public string LayerHint { get; init; } = string.Empty;

        [JsonPropertyName("sheet_role")]
        public string SheetRole { get; init; } = string.Empty;

        [JsonPropertyName("sheet_uuid")]
        public string SheetUuid { get; init; } = string.Empty;
    }

    public sealed class KiCadPcb
    {
        [JsonPropertyName("filename")]
        public string Filename { get; init; } = string.Empty;

        [JsonPropertyName("nets")]
        public KiCadPcbNets Nets { get; init; } = new();

        [JsonPropertyName("footprints")]
        public List<KiCadPcbFootprint> Footprints { get; init; } = new();

        [JsonPropertyName("routing")]
        public KiCadPcbRouting Routing { get; init; } = new();

        [JsonPropertyName("highlight_index")]
        [JsonConverter(typeof(KiCadPcbHighlightIndexJsonConverter))]
        public Dictionary<string, KiCadPcbHighlightBucket> HighlightIndex { get; init; }
            = new(StringComparer.OrdinalIgnoreCase);
    }

    public sealed class KiCadPcbNets
    {
        [JsonPropertyName("list")]
        public List<KiCadNetRef> List { get; init; } = new();
    }

    public sealed class KiCadNetRef
    {
        [JsonPropertyName("id")]
        [JsonConverter(typeof(KiCadStringOrNumberJsonConverter))]
        public string? Id { get; init; }

        [JsonPropertyName("name")]
        public string? Name { get; init; }

        [JsonPropertyName("normalized_name")]
        public string? NormalizedName { get; init; }
    }

    public sealed class KiCadPcbFootprint
    {
        [JsonPropertyName("reference")]
        public string? Reference { get; init; }

        [JsonPropertyName("layer")]
        public string? Layer { get; init; }

        [JsonPropertyName("pads")]
        public List<KiCadPcbPad> Pads { get; init; } = new();
    }

    public sealed class KiCadPcbPad
    {
        [JsonPropertyName("number")]
        [JsonConverter(typeof(KiCadStringOrNumberJsonConverter))]
        public string? Number { get; init; }

        [JsonPropertyName("shape")]
        public string? Shape { get; init; }

        [JsonPropertyName("absolute_center")]
        public KiCadPoint2D? AbsoluteCenter { get; init; }

        [JsonPropertyName("size")]
        public KiCadSize2D? Size { get; init; }

        // The pad's rotation in degrees, as KiCad stores it: already absolute, meaning it includes
        // the parent footprint's rotation. Without this a rotated rect/oval pad is drawn with its
        // width and height swapped - a vertical pad appears horizontal.
        [JsonPropertyName("rotation")]
        public double RotationDegrees { get; init; }

        [JsonPropertyName("layers")]
        public List<string> Layers { get; init; } = new();

        [JsonPropertyName("net")]
        public KiCadNetRef? Net { get; init; }
    }

    public sealed class KiCadPcbRouting
    {
        [JsonPropertyName("segments")]
        public List<KiCadPcbSegment> Segments { get; init; } = new();

        [JsonPropertyName("vias")]
        public List<KiCadPcbVia> Vias { get; init; } = new();

        [JsonPropertyName("arcs")]
        public List<KiCadPcbArc> Arcs { get; init; } = new();

        [JsonPropertyName("zones")]
        public List<KiCadPcbZone> Zones { get; init; } = new();
    }

    public sealed class KiCadPcbZone
    {
        [JsonPropertyName("layers")]
        public List<string> Layers { get; init; } = new();

        [JsonPropertyName("net")]
        public KiCadNetRef? Net { get; init; }

        [JsonPropertyName("outline_polygons")]
        public List<KiCadPcbZonePolygon> OutlinePolygons { get; init; } = new();

        [JsonPropertyName("filled_polygons")]
        public List<KiCadPcbZonePolygon> FilledPolygons { get; init; } = new();
    }

    public sealed class KiCadPcbZonePolygon
    {
        [JsonPropertyName("points")]
        public List<KiCadPoint2D> Points { get; init; } = new();
    }

    public sealed class KiCadPcbSegment
    {
        [JsonPropertyName("start")]
        public KiCadPoint2D? Start { get; init; }

        [JsonPropertyName("end")]
        public KiCadPoint2D? End { get; init; }

        [JsonPropertyName("width")]
        public double? Width { get; init; }

        [JsonPropertyName("layer")]
        public string? Layer { get; init; }

        [JsonPropertyName("net")]
        public KiCadNetRef? Net { get; init; }
    }

    public sealed class KiCadPcbVia
    {
        [JsonPropertyName("at")]
        public KiCadPoint2DAngle? At { get; init; }

        [JsonPropertyName("size")]
        public double? Size { get; init; }

        [JsonPropertyName("layers")]
        public List<string> Layers { get; init; } = new();

        [JsonPropertyName("net")]
        public KiCadNetRef? Net { get; init; }
    }

    public sealed class KiCadPcbArc
    {
        [JsonPropertyName("start")]
        public KiCadPoint2D? Start { get; init; }

        [JsonPropertyName("mid")]
        public KiCadPoint2D? Mid { get; init; }

        [JsonPropertyName("end")]
        public KiCadPoint2D? End { get; init; }

        [JsonPropertyName("width")]
        public double? Width { get; init; }

        [JsonPropertyName("layer")]
        public string? Layer { get; init; }

        [JsonPropertyName("net")]
        public KiCadNetRef? Net { get; init; }
    }

    [JsonConverter(typeof(KiCadPcbHighlightBucketJsonConverter))]
    public sealed class KiCadPcbHighlightBucket
    {
        [JsonPropertyName("segments")]
        public List<int> Segments { get; init; } = new();

        [JsonPropertyName("vias")]
        public List<int> Vias { get; init; } = new();

        [JsonPropertyName("arcs")]
        public List<int> Arcs { get; init; } = new();

        [JsonPropertyName("pads")]
        public List<KiCadPcbHighlightPadRef> Pads { get; init; } = new();

        [JsonPropertyName("zones")]
        public List<int> Zones { get; init; } = new();
    }

    public sealed class KiCadPcbHighlightPadRef
    {
        [JsonPropertyName("footprint_index")]
        public int FootprintIndex { get; init; }

        [JsonPropertyName("pad_index")]
        public int PadIndex { get; init; }

        [JsonPropertyName("reference")]
        public string? Reference { get; init; }

        [JsonPropertyName("pad_number")]
        [JsonConverter(typeof(KiCadStringOrNumberJsonConverter))]
        public string? PadNumber { get; init; }
    }

    public sealed class KiCadSchematic
    {
        [JsonPropertyName("filename")]
        public string Filename { get; init; } = string.Empty;

        [JsonPropertyName("display_name")]
        public string DisplayName { get; init; } = string.Empty;

        [JsonPropertyName("wires")]
        public List<KiCadSchematicPathItem> Wires { get; init; } = new();

        [JsonPropertyName("polylines")]
        public List<KiCadSchematicPathItem> Polylines { get; init; } = new();

        [JsonPropertyName("labels")]
        public KiCadSchematicLabels Labels { get; init; } = new();

        [JsonPropertyName("symbols")]
        public List<KiCadSchematicSymbol> Symbols { get; init; } = new();
    }

    public sealed class KiCadSchematicSymbol
    {
        [JsonPropertyName("reference")]
        public string? Reference { get; init; }

        [JsonPropertyName("value")]
        public string? Value { get; init; }

        [JsonPropertyName("at")]
        public KiCadPoint2DAngle? At { get; init; }

        [JsonPropertyName("properties_detailed")]
        public List<KiCadSchematicSymbolPropertyDetailed> PropertiesDetailed { get; init; } = new();
    }

    public sealed class KiCadSchematicSymbolPropertyDetailed
    {
        [JsonPropertyName("name")]
        public string? Name { get; init; }

        [JsonPropertyName("value")]
        public string? Value { get; init; }

        [JsonPropertyName("at")]
        public KiCadPoint2DAngle? At { get; init; }

        [JsonPropertyName("effects")]
        public KiCadSchematicTextEffects? Effects { get; init; }
    }

    public sealed class KiCadSchematicTextEffects
    {
        [JsonPropertyName("hide")]
        public bool Hide { get; init; }

        [JsonPropertyName("justify")]
        public List<string> Justify { get; init; } = new();
    }

    public sealed class KiCadSchematicPathItem
    {
        [JsonPropertyName("type")]
        public string? Type { get; init; }

        [JsonPropertyName("points")]
        public List<KiCadPoint2D> Points { get; init; } = new();
    }

    public sealed class KiCadSchematicLabels
    {
        [JsonPropertyName("local")]
        public List<KiCadSchematicLabel> Local { get; init; } = new();

        [JsonPropertyName("global")]
        public List<KiCadSchematicLabel> Global { get; init; } = new();

        [JsonPropertyName("hierarchical")]
        public List<KiCadSchematicLabel> Hierarchical { get; init; } = new();
    }

    public sealed class KiCadSchematicLabel
    {
        [JsonPropertyName("type")]
        public string? Type { get; init; }

        [JsonPropertyName("text")]
        public string? Text { get; init; }

        [JsonPropertyName("normalized_text")]
        public string? NormalizedText { get; init; }

        [JsonPropertyName("at")]
        public KiCadPoint2DAngle? At { get; init; }
    }

    public sealed class KiCadResolvedPath
    {
        public List<KiCadPoint2D> Points { get; init; } = new();
    }

    public sealed class KiCadPoint2D
    {
        [JsonPropertyName("x")]
        public double X { get; init; }

        [JsonPropertyName("y")]
        public double Y { get; init; }
    }

    public sealed class KiCadPoint2DAngle
    {
        [JsonPropertyName("x")]
        public double X { get; init; }

        [JsonPropertyName("y")]
        public double Y { get; init; }

        [JsonPropertyName("angle")]
        public double? Angle { get; init; }
    }

    [JsonConverter(typeof(KiCadSize2DJsonConverter))]
    public sealed class KiCadSize2D
    {
        [JsonPropertyName("x")]
        public double X { get; init; }

        [JsonPropertyName("y")]
        public double Y { get; init; }
    }

    // ###########################################################################################
    // Deserializes one scalar JSON value that may be either a string or a number into text.
    // This keeps DTOs tolerant of KiCad exporter revisions that change identifier types.
    // ###########################################################################################
    internal sealed class KiCadStringOrNumberJsonConverter : JsonConverter<string?>
    {
        // ###########################################################################################
        // Reads one JSON scalar as text, accepting null, string, integer, or floating-point tokens.
        // ###########################################################################################
        public override string? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            if (reader.TokenType == JsonTokenType.Null)
            {
                return null;
            }

            if (reader.TokenType == JsonTokenType.String)
            {
                return reader.GetString();
            }

            if (reader.TokenType == JsonTokenType.Number)
            {
                using var document = JsonDocument.ParseValue(ref reader);
                return document.RootElement.GetRawText();
            }

            if (reader.TokenType == JsonTokenType.True)
            {
                return bool.TrueString;
            }

            if (reader.TokenType == JsonTokenType.False)
            {
                return bool.FalseString;
            }

            throw new JsonException($"Unsupported scalar JSON token [{reader.TokenType}] for string-backed value.");
        }

        // ###########################################################################################
        // Writes the scalar value back as JSON text.
        // ###########################################################################################
        public override void Write(Utf8JsonWriter writer, string? value, JsonSerializerOptions options)
        {
            if (value == null)
            {
                writer.WriteNullValue();
                return;
            }

            writer.WriteStringValue(value);
        }
    }

    // ###########################################################################################
    // Deserializes KiCad size values from either an object form { x, y } or an array form [x, y].
    // This keeps the DTO tolerant of small export-format differences between converter versions.
    // ###########################################################################################
    internal sealed class KiCadSize2DJsonConverter : JsonConverter<KiCadSize2D>
    {
        // ###########################################################################################
        // Reads one KiCad size value from JSON, accepting either object or array syntax.
        // ###########################################################################################
        public override KiCadSize2D Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            if (reader.TokenType == JsonTokenType.Null)
            {
                return new KiCadSize2D();
            }

            if (reader.TokenType == JsonTokenType.StartArray)
            {
                double x = 0.0;
                double y = 0.0;

                reader.Read();
                if (reader.TokenType == JsonTokenType.Number)
                {
                    x = reader.GetDouble();
                }

                reader.Read();
                if (reader.TokenType == JsonTokenType.Number)
                {
                    y = reader.GetDouble();
                }

                while (reader.TokenType != JsonTokenType.EndArray && reader.Read())
                {
                }

                return new KiCadSize2D
                {
                    X = x,
                    Y = y
                };
            }

            if (reader.TokenType == JsonTokenType.StartObject)
            {
                double x = 0.0;
                double y = 0.0;

                while (reader.Read())
                {
                    if (reader.TokenType == JsonTokenType.EndObject)
                    {
                        break;
                    }

                    if (reader.TokenType != JsonTokenType.PropertyName)
                    {
                        continue;
                    }

                    string propertyName = reader.GetString() ?? string.Empty;

                    if (!reader.Read())
                    {
                        break;
                    }

                    if (string.Equals(propertyName, "x", StringComparison.OrdinalIgnoreCase) &&
                        reader.TokenType == JsonTokenType.Number)
                    {
                        x = reader.GetDouble();
                    }
                    else if (string.Equals(propertyName, "y", StringComparison.OrdinalIgnoreCase) &&
                             reader.TokenType == JsonTokenType.Number)
                    {
                        y = reader.GetDouble();
                    }
                    else
                    {
                        using var ignored = JsonDocument.ParseValue(ref reader);
                    }
                }

                return new KiCadSize2D
                {
                    X = x,
                    Y = y
                };
            }

            throw new JsonException($"Unsupported KiCad size JSON token [{reader.TokenType}].");
        }

        // ###########################################################################################
        // Writes KiCad size values as a stable object with x and y properties.
        // ###########################################################################################
        public override void Write(Utf8JsonWriter writer, KiCadSize2D value, JsonSerializerOptions options)
        {
            writer.WriteStartObject();
            writer.WriteNumber("x", value.X);
            writer.WriteNumber("y", value.Y);
            writer.WriteEndObject();
        }
    }

    // ###########################################################################################
    // Deserializes PCB highlight_index from either an object map or an array-based form.
    // This keeps the loader tolerant of converter revisions without changing the rendering code.
    // ###########################################################################################
    internal sealed class KiCadPcbHighlightIndexJsonConverter
        : JsonConverter<Dictionary<string, KiCadPcbHighlightBucket>>
    {
        // ###########################################################################################
        // Reads one PCB highlight index from JSON, accepting object, array, or null syntax.
        // ###########################################################################################
        public override Dictionary<string, KiCadPcbHighlightBucket> Read(
            ref Utf8JsonReader reader,
            Type typeToConvert,
            JsonSerializerOptions options)
        {
            var result = new Dictionary<string, KiCadPcbHighlightBucket>(StringComparer.OrdinalIgnoreCase);

            if (reader.TokenType == JsonTokenType.Null)
            {
                return result;
            }

            using var document = JsonDocument.ParseValue(ref reader);
            JsonElement rootElement = document.RootElement;

            if (rootElement.ValueKind == JsonValueKind.Object)
            {
                foreach (JsonProperty property in rootElement.EnumerateObject())
                {
                    KiCadPcbHighlightBucket bucket =
                        KiCadPcbHighlightIndexJsonConverter.DeserializeBucket(property.Value, options);

                    result[property.Name] = bucket;
                }

                return result;
            }

            if (rootElement.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement item in rootElement.EnumerateArray())
                {
                    if (item.ValueKind == JsonValueKind.Null)
                    {
                        continue;
                    }

                    if (item.ValueKind == JsonValueKind.Object)
                    {
                        string key =
                            KiCadPcbHighlightIndexJsonConverter.TryGetBucketKey(item) ??
                            string.Empty;

                        if (!string.IsNullOrWhiteSpace(key))
                        {
                            result[key] = KiCadPcbHighlightIndexJsonConverter.DeserializeBucket(item, options);
                            continue;
                        }

                        JsonProperty? singleProperty = item.EnumerateObject().FirstOrDefault();
                        if (singleProperty.HasValue)
                        {
                            JsonProperty property = singleProperty.Value;
                            result[property.Name] =
                                KiCadPcbHighlightIndexJsonConverter.DeserializeBucket(property.Value, options);
                        }
                    }
                }

                return result;
            }

            return result;
        }

        // ###########################################################################################
        // Writes the PCB highlight index back as a normal JSON object map.
        // ###########################################################################################
        public override void Write(
            Utf8JsonWriter writer,
            Dictionary<string, KiCadPcbHighlightBucket> value,
            JsonSerializerOptions options)
        {
            writer.WriteStartObject();

            foreach (KeyValuePair<string, KiCadPcbHighlightBucket> pair in value)
            {
                writer.WritePropertyName(pair.Key);
                JsonSerializer.Serialize(writer, pair.Value, options);
            }

            writer.WriteEndObject();
        }

        // ###########################################################################################
        // Deserializes one bucket value while tolerating null or malformed bucket payloads.
        // ###########################################################################################
        private static KiCadPcbHighlightBucket DeserializeBucket(JsonElement element, JsonSerializerOptions options)
        {
            try
            {
                if (element.ValueKind == JsonValueKind.Null)
                {
                    return new KiCadPcbHighlightBucket();
                }

                KiCadPcbHighlightBucket? bucket = element.Deserialize<KiCadPcbHighlightBucket>(options);
                return bucket ?? new KiCadPcbHighlightBucket();
            }
            catch
            {
                return new KiCadPcbHighlightBucket();
            }
        }

        // ###########################################################################################
        // Tries to derive the string key for one array-style highlight bucket entry.
        // ###########################################################################################
        private static string? TryGetBucketKey(JsonElement item)
        {
            if (item.TryGetProperty("net_id", out JsonElement netIdElement))
            {
                return KiCadPcbHighlightIndexJsonConverter.GetElementScalarText(netIdElement);
            }

            if (item.TryGetProperty("id", out JsonElement idElement))
            {
                return KiCadPcbHighlightIndexJsonConverter.GetElementScalarText(idElement);
            }

            if (item.TryGetProperty("key", out JsonElement keyElement))
            {
                return KiCadPcbHighlightIndexJsonConverter.GetElementScalarText(keyElement);
            }

            if (item.TryGetProperty("net", out JsonElement netElement) &&
                netElement.ValueKind == JsonValueKind.Object &&
                netElement.TryGetProperty("id", out JsonElement nestedIdElement))
            {
                return KiCadPcbHighlightIndexJsonConverter.GetElementScalarText(nestedIdElement);
            }

            return null;
        }

        // ###########################################################################################
        // Converts a scalar JSON value into text suitable for use as a dictionary key.
        // ###########################################################################################
        private static string? GetElementScalarText(JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.String => element.GetString(),
                JsonValueKind.Number => element.GetRawText(),
                _ => null
            };
        }
    }

    // ###########################################################################################
    // Deserializes one PCB highlight bucket while tolerating null list fields and missing members.
    // ###########################################################################################
    internal sealed class KiCadPcbHighlightBucketJsonConverter
        : JsonConverter<KiCadPcbHighlightBucket>
    {
        // ###########################################################################################
        // Reads one PCB highlight bucket from JSON and normalizes missing arrays to empty lists.
        // ###########################################################################################
        public override KiCadPcbHighlightBucket Read(
            ref Utf8JsonReader reader,
            Type typeToConvert,
            JsonSerializerOptions options)
        {
            if (reader.TokenType == JsonTokenType.Null)
            {
                return new KiCadPcbHighlightBucket();
            }

            using var document = JsonDocument.ParseValue(ref reader);
            JsonElement rootElement = document.RootElement;

            if (rootElement.ValueKind != JsonValueKind.Object)
            {
                return new KiCadPcbHighlightBucket();
            }

            List<int> segments = new();
            List<int> vias = new();
            List<int> arcs = new();
            List<int> zones = new();
            List<KiCadPcbHighlightPadRef> pads = new();

            if (rootElement.TryGetProperty("segments", out JsonElement segmentsElement))
            {
                segments = KiCadPcbHighlightBucketJsonConverter.ReadIntList(segmentsElement);
            }

            if (rootElement.TryGetProperty("vias", out JsonElement viasElement))
            {
                vias = KiCadPcbHighlightBucketJsonConverter.ReadIntList(viasElement);
            }

            if (rootElement.TryGetProperty("arcs", out JsonElement arcsElement))
            {
                arcs = KiCadPcbHighlightBucketJsonConverter.ReadIntList(arcsElement);
            }

            if (rootElement.TryGetProperty("zones", out JsonElement zonesElement))
            {
                zones = KiCadPcbHighlightBucketJsonConverter.ReadIntList(zonesElement);
            }

            if (rootElement.TryGetProperty("pads", out JsonElement padsElement) &&
                padsElement.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement padElement in padsElement.EnumerateArray())
                {
                    try
                    {
                        KiCadPcbHighlightPadRef? pad = padElement.Deserialize<KiCadPcbHighlightPadRef>(options);
                        if (pad != null)
                        {
                            pads.Add(pad);
                        }
                    }
                    catch
                    {
                        // One malformed pad entry should not discard the rest of the bucket.
                    }
                }
            }

            return new KiCadPcbHighlightBucket
            {
                Segments = segments,
                Vias = vias,
                Arcs = arcs,
                Pads = pads,
                Zones = zones
            };
        }

        // ###########################################################################################
        // Writes one PCB highlight bucket as a stable JSON object.
        // ###########################################################################################
        public override void Write(
            Utf8JsonWriter writer,
            KiCadPcbHighlightBucket value,
            JsonSerializerOptions options)
        {
            writer.WriteStartObject();

            writer.WritePropertyName("segments");
            JsonSerializer.Serialize(writer, value.Segments, options);

            writer.WritePropertyName("vias");
            JsonSerializer.Serialize(writer, value.Vias, options);

            writer.WritePropertyName("arcs");
            JsonSerializer.Serialize(writer, value.Arcs, options);

            writer.WritePropertyName("pads");
            JsonSerializer.Serialize(writer, value.Pads, options);

            writer.WritePropertyName("zones");
            JsonSerializer.Serialize(writer, value.Zones, options);

            writer.WriteEndObject();
        }

        // ###########################################################################################
        // Reads one integer list from a JSON array while ignoring invalid entries.
        // ###########################################################################################
        private static List<int> ReadIntList(JsonElement element)
        {
            List<int> result = new();

            if (element.ValueKind != JsonValueKind.Array)
            {
                return result;
            }

            foreach (JsonElement item in element.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.Number && item.TryGetInt32(out int value))
                {
                    result.Add(value);
                }
            }

            return result;
        }
    }
}