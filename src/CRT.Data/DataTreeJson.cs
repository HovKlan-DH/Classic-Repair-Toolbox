using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // READS AND WRITES THE CRT 4.0 DATA TREE'S JSON FILES - board.json and the main file.
    //
    // *** THE SAME CONTENT IS ALWAYS THE SAME BYTES. *** The manifest, the sync and git all compare
    // files by their bytes, so a board written twice with nothing changed must not "change". The
    // layout is fixed here, not left to a serializer setting:
    //
    //   - the top-level fields in the order the document declares them, then any field this CRT
    //     does not know, sorted by name;
    //   - every list one ROW PER LINE - so a changed row is one changed line in a diff, and a
    //     board of a thousand highlights stays readable;
    //   - text as it is (no \u escapes for "ø" or "µ"), "\n" line endings, a final newline,
    //     UTF-8 without a byte order mark (FileReplacer.WriteAllText's encoding).
    //
    // *** READING IS STRICT ABOUT SHAPE AND LENIENT ABOUT ADDITIONS. *** A field this CRT does not
    // know is kept (the documents' Unknown) and written back. A file of a HIGHER format version, a
    // wrong type, a null where a row should be, or no object at all is refused with
    // InvalidDataException, saying what - a board silently read as empty is how data is lost.
    // A null text reads as empty, as a blank cell did.
    // ###########################################################################################
    public static class DataTreeJson
    {
        private static readonly JsonSerializerOptions WriteOptions = new()
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            WriteIndented = false
        };

        private static readonly JsonSerializerOptions ReadOptions = new()
        {
            Converters = { new NullAsEmptyText() }
        };

        // ---- board.json ------------------------------------------------------------------------

        public static string Write(BoardDocument document)
        {
            ArgumentNullException.ThrowIfNull(document);

            return DataTreeJson.WriteObject(
                [
                    ("formatVersion", document.FormatVersion),
                    ("hardwareName", document.HardwareName),
                    ("boardName", document.BoardName),
                    ("revisionDate", document.RevisionDate),
                    ("schematics", document.Schematics),
                    ("components", document.Components),
                    ("componentImages", document.ComponentImages),
                    ("scopeBaselines", document.ScopeBaselines),
                    ("componentFiles", document.ComponentFiles),
                    ("componentLinks", document.ComponentLinks),
                    ("componentPins", document.ComponentPins),
                    ("componentReplacements", document.ComponentReplacements),
                    ("boardFiles", document.BoardFiles),
                    ("boardLinks", document.BoardLinks),
                    ("credits", document.Credits),
                    ("importantSignals", document.ImportantSignals),
                    ("highlights", document.Highlights),
                    ("kiCadCalibrations", document.KiCadCalibrations),
                    ("schematicRenames", document.SchematicRenames),
                    ("files", document.Files),
                ],
                document.Unknown);
        }

        public static BoardDocument ReadBoard(string json)
        {
            BoardDocument document = DataTreeJson.Read<BoardDocument>(json, "board file");
            DataTreeJson.CheckVersion(document.FormatVersion, "board file");

            document.Schematics = DataTreeJson.Rows(document.Schematics, "schematics");
            document.Components = DataTreeJson.Rows(document.Components, "components");
            document.ComponentImages = DataTreeJson.Rows(document.ComponentImages, "componentImages");
            document.ScopeBaselines = DataTreeJson.Rows(document.ScopeBaselines, "scopeBaselines");
            document.ComponentFiles = DataTreeJson.Rows(document.ComponentFiles, "componentFiles");
            document.ComponentLinks = DataTreeJson.Rows(document.ComponentLinks, "componentLinks");
            document.ComponentPins = DataTreeJson.Rows(document.ComponentPins, "componentPins");
            document.ComponentReplacements = DataTreeJson.Rows(document.ComponentReplacements, "componentReplacements");
            document.BoardFiles = DataTreeJson.Rows(document.BoardFiles, "boardFiles");
            document.BoardLinks = DataTreeJson.Rows(document.BoardLinks, "boardLinks");
            document.Credits = DataTreeJson.Rows(document.Credits, "credits");
            document.ImportantSignals = DataTreeJson.Rows(document.ImportantSignals, "importantSignals");
            document.Highlights = DataTreeJson.Rows(document.Highlights, "highlights");
            document.KiCadCalibrations = DataTreeJson.Rows(document.KiCadCalibrations, "kiCadCalibrations");
            document.SchematicRenames = DataTreeJson.Rows(document.SchematicRenames, "schematicRenames");
            document.Files = DataTreeJson.Rows(document.Files, "files");

            return document;
        }

        // ---- the main file ---------------------------------------------------------------------

        public static string Write(MainDocument document)
        {
            ArgumentNullException.ThrowIfNull(document);

            return DataTreeJson.WriteObject(
                [
                    ("formatVersion", document.FormatVersion),
                    ("boardsRevisionDate", document.BoardsRevisionDate),
                    ("boards", document.Boards),
                    ("oscilloscopesRevisionDate", document.OscilloscopesRevisionDate),
                    ("oscilloscopes", document.Oscilloscopes),
                ],
                document.Unknown);
        }

        public static MainDocument ReadMain(string json)
        {
            MainDocument document = DataTreeJson.Read<MainDocument>(json, "main file");
            DataTreeJson.CheckVersion(document.FormatVersion, "main file");

            document.Boards = DataTreeJson.Rows(document.Boards, "boards");
            document.Oscilloscopes = DataTreeJson.Rows(document.Oscilloscopes, "oscilloscopes");

            return document;
        }

        // ---- the shared layout -----------------------------------------------------------------

        private static string WriteObject(
            IReadOnlyList<(string Name, object Value)> fields,
            Dictionary<string, JsonElement>? unknown)
        {
            var lines = new List<string>();

            foreach ((string name, object value) in fields)
                lines.Add($"  {DataTreeJson.Text(name)}: {DataTreeJson.ValueText(value)}");

            foreach (KeyValuePair<string, JsonElement> field in (unknown ?? []).OrderBy(pair => pair.Key, StringComparer.Ordinal))
                lines.Add($"  {DataTreeJson.Text(field.Key)}: {JsonSerializer.Serialize(field.Value, DataTreeJson.WriteOptions)}");

            var text = new StringBuilder();
            text.Append("{\n");
            text.Append(string.Join(",\n", lines));
            text.Append("\n}\n");

            return text.ToString();
        }

        private static string ValueText(object value)
        {
            if (value is not IList rows)
                return JsonSerializer.Serialize(value, value.GetType(), DataTreeJson.WriteOptions);

            if (rows.Count == 0)
                return "[]";

            IEnumerable<string> rowLines = rows
                .Cast<object>()
                .Select(row => "    " + JsonSerializer.Serialize(row, row.GetType(), DataTreeJson.WriteOptions));

            return "[\n" + string.Join(",\n", rowLines) + "\n  ]";
        }

        private static string Text(string value) => JsonSerializer.Serialize(value, DataTreeJson.WriteOptions);

        private static T Read<T>(string json, string what) where T : class
        {
            if (string.IsNullOrWhiteSpace(json))
                throw new InvalidDataException($"The {what} is empty.");

            T? document;

            try
            {
                document = JsonSerializer.Deserialize<T>(json, DataTreeJson.ReadOptions);
            }
            catch (JsonException ex)
            {
                throw new InvalidDataException($"The {what} could not be read: {ex.Message}", ex);
            }

            return document ?? throw new InvalidDataException($"The {what} holds no object.");
        }

        private static void CheckVersion(int formatVersion, string what)
        {
            if (formatVersion > DataTreeLayout.FormatVersion)
                throw new InvalidDataException(
                    $"The {what} is format {formatVersion}, written by a newer CRT - this one reads format {DataTreeLayout.FormatVersion}.");

            if (formatVersion < 1)
                throw new InvalidDataException($"The {what} names format {formatVersion}, which does not exist.");
        }

        private static List<T> Rows<T>(List<T>? rows, string listName) where T : class
        {
            if (rows is null)
                return new List<T>();

            if (rows.Any(row => row is null))
                throw new InvalidDataException($"The list \"{listName}\" holds an empty (null) row.");

            return rows;
        }

        // A null text in a file reads as empty text - a blank cell's value in the workbooks.
        private sealed class NullAsEmptyText : JsonConverter<string>
        {
            public override bool HandleNull => true;

            public override string Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) =>
                reader.TokenType switch
                {
                    JsonTokenType.Null => string.Empty,
                    JsonTokenType.String => reader.GetString() ?? string.Empty,
                    _ => throw new JsonException($"Expected text, found {reader.TokenType}.")
                };

            public override void Write(Utf8JsonWriter writer, string value, JsonSerializerOptions options) =>
                writer.WriteStringValue(value);
        }
    }
}
