using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // ONE BOARD IN THE CRT 4.0 DATA TREE - "<board id>/board.json" (DataTreeLayout).
    //
    // Everything the Excel generations kept in two files is here in one: the nine sheets of the
    // board workbook and the three roots of its JSON sidecar (highlights, KiCad calibrations,
    // schematic renames) - plus the board's own list of files, by name and checksum.
    //
    // *** THE SHAPE FOLLOWS THE WORKBOOK, WITH ONE DELIBERATE SPLIT. *** Rows keep the workbook's
    // columns, in its order, as text (opacities and trigger levels are parsed where their meaning
    // is known, as before). The "Component images" sheet becomes TWO lists: pictures and scope
    // baselines, split by ComponentImageKinds - the popup's own rule - so each can be edited as
    // what it is. A picture keeps any scope setting it carries (a pin-less row with a T/DIV is
    // still a picture), written only when there is one.
    //
    // *** A FILE IS NAMED BY THE BOARD. *** Every "file" field - a schematic's image, a picture,
    // a baseline, a component or board file, a rename's image - holds a name from this board's
    // own `files` list ("Component images/7406.png"), never a path into the tree. `files` gives
    // each name its SHA-256 and size; the bytes are at DataTreeLayout.StoredFilePath.
    //
    // *** WHAT THIS CRT DOES NOT KNOW, IT KEEPS. *** Every object carries the fields it did not
    // recognise in `Unknown` (JsonExtensionData), and BoardDocumentJson writes them back - so a
    // field a later CRT adds survives an edit made with this one. FormatVersion is what a later
    // CRT raises when it changes something this one could not read.
    // ###########################################################################################
    public sealed class BoardDocument : IKeepsUnknownFields
    {
        [JsonPropertyName("formatVersion")] public int FormatVersion { get; set; } = DataTreeLayout.FormatVersion;

        // The caption the board's workbook carried ("# Hardware:", "# Board:"). Presentation only,
        // as before - the board id and the main file's names are what CRT keys off.
        [JsonPropertyName("hardwareName")] public string HardwareName { get; set; } = string.Empty;

        [JsonPropertyName("boardName")] public string BoardName { get; set; } = string.Empty;

        // The board's hand-maintained revision date ("2026-October-4"), the server's at a publish.
        [JsonPropertyName("revisionDate")] public string RevisionDate { get; set; } = string.Empty;

        [JsonPropertyName("schematics")] public List<SchematicRow> Schematics { get; set; } = new();

        [JsonPropertyName("components")] public List<ComponentRow> Components { get; set; } = new();

        [JsonPropertyName("componentImages")] public List<ComponentImageRow> ComponentImages { get; set; } = new();

        [JsonPropertyName("scopeBaselines")] public List<ScopeBaselineRow> ScopeBaselines { get; set; } = new();

        [JsonPropertyName("componentFiles")] public List<ComponentFileRow> ComponentFiles { get; set; } = new();

        [JsonPropertyName("componentLinks")] public List<ComponentLinkRow> ComponentLinks { get; set; } = new();

        [JsonPropertyName("boardFiles")] public List<BoardFileRow> BoardFiles { get; set; } = new();

        [JsonPropertyName("boardLinks")] public List<BoardLinkRow> BoardLinks { get; set; } = new();

        [JsonPropertyName("credits")] public List<CreditRow> Credits { get; set; } = new();

        [JsonPropertyName("importantSignals")] public List<ImportantSignalRow> ImportantSignals { get; set; } = new();

        // CRT 4.0's own data (step 2, 2026-10-25): each component's pins and replacement parts. A
        // board converted from Excel has none - two empty lists.
        [JsonPropertyName("componentPins")] public List<ComponentPinRow> ComponentPins { get; set; } = new();

        [JsonPropertyName("componentReplacements")] public List<ComponentReplacementRow> ComponentReplacements { get; set; } = new();

        [JsonPropertyName("highlights")] public List<HighlightRow> Highlights { get; set; } = new();

        [JsonPropertyName("kiCadCalibrations")] public List<KiCadCalibrationRow> KiCadCalibrations { get; set; } = new();

        [JsonPropertyName("schematicRenames")] public List<SchematicRenameRow> SchematicRenames { get; set; } = new();

        [JsonPropertyName("files")] public List<StoredFileRow> Files { get; set; } = new();

        [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }

        // ---- The rows ------------------------------------------------------------------------

        public sealed class SchematicRow : IKeepsUnknownFields
        {
            [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
            [JsonPropertyName("imageFile")] public string ImageFile { get; set; } = string.Empty;
            [JsonPropertyName("highlightColor")] public string HighlightColor { get; set; } = string.Empty;
            [JsonPropertyName("highlightOpacity")] public string HighlightOpacity { get; set; } = string.Empty;
            [JsonPropertyName("oppositeTraceHighlightColor")] public string OppositeTraceHighlightColor { get; set; } = string.Empty;
            [JsonPropertyName("thumbnailHighlightColor")] public string ThumbnailHighlightColor { get; set; } = string.Empty;
            [JsonPropertyName("thumbnailHighlightOpacity")] public string ThumbnailHighlightOpacity { get; set; } = string.Empty;
            [JsonPropertyName("cadName")] public string CadName { get; set; } = string.Empty;
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        public sealed class ComponentRow : IKeepsUnknownFields
        {
            [JsonPropertyName("boardLabel")] public string BoardLabel { get; set; } = string.Empty;
            [JsonPropertyName("friendlyName")] public string FriendlyName { get; set; } = string.Empty;
            [JsonPropertyName("technicalNameOrValue")] public string TechnicalNameOrValue { get; set; } = string.Empty;
            [JsonPropertyName("partNumber")] public string PartNumber { get; set; } = string.Empty;
            [JsonPropertyName("category")] public string Category { get; set; } = string.Empty;
            [JsonPropertyName("region")] public string Region { get; set; } = string.Empty;
            [JsonPropertyName("description")] public string Description { get; set; } = string.Empty;

            // CRT 4.0's own (2026-10-25), written only when filled - so the rows of a board that has
            // none read as they always did, line for line.
            [JsonPropertyName("package"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
            public string? Package { get; set; }

            [JsonPropertyName("repairNotes"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
            public string? RepairNotes { get; set; }

            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        // A picture of a component - a photo, a render, a pinout. The three scope settings are
        // written only when the row has one (see the class header).
        public sealed class ComponentImageRow : IKeepsUnknownFields
        {
            [JsonPropertyName("boardLabel")] public string BoardLabel { get; set; } = string.Empty;
            [JsonPropertyName("region")] public string Region { get; set; } = string.Empty;
            [JsonPropertyName("pin")] public string Pin { get; set; } = string.Empty;
            [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
            [JsonPropertyName("expectedOscilloscopeReading")] public string ExpectedOscilloscopeReading { get; set; } = string.Empty;
            [JsonPropertyName("file")] public string File { get; set; } = string.Empty;
            [JsonPropertyName("note")] public string Note { get; set; } = string.Empty;

            [JsonPropertyName("timeDiv"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
            public string? TimeDiv { get; set; }

            [JsonPropertyName("voltsDiv"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
            public string? VoltsDiv { get; set; }

            [JsonPropertyName("triggerLevel"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
            public string? TriggerLevel { get; set; }

            // A baseline's two conditions (CRT 4.0, 2026-10-25), kept on a picture row that has them
            // so nothing typed in the table is lost - written only when filled.
            [JsonPropertyName("machineState"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
            public string? MachineState { get; set; }

            [JsonPropertyName("source"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
            public string? Source { get; set; }

            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        // An oscilloscope reading of one pin, from a known-good board.
        public sealed class ScopeBaselineRow : IKeepsUnknownFields
        {
            [JsonPropertyName("boardLabel")] public string BoardLabel { get; set; } = string.Empty;
            [JsonPropertyName("region")] public string Region { get; set; } = string.Empty;
            [JsonPropertyName("pin")] public string Pin { get; set; } = string.Empty;
            [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
            [JsonPropertyName("expectedOscilloscopeReading")] public string ExpectedOscilloscopeReading { get; set; } = string.Empty;
            [JsonPropertyName("file")] public string File { get; set; } = string.Empty;
            [JsonPropertyName("note")] public string Note { get; set; } = string.Empty;
            [JsonPropertyName("timeDiv")] public string TimeDiv { get; set; } = string.Empty;
            [JsonPropertyName("voltsDiv")] public string VoltsDiv { get; set; } = string.Empty;
            [JsonPropertyName("triggerLevel")] public string TriggerLevel { get; set; } = string.Empty;

            // CRT 4.0's own (2026-10-25): what the machine was doing, and where the reading comes
            // from (ComponentVocabulary.BaselineSources) - written only when filled.
            [JsonPropertyName("machineState"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
            public string? MachineState { get; set; }

            [JsonPropertyName("source"), JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
            public string? Source { get; set; }

            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        // One pin of a component - see ComponentPinEntry.
        public sealed class ComponentPinRow : IKeepsUnknownFields
        {
            [JsonPropertyName("boardLabel")] public string BoardLabel { get; set; } = string.Empty;
            [JsonPropertyName("region")] public string Region { get; set; } = string.Empty;
            [JsonPropertyName("pin")] public string Pin { get; set; } = string.Empty;
            [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
            [JsonPropertyName("direction")] public string Direction { get; set; } = string.Empty;
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        // A part that can stand in for a component - see ComponentReplacementEntry.
        public sealed class ComponentReplacementRow : IKeepsUnknownFields
        {
            [JsonPropertyName("boardLabel")] public string BoardLabel { get; set; } = string.Empty;
            [JsonPropertyName("region")] public string Region { get; set; } = string.Empty;
            [JsonPropertyName("partNumber")] public string PartNumber { get; set; } = string.Empty;
            [JsonPropertyName("kind")] public string Kind { get; set; } = string.Empty;
            [JsonPropertyName("changes")] public string Changes { get; set; } = string.Empty;
            [JsonPropertyName("source")] public string Source { get; set; } = string.Empty;
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        public sealed class ComponentFileRow : IKeepsUnknownFields
        {
            [JsonPropertyName("boardLabel")] public string BoardLabel { get; set; } = string.Empty;
            [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
            [JsonPropertyName("file")] public string File { get; set; } = string.Empty;
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        public sealed class ComponentLinkRow : IKeepsUnknownFields
        {
            [JsonPropertyName("boardLabel")] public string BoardLabel { get; set; } = string.Empty;
            [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
            [JsonPropertyName("url")] public string Url { get; set; } = string.Empty;
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        public sealed class BoardFileRow : IKeepsUnknownFields
        {
            [JsonPropertyName("category")] public string Category { get; set; } = string.Empty;
            [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
            [JsonPropertyName("file")] public string File { get; set; } = string.Empty;
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        public sealed class BoardLinkRow : IKeepsUnknownFields
        {
            [JsonPropertyName("category")] public string Category { get; set; } = string.Empty;
            [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
            [JsonPropertyName("url")] public string Url { get; set; } = string.Empty;
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        public sealed class CreditRow : IKeepsUnknownFields
        {
            [JsonPropertyName("category")] public string Category { get; set; } = string.Empty;
            [JsonPropertyName("subCategory")] public string SubCategory { get; set; } = string.Empty;
            [JsonPropertyName("nameOrHandle")] public string NameOrHandle { get; set; } = string.Empty;
            [JsonPropertyName("contact")] public string Contact { get; set; } = string.Empty;
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        public sealed class ImportantSignalRow : IKeepsUnknownFields
        {
            [JsonPropertyName("displayName")] public string DisplayName { get; set; } = string.Empty;
            [JsonPropertyName("kiCadNetName")] public string KiCadNetName { get; set; } = string.Empty;
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        // One rectangle marking a component on a schematic image, in the image's own pixels.
        public sealed class HighlightRow : IKeepsUnknownFields
        {
            [JsonPropertyName("schematic")] public string Schematic { get; set; } = string.Empty;
            [JsonPropertyName("boardLabel")] public string BoardLabel { get; set; } = string.Empty;
            [JsonPropertyName("x")] public int X { get; set; }
            [JsonPropertyName("y")] public int Y { get; set; }
            [JsonPropertyName("width")] public int Width { get; set; }
            [JsonPropertyName("height")] public int Height { get; set; }
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        // How one schematic image lines up with a KiCad view - see KiCadCalibrationEntry.
        public sealed class KiCadCalibrationRow : IKeepsUnknownFields
        {
            [JsonPropertyName("schematic")] public string Schematic { get; set; } = string.Empty;
            [JsonPropertyName("cadName")] public string CadName { get; set; } = string.Empty;
            [JsonPropertyName("offsetX")] public double OffsetX { get; set; }
            [JsonPropertyName("offsetY")] public double OffsetY { get; set; }
            [JsonPropertyName("scaleX")] public double ScaleX { get; set; } = 1.0;
            [JsonPropertyName("scaleY")] public double ScaleY { get; set; } = 1.0;
            [JsonPropertyName("mirrorX")] public bool MirrorX { get; set; }
            [JsonPropertyName("mirrorY")] public bool MirrorY { get; set; }
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        // A schematic image renamed by a publish, oldest first - see SchematicRenameLog.
        public sealed class SchematicRenameRow : IKeepsUnknownFields
        {
            [JsonPropertyName("from")] public string From { get; set; } = string.Empty;
            [JsonPropertyName("to")] public string To { get; set; } = string.Empty;
            [JsonPropertyName("file")] public string File { get; set; } = string.Empty;
            [JsonPropertyName("date")] public string Date { get; set; } = string.Empty;
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        // One file of this board: the board's name for it, and the bytes it names.
        public sealed class StoredFileRow : IKeepsUnknownFields
        {
            [JsonPropertyName("name")] public string Name { get; set; } = string.Empty;
            [JsonPropertyName("sha256")] public string Sha256 { get; set; } = string.Empty;
            [JsonPropertyName("size")] public long Size { get; set; }
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }
    }

    // Every object of a data tree file keeps the fields this CRT does not know - what
    // BoardDocumentUnknowns carries from a board file to the one written over it.
    public interface IKeepsUnknownFields
    {
        Dictionary<string, JsonElement>? Unknown { get; set; }
    }
}
