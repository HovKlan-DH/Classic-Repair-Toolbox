using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE MAIN FILE OF THE CRT 4.0 DATA TREE - "Classic-Repair-Toolbox.v4.0.0.json"
    // (DataTreeLayout), what the main Excel file's two sheets were.
    //
    // `boards` is "Hardware & Board": the drop-down lists, top to bottom - each board by its id
    // (its folders, "Commodore/C64/250407"), the names it is listed under and its hardware notes.
    // Its file is found by the id (DataTreeLayout.BoardFilePath), so no path is written here.
    //
    // `oscilloscopes` is "Oscilloscope": one entry per series or model, with every SCPI command
    // the sheet's columns held. The brand is written on every entry - the sheet left it blank
    // under the row that named it, and CRT carried it down when reading.
    //
    // Each sheet's "# Revision date:" line is kept as its own field. Unknown fields are kept and
    // written back, as in BoardDocument.
    // ###########################################################################################
    public sealed class MainDocument
    {
        [JsonPropertyName("formatVersion")] public int FormatVersion { get; set; } = DataTreeLayout.FormatVersion;

        [JsonPropertyName("boardsRevisionDate")] public string BoardsRevisionDate { get; set; } = string.Empty;

        [JsonPropertyName("boards")] public List<BoardListingRow> Boards { get; set; } = new();

        [JsonPropertyName("oscilloscopesRevisionDate")] public string OscilloscopesRevisionDate { get; set; } = string.Empty;

        [JsonPropertyName("oscilloscopes")] public List<OscilloscopeRow> Oscilloscopes { get; set; } = new();

        [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }

        public sealed class BoardListingRow
        {
            [JsonPropertyName("boardId")] public string BoardId { get; set; } = string.Empty;
            [JsonPropertyName("hardwareName")] public string HardwareName { get; set; } = string.Empty;
            [JsonPropertyName("boardName")] public string BoardName { get; set; } = string.Empty;
            [JsonPropertyName("hardwareNotes")] public string HardwareNotes { get; set; } = string.Empty;
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }

        // The "Oscilloscope" sheet's columns, in its order (MasterWorkbookSchema names them).
        public sealed class OscilloscopeRow
        {
            [JsonPropertyName("brand")] public string Brand { get; set; } = string.Empty;
            [JsonPropertyName("seriesOrModel")] public string SeriesOrModel { get; set; } = string.Empty;
            [JsonPropertyName("port")] public string Port { get; set; } = string.Empty;
            [JsonPropertyName("debounceTime")] public string DebounceTime { get; set; } = string.Empty;
            [JsonPropertyName("identify")] public string Identify { get; set; } = string.Empty;
            [JsonPropertyName("drainErrorQueue")] public string DrainErrorQueue { get; set; } = string.Empty;
            [JsonPropertyName("operationComplete")] public string OperationComplete { get; set; } = string.Empty;
            [JsonPropertyName("clearStatistics")] public string ClearStatistics { get; set; } = string.Empty;
            [JsonPropertyName("queryActiveTrigger")] public string QueryActiveTrigger { get; set; } = string.Empty;
            [JsonPropertyName("stop")] public string Stop { get; set; } = string.Empty;
            [JsonPropertyName("single")] public string Single { get; set; } = string.Empty;
            [JsonPropertyName("run")] public string Run { get; set; } = string.Empty;
            [JsonPropertyName("queryTriggerMode")] public string QueryTriggerMode { get; set; } = string.Empty;
            [JsonPropertyName("queryTriggerLevel")] public string QueryTriggerLevel { get; set; } = string.Empty;
            [JsonPropertyName("setTriggerLevel")] public string SetTriggerLevel { get; set; } = string.Empty;
            [JsonPropertyName("queryTimeDiv")] public string QueryTimeDiv { get; set; } = string.Empty;
            [JsonPropertyName("setTimeDiv")] public string SetTimeDiv { get; set; } = string.Empty;
            [JsonPropertyName("queryVoltsDiv")] public string QueryVoltsDiv { get; set; } = string.Empty;
            [JsonPropertyName("setVoltsDiv")] public string SetVoltsDiv { get; set; } = string.Empty;
            [JsonPropertyName("dumpImage")] public string DumpImage { get; set; } = string.Empty;
            [JsonPropertyName("timeDivList")] public string TimeDivList { get; set; } = string.Empty;
            [JsonPropertyName("voltsDivList")] public string VoltsDivList { get; set; } = string.Empty;
            [JsonExtensionData] public Dictionary<string, JsonElement>? Unknown { get; set; }
        }
    }
}
