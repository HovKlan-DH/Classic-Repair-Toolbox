using System.Collections.Generic;

namespace Handlers.DataHandling
{
    public class BoardSchematicEntry
    {
        public string UuidV4 { get; init; } = string.Empty;
        public string SchematicName { get; init; } = string.Empty;
        public string SchematicImageFile { get; init; } = string.Empty;
        public string MainImageHighlightColor { get; init; } = string.Empty;
        public string MainHighlightOpacity { get; init; } = string.Empty;
        public string ThumbnailImageHighlightColor { get; init; } = string.Empty;
        public string ThumbnailHighlightOpacity { get; init; } = string.Empty;
    }

    public class ComponentEntry
    {
        public string UuidV4 { get; init; } = string.Empty;
        public string BoardLabel { get; init; } = string.Empty;
        public string FriendlyName { get; init; } = string.Empty;
        public string TechnicalNameOrValue { get; init; } = string.Empty;
        public string PartNumber { get; init; } = string.Empty;
        public string Category { get; init; } = string.Empty;
        public string Region { get; init; } = string.Empty;
        public string Description { get; init; } = string.Empty;
    }

    public class ComponentImageEntry
    {
        public string UuidV4 { get; init; } = string.Empty;
        public string BoardLabel { get; init; } = string.Empty;
        public string Region { get; init; } = string.Empty;
        public string Pin { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string ExpectedOscilloscopeReading { get; init; } = string.Empty;
        public string File { get; init; } = string.Empty;
        public string Note { get; init; } = string.Empty;
        public string TimeDiv { get; init; } = string.Empty;
        public string VoltsDiv { get; init; } = string.Empty;
        public string TriggerLevelVolts { get; init; } = string.Empty;
    }

    public class ComponentHighlightEntry
    {
        public string SchematicName { get; init; } = string.Empty;
        public string BoardLabel { get; init; } = string.Empty;
        public string X { get; init; } = string.Empty;
        public string Y { get; init; } = string.Empty;
        public string Width { get; init; } = string.Empty;
        public string Height { get; init; } = string.Empty;
    }

    public class ComponentLocalFileEntry
    {
        public string UuidV4 { get; init; } = string.Empty;
        public string BoardLabel { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string File { get; init; } = string.Empty;
    }

    public class ComponentLinkEntry
    {
        public string UuidV4 { get; init; } = string.Empty;
        public string BoardLabel { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string Url { get; init; } = string.Empty;
    }

    public class BoardLocalFileEntry
    {
        public string UuidV4 { get; init; } = string.Empty;
        public string Category { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string File { get; init; } = string.Empty;
    }

    public class BoardLinkEntry
    {
        public string UuidV4 { get; init; } = string.Empty;
        public string Category { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string Url { get; init; } = string.Empty;
    }

    public class CreditEntry
    {
        public string UuidV4 { get; init; } = string.Empty;
        public string Category { get; init; } = string.Empty;
        public string SubCategory { get; init; } = string.Empty;
        public string NameOrHandle { get; init; } = string.Empty;
        public string Contact { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // Container for all data loaded from a board-specific Excel file.
    // IsLoaded is true only when the file was read successfully.
    // ###########################################################################################
    public class BoardData
    {
        /*
        public List<BoardSchematicEntry> Schematics { get; init; } = [];
        public List<ComponentEntry> Components { get; init; } = [];
        public List<ComponentImageEntry> ComponentImages { get; init; } = [];
        public List<ComponentHighlightEntry> ComponentHighlights { get; init; } = [];
        public List<ComponentLocalFileEntry> ComponentLocalFiles { get; init; } = [];
        public List<ComponentLinkEntry> ComponentLinks { get; init; } = [];
        public List<BoardLocalFileEntry> BoardLocalFiles { get; init; } = [];
        public List<BoardLinkEntry> BoardLinks { get; init; } = [];
        public List<CreditEntry> Credits { get; init; } = [];
        public bool IsLoaded { get; init; }
        */
        // Below is supporting .NET6
        public string RevisionDate { get; init; } = string.Empty;
        public List<BoardSchematicEntry> Schematics { get; init; } = new();
        public List<ComponentEntry> Components { get; init; } = new();
        public List<ComponentImageEntry> ComponentImages { get; init; } = new();
        public List<ComponentHighlightEntry> ComponentHighlights { get; init; } = new();
        public List<ComponentLocalFileEntry> ComponentLocalFiles { get; init; } = new();
        public List<ComponentLinkEntry> ComponentLinks { get; init; } = new();
        public List<BoardLocalFileEntry> BoardLocalFiles { get; init; } = new();
        public List<BoardLinkEntry> BoardLinks { get; init; } = new();
        public List<CreditEntry> Credits { get; init; } = new();
        public bool IsLoaded { get; init; }
    }

    public class OscilloscopeEntry
    {
        public string Brand { get; init; } = string.Empty;
        public string SeriesOrModel { get; init; } = string.Empty;
        public string Port { get; init; } = string.Empty;
        public string Identify { get; init; } = string.Empty;
        public string DrainErrorQueue { get; init; } = string.Empty;
        public string OperationComplete { get; init; } = string.Empty;
        public string ClearStatistics { get; init; } = string.Empty;
        public string QueryActiveTrigger { get; init; } = string.Empty;
        public string Stop { get; init; } = string.Empty;
        public string Single { get; init; } = string.Empty;
        public string Run { get; init; } = string.Empty;
        public string QueryTriggerMode { get; init; } = string.Empty;
        public string QueryTriggerLevel { get; init; } = string.Empty;
        public string SetTriggerLevel { get; init; } = string.Empty;
        public string QueryTimeDiv { get; init; } = string.Empty;
        public string SetTimeDiv { get; init; } = string.Empty;
        public string QueryVoltsDiv { get; init; } = string.Empty;
        public string SetVoltsDiv { get; init; } = string.Empty;
        public string DumpImage { get; init; } = string.Empty;
        public string TimeDivList { get; init; } = string.Empty;
        public string VoltsDivList { get; init; } = string.Empty;
        public string DebounceTime { get; init; } = string.Empty;
    }
}