using System.Collections.Generic;

namespace Handlers.DataHandling
{
    public class BoardSchematicEntry
    {
        public string SchematicName { get; init; } = string.Empty;
        public string CadName { get; init; } = string.Empty;
        public string SchematicImageFile { get; init; } = string.Empty;
        public string SchematicHighlightColor { get; init; } = string.Empty;
        public string SchematicHighlightOpacity { get; init; } = string.Empty;
        public string OppositeTraceHighlightColor { get; init; } = string.Empty;
        public string ThumbnailHighlightColor { get; init; } = string.Empty;
        public string ThumbnailHighlightOpacity { get; init; } = string.Empty;
    }

    public class ComponentEntry
    {
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
        public string BoardLabel { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string File { get; init; } = string.Empty;
    }

    public class ComponentLinkEntry
    {
        public string BoardLabel { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string Url { get; init; } = string.Empty;
    }

    public class BoardLocalFileEntry
    {
        public string Category { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string File { get; init; } = string.Empty;
    }

    public class BoardLinkEntry
    {
        public string Category { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string Url { get; init; } = string.Empty;
    }

    public class CreditEntry
    {
        public string Category { get; init; } = string.Empty;
        public string SubCategory { get; init; } = string.Empty;
        public string NameOrHandle { get; init; } = string.Empty;
        public string Contact { get; init; } = string.Empty;
    }

    public class KiCadImportantSignalEntry
    {
        public string DisplayName { get; init; } = string.Empty;
        public string KiCadNetName { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // One schematic's KiCad trace-overlay calibration - the offset/scale/mirror box that maps the
    // schematic image onto the KiCad board's own coordinate space. Mirrors
    // BoardComponentHighlightStorage.TryLoadKiCadCalibration/SaveKiCadCalibration's own field
    // shape, which stores this under a separate "KiCad calibration points" JSON root rather than
    // as a BoardData section - added here as a BoardDraft section only (session 2c task 8's KiCad
    // calibration piece); BoardDataReader does not yet load this as part of an official BoardData
    // read, so it is not populated by the ordinary board-load path today. See
    // Handlers/Geometry/KiCadCalibrationGeometry.cs for why mirroring is edge ORDERING
    // (Left/Right/Top/Bottom as doubles) rather than a separate flag - the same reasoning applies
    // here, and this entry stores the same four edges plus the two derived mirror flags the
    // existing JSON format already persists.
    // ###########################################################################################
    public class KiCadCalibrationEntry
    {
        public string SchematicName { get; init; } = string.Empty;
        public string CadName { get; init; } = string.Empty;
        public double OffsetX { get; init; }
        public double OffsetY { get; init; }
        public double ScaleX { get; init; } = 1.0;
        public double ScaleY { get; init; } = 1.0;
        public bool MirrorX { get; init; }
        public bool MirrorY { get; init; }
    }

    // ###########################################################################################
    // Container for all data loaded from a board-specific Excel file.
    // IsLoaded is true only when the file was read successfully.
    // ###########################################################################################
    public class BoardData
    {
        public string RevisionDate { get; init; } = string.Empty;

        // ###########################################################################################
        // The hardware and board as the workbook's own preamble names them ("Commodore 64",
        // "250407") - added 2026-09-23 so a written workbook can reproduce that header block.
        //
        // *** READ FOR PRESENTATION ONLY. NOTHING IN THE APPLICATION KEYS OFF THEM. *** Which
        // system a board belongs to is decided by its ExcelDataFile path everywhere it matters, so
        // these are the human-facing caption at the top of each sheet and nothing more. They are
        // carried on BoardData rather than passed to the writer separately because the writer
        // already receives the board and every caller would otherwise have to find the names
        // itself - four call sites, each with a different idea of where to look.
        //
        // Empty for a board whose preamble does not carry them, which the writer then simply
        // omits rather than emitting "# Hardware: ".
        // ###########################################################################################
        public string HardwareName { get; init; } = string.Empty;

        public string BoardName { get; init; } = string.Empty;

        public List<BoardSchematicEntry> Schematics { get; init; } = new();
        public List<ComponentEntry> Components { get; init; } = new();
        public List<ComponentImageEntry> ComponentImages { get; init; } = new();
        public List<ComponentHighlightEntry> ComponentHighlights { get; init; } = new();
        public List<ComponentLocalFileEntry> ComponentLocalFiles { get; init; } = new();
        public List<ComponentLinkEntry> ComponentLinks { get; init; } = new();
        public List<BoardLocalFileEntry> BoardLocalFiles { get; init; } = new();
        public List<BoardLinkEntry> BoardLinks { get; init; } = new();
        public List<CreditEntry> Credits { get; init; } = new();
        public List<KiCadImportantSignalEntry> KiCadImportantSignals { get; init; } = new();

        // ###########################################################################################
        // The same board with a different revision date - used by the publish, which stamps the
        // date it actually published on (added 2026-09-23).
        //
        // *** THE LISTS ARE SHARED, NOT COPIED, and that is safe HERE for one specific reason: ***
        // every writer in this codebase builds a NEW BoardData rather than mutating one, so no
        // caller holds a board whose rows can change underneath it. A deep copy would duplicate
        // upwards of 1,500 rows for a one-string change. If that convention ever stops holding,
        // this is the method to revisit.
        // ###########################################################################################
        public BoardData WithRevisionDate(string revisionDate) => new()
        {
            RevisionDate = revisionDate ?? string.Empty,
            HardwareName = this.HardwareName,
            BoardName = this.BoardName,
            Schematics = this.Schematics,
            Components = this.Components,
            ComponentImages = this.ComponentImages,
            ComponentHighlights = this.ComponentHighlights,
            ComponentLocalFiles = this.ComponentLocalFiles,
            ComponentLinks = this.ComponentLinks,
            BoardLocalFiles = this.BoardLocalFiles,
            BoardLinks = this.BoardLinks,
            Credits = this.Credits,
            KiCadImportantSignals = this.KiCadImportantSignals,
        };

        // The same board with different highlights - used when a save drops the highlights of a
        // deleted component (BoardTableDocument.ApplyTo). Lists shared, as WithRevisionDate's are.
        public BoardData WithComponentHighlights(List<ComponentHighlightEntry> highlights) => new()
        {
            RevisionDate = this.RevisionDate,
            HardwareName = this.HardwareName,
            BoardName = this.BoardName,
            Schematics = this.Schematics,
            Components = this.Components,
            ComponentImages = this.ComponentImages,
            ComponentHighlights = highlights ?? new(),
            ComponentLocalFiles = this.ComponentLocalFiles,
            ComponentLinks = this.ComponentLinks,
            BoardLocalFiles = this.BoardLocalFiles,
            BoardLinks = this.BoardLinks,
            Credits = this.Credits,
            KiCadImportantSignals = this.KiCadImportantSignals,
        };
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
        public string QueryTriggerSource { get; init; } = string.Empty;
        public string SetTriggerSource { get; init; } = string.Empty;
        public string QueryTriggerSlope { get; init; } = string.Empty;
        public string SetTriggerSlope { get; init; } = string.Empty;
        public string QueryTriggerLevel { get; init; } = string.Empty;
        public string SetTriggerLevel { get; init; } = string.Empty;

        public string QueryAvgCount { get; init; } = string.Empty;
        public string SetAvgCount { get; init; } = string.Empty;
        public string QueryMemoryDepth { get; init; } = string.Empty;
        public string SetMemoryDepth { get; init; } = string.Empty;
        public string QuerySampleRate { get; init; } = string.Empty;

        public string QueryProbeAttenuation { get; init; } = string.Empty;
        public string SetProbeAttenuation { get; init; } = string.Empty;
        public string QueryChannelScale { get; init; } = string.Empty;
        public string SetChannelScale { get; init; } = string.Empty;
        public string QueryChannelOffset { get; init; } = string.Empty;
        public string SetChannelOffset { get; init; } = string.Empty;

        public string QueryTimeScale { get; init; } = string.Empty;
        public string SetTimeScale { get; init; } = string.Empty;
        public string QueryTimeOffset { get; init; } = string.Empty;
        public string SetTimeOffset { get; init; } = string.Empty;

        public string QueryTimeDiv { get; init; } = string.Empty;
        public string SetTimeDiv { get; init; } = string.Empty;
        public string QueryVoltsDiv { get; init; } = string.Empty;
        public string SetVoltsDiv { get; init; } = string.Empty;

        public string QueryMeasureFrequency { get; init; } = string.Empty;
        public string QueryMeasurePeriod { get; init; } = string.Empty;
        public string QueryMeasureDutyCycle { get; init; } = string.Empty;
        public string QueryMeasureRiseTime { get; init; } = string.Empty;
        public string QueryMeasureFallTime { get; init; } = string.Empty;
        public string QueryMeasureOvershoot { get; init; } = string.Empty;
        public string QueryMeasurePreshoot { get; init; } = string.Empty;
        public string QueryMeasureAmplitude { get; init; } = string.Empty;
        public string QueryMeasurePkToPk { get; init; } = string.Empty;
        public string QueryMeasureMaximum { get; init; } = string.Empty;
        public string QueryMeasureMinimum { get; init; } = string.Empty;
        public string QueryMeasureMean { get; init; } = string.Empty;
        public string QueryMeasureRms { get; init; } = string.Empty;

        public string QueryCursorHorizontal { get; init; } = string.Empty;
        public string SetCursorHorizontal { get; init; } = string.Empty;
        public string QueryCursorVertical { get; init; } = string.Empty;
        public string SetCursorVertical { get; init; } = string.Empty;

        public string ReadWaveformPreamble { get; init; } = string.Empty;
        public string ReadWaveformData { get; init; } = string.Empty;
        public string DumpImage { get; init; } = string.Empty;
        public string ScreenshotCommand { get; init; } = string.Empty;

        public string TimeDivList { get; init; } = string.Empty;
        public string VoltsDivList { get; init; } = string.Empty;
        public string DebounceTime { get; init; } = string.Empty;

        public string DefaultFileExtension { get; init; } = string.Empty;
        public string Notes { get; init; } = string.Empty;
    }
}