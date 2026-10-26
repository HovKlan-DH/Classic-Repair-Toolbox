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

        // This row with another CAD name and every other column as it was - what picking a KiCad
        // name in the "Schematic images" window writes (2026-10-10). A column left out of this
        // copy would be erased from the workbook, silently; BoardSchematicEntryTests compares
        // every property by reflection, so a new one cannot be forgotten here.
        public BoardSchematicEntry WithCadName(string cadName) => this.With(this.SchematicName, cadName);

        // This row under another schematic name - what renaming an image in the "Schematic images"
        // window writes (2026-10-10). Same rule as WithCadName.
        public BoardSchematicEntry WithSchematicName(string schematicName) => this.With(schematicName, this.CadName);

        private BoardSchematicEntry With(string schematicName, string cadName) => new()
        {
            SchematicName = schematicName,
            CadName = cadName,
            SchematicImageFile = this.SchematicImageFile,
            SchematicHighlightColor = this.SchematicHighlightColor,
            SchematicHighlightOpacity = this.SchematicHighlightOpacity,
            OppositeTraceHighlightColor = this.OppositeTraceHighlightColor,
            ThumbnailHighlightColor = this.ThumbnailHighlightColor,
            ThumbnailHighlightOpacity = this.ThumbnailHighlightOpacity,
        };
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

        // ###########################################################################################
        // New in CRT 4.0 (step 2, 2026-10-25 - the plan's component pane: "a one-line description
        // with a counter, plus repair notes with clickable links ... Pick a package"). Never in the
        // Excel generations: a board file writes each only when it is filled.
        //
        // Package is the chip's case, "DIP-40" (ComponentVocabulary.Packages offers the usual ones);
        // the pins are rows of their own (ComponentPinEntry). RepairNotes is the long text the
        // one-line Description has no room for - several lines, links clickable where it is shown.
        // ###########################################################################################
        public string Package { get; init; } = string.Empty;
        public string RepairNotes { get; init; } = string.Empty;
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

        // ###########################################################################################
        // A SCOPE BASELINE's conditions, new in CRT 4.0 (step 2, 2026-10-25 - the plan: "region and
        // machine state ... where it comes from: measured on a known-good board, datasheet, service
        // manual, or other"). MachineState is what the machine was doing ("READY prompt, nothing
        // pressed"); Source is one of ComponentVocabulary.BaselineSources. A picture leaves both
        // empty, and a board file writes each only when it is filled.
        // ###########################################################################################
        public string MachineState { get; init; } = string.Empty;
        public string Source { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // ONE PIN OF A COMPONENT - new in CRT 4.0 (step 2, 2026-10-25 - the plan's Pinout: "name the
    // pins: number, signal name, direction"). Pin is its number on the case ("1", "40"), Name the
    // signal ("VCC", "~{RESET}"), Direction one of ComponentVocabulary.PinDirections. A pin is
    // filed by label and region, as a component's pictures are: a blank region is every region.
    // ###########################################################################################
    public class ComponentPinEntry
    {
        public string BoardLabel { get; init; } = string.Empty;
        public string Region { get; init; } = string.Empty;
        public string Pin { get; init; } = string.Empty;
        public string Name { get; init; } = string.Empty;
        public string Direction { get; init; } = string.Empty;
    }

    // ###########################################################################################
    // A PART THAT CAN STAND IN FOR A COMPONENT - new in CRT 4.0 (step 2, 2026-10-25 - the plan:
    // "Replacements: part number; kind (drop-in, or needs changes); what to change; where the advice
    // comes from"). Kind is one of ComponentVocabulary.ReplacementKinds; Changes says what has to
    // be done for one that is not a drop-in; Source is where the advice comes from, in words (a
    // forum thread, a magazine, "tried on my own board").
    // ###########################################################################################
    public class ComponentReplacementEntry
    {
        public string BoardLabel { get; init; } = string.Empty;
        public string Region { get; init; } = string.Empty;
        public string PartNumber { get; init; } = string.Empty;
        public string Kind { get; init; } = string.Empty;
        public string Changes { get; init; } = string.Empty;
        public string Source { get; init; } = string.Empty;
    }

    public class ComponentHighlightEntry
    {
        public string SchematicName { get; init; } = string.Empty;
        public string BoardLabel { get; init; } = string.Empty;
        public string X { get; init; } = string.Empty;
        public string Y { get; init; } = string.Empty;
        public string Width { get; init; } = string.Empty;
        public string Height { get; init; } = string.Empty;

        // This rectangle on a renamed schematic, every other value as it was (2026-10-10).
        // BoardSchematicEntryTests compares every property by reflection.
        public ComponentHighlightEntry WithSchematicName(string schematicName) => new()
        {
            SchematicName = schematicName,
            BoardLabel = this.BoardLabel,
            X = this.X,
            Y = this.Y,
            Width = this.Width,
            Height = this.Height,
        };

        // This rectangle on a relabelled component, every other value as it was (ComponentRelabel,
        // code review 2026-10-13).
        public ComponentHighlightEntry WithBoardLabel(string boardLabel) => new()
        {
            SchematicName = this.SchematicName,
            BoardLabel = boardLabel,
            X = this.X,
            Y = this.Y,
            Width = this.Width,
            Height = this.Height,
        };
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

        // This calibration filed under a renamed schematic, every other value as it was
        // (2026-10-10). BoardSchematicEntryTests compares every property by reflection.
        public KiCadCalibrationEntry WithSchematicName(string schematicName) => new()
        {
            SchematicName = schematicName,
            CadName = this.CadName,
            OffsetX = this.OffsetX,
            OffsetY = this.OffsetY,
            ScaleX = this.ScaleX,
            ScaleY = this.ScaleY,
            MirrorX = this.MirrorX,
            MirrorY = this.MirrorY,
        };
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
        // board a board belongs to is decided by its ExcelDataFile path everywhere it matters, so
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
        // New in CRT 4.0 (step 2, 2026-10-25): each component's pins and replacement parts.
        //
        // *** A BOARD COPIED BY HAND NAMES EVERY ONE OF THESE. *** BoardData is copied by listing its
        // properties, and a list left out of a copy is dropped without a sound - BoardDataCaptionTests
        // fails on a copy that carries the board's sections without all of them.
        // ###########################################################################################
        public List<ComponentPinEntry> ComponentPins { get; init; } = new();
        public List<ComponentReplacementEntry> ComponentReplacements { get; init; } = new();

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
            ComponentPins = this.ComponentPins,
            ComponentReplacements = this.ComponentReplacements,
        };

        // ###########################################################################################
        // The same board with a different Schematics list - importing, removing or reordering
        // board images (BoardFilesWindow). Lists shared, as WithRevisionDate's are.
        //
        // *** IT EXISTS BECAUSE THAT WINDOW'S OWN COPY DROPPED THE CAPTION (2026-09-27). *** It
        // listed the sections by hand and left out HardwareName and BoardName, so every image
        // import or removal wrote the draft back without its "# Hardware:" / "# Board:" lines - the
        // caption a brand-new board was already reported once for missing (see DraftSeeder).
        // A copy here sits beside the other two, where a new property is seen by whoever adds it.
        // ###########################################################################################
        public BoardData WithSchematics(List<BoardSchematicEntry> schematics) => new()
        {
            RevisionDate = this.RevisionDate,
            HardwareName = this.HardwareName,
            BoardName = this.BoardName,
            Schematics = schematics ?? new(),
            Components = this.Components,
            ComponentImages = this.ComponentImages,
            ComponentHighlights = this.ComponentHighlights,
            ComponentLocalFiles = this.ComponentLocalFiles,
            ComponentLinks = this.ComponentLinks,
            BoardLocalFiles = this.BoardLocalFiles,
            BoardLinks = this.BoardLinks,
            Credits = this.Credits,
            KiCadImportantSignals = this.KiCadImportantSignals,
            ComponentPins = this.ComponentPins,
            ComponentReplacements = this.ComponentReplacements,
        };

        // The same board under another "# Hardware:" / "# Board:" caption - the publish, for a board
        // whose caption it has to supply (PublishMerge.CaptionedAs). Lists shared, as
        // WithRevisionDate's are.
        public BoardData WithCaption(string hardwareName, string boardName) => new()
        {
            RevisionDate = this.RevisionDate,
            HardwareName = hardwareName ?? string.Empty,
            BoardName = boardName ?? string.Empty,
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
            ComponentPins = this.ComponentPins,
            ComponentReplacements = this.ComponentReplacements,
        };

        // ###########################################################################################
        // The same board with some of a component's sections replaced - the component pane's edits
        // (ComponentPaneEdits, CRT 4.0 step 2). A null list keeps the board's own. Lists shared, as
        // WithRevisionDate's are.
        // ###########################################################################################
        public BoardData WithComponentSections(
            List<ComponentEntry>? components = null,
            List<ComponentImageEntry>? componentImages = null,
            List<ComponentLocalFileEntry>? componentLocalFiles = null,
            List<ComponentLinkEntry>? componentLinks = null,
            List<ComponentPinEntry>? componentPins = null,
            List<ComponentReplacementEntry>? componentReplacements = null) => new()
        {
            RevisionDate = this.RevisionDate,
            HardwareName = this.HardwareName,
            BoardName = this.BoardName,
            Schematics = this.Schematics,
            Components = components ?? this.Components,
            ComponentImages = componentImages ?? this.ComponentImages,
            ComponentHighlights = this.ComponentHighlights,
            ComponentLocalFiles = componentLocalFiles ?? this.ComponentLocalFiles,
            ComponentLinks = componentLinks ?? this.ComponentLinks,
            BoardLocalFiles = this.BoardLocalFiles,
            BoardLinks = this.BoardLinks,
            Credits = this.Credits,
            KiCadImportantSignals = this.KiCadImportantSignals,
            ComponentPins = componentPins ?? this.ComponentPins,
            ComponentReplacements = componentReplacements ?? this.ComponentReplacements,
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
            ComponentPins = this.ComponentPins,
            ComponentReplacements = this.ComponentReplacements,
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