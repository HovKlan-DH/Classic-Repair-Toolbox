using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A board as the Excel generations hold it: the workbook's rows (BoardData, highlights
    // included) and the sidecar's other two roots.
    // ###########################################################################################
    public sealed record BoardContent(
        BoardData Board,
        IReadOnlyList<KiCadCalibrationEntry> Calibrations,
        IReadOnlyList<SchematicRenameRecord> Renames);

    // ###########################################################################################
    // BETWEEN BoardContent AND THE CRT 4.0 BOARD FILE (BoardDocument), both ways.
    //
    // Every value is carried as it is, with two deliberate changes, both undone on the way back:
    //
    //   - "Component images" is split into pictures and scope baselines by ComponentImageKinds.
    //     Back again, the pictures come first and the baselines after, each in its own order - the
    //     popup shows the two apart, so their interleaving in the sheet was never seen;
    //   - a highlight's coordinates are whole numbers in the file (the sidecar's own form), text in
    //     BoardData (as BoardComponentHighlightStorage reads them).
    //
    // *** FILES ARE NAMED BY THE CALLER. *** ToDocument asks `fileName` for the board's name of
    // every file a row cites (BoardFileNames), and is handed the board's `files` list - so the
    // mapping itself never decides where anything is. ToContent leaves the names in the rows.
    // ###########################################################################################
    public static class BoardDocumentMapping
    {
        public static BoardDocument ToDocument(
            BoardContent content,
            Func<string, string> fileName,
            IReadOnlyList<BoardDocument.StoredFileRow> files)
        {
            ArgumentNullException.ThrowIfNull(content);
            ArgumentNullException.ThrowIfNull(fileName);
            ArgumentNullException.ThrowIfNull(files);

            BoardData board = content.Board;
            string Name(string? cited) => string.IsNullOrWhiteSpace(cited) ? string.Empty : fileName(cited);

            return new BoardDocument
            {
                HardwareName = board.HardwareName,
                BoardName = board.BoardName,
                RevisionDate = board.RevisionDate,
                Schematics = board.Schematics.Select(row => new BoardDocument.SchematicRow
                {
                    Name = row.SchematicName,
                    ImageFile = Name(row.SchematicImageFile),
                    HighlightColor = row.SchematicHighlightColor,
                    HighlightOpacity = row.SchematicHighlightOpacity,
                    OppositeTraceHighlightColor = row.OppositeTraceHighlightColor,
                    ThumbnailHighlightColor = row.ThumbnailHighlightColor,
                    ThumbnailHighlightOpacity = row.ThumbnailHighlightOpacity,
                    CadName = row.CadName,
                }).ToList(),
                Components = board.Components.Select(row => new BoardDocument.ComponentRow
                {
                    BoardLabel = row.BoardLabel,
                    FriendlyName = row.FriendlyName,
                    TechnicalNameOrValue = row.TechnicalNameOrValue,
                    PartNumber = row.PartNumber,
                    Category = row.Category,
                    Region = row.Region,
                    Description = row.Description,
                    Package = BoardDocumentMapping.WhenGiven(row.Package),
                    RepairNotes = BoardDocumentMapping.WhenGiven(row.RepairNotes),
                }).ToList(),
                ComponentImages = board.ComponentImages
                    .Where(row => !ComponentImageKinds.IsScopeBaseline(row))
                    .Select(row => new BoardDocument.ComponentImageRow
                    {
                        BoardLabel = row.BoardLabel,
                        Region = row.Region,
                        Pin = row.Pin,
                        Name = row.Name,
                        ExpectedOscilloscopeReading = row.ExpectedOscilloscopeReading,
                        File = Name(row.File),
                        Note = row.Note,
                        TimeDiv = BoardDocumentMapping.WhenGiven(row.TimeDiv),
                        VoltsDiv = BoardDocumentMapping.WhenGiven(row.VoltsDiv),
                        TriggerLevel = BoardDocumentMapping.WhenGiven(row.TriggerLevelVolts),
                        MachineState = BoardDocumentMapping.WhenGiven(row.MachineState),
                        Source = BoardDocumentMapping.WhenGiven(row.Source),
                    }).ToList(),
                ScopeBaselines = board.ComponentImages
                    .Where(ComponentImageKinds.IsScopeBaseline)
                    .Select(row => new BoardDocument.ScopeBaselineRow
                    {
                        BoardLabel = row.BoardLabel,
                        Region = row.Region,
                        Pin = row.Pin,
                        Name = row.Name,
                        ExpectedOscilloscopeReading = row.ExpectedOscilloscopeReading,
                        File = Name(row.File),
                        Note = row.Note,
                        TimeDiv = row.TimeDiv,
                        VoltsDiv = row.VoltsDiv,
                        TriggerLevel = row.TriggerLevelVolts,
                        MachineState = BoardDocumentMapping.WhenGiven(row.MachineState),
                        Source = BoardDocumentMapping.WhenGiven(row.Source),
                    }).ToList(),
                ComponentFiles = board.ComponentLocalFiles.Select(row => new BoardDocument.ComponentFileRow
                {
                    BoardLabel = row.BoardLabel,
                    Name = row.Name,
                    File = Name(row.File),
                }).ToList(),
                ComponentLinks = board.ComponentLinks.Select(row => new BoardDocument.ComponentLinkRow
                {
                    BoardLabel = row.BoardLabel,
                    Name = row.Name,
                    Url = row.Url,
                }).ToList(),
                ComponentPins = board.ComponentPins.Select(row => new BoardDocument.ComponentPinRow
                {
                    BoardLabel = row.BoardLabel,
                    Region = row.Region,
                    Pin = row.Pin,
                    Name = row.Name,
                    Direction = row.Direction,
                }).ToList(),
                ComponentReplacements = board.ComponentReplacements.Select(row => new BoardDocument.ComponentReplacementRow
                {
                    BoardLabel = row.BoardLabel,
                    Region = row.Region,
                    PartNumber = row.PartNumber,
                    Kind = row.Kind,
                    Changes = row.Changes,
                    Source = row.Source,
                }).ToList(),
                BoardFiles = board.BoardLocalFiles.Select(row => new BoardDocument.BoardFileRow
                {
                    Category = row.Category,
                    Name = row.Name,
                    File = Name(row.File),
                }).ToList(),
                BoardLinks = board.BoardLinks.Select(row => new BoardDocument.BoardLinkRow
                {
                    Category = row.Category,
                    Name = row.Name,
                    Url = row.Url,
                }).ToList(),
                Credits = board.Credits.Select(row => new BoardDocument.CreditRow
                {
                    Category = row.Category,
                    SubCategory = row.SubCategory,
                    NameOrHandle = row.NameOrHandle,
                    Contact = row.Contact,
                }).ToList(),
                ImportantSignals = board.KiCadImportantSignals.Select(row => new BoardDocument.ImportantSignalRow
                {
                    DisplayName = row.DisplayName,
                    KiCadNetName = row.KiCadNetName,
                }).ToList(),
                Highlights = board.ComponentHighlights.Select(row => new BoardDocument.HighlightRow
                {
                    Schematic = row.SchematicName,
                    BoardLabel = row.BoardLabel,
                    X = BoardDocumentMapping.WholeNumber(row.X, row),
                    Y = BoardDocumentMapping.WholeNumber(row.Y, row),
                    Width = BoardDocumentMapping.WholeNumber(row.Width, row),
                    Height = BoardDocumentMapping.WholeNumber(row.Height, row),
                }).ToList(),
                KiCadCalibrations = content.Calibrations.Select(row => new BoardDocument.KiCadCalibrationRow
                {
                    Schematic = row.SchematicName,
                    CadName = row.CadName,
                    OffsetX = row.OffsetX,
                    OffsetY = row.OffsetY,
                    ScaleX = row.ScaleX,
                    ScaleY = row.ScaleY,
                    MirrorX = row.MirrorX,
                    MirrorY = row.MirrorY,
                }).ToList(),
                SchematicRenames = content.Renames.Select(row => new BoardDocument.SchematicRenameRow
                {
                    From = row.From,
                    To = row.To,
                    File = Name(row.File),
                    Date = row.Date,
                }).ToList(),
                Files = files.ToList(),
            };
        }

        public static BoardContent ToContent(BoardDocument document)
        {
            ArgumentNullException.ThrowIfNull(document);

            var board = new BoardData
            {
                HardwareName = document.HardwareName,
                BoardName = document.BoardName,
                RevisionDate = document.RevisionDate,
                Schematics = document.Schematics.Select(row => new BoardSchematicEntry
                {
                    SchematicName = row.Name,
                    SchematicImageFile = row.ImageFile,
                    SchematicHighlightColor = row.HighlightColor,
                    SchematicHighlightOpacity = row.HighlightOpacity,
                    OppositeTraceHighlightColor = row.OppositeTraceHighlightColor,
                    ThumbnailHighlightColor = row.ThumbnailHighlightColor,
                    ThumbnailHighlightOpacity = row.ThumbnailHighlightOpacity,
                    CadName = row.CadName,
                }).ToList(),
                Components = document.Components.Select(row => new ComponentEntry
                {
                    BoardLabel = row.BoardLabel,
                    FriendlyName = row.FriendlyName,
                    TechnicalNameOrValue = row.TechnicalNameOrValue,
                    PartNumber = row.PartNumber,
                    Category = row.Category,
                    Region = row.Region,
                    Description = row.Description,
                    Package = row.Package ?? string.Empty,
                    RepairNotes = row.RepairNotes ?? string.Empty,
                }).ToList(),
                ComponentImages = document.ComponentImages
                    .Select(row => new ComponentImageEntry
                    {
                        BoardLabel = row.BoardLabel,
                        Region = row.Region,
                        Pin = row.Pin,
                        Name = row.Name,
                        ExpectedOscilloscopeReading = row.ExpectedOscilloscopeReading,
                        File = row.File,
                        Note = row.Note,
                        TimeDiv = row.TimeDiv ?? string.Empty,
                        VoltsDiv = row.VoltsDiv ?? string.Empty,
                        TriggerLevelVolts = row.TriggerLevel ?? string.Empty,
                        MachineState = row.MachineState ?? string.Empty,
                        Source = row.Source ?? string.Empty,
                    })
                    .Concat(document.ScopeBaselines.Select(row => new ComponentImageEntry
                    {
                        BoardLabel = row.BoardLabel,
                        Region = row.Region,
                        Pin = row.Pin,
                        Name = row.Name,
                        ExpectedOscilloscopeReading = row.ExpectedOscilloscopeReading,
                        File = row.File,
                        Note = row.Note,
                        TimeDiv = row.TimeDiv,
                        VoltsDiv = row.VoltsDiv,
                        TriggerLevelVolts = row.TriggerLevel,
                        MachineState = row.MachineState ?? string.Empty,
                        Source = row.Source ?? string.Empty,
                    }))
                    .ToList(),
                ComponentHighlights = document.Highlights.Select(row => new ComponentHighlightEntry
                {
                    SchematicName = row.Schematic,
                    BoardLabel = row.BoardLabel,
                    X = row.X.ToString(CultureInfo.InvariantCulture),
                    Y = row.Y.ToString(CultureInfo.InvariantCulture),
                    Width = row.Width.ToString(CultureInfo.InvariantCulture),
                    Height = row.Height.ToString(CultureInfo.InvariantCulture),
                }).ToList(),
                ComponentLocalFiles = document.ComponentFiles.Select(row => new ComponentLocalFileEntry
                {
                    BoardLabel = row.BoardLabel,
                    Name = row.Name,
                    File = row.File,
                }).ToList(),
                ComponentLinks = document.ComponentLinks.Select(row => new ComponentLinkEntry
                {
                    BoardLabel = row.BoardLabel,
                    Name = row.Name,
                    Url = row.Url,
                }).ToList(),
                BoardLocalFiles = document.BoardFiles.Select(row => new BoardLocalFileEntry
                {
                    Category = row.Category,
                    Name = row.Name,
                    File = row.File,
                }).ToList(),
                BoardLinks = document.BoardLinks.Select(row => new BoardLinkEntry
                {
                    Category = row.Category,
                    Name = row.Name,
                    Url = row.Url,
                }).ToList(),
                Credits = document.Credits.Select(row => new CreditEntry
                {
                    Category = row.Category,
                    SubCategory = row.SubCategory,
                    NameOrHandle = row.NameOrHandle,
                    Contact = row.Contact,
                }).ToList(),
                KiCadImportantSignals = document.ImportantSignals.Select(row => new KiCadImportantSignalEntry
                {
                    DisplayName = row.DisplayName,
                    KiCadNetName = row.KiCadNetName,
                }).ToList(),
                ComponentPins = document.ComponentPins.Select(row => new ComponentPinEntry
                {
                    BoardLabel = row.BoardLabel,
                    Region = row.Region,
                    Pin = row.Pin,
                    Name = row.Name,
                    Direction = row.Direction,
                }).ToList(),
                ComponentReplacements = document.ComponentReplacements.Select(row => new ComponentReplacementEntry
                {
                    BoardLabel = row.BoardLabel,
                    Region = row.Region,
                    PartNumber = row.PartNumber,
                    Kind = row.Kind,
                    Changes = row.Changes,
                    Source = row.Source,
                }).ToList(),
            };

            IReadOnlyList<KiCadCalibrationEntry> calibrations = document.KiCadCalibrations.Select(row => new KiCadCalibrationEntry
            {
                SchematicName = row.Schematic,
                CadName = row.CadName,
                OffsetX = row.OffsetX,
                OffsetY = row.OffsetY,
                ScaleX = row.ScaleX,
                ScaleY = row.ScaleY,
                MirrorX = row.MirrorX,
                MirrorY = row.MirrorY,
            }).ToList();

            IReadOnlyList<SchematicRenameRecord> renames = document.SchematicRenames
                .Select(row => new SchematicRenameRecord(row.From, row.To, row.File, row.Date))
                .ToList();

            return new BoardContent(board, calibrations, renames);
        }

        // A picture's scope setting, written only when it has one - see BoardDocument.ComponentImageRow.
        private static string? WhenGiven(string? value) => string.IsNullOrEmpty(value) ? null : value;

        // A highlight's coordinate in whole pixels. The label editor hands over real geometry
        // ("123.4567" at any zoom), so a fraction is ROUNDED - away from zero, as BoardSidecarWriter
        // and BoardComponentHighlightStorage round it - never refused: refused, every rectangle drawn
        // or moved in a 4.0 draft failed to save (code review, 2026-10-26). Read invariant, always.
        // What is no number is refused, never a zero, which is a real place on the board.
        private static int WholeNumber(string? value, ComponentHighlightEntry highlight)
        {
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) &&
                double.IsFinite(number) && Math.Abs(number) <= int.MaxValue)
            {
                return (int)Math.Round(number, MidpointRounding.AwayFromZero);
            }

            throw new FormatException(
                $"The highlight of [{highlight.BoardLabel}] on [{highlight.SchematicName}] has the coordinate [{value}], which is not a number.");
        }
    }
}
