using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE ONE DEFINITION of a board workbook's shape: which sheets exist, what their columns are
    // called, which columns a sheet must carry to be recognised, and how one row of BoardData
    // becomes a row of cells.
    //
    // *** WHY THIS EXISTS AT ALL. *** BoardDataReader held all of this privately, which was
    // correct while reading was the only thing anyone did. Phase 5 adds a WRITER (publishing), and
    // a writer with its own copy of the column names is precisely the defect this project has
    // already been bitten by twice and written down both times: the old ComponentContributionPayload
    // kept the app's and the old server's idea of a shape in step BY HAND, and SubmissionContract.cs
    // exists because that failed. A workbook written with "Part number" and read with
    // "Part-number" produces a board that loads with every part number silently blank - nothing
    // throws, and the data is simply gone.
    //
    // So: BoardDataReader reads THROUGH this, BoardWorkbookWriter writes THROUGH this, and a
    // column renamed here changes both at once. Round-trip tests assert the pair, which is only
    // meaningful because they share this file.
    //
    // *** EVERY VALUE IS A STRING, DELIBERATELY. *** The reader takes each cell's .Text and does
    // no conversion: opacities, coordinates and trigger levels all arrive as text and are parsed
    // (invariant-culture) much later, where the meaning is known. The writer must match that - see
    // BoardWorkbookWriter's header for why writing a real number into these cells would corrupt
    // data under a comma-decimal locale.
    //
    // THE REQUIRED HEADERS ARE A SUBSET, NOT THE WHOLE ROW. The reader finds a sheet's header row
    // by scanning for the first row carrying ALL of the required headers, so those are the ones
    // that must never be absent or misspelled. The rest are read when present and blank when not,
    // which is what lets an older workbook lacking a newer column still load.
    // ###########################################################################################
    public static class BoardWorkbookSchema
    {
        // ---- Sheet names -------------------------------------------------------------------
        public const string SheetBoardSchematics = "Board schematics";
        public const string SheetComponents = "Components";
        public const string SheetComponentImages = "Component images";
        public const string SheetComponentLocalFiles = "Component local files";
        public const string SheetComponentLinks = "Component links";
        public const string SheetBoardLocalFiles = "Board local files";
        public const string SheetBoardLinks = "Board links";
        public const string SheetCredits = "Credits";
        public const string SheetKiCadImportantSignals = "Important signals";

        // ---- Board schematics columns ------------------------------------------------------
        public const string ColSchematicName = "Schematic name";
        public const string ColCadName = "CAD name";
        public const string ColSchematicImageFile = "Schematic image file";
        public const string ColSchematicHighlightColor = "Schematic highlight color";
        public const string ColSchematicHighlightOpacity = "Schematic highlight opacity";
        public const string ColOppositeTraceHighlightColor = "Opposite trace highlight color";
        public const string ColThumbnailHighlightColor = "Thumbnail highlight color";
        public const string ColThumbnailHighlightOpacity = "Thumbnail highlight opacity";

        // ---- Shared columns ----------------------------------------------------------------
        //
        public const string ColBoardLabel = "Board label";
        public const string ColCategory = "Category";
        public const string ColName = "Name";
        public const string ColFile = "File";
        public const string ColUrl = "URL";
        public const string ColRegion = "Region";

        // ---- Components columns ------------------------------------------------------------
        public const string ColFriendlyName = "Friendly name";
        public const string ColTechnicalNameOrValue = "Technical name or value";
        public const string ColPartNumber = "Part-number";

        // The parenthetical IS part of the header in every shipped workbook. It reads like a note
        // to the author and is not one - dropping it renames the column.
        public const string ColDescription = "Short one-liner description (one short line only!)";

        // ---- Component images columns ------------------------------------------------------
        public const string ColPin = "Pin";
        public const string ColExpectedOscilloscopeReading = "Expected oscilloscope reading";
        public const string ColNote = "Note";
        public const string ColTimeDiv = "T/DIV";
        public const string ColVoltsDiv = "V/DIV";
        public const string ColTriggerLevelVolts = "T.LVL";

        // ---- Credits columns ---------------------------------------------------------------
        public const string ColSubCategory = "Sub-category";
        public const string ColNameOrHandle = "Name or handle";
        public const string ColContact = "Contact (email or web page)";

        // ---- Important signals columns -----------------------------------------------------
        public const string ColDisplayName = "Display name";
        public const string ColKiCadNetName = "KiCad net name";

        // ###########################################################################################
        // The marker cell carrying the board's hand-maintained revision date, in the top-left
        // 10x10 of the schematics sheet. The reader scans for a cell whose text STARTS with this
        // prefix and takes the remainder; the writer emits prefix + value so the two agree by
        // construction.
        // ###########################################################################################
        public const string RevisionDateMarkerPrefix = "# Revision date:";

        // The other two preamble markers, named here beside the revision one so the writer and the
        // reader cannot spell them differently. Presentation only - see BoardData.HardwareName.
        public const string HardwareMarkerPrefix = "# Hardware:";

        public const string BoardMarkerPrefix = "# Board:";

        public static string BuildRevisionDateMarker(string? revisionDate) =>
            $"{BoardWorkbookSchema.RevisionDateMarkerPrefix} {(revisionDate ?? string.Empty).Trim()}".TrimEnd();

        // ###########################################################################################
        // One sheet's definition: its name, the columns it writes in order, and the subset the
        // reader requires in order to recognise a header row.
        //
        // ColumnOrder is the WRITE order and also the order a freshly written workbook presents to
        // a human opening it. The reader does not care - it maps by header name - so reordering
        // here cannot break reading, only appearance.
        // ###########################################################################################
        public sealed record SheetDefinition(
            string SheetName,
            IReadOnlyList<string> ColumnOrder,
            IReadOnlyList<string> RequiredHeaders);

        public static readonly SheetDefinition BoardSchematics = new(
            BoardWorkbookSchema.SheetBoardSchematics,
            [
                // ###########################################################################################
                // *** THIS ORDER MATCHES THE SHIPPED BOARDS, and CAD name sits LAST for a reason
                // (corrected 2026-09-24). ***
                //
                // It used to be written second, which put it between the schematic name and its
                // image file - nowhere near where any shipped board has it. The reference groups
                // the five HIGHLIGHT columns together under one heading and then puts CAD name
                // after them, because that column is filled in by the KiCad import rather than by
                // hand and is banded separately to say so.
                //
                // Reading is unaffected either way - the reader maps by header name - so this is
                // about a workbook a person opens matching the ones they already know.
                // ###########################################################################################
                BoardWorkbookSchema.ColSchematicName,
                BoardWorkbookSchema.ColSchematicImageFile,
                BoardWorkbookSchema.ColSchematicHighlightColor,
                BoardWorkbookSchema.ColSchematicHighlightOpacity,
                BoardWorkbookSchema.ColOppositeTraceHighlightColor,
                BoardWorkbookSchema.ColThumbnailHighlightColor,
                BoardWorkbookSchema.ColThumbnailHighlightOpacity,
                BoardWorkbookSchema.ColCadName
            ],
            [BoardWorkbookSchema.ColSchematicName, BoardWorkbookSchema.ColSchematicImageFile]);

        public static readonly SheetDefinition Components = new(
            BoardWorkbookSchema.SheetComponents,
            [
                BoardWorkbookSchema.ColBoardLabel,
                BoardWorkbookSchema.ColFriendlyName,
                BoardWorkbookSchema.ColTechnicalNameOrValue,
                BoardWorkbookSchema.ColPartNumber,
                BoardWorkbookSchema.ColCategory,
                BoardWorkbookSchema.ColRegion,
                BoardWorkbookSchema.ColDescription
            ],
            [
                BoardWorkbookSchema.ColBoardLabel,
                BoardWorkbookSchema.ColFriendlyName,
                BoardWorkbookSchema.ColTechnicalNameOrValue
            ]);

        public static readonly SheetDefinition ComponentImages = new(
            BoardWorkbookSchema.SheetComponentImages,
            [
                BoardWorkbookSchema.ColBoardLabel,
                BoardWorkbookSchema.ColRegion,
                BoardWorkbookSchema.ColPin,
                BoardWorkbookSchema.ColName,
                BoardWorkbookSchema.ColExpectedOscilloscopeReading,
                BoardWorkbookSchema.ColFile,
                BoardWorkbookSchema.ColNote,
                BoardWorkbookSchema.ColTimeDiv,
                BoardWorkbookSchema.ColVoltsDiv,
                BoardWorkbookSchema.ColTriggerLevelVolts
            ],
            [
                BoardWorkbookSchema.ColBoardLabel,
                BoardWorkbookSchema.ColPin,
                BoardWorkbookSchema.ColName,
                BoardWorkbookSchema.ColFile
            ]);

        public static readonly SheetDefinition ComponentLocalFiles = new(
            BoardWorkbookSchema.SheetComponentLocalFiles,
            [
                BoardWorkbookSchema.ColBoardLabel,
                BoardWorkbookSchema.ColName,
                BoardWorkbookSchema.ColFile
            ],
            [BoardWorkbookSchema.ColBoardLabel, BoardWorkbookSchema.ColName, BoardWorkbookSchema.ColFile]);

        public static readonly SheetDefinition ComponentLinks = new(
            BoardWorkbookSchema.SheetComponentLinks,
            [
                BoardWorkbookSchema.ColBoardLabel,
                BoardWorkbookSchema.ColName,
                BoardWorkbookSchema.ColUrl
            ],
            [BoardWorkbookSchema.ColBoardLabel, BoardWorkbookSchema.ColName, BoardWorkbookSchema.ColUrl]);

        public static readonly SheetDefinition BoardLocalFiles = new(
            BoardWorkbookSchema.SheetBoardLocalFiles,
            [
                BoardWorkbookSchema.ColCategory,
                BoardWorkbookSchema.ColName,
                BoardWorkbookSchema.ColFile
            ],
            [BoardWorkbookSchema.ColCategory, BoardWorkbookSchema.ColName, BoardWorkbookSchema.ColFile]);

        public static readonly SheetDefinition BoardLinks = new(
            BoardWorkbookSchema.SheetBoardLinks,
            [
                BoardWorkbookSchema.ColCategory,
                BoardWorkbookSchema.ColName,
                BoardWorkbookSchema.ColUrl
            ],
            [BoardWorkbookSchema.ColCategory, BoardWorkbookSchema.ColName, BoardWorkbookSchema.ColUrl]);

        public static readonly SheetDefinition Credits = new(
            BoardWorkbookSchema.SheetCredits,
            [
                BoardWorkbookSchema.ColCategory,
                BoardWorkbookSchema.ColSubCategory,
                BoardWorkbookSchema.ColNameOrHandle,
                BoardWorkbookSchema.ColContact
            ],
            [BoardWorkbookSchema.ColCategory, BoardWorkbookSchema.ColNameOrHandle]);

        public static readonly SheetDefinition KiCadImportantSignals = new(
            BoardWorkbookSchema.SheetKiCadImportantSignals,
            [BoardWorkbookSchema.ColDisplayName, BoardWorkbookSchema.ColKiCadNetName],
            [BoardWorkbookSchema.ColDisplayName, BoardWorkbookSchema.ColKiCadNetName]);

        // Write order across the workbook - and the order the Drafts tab's table editor shows the
        // sheets in. Schematics first because that sheet also carries the revision-date marker, so
        // a reader opening the file sees the board's identity first.
        //
        // *** "Credits" IS LAST, AS IN EVERY PUBLISHED WORKBOOK (2026-09-24). *** All thirteen
        // published boards with an "Important signals" sheet put it before "Credits"; this list had
        // them the other way round, so every workbook the app wrote ended in "Important signals"
        // and the table's sheet tabs disagreed with the file the project owner knows (reported).
        public static readonly IReadOnlyList<SheetDefinition> AllSheets =
        [
            BoardWorkbookSchema.BoardSchematics,
            BoardWorkbookSchema.Components,
            BoardWorkbookSchema.ComponentImages,
            BoardWorkbookSchema.ComponentLocalFiles,
            BoardWorkbookSchema.ComponentLinks,
            BoardWorkbookSchema.BoardLocalFiles,
            BoardWorkbookSchema.BoardLinks,
            BoardWorkbookSchema.KiCadImportantSignals,
            BoardWorkbookSchema.Credits
        ];

        // ###########################################################################################
        // One row of a sheet as column-name to cell-text. Built from BoardData's own entry types,
        // so the writer never reaches into an entry itself and a new field is added here once.
        //
        // Null is normalised to empty: the reader's Val() answers "" for an absent column, so a
        // round trip has to produce "" rather than null or the two disagree on a blank cell.
        // ###########################################################################################
        public static IReadOnlyList<IReadOnlyDictionary<string, string>> BuildRows(
            SheetDefinition sheet,
            BoardData data)
        {
            ArgumentNullException.ThrowIfNull(sheet);

            if (data == null)
                return [];

            if (sheet.SheetName == BoardWorkbookSchema.SheetBoardSchematics)
            {
                return [.. data.Schematics.Select(entry => Row(
                    (BoardWorkbookSchema.ColSchematicName, entry.SchematicName),
                    (BoardWorkbookSchema.ColCadName, entry.CadName),
                    (BoardWorkbookSchema.ColSchematicImageFile, entry.SchematicImageFile),
                    (BoardWorkbookSchema.ColSchematicHighlightColor, entry.SchematicHighlightColor),
                    (BoardWorkbookSchema.ColSchematicHighlightOpacity, entry.SchematicHighlightOpacity),
                    (BoardWorkbookSchema.ColOppositeTraceHighlightColor, entry.OppositeTraceHighlightColor),
                    (BoardWorkbookSchema.ColThumbnailHighlightColor, entry.ThumbnailHighlightColor),
                    (BoardWorkbookSchema.ColThumbnailHighlightOpacity, entry.ThumbnailHighlightOpacity)))];
            }

            if (sheet.SheetName == BoardWorkbookSchema.SheetComponents)
            {
                return [.. data.Components.Select(entry => Row(
                    (BoardWorkbookSchema.ColBoardLabel, entry.BoardLabel),
                    (BoardWorkbookSchema.ColFriendlyName, entry.FriendlyName),
                    (BoardWorkbookSchema.ColTechnicalNameOrValue, entry.TechnicalNameOrValue),
                    (BoardWorkbookSchema.ColPartNumber, entry.PartNumber),
                    (BoardWorkbookSchema.ColCategory, entry.Category),
                    (BoardWorkbookSchema.ColRegion, entry.Region),
                    (BoardWorkbookSchema.ColDescription, entry.Description)))];
            }

            if (sheet.SheetName == BoardWorkbookSchema.SheetComponentImages)
            {
                return [.. data.ComponentImages.Select(entry => Row(
                    (BoardWorkbookSchema.ColBoardLabel, entry.BoardLabel),
                    (BoardWorkbookSchema.ColRegion, entry.Region),
                    (BoardWorkbookSchema.ColPin, entry.Pin),
                    (BoardWorkbookSchema.ColName, entry.Name),
                    (BoardWorkbookSchema.ColExpectedOscilloscopeReading, entry.ExpectedOscilloscopeReading),
                    (BoardWorkbookSchema.ColFile, entry.File),
                    (BoardWorkbookSchema.ColNote, entry.Note),
                    (BoardWorkbookSchema.ColTimeDiv, entry.TimeDiv),
                    (BoardWorkbookSchema.ColVoltsDiv, entry.VoltsDiv),
                    (BoardWorkbookSchema.ColTriggerLevelVolts, entry.TriggerLevelVolts)))];
            }

            if (sheet.SheetName == BoardWorkbookSchema.SheetComponentLocalFiles)
            {
                return [.. data.ComponentLocalFiles.Select(entry => Row(
                    (BoardWorkbookSchema.ColBoardLabel, entry.BoardLabel),
                    (BoardWorkbookSchema.ColName, entry.Name),
                    (BoardWorkbookSchema.ColFile, entry.File)))];
            }

            if (sheet.SheetName == BoardWorkbookSchema.SheetComponentLinks)
            {
                return [.. data.ComponentLinks.Select(entry => Row(
                    (BoardWorkbookSchema.ColBoardLabel, entry.BoardLabel),
                    (BoardWorkbookSchema.ColName, entry.Name),
                    (BoardWorkbookSchema.ColUrl, entry.Url)))];
            }

            if (sheet.SheetName == BoardWorkbookSchema.SheetBoardLocalFiles)
            {
                return [.. data.BoardLocalFiles.Select(entry => Row(
                    (BoardWorkbookSchema.ColCategory, entry.Category),
                    (BoardWorkbookSchema.ColName, entry.Name),
                    (BoardWorkbookSchema.ColFile, entry.File)))];
            }

            if (sheet.SheetName == BoardWorkbookSchema.SheetBoardLinks)
            {
                return [.. data.BoardLinks.Select(entry => Row(
                    (BoardWorkbookSchema.ColCategory, entry.Category),
                    (BoardWorkbookSchema.ColName, entry.Name),
                    (BoardWorkbookSchema.ColUrl, entry.Url)))];
            }

            if (sheet.SheetName == BoardWorkbookSchema.SheetCredits)
            {
                return [.. data.Credits.Select(entry => Row(
                    (BoardWorkbookSchema.ColCategory, entry.Category),
                    (BoardWorkbookSchema.ColSubCategory, entry.SubCategory),
                    (BoardWorkbookSchema.ColNameOrHandle, entry.NameOrHandle),
                    (BoardWorkbookSchema.ColContact, entry.Contact)))];
            }

            if (sheet.SheetName == BoardWorkbookSchema.SheetKiCadImportantSignals)
            {
                return [.. data.KiCadImportantSignals.Select(entry => Row(
                    (BoardWorkbookSchema.ColDisplayName, entry.DisplayName),
                    (BoardWorkbookSchema.ColKiCadNetName, entry.KiCadNetName)))];
            }

            // An unknown sheet is a programming error here, not contributed data: AllSheets is the
            // only source of SheetDefinitions. Throwing beats returning empty, which would write a
            // header-only sheet and silently drop a whole section.
            throw new ArgumentOutOfRangeException(
                nameof(sheet),
                sheet.SheetName,
                "No row builder for this sheet - add one alongside its SheetDefinition.");
        }

        private static Dictionary<string, string> Row(params (string Column, string? Value)[] cells)
        {
            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach ((string column, string? value) in cells)
            {
                row[column] = value ?? string.Empty;
            }

            return row;
        }

        // ###########################################################################################
        // THE INVERSE OF BuildRows: rows of column-name-to-cell-text back into BoardData's entry
        // types (moved here from BoardDataReader, 2026-09-24).
        //
        // *** WHY THE READER'S MAPPING MOVED. *** It was private to BoardDataReader, which was fine
        // while a workbook was the only thing rows came from. The Drafts tab's table editor
        // (BoardTableDocument) now produces rows too, from cells a contributor typed, and turning
        // those into a BoardData with a second, hand-copied mapping is the exact two-copies defect
        // this file's header describes. So the reader reads THROUGH these, the table editor saves
        // THROUGH these, and BuildRows above is the other direction - one definition each way.
        //
        // A missing column reads as "" rather than throwing, because an older workbook lacking a
        // newer column must still load (see "THE REQUIRED HEADERS ARE A SUBSET" above). Lookups are
        // by whatever comparer the row dictionary carries; the reader's and the table's are both
        // OrdinalIgnoreCase.
        // ###########################################################################################
        private static string Val(IReadOnlyDictionary<string, string> row, string key)
            => row.TryGetValue(key, out var value) ? value ?? string.Empty : string.Empty;

        // KiCad calibration is not read here: it lives in the board's JSON sidecar, not the
        // workbook. Opposite trace highlight OPACITY is not a column either - the renderer follows
        // the main schematic highlight opacity for it.
        public static List<BoardSchematicEntry> MapSchematics(IEnumerable<IReadOnlyDictionary<string, string>> rows)
            => rows.Select(r => new BoardSchematicEntry
            {
                SchematicName = Val(r, BoardWorkbookSchema.ColSchematicName),
                CadName = Val(r, BoardWorkbookSchema.ColCadName),
                SchematicImageFile = Val(r, BoardWorkbookSchema.ColSchematicImageFile),
                SchematicHighlightColor = Val(r, BoardWorkbookSchema.ColSchematicHighlightColor),
                SchematicHighlightOpacity = Val(r, BoardWorkbookSchema.ColSchematicHighlightOpacity),
                OppositeTraceHighlightColor = Val(r, BoardWorkbookSchema.ColOppositeTraceHighlightColor),
                ThumbnailHighlightColor = Val(r, BoardWorkbookSchema.ColThumbnailHighlightColor),
                ThumbnailHighlightOpacity = Val(r, BoardWorkbookSchema.ColThumbnailHighlightOpacity)
            }).ToList();

        public static List<ComponentEntry> MapComponents(IEnumerable<IReadOnlyDictionary<string, string>> rows)
            => rows.Select(r => new ComponentEntry
            {
                BoardLabel = Val(r, BoardWorkbookSchema.ColBoardLabel),
                FriendlyName = Val(r, BoardWorkbookSchema.ColFriendlyName),
                TechnicalNameOrValue = Val(r, BoardWorkbookSchema.ColTechnicalNameOrValue),
                PartNumber = Val(r, BoardWorkbookSchema.ColPartNumber),
                Category = Val(r, BoardWorkbookSchema.ColCategory),
                Region = Val(r, BoardWorkbookSchema.ColRegion),
                Description = Val(r, BoardWorkbookSchema.ColDescription)
            }).ToList();

        public static List<ComponentImageEntry> MapComponentImages(IEnumerable<IReadOnlyDictionary<string, string>> rows)
            => rows.Select(r => new ComponentImageEntry
            {
                BoardLabel = Val(r, BoardWorkbookSchema.ColBoardLabel),
                Region = Val(r, BoardWorkbookSchema.ColRegion),
                Pin = Val(r, BoardWorkbookSchema.ColPin),
                Name = Val(r, BoardWorkbookSchema.ColName),
                ExpectedOscilloscopeReading = Val(r, BoardWorkbookSchema.ColExpectedOscilloscopeReading),
                File = Val(r, BoardWorkbookSchema.ColFile),
                Note = Val(r, BoardWorkbookSchema.ColNote),
                TimeDiv = Val(r, BoardWorkbookSchema.ColTimeDiv),
                VoltsDiv = Val(r, BoardWorkbookSchema.ColVoltsDiv),
                TriggerLevelVolts = Val(r, BoardWorkbookSchema.ColTriggerLevelVolts)
            }).ToList();

        public static List<ComponentLocalFileEntry> MapComponentLocalFiles(IEnumerable<IReadOnlyDictionary<string, string>> rows)
            => rows.Select(r => new ComponentLocalFileEntry
            {
                BoardLabel = Val(r, BoardWorkbookSchema.ColBoardLabel),
                Name = Val(r, BoardWorkbookSchema.ColName),
                File = Val(r, BoardWorkbookSchema.ColFile)
            }).ToList();

        public static List<ComponentLinkEntry> MapComponentLinks(IEnumerable<IReadOnlyDictionary<string, string>> rows)
            => rows.Select(r => new ComponentLinkEntry
            {
                BoardLabel = Val(r, BoardWorkbookSchema.ColBoardLabel),
                Name = Val(r, BoardWorkbookSchema.ColName),
                Url = Val(r, BoardWorkbookSchema.ColUrl)
            }).ToList();

        public static List<BoardLocalFileEntry> MapBoardLocalFiles(IEnumerable<IReadOnlyDictionary<string, string>> rows)
            => rows.Select(r => new BoardLocalFileEntry
            {
                Category = Val(r, BoardWorkbookSchema.ColCategory),
                Name = Val(r, BoardWorkbookSchema.ColName),
                File = Val(r, BoardWorkbookSchema.ColFile)
            }).ToList();

        public static List<BoardLinkEntry> MapBoardLinks(IEnumerable<IReadOnlyDictionary<string, string>> rows)
            => rows.Select(r => new BoardLinkEntry
            {
                Category = Val(r, BoardWorkbookSchema.ColCategory),
                Name = Val(r, BoardWorkbookSchema.ColName),
                Url = Val(r, BoardWorkbookSchema.ColUrl)
            }).ToList();

        public static List<CreditEntry> MapCredits(IEnumerable<IReadOnlyDictionary<string, string>> rows)
            => rows.Select(r => new CreditEntry
            {
                Category = Val(r, BoardWorkbookSchema.ColCategory),
                SubCategory = Val(r, BoardWorkbookSchema.ColSubCategory),
                NameOrHandle = Val(r, BoardWorkbookSchema.ColNameOrHandle),
                Contact = Val(r, BoardWorkbookSchema.ColContact)
            }).ToList();

        // ###########################################################################################
        // The one mapper that DROPS rows: an important signal needs both halves to mean anything,
        // so a row missing either is skipped rather than loaded half-empty. The table editor relies
        // on this being the only such rule - it marks a row this would drop as "incomplete" by
        // asking MapRows whether the row survives, rather than restating the condition.
        // ###########################################################################################
        //
        // RequiredColumns below names the same two columns, for the table's warning on a row this
        // would drop ("This row is left out when saving, because KiCad net name is empty",
        // 2026-10-03) - BoardWorkbookSchemaTests holds the two to each other.
        public static List<KiCadImportantSignalEntry> MapKiCadImportantSignals(IEnumerable<IReadOnlyDictionary<string, string>> rows)
            => rows
                .Select(r => new KiCadImportantSignalEntry
                {
                    DisplayName = Val(r, BoardWorkbookSchema.ColDisplayName),
                    KiCadNetName = Val(r, BoardWorkbookSchema.ColKiCadNetName)
                })
                .Where(entry =>
                    !string.IsNullOrWhiteSpace(entry.DisplayName) &&
                    !string.IsNullOrWhiteSpace(entry.KiCadNetName))
                .ToList();

        // The columns a row of this sheet cannot be saved without - empty for every sheet but the
        // one above, whose mapper drops a row missing either.
        public static IReadOnlyList<string> RequiredColumns(string sheetName) =>
            sheetName == BoardWorkbookSchema.SheetKiCadImportantSignals
                ? [BoardWorkbookSchema.ColDisplayName, BoardWorkbookSchema.ColKiCadNetName]
                : [];

        // ###########################################################################################
        // One sheet's rows mapped to its entry type, dispatched on the sheet - for a caller that
        // works with sheets generically (the table editor) and needs each row's natural key, which
        // BoardDraftNaturalKeys builds from an ENTRY, not from cells.
        //
        // Throws for an unknown sheet, for the same reason BuildRows does.
        // ###########################################################################################
        public static IReadOnlyList<object> MapRows(
            SheetDefinition sheet,
            IEnumerable<IReadOnlyDictionary<string, string>> rows)
        {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(rows);

            return sheet.SheetName switch
            {
                BoardWorkbookSchema.SheetBoardSchematics => BoardWorkbookSchema.MapSchematics(rows),
                BoardWorkbookSchema.SheetComponents => BoardWorkbookSchema.MapComponents(rows),
                BoardWorkbookSchema.SheetComponentImages => BoardWorkbookSchema.MapComponentImages(rows),
                BoardWorkbookSchema.SheetComponentLocalFiles => BoardWorkbookSchema.MapComponentLocalFiles(rows),
                BoardWorkbookSchema.SheetComponentLinks => BoardWorkbookSchema.MapComponentLinks(rows),
                BoardWorkbookSchema.SheetBoardLocalFiles => BoardWorkbookSchema.MapBoardLocalFiles(rows),
                BoardWorkbookSchema.SheetBoardLinks => BoardWorkbookSchema.MapBoardLinks(rows),
                BoardWorkbookSchema.SheetCredits => BoardWorkbookSchema.MapCredits(rows),
                BoardWorkbookSchema.SheetKiCadImportantSignals => BoardWorkbookSchema.MapKiCadImportantSignals(rows),
                _ => throw new ArgumentOutOfRangeException(
                    nameof(sheet),
                    sheet.SheetName,
                    "No row mapper for this sheet - add one alongside its SheetDefinition.")
            };
        }

        // ###########################################################################################
        // The BoardData section a sheet holds, as the entries themselves - the entry-side twin of
        // BuildRows, in the same order, so index N here is row N there.
        // ###########################################################################################
        public static IReadOnlyList<object> EntriesOf(SheetDefinition sheet, BoardData data)
        {
            ArgumentNullException.ThrowIfNull(sheet);

            if (data == null)
                return [];

            return sheet.SheetName switch
            {
                BoardWorkbookSchema.SheetBoardSchematics => data.Schematics,
                BoardWorkbookSchema.SheetComponents => data.Components,
                BoardWorkbookSchema.SheetComponentImages => data.ComponentImages,
                BoardWorkbookSchema.SheetComponentLocalFiles => data.ComponentLocalFiles,
                BoardWorkbookSchema.SheetComponentLinks => data.ComponentLinks,
                BoardWorkbookSchema.SheetBoardLocalFiles => data.BoardLocalFiles,
                BoardWorkbookSchema.SheetBoardLinks => data.BoardLinks,
                BoardWorkbookSchema.SheetCredits => data.Credits,
                BoardWorkbookSchema.SheetKiCadImportantSignals => data.KiCadImportantSignals,
                _ => throw new ArgumentOutOfRangeException(
                    nameof(sheet),
                    sheet.SheetName,
                    "No section for this sheet - add one alongside its SheetDefinition.")
            };
        }

        // ###########################################################################################
        // The same board with ONE sheet's section replaced by the given rows.
        //
        // Everything else is carried across untouched - the other sections, the revision date, the
        // preamble names and the highlights, which live in the JSON sidecar rather than any sheet.
        // The lists are shared rather than copied, on BoardData.WithRevisionDate's reasoning: every
        // writer here builds a new BoardData instead of mutating one.
        // ###########################################################################################
        public static BoardData WithRows(
            BoardData data,
            SheetDefinition sheet,
            IEnumerable<IReadOnlyDictionary<string, string>> rows)
        {
            ArgumentNullException.ThrowIfNull(data);
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(rows);

            string name = sheet.SheetName;

            if (!BoardWorkbookSchema.AllSheets.Any(known => known.SheetName == name))
            {
                throw new ArgumentOutOfRangeException(
                    nameof(sheet),
                    name,
                    "No section for this sheet - add one alongside its SheetDefinition.");
            }

            var materialised = rows.ToList();

            return new BoardData
            {
                RevisionDate = data.RevisionDate,
                HardwareName = data.HardwareName,
                BoardName = data.BoardName,
                Schematics = name == BoardWorkbookSchema.SheetBoardSchematics
                    ? BoardWorkbookSchema.MapSchematics(materialised)
                    : data.Schematics,
                Components = name == BoardWorkbookSchema.SheetComponents
                    ? BoardWorkbookSchema.MapComponents(materialised)
                    : data.Components,
                ComponentImages = name == BoardWorkbookSchema.SheetComponentImages
                    ? BoardWorkbookSchema.MapComponentImages(materialised)
                    : data.ComponentImages,
                ComponentHighlights = data.ComponentHighlights,
                ComponentLocalFiles = name == BoardWorkbookSchema.SheetComponentLocalFiles
                    ? BoardWorkbookSchema.MapComponentLocalFiles(materialised)
                    : data.ComponentLocalFiles,
                ComponentLinks = name == BoardWorkbookSchema.SheetComponentLinks
                    ? BoardWorkbookSchema.MapComponentLinks(materialised)
                    : data.ComponentLinks,
                BoardLocalFiles = name == BoardWorkbookSchema.SheetBoardLocalFiles
                    ? BoardWorkbookSchema.MapBoardLocalFiles(materialised)
                    : data.BoardLocalFiles,
                BoardLinks = name == BoardWorkbookSchema.SheetBoardLinks
                    ? BoardWorkbookSchema.MapBoardLinks(materialised)
                    : data.BoardLinks,
                Credits = name == BoardWorkbookSchema.SheetCredits
                    ? BoardWorkbookSchema.MapCredits(materialised)
                    : data.Credits,
                KiCadImportantSignals = name == BoardWorkbookSchema.SheetKiCadImportantSignals
                    ? BoardWorkbookSchema.MapKiCadImportantSignals(materialised)
                    : data.KiCadImportantSignals,
            };
        }
    }
}
