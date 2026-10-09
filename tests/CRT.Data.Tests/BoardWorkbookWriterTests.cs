using System;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Covers BoardWorkbookWriter - the publishing step's board workbook writer (Phase 5, task 6).
//
// THE CENTRAL TEST HERE IS THE ROUND TRIP: write a BoardData, read it back with the real
// BoardDataReader, and assert every field survived. That is the writer's entire contract, and it
// is only meaningful because both sides now go through BoardWorkbookSchema - before that they
// were two hand-maintained copies of the same column names, which is the defect
// SubmissionContract.cs exists to prevent on the submission side.
//
// The failure this guards against is SILENT. A misspelled header does not throw: the reader finds
// no such column, answers "" for it, and the board loads with that field blank on every row. So
// these tests assert on VALUES read back, never on the file's bytes or its internal layout.
//
// Shares the "BoardData" collection because BoardDataReader caches loaded boards in shared static
// state - see BoardDataReaderTests' own header.
[Collection("BoardData")]
public sealed class BoardWorkbookWriterTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();
    private readonly List<string> thisCacheKeys = new();

    public void Dispose()
    {
        foreach (string key in this.thisCacheKeys)
        {
            BoardDataReader.ClearCache(key);
        }

        this.thisWorkspace.Dispose();
    }

    private string NewCacheKey()
    {
        string key = "writer-" + Guid.NewGuid().ToString("N");
        this.thisCacheKeys.Add(key);
        return key;
    }

    private Task<BoardData?> LoadAsync(string excelPath) =>
        BoardDataReader.LoadAsync(excelPath, this.NewCacheKey());

    // A board exercising every sheet and every column, including the awkward values: a
    // numeric-looking opacity, a negative trigger level, an embedded comma, and a blank field.
    // ###########################################################################################
    // The header row of one sheet, read straight out of the .xlsx.
    //
    // Needed because a read-back through BoardDataReader cannot tell "the column is absent" from
    // "the column is present and ignored" - the reader simply has no property for it any more.
    // Asserting a column is GONE has to look at the file.
    // ###########################################################################################
    private static IReadOnlyList<string> HeaderRowOf(string excelPath, string sheetName)
    {
        EpplusLicense.Ensure();

        using var package = new OfficeOpenXml.ExcelPackage(new FileInfo(excelPath));
        OfficeOpenXml.ExcelWorksheet sheet = package.Workbook.Worksheets[sheetName];

        var headers = new List<string>();
        if (sheet.Dimension is null)
        {
            return headers;
        }

        // Row 2 throughout - row 1 carries the revision-date marker. See BoardWorkbookWriter.
        for (int column = 1; column <= sheet.Dimension.End.Column; column++)
        {
            headers.Add(sheet.Cells[2, column].Text ?? string.Empty);
        }

        return headers;
    }

    private static BoardData BuildFullBoard() => new()
    {
        RevisionDate = "2026-August-21",
        Schematics =
        [
            new BoardSchematicEntry
            {
                SchematicName = "Sheet 1",
                CadName = "CAD-1",
                SchematicImageFile = "Images/sheet1.png",
                SchematicHighlightColor = "#FF0000",
                SchematicHighlightOpacity = "0.35",
                OppositeTraceHighlightColor = "#00FF00",
                ThumbnailHighlightColor = "#0000FF",
                ThumbnailHighlightOpacity = "0.5"
            }
        ],
        Components =
        [
            new ComponentEntry
            {
                BoardLabel = "U8",
                FriendlyName = "PLA",
                TechnicalNameOrValue = "906114-01",
                PartNumber = "251715-01",
                Category = "IC",
                Region = "ASSY 250407",
                Description = "Programmable logic array, decodes memory map"
            },
            new ComponentEntry
            {
                // Deliberately sparse: only the required columns. A row like this must not pick
                // up stray text from the row above it.
                BoardLabel = "R1",
                FriendlyName = "Resistor",
                TechnicalNameOrValue = "1k"
            }
        ],
        ComponentImages =
        [
            new ComponentImageEntry
            {
                BoardLabel = "U8",
                Region = "ASSY 250407",
                Pin = "1",
                Name = "Pin 1 baseline",
                ExpectedOscilloscopeReading = "5V square, 1 MHz",
                File = "Images/u8-pin1.png",
                Note = "Measured at room temperature",
                TimeDiv = "0.5",
                VoltsDiv = "2",
                TriggerLevelVolts = "-1.2"
            }
        ],
        ComponentLocalFiles =
        [
            new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "Files/pla.pdf" }
        ],
        ComponentLinks =
        [
            new ComponentLinkEntry { BoardLabel = "U8", Name = "Reference", Url = "https://example.com/pla" }
        ],
        BoardLocalFiles =
        [
            new BoardLocalFileEntry { Category = "Manuals", Name = "Service manual", File = "Files/service.pdf" }
        ],
        BoardLinks =
        [
            new BoardLinkEntry { Category = "Community", Name = "Forum", Url = "https://example.com/forum" }
        ],
        Credits =
        [
            new CreditEntry
            {
                Category = "Photography",
                SubCategory = "Board scans",
                NameOrHandle = "Someone",
                Contact = "someone@example.com"
            }
        ],
        KiCadImportantSignals =
        [
            new KiCadImportantSignalEntry { DisplayName = "Clock", KiCadNetName = "/CLK" }
        ]
    };

    // -----------------------------------------------------------------------------------
    // The round trip
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** THE WORKBOOK IS REPLACED, NEVER WRITTEN IN PLACE (owner report, 2026-09-28). *** Written
    // complete beside the target and renamed over it, so a reader mid-sync never sees half a
    // workbook and no temporary is left behind - including the one MakeDeterministic uses.
    // ###########################################################################################
    [Fact]
    public async Task Rewriting_a_workbook_replaces_it_and_leaves_no_temporary_file()
    {
        string path = Path.Combine(this.thisWorkspace.Root, "Data Test Board v2.0.0.xlsx");

        BoardWorkbookWriter.Write(path, BoardWorkbookWriterTests.BuildFullBoard());
        BoardWorkbookWriter.Write(path, BoardWorkbookWriterTests.BuildFullBoard().WithRevisionDate("2026-September-28"));

        Assert.Equal([path], Directory.GetFiles(this.thisWorkspace.Root));
        Assert.Equal("2026-September-28", (await this.LoadAsync(path))!.RevisionDate);
    }

    // A workbook copied into the tree by hand as root may be replaced but not opened for writing.
    // Made here as a file its owner may not write; not on Windows (no such mode) or as root (which
    // ignores it).
    [Fact]
    public async Task A_workbook_that_may_not_be_opened_for_writing_is_still_replaced()
    {
        // The return is for the platform analyzer, which cannot see that Skip throws.
        if (OperatingSystem.IsWindows())
        {
            Assert.Skip("Unix permissions only.");
            return;
        }
        Assert.SkipWhen(Environment.UserName == "root", "root may write anything.");

        string path = Path.Combine(this.thisWorkspace.Root, "Data Test Board v2.0.0.xlsx");
        File.WriteAllText(path, "copied in by hand");
        File.SetUnixFileMode(path, UnixFileMode.UserRead | UnixFileMode.GroupRead | UnixFileMode.OtherRead);

        BoardWorkbookWriter.Write(path, BoardWorkbookWriterTests.BuildFullBoard());

        Assert.Equal("2026-August-21", (await this.LoadAsync(path))!.RevisionDate);
    }

    [Fact]
    public async Task A_written_board_reads_back_with_every_field_intact()
    {
        string path = Path.Combine(this.thisWorkspace.Root, "Data Test Board v2.0.0.xlsx");
        BoardData original = BoardWorkbookWriterTests.BuildFullBoard();

        BoardWorkbookWriter.Write(path, original);
        BoardData? read = await this.LoadAsync(path);

        Assert.NotNull(read);

        // The revision date lives in a marker cell, not a column - its own mechanism, so worth
        // asserting explicitly rather than trusting the sheet comparison below.
        Assert.Equal("2026-August-21", read!.RevisionDate);

        BoardSchematicEntry schematic = Assert.Single(read.Schematics);
        Assert.Equal("Sheet 1", schematic.SchematicName);
        Assert.Equal("CAD-1", schematic.CadName);
        Assert.Equal("Images/sheet1.png", schematic.SchematicImageFile);
        Assert.Equal("#FF0000", schematic.SchematicHighlightColor);
        Assert.Equal("0.35", schematic.SchematicHighlightOpacity);
        Assert.Equal("#00FF00", schematic.OppositeTraceHighlightColor);
        Assert.Equal("#0000FF", schematic.ThumbnailHighlightColor);
        Assert.Equal("0.5", schematic.ThumbnailHighlightOpacity);

        Assert.Equal(2, read.Components.Count);
        ComponentEntry u8 = read.Components[0];
        Assert.Equal("U8", u8.BoardLabel);
        Assert.Equal("PLA", u8.FriendlyName);
        Assert.Equal("906114-01", u8.TechnicalNameOrValue);
        Assert.Equal("251715-01", u8.PartNumber);
        Assert.Equal("IC", u8.Category);
        Assert.Equal("ASSY 250407", u8.Region);
        Assert.Equal("Programmable logic array, decodes memory map", u8.Description);

        ComponentImageEntry image = Assert.Single(read.ComponentImages);
        Assert.Equal("U8", image.BoardLabel);
        Assert.Equal("ASSY 250407", image.Region);
        Assert.Equal("1", image.Pin);
        Assert.Equal("Pin 1 baseline", image.Name);
        Assert.Equal("5V square, 1 MHz", image.ExpectedOscilloscopeReading);
        Assert.Equal("Images/u8-pin1.png", image.File);
        Assert.Equal("Measured at room temperature", image.Note);
        Assert.Equal("0.5", image.TimeDiv);
        Assert.Equal("2", image.VoltsDiv);
        Assert.Equal("-1.2", image.TriggerLevelVolts);

        ComponentLocalFileEntry componentFile = Assert.Single(read.ComponentLocalFiles);
        Assert.Equal("U8", componentFile.BoardLabel);
        Assert.Equal("Datasheet", componentFile.Name);
        Assert.Equal("Files/pla.pdf", componentFile.File);

        ComponentLinkEntry componentLink = Assert.Single(read.ComponentLinks);
        Assert.Equal("U8", componentLink.BoardLabel);
        Assert.Equal("Reference", componentLink.Name);
        Assert.Equal("https://example.com/pla", componentLink.Url);

        BoardLocalFileEntry boardFile = Assert.Single(read.BoardLocalFiles);
        Assert.Equal("Manuals", boardFile.Category);
        Assert.Equal("Service manual", boardFile.Name);
        Assert.Equal("Files/service.pdf", boardFile.File);

        BoardLinkEntry boardLink = Assert.Single(read.BoardLinks);
        Assert.Equal("Community", boardLink.Category);
        Assert.Equal("Forum", boardLink.Name);
        Assert.Equal("https://example.com/forum", boardLink.Url);

        CreditEntry credit = Assert.Single(read.Credits);
        Assert.Equal("Photography", credit.Category);
        Assert.Equal("Board scans", credit.SubCategory);
        Assert.Equal("Someone", credit.NameOrHandle);
        Assert.Equal("someone@example.com", credit.Contact);

        KiCadImportantSignalEntry signal = Assert.Single(read.KiCadImportantSignals);
        Assert.Equal("Clock", signal.DisplayName);
        Assert.Equal("/CLK", signal.KiCadNetName);
    }

    [Fact]
    public async Task A_sparse_row_reads_back_blank_rather_than_inheriting_its_neighbour()
    {
        // A row carrying only the required columns sits below a fully populated one. If the
        // writer skipped blanks by shifting cells left, or the reader mapped by position, this
        // row would come back wearing the previous row's part number and description.
        string path = Path.Combine(this.thisWorkspace.Root, "sparse.xlsx");

        BoardWorkbookWriter.Write(path, BoardWorkbookWriterTests.BuildFullBoard());
        BoardData? read = await this.LoadAsync(path);

        ComponentEntry sparse = read!.Components[1];
        Assert.Equal("R1", sparse.BoardLabel);
        Assert.Equal("Resistor", sparse.FriendlyName);
        Assert.Equal("1k", sparse.TechnicalNameOrValue);
        Assert.Equal(string.Empty, sparse.PartNumber);
        Assert.Equal(string.Empty, sparse.Category);
        Assert.Equal(string.Empty, sparse.Region);
        Assert.Equal(string.Empty, sparse.Description);
    }

    // -----------------------------------------------------------------------------------
    // Numeric-looking text - the locale trap
    // -----------------------------------------------------------------------------------

    [Fact]
    public async Task A_decimal_value_survives_a_comma_decimal_culture()
    {
        // *** THE BUG THIS EXISTS FOR. *** BoardData holds "0.35" as TEXT. If the writer let
        // Excel store it as a NUMBER, the reader's .Text would format it with the CURRENT
        // culture, giving "0,35" here - and the invariant-culture parse that reads an opacity
        // later would then fail or read it as 35. Nothing would throw; highlights would just be
        // wrong. Written as text, the string is returned exactly as supplied.
        CultureInfo original = CultureInfo.CurrentCulture;

        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("da-DK");

            string path = Path.Combine(this.thisWorkspace.Root, "culture.xlsx");
            BoardWorkbookWriter.Write(path, BoardWorkbookWriterTests.BuildFullBoard());

            BoardData? read = await this.LoadAsync(path);

            Assert.Equal("0.35", read!.Schematics[0].SchematicHighlightOpacity);
            Assert.Equal("-1.2", read.ComponentImages[0].TriggerLevelVolts);
            Assert.Equal("0.5", read.ComponentImages[0].TimeDiv);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Fact]
    public async Task A_value_that_looks_like_a_date_is_not_reformatted()
    {
        // Excel's eagerness to interpret text is not limited to decimals. A part number like
        // "1-2-3", or a revision date, must come back as typed.
        string path = Path.Combine(this.thisWorkspace.Root, "dates.xlsx");

        var board = new BoardData
        {
            RevisionDate = "2026-August-21",
            Components =
            [
                new ComponentEntry
                {
                    BoardLabel = "U1",
                    FriendlyName = "Thing",
                    TechnicalNameOrValue = "1-2-3",
                    PartNumber = "01/02/2026"
                }
            ]
        };

        BoardWorkbookWriter.Write(path, board);
        BoardData? read = await this.LoadAsync(path);

        Assert.Equal("1-2-3", read!.Components[0].TechnicalNameOrValue);
        Assert.Equal("01/02/2026", read.Components[0].PartNumber);
        Assert.Equal("2026-August-21", read.RevisionDate);
    }

    // -----------------------------------------------------------------------------------
    // Edges
    // -----------------------------------------------------------------------------------

    [Fact]
    public async Task An_empty_board_writes_and_reads_back_empty_without_error()
    {
        // A brand-new board registered with no rows yet is a real state - "Add a new board"
        // creates exactly this. It must produce a valid workbook, not a malformed one.
        string path = Path.Combine(this.thisWorkspace.Root, "empty.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData());
        BoardData? read = await this.LoadAsync(path);

        Assert.NotNull(read);
        Assert.Empty(read!.Schematics);
        Assert.Empty(read.Components);
        Assert.Empty(read.Credits);
        Assert.Equal(string.Empty, read.RevisionDate);
    }

    [Fact]
    public async Task A_blank_revision_date_still_leaves_a_readable_workbook()
    {
        string path = Path.Combine(this.thisWorkspace.Root, "norev.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "A", TechnicalNameOrValue = "B" }]
        });

        BoardData? read = await this.LoadAsync(path);

        Assert.Equal(string.Empty, read!.RevisionDate);
        Assert.Single(read.Components);
    }

    [Fact]
    public void The_target_folder_is_created_when_it_does_not_exist()
    {
        // Publishing a brand-new board writes into a folder that has never existed.
        string path = Path.Combine(this.thisWorkspace.Root, "Commodore", "C64", "250407", "board.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData());

        Assert.True(File.Exists(path));
    }

    [Fact]
    public async Task Writing_over_an_existing_workbook_replaces_it_entirely()
    {
        // Publishing overwrites in place (no revision history - owner's decision). The old
        // contents must not survive as trailing rows below the new ones.
        string path = Path.Combine(this.thisWorkspace.Root, "replace.xlsx");

        BoardWorkbookWriter.Write(path, BoardWorkbookWriterTests.BuildFullBoard());
        BoardWorkbookWriter.Write(path, new BoardData
        {
            Components = [new ComponentEntry { BoardLabel = "Z9", FriendlyName = "Only", TechnicalNameOrValue = "One" }]
        });

        BoardData? read = await this.LoadAsync(path);

        ComponentEntry only = Assert.Single(read!.Components);
        Assert.Equal("Z9", only.BoardLabel);
        Assert.Empty(read.Schematics);
    }

    [Fact]
    public void A_missing_path_is_refused_rather_than_guessed_at()
    {
        Assert.Throws<ArgumentException>(() => BoardWorkbookWriter.Write("", new BoardData()));
        Assert.Throws<ArgumentException>(() => BoardWorkbookWriter.Write("   ", new BoardData()));
    }

    [Fact]
    public void A_null_board_is_refused()
    {
        string path = Path.Combine(this.thisWorkspace.Root, "null.xlsx");

        Assert.Throws<ArgumentNullException>(() => BoardWorkbookWriter.Write(path, null!));
    }

    // -----------------------------------------------------------------------------------
    // The UUID transition rule
    // -----------------------------------------------------------------------------------

    // ###########################################################################################
    // *** NO UUID COLUMN IS WRITTEN AT ALL (2026-09-23). ***
    //
    // This test REPLACES An_existing_uuid_is_written_back_and_a_missing_one_is_not_invented,
    // which pinned the transition behaviour: Phase 4 retired UuidV4 as an identity but kept
    // READING it, so the writer carried whatever a row held through verbatim.
    //
    // The project owner has since dropped the column from the data outright - nothing generated
    // one, nothing compared one, and the only remaining consumer was a validator warning that a
    // dead field needed fixing. The old test is gone because the behaviour it described is gone,
    // and this one is here because "the column is absent" is now the thing worth pinning: a
    // future edit that reintroduced it would otherwise pass silently.
    // ###########################################################################################
    [Fact]
    public async Task NO_uuid_column_is_written_to_the_workbook()
    {
        string path = Path.Combine(this.thisWorkspace.Root, "uuid.xlsx");

        BoardWorkbookWriter.Write(path, BoardWorkbookWriterTests.BuildFullBoard());

        // Asserted against the SHEET ITSELF rather than against a round-trip: the reader no
        // longer has a UuidV4 property to report, so a read-back could not tell the difference
        // between "the column is absent" and "the column is there and ignored".
        Assert.DoesNotContain(
            "UUID v4",
            BoardWorkbookWriterTests.HeaderRowOf(path, "Components"),
            StringComparer.OrdinalIgnoreCase);

        // And the board still loads, which is the point of removing it rather than blanking it.
        BoardData? read = await this.LoadAsync(path);
        Assert.NotNull(read);
        Assert.NotEmpty(read!.Components);
    }

    // -----------------------------------------------------------------------------------
    // The schema is genuinely shared
    // -----------------------------------------------------------------------------------

    [Fact]
    public void Every_sheet_definition_can_build_its_rows()
    {
        // BuildRows throws for a sheet it has no builder for, which is what stops a
        // SheetDefinition being added without its mapping and silently writing a header-only
        // sheet - dropping a whole section of a published board with nothing failing.
        BoardData board = BoardWorkbookWriterTests.BuildFullBoard();

        foreach (BoardWorkbookSchema.SheetDefinition sheet in BoardWorkbookSchema.AllSheets)
        {
            IReadOnlyList<IReadOnlyDictionary<string, string>> rows =
                BoardWorkbookSchema.BuildRows(sheet, board);

            Assert.NotEmpty(rows);

            // Every required header must actually be produced, or the reader will not recognise
            // the sheet's header row at all.
            foreach (string required in sheet.RequiredHeaders)
            {
                Assert.Contains(required, sheet.ColumnOrder);
            }
        }
    }

    // ###########################################################################################
    // The cross-check that matters most: a workbook built by a DIFFERENT author, through
    // BoardWorkbookBuilder, which hand-writes the sheets in the shape the reader has always
    // expected and knows nothing about BoardWorkbookSchema.
    //
    // Everything else here writes with the schema and reads with the schema, so a column renamed
    // on BOTH sides at once would still pass. This one cannot be fooled that way: the source
    // workbook's headers are typed out independently, so if the schema's spelling drifts from the
    // real one, the FIRST read produces blank fields and the comparison fails.
    // ###########################################################################################
    [Fact]
    public async Task A_workbook_built_independently_survives_a_read_write_read_cycle()
    {
        string source = BoardWorkbookBuilder.WriteCompleteBoard(
            this.thisWorkspace.Path_("independent-source.xlsx"),
            "2026-August-21");

        BoardData? first = await this.LoadAsync(source);
        Assert.NotNull(first);

        // *** ANTI-VACUITY, AND IT HAS TO BE THIS SPECIFIC. *** Comparing first-read against
        // second-read is not enough on its own: if the schema misspells a column, the FIRST read
        // already yields "" for it, the second does too, and they agree happily while the field
        // is silently gone. That exact hole was found by renaming ColPartNumber to "Part number" -
        // all the other assertions still passed.
        //
        // So every column the independent fixture populates is asserted against its KNOWN value
        // here, before any comparison. These literals are the fixture's, not the schema's, so a
        // schema spelling that drifts from the real workbook fails right here.
        Assert.NotEmpty(first!.Components);
        Assert.NotEmpty(first.Schematics);
        Assert.Equal("2026-August-21", first.RevisionDate);

        ComponentEntry sourceComponent = first.Components[0];
        Assert.Equal("U1", sourceComponent.BoardLabel);
        Assert.Equal("PLA", sourceComponent.FriendlyName);
        Assert.Equal("906114-01", sourceComponent.TechnicalNameOrValue);
        Assert.Equal("906114", sourceComponent.PartNumber);
        Assert.Equal("IC", sourceComponent.Category);
        Assert.Equal("PAL", sourceComponent.Region);
        Assert.Equal("Programmable logic array", sourceComponent.Description);

        ComponentImageEntry sourceImage = first.ComponentImages[0];
        Assert.Equal("U1", sourceImage.BoardLabel);
        Assert.Equal("PAL", sourceImage.Region);
        Assert.Equal("1", sourceImage.Pin);
        Assert.Equal("Clock in", sourceImage.Name);
        Assert.Equal("1MHz square", sourceImage.ExpectedOscilloscopeReading);
        Assert.Equal("u1-pin1.png", sourceImage.File);
        Assert.Equal("note", sourceImage.Note);
        Assert.Equal("5ms", sourceImage.TimeDiv);
        Assert.Equal("500mV", sourceImage.VoltsDiv);
        Assert.Equal("1.65V", sourceImage.TriggerLevelVolts);

        BoardSchematicEntry sourceSchematic = first.Schematics[0];
        Assert.Equal("Sheet 1", sourceSchematic.SchematicName);
        Assert.Equal("board.kicad_pcb", sourceSchematic.CadName);
        Assert.Equal("sheet1.png", sourceSchematic.SchematicImageFile);
        Assert.Equal("#FF0000", sourceSchematic.SchematicHighlightColor);
        Assert.Equal("0.5", sourceSchematic.SchematicHighlightOpacity);
        Assert.Equal("#00FF00", sourceSchematic.OppositeTraceHighlightColor);
        Assert.Equal("#0000FF", sourceSchematic.ThumbnailHighlightColor);
        Assert.Equal("0.25", sourceSchematic.ThumbnailHighlightOpacity);

        Assert.Equal("Someone", first.Credits[0].NameOrHandle);
        Assert.Equal("someone@example.org", first.Credits[0].Contact);
        Assert.Equal("Schematics", first.Credits[0].SubCategory);
        Assert.Equal("Clock", first.KiCadImportantSignals[0].DisplayName);
        Assert.Equal("/Sheet1/CLK", first.KiCadImportantSignals[0].KiCadNetName);

        string rewritten = Path.Combine(this.thisWorkspace.Root, "rewritten.xlsx");
        BoardWorkbookWriter.Write(rewritten, first);

        BoardData? second = await this.LoadAsync(rewritten);
        Assert.NotNull(second);

        Assert.Equal(first.RevisionDate, second!.RevisionDate);
        Assert.Equal(first.Schematics.Count, second.Schematics.Count);
        Assert.Equal(first.Components.Count, second.Components.Count);
        Assert.Equal(first.ComponentImages.Count, second.ComponentImages.Count);
        Assert.Equal(first.ComponentLocalFiles.Count, second.ComponentLocalFiles.Count);
        Assert.Equal(first.ComponentLinks.Count, second.ComponentLinks.Count);
        Assert.Equal(first.BoardLocalFiles.Count, second.BoardLocalFiles.Count);
        Assert.Equal(first.BoardLinks.Count, second.BoardLinks.Count);
        Assert.Equal(first.Credits.Count, second.Credits.Count);
        Assert.Equal(first.KiCadImportantSignals.Count, second.KiCadImportantSignals.Count);

        for (int i = 0; i < first.Components.Count; i++)
        {
            ComponentEntry before = first.Components[i];
            ComponentEntry after = second.Components[i];

            Assert.Equal(before.BoardLabel, after.BoardLabel);
            Assert.Equal(before.FriendlyName, after.FriendlyName);
            Assert.Equal(before.TechnicalNameOrValue, after.TechnicalNameOrValue);
            Assert.Equal(before.PartNumber, after.PartNumber);
            Assert.Equal(before.Category, after.Category);
            Assert.Equal(before.Region, after.Region);
            Assert.Equal(before.Description, after.Description);
        }

        for (int i = 0; i < first.ComponentImages.Count; i++)
        {
            ComponentImageEntry before = first.ComponentImages[i];
            ComponentImageEntry after = second.ComponentImages[i];

            Assert.Equal(before.BoardLabel, after.BoardLabel);
            Assert.Equal(before.Pin, after.Pin);
            Assert.Equal(before.Name, after.Name);
            Assert.Equal(before.File, after.File);
            Assert.Equal(before.Region, after.Region);
            Assert.Equal(before.ExpectedOscilloscopeReading, after.ExpectedOscilloscopeReading);
            Assert.Equal(before.Note, after.Note);
            Assert.Equal(before.TimeDiv, after.TimeDiv);
            Assert.Equal(before.VoltsDiv, after.VoltsDiv);
            Assert.Equal(before.TriggerLevelVolts, after.TriggerLevelVolts);
        }

        for (int i = 0; i < first.Schematics.Count; i++)
        {
            BoardSchematicEntry before = first.Schematics[i];
            BoardSchematicEntry after = second.Schematics[i];

            Assert.Equal(before.SchematicName, after.SchematicName);
            Assert.Equal(before.CadName, after.CadName);
            Assert.Equal(before.SchematicImageFile, after.SchematicImageFile);
            Assert.Equal(before.SchematicHighlightColor, after.SchematicHighlightColor);
            Assert.Equal(before.SchematicHighlightOpacity, after.SchematicHighlightOpacity);
            Assert.Equal(before.OppositeTraceHighlightColor, after.OppositeTraceHighlightColor);
            Assert.Equal(before.ThumbnailHighlightColor, after.ThumbnailHighlightColor);
            Assert.Equal(before.ThumbnailHighlightOpacity, after.ThumbnailHighlightOpacity);
        }
    }

    [Fact]
    public void The_revision_date_marker_round_trips_through_the_schema()
    {
        // The writer emits prefix + value and the reader takes everything after the prefix. This
        // asserts the pair agrees without going near a file.
        string marker = BoardWorkbookSchema.BuildRevisionDateMarker("2026-August-21");

        Assert.StartsWith(BoardWorkbookSchema.RevisionDateMarkerPrefix, marker);
        Assert.Equal(
            "2026-August-21",
            marker.Substring(BoardWorkbookSchema.RevisionDateMarkerPrefix.Length).Trim());
    }

    // -----------------------------------------------------------------------------------
    // DETERMINISM - the same board must produce the same BYTES (found 2026-09-22).
    // -----------------------------------------------------------------------------------

    [Fact]
    public void Writing_the_SAME_BOARD_TWICE_produces_byte_identical_files()
    {
        // *** AN .xlsx IS A ZIP, AND EPPlus STAMPS EVERY ENTRY WITH THE CURRENT CLOCK. *** Without
        // BoardWorkbookWriter.MakeDeterministic, two publishes of an identical board seconds apart
        // produced different bytes - and therefore different SHA-256 hashes - with nothing in the
        // CONTENT different at all.
        //
        // That is not cosmetic. The workbook's hash is folded into the board's ContentHash
        // (PublishPlan.DescriptorWithWorkbook), which is what every CRT client uses to decide
        // whether to re-download a board. So re-publishing an unchanged board made every user on
        // every machine re-download it.
        //
        // *** THE SLEEP IS LOAD-BEARING and is why this test is worth its second. *** Two writes
        // inside the same clock second agree even with the bug present, which is exactly why it
        // surfaced only as an intermittent failure in PublishExecutorTests rather than as a
        // reproducible one. Crossing a second boundary is what makes this deterministic.
        var data = new BoardData();
        data.Components.Add(new ComponentEntry { BoardLabel = "U8" });

        string first = this.thisWorkspace.Path_("first.xlsx");
        string second = this.thisWorkspace.Path_("second.xlsx");

        BoardWorkbookWriter.Write(first, data);
        Thread.Sleep(1100);
        BoardWorkbookWriter.Write(second, data);

        Assert.Equal(
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(first))),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(second))));
    }

    [Fact]
    public async Task A_normalised_workbook_still_READS_BACK_correctly()
    {
        // MakeDeterministic rewrites the archive entry by entry, so this pins that the result is
        // still a valid workbook rather than merely a consistent one. An .xlsx whose entry ORDER
        // was disturbed ([Content_Types].xml must come first) can hash beautifully and fail to
        // open - which no hash comparison would ever catch.
        var data = new BoardData();
        data.Components.Add(new ComponentEntry { BoardLabel = "U8", Description = "CPU" });

        string path = this.thisWorkspace.Path_("normalised.xlsx");
        BoardWorkbookWriter.Write(path, data);

        string key = "deterministic-" + Guid.NewGuid().ToString("N");
        this.thisCacheKeys.Add(key);

        BoardData? read = await BoardDataReader.LoadAsync(path, key);

        Assert.Equal("U8", Assert.Single(read!.Components).BoardLabel);
    }

    [Fact]
    public void A_CHANGED_board_still_produces_DIFFERENT_bytes()
    {
        // Anti-vacuity: a "fix" that made every workbook identical regardless of content would
        // pass the determinism test above and break publishing entirely - no client would ever
        // re-download anything.
        var one = new BoardData();
        one.Components.Add(new ComponentEntry { BoardLabel = "U8" });

        var two = new BoardData();
        two.Components.Add(new ComponentEntry { BoardLabel = "U9" });

        string firstPath = this.thisWorkspace.Path_("one.xlsx");
        string secondPath = this.thisWorkspace.Path_("two.xlsx");

        BoardWorkbookWriter.Write(firstPath, one);
        BoardWorkbookWriter.Write(secondPath, two);

        Assert.NotEqual(
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(firstPath))),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(secondPath))));
    }

    // ------------------------------------------------------------------ the PRESENTATION

    // ###########################################################################################
    // *** A GENERATED WORKBOOK CARRIES THE SAME HEADER BLOCK AS A HAND-MAINTAINED ONE
    // (owner request, 2026-09-23). ***
    //
    // The writer used to emit data and nothing else, so the first publish of the C64 250407 turned
    // a 190 KB owner-authored workbook into an 81 KB one: no identity lines, no documentation
    // link, no shaded header, no column widths. Nothing was deleted - it was never written - but
    // the result is a file nobody wants to open by hand.
    //
    // These pin the parts a reader would notice. The exact colours and widths are
    // BoardWorkbookStyle's business and are not re-asserted here; what matters is that the block
    // EXISTS and that the data still reads back through the ordinary reader underneath it.
    // ###########################################################################################
    // ###########################################################################################
    // *** THE PREAMBLE IS THE SAME ON EVERY SHEET, WITH NO DOCUMENTATION LINES (owner request,
    // 2026-10-02). *** Identity block (hardware, board, revision date), one blank line, the title
    // band, the headers. The "Documentation of columns in this worksheet is availble here:" line,
    // its link to the old Wiki and the blank line under them were removed from every published
    // board by hand - the columns are maintained from within CRT now - and CRT stopped writing them.
    // ###########################################################################################
    [Fact]
    public void Every_sheet_has_the_identity_block_then_a_blank_line_then_its_title_and_no_documentation_lines()
    {
        string path = this.thisWorkspace.Path_("styled.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData
        {
            RevisionDate = "2026-May-12",
            HardwareName = "Commodore 64",
            BoardName = "250407",
            Components = { new ComponentEntry { BoardLabel = "U8" } },
        });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));

        foreach (BoardWorkbookSchema.SheetDefinition definition in BoardWorkbookSchema.AllSheets)
        {
            OfficeOpenXml.ExcelWorksheet sheet = package.Workbook.Worksheets[definition.SheetName];

            Assert.Equal("# Hardware: Commodore 64", sheet.Cells[1, 1].Text);
            Assert.Equal("# Board: 250407", sheet.Cells[2, 1].Text);
            Assert.Equal("# Revision date: 2026-May-12", sheet.Cells[3, 1].Text);
            Assert.Equal(string.Empty, sheet.Cells[4, 1].Text);
            Assert.Equal(BoardWorkbookStyle.SectionTitleFor(definition.SheetName), sheet.Cells[5, 1].Text);
            Assert.Equal(definition.ColumnOrder[0], sheet.Cells[6, 1].Text);

            // Nothing anywhere above the data points at documentation.
            for (int row = 1; row <= 6; row++)
            {
                for (int column = 1; column <= definition.ColumnOrder.Count; column++)
                {
                    OfficeOpenXml.ExcelRange cell = sheet.Cells[row, column];

                    Assert.DoesNotContain("Documentation", cell.Text ?? string.Empty, StringComparison.OrdinalIgnoreCase);
                    Assert.Null(cell.Hyperlink);
                }
            }
        }
    }

    // ###########################################################################################
    // *** THE REVISION DATE IS ON EVERY SHEET, NOT ONLY BOARD SCHEMATICS (owner request,
    // 2026-10-02). *** The same rich-text line on all of them - and the reader still takes the
    // board's date from Board schematics, which this proves by reading the workbook back.
    // ###########################################################################################
    [Fact]
    public void Every_sheet_carries_the_revision_date_and_the_board_still_reads_it_back()
    {
        string path = this.thisWorkspace.Path_("revision-everywhere.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData
        {
            RevisionDate = "2026-October-2",
            HardwareName = "ZX Spectrum 16K/48K",
            BoardName = "Issue 4B",
        });

        using (var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path)))
        {
            foreach (BoardWorkbookSchema.SheetDefinition definition in BoardWorkbookSchema.AllSheets)
            {
                OfficeOpenXml.ExcelRange cell =
                    package.Workbook.Worksheets[definition.SheetName].Cells[BoardWorkbookStyle.PreambleRevisionDateRow, 1];

                Assert.Equal("# Revision date: 2026-October-2", cell.Text);
                Assert.True(cell.RichText[1].Bold, $"the date on {definition.SheetName} is not bold");
            }
        }

        Assert.Equal("2026-October-2", BoardDataReader.ReadWorkbookUncached(path)!.RevisionDate);
    }

    // ###########################################################################################
    // *** EVERY SHEET IS FROZEN UNDER ITS HEADER ROW (owner request, 2026-10-02). *** Excel's
    // "Freeze Panes" with A7 selected: everything down to the column headers stays on screen while
    // the data scrolls, and no column is frozen.
    // ###########################################################################################
    [Fact]
    public void Every_sheet_is_frozen_under_its_header_row_and_no_column_is_frozen()
    {
        string path = this.thisWorkspace.Path_("frozen.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData
        {
            HardwareName = "C64",
            BoardName = "250407",
            Components = { new ComponentEntry { BoardLabel = "U8" } },
        });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));

        foreach (BoardWorkbookSchema.SheetDefinition definition in BoardWorkbookSchema.AllSheets)
        {
            OfficeOpenXml.ExcelWorksheetView.ExcelWorksheetViewPaneSettings panes =
                package.Workbook.Worksheets[definition.SheetName].View.PaneSettings;

            Assert.Equal(OfficeOpenXml.ePaneState.Frozen, panes.State);
            Assert.Equal(BoardWorkbookStyle.HeaderRow, panes.YSplit);
            // No column split: the attribute is left out of the file (EPPlus reads that as NaN),
            // exactly as Excel writes a rows-only freeze.
            Assert.True(double.IsNaN(panes.XSplit) || panes.XSplit == 0, $"{definition.SheetName} freezes {panes.XSplit} column(s)");
            Assert.Equal($"A{BoardWorkbookStyle.HeaderRow + 1}", panes.TopLeftCell);
        }
    }

    // ###########################################################################################
    // *** THE REVISION DATE IS RICH TEXT: a plain label and a BOLD date. ***
    //
    // The project owner asked for the date itself to stay bold, which is how the shipped boards store
    // it - two runs in one cell rather than one bolded cell. A single run would be visibly wrong
    // whichever weight it took.
    //
    // The reader is unaffected either way, which this also proves: it reads the cell's TEXT, which
    // is the runs concatenated.
    // ###########################################################################################
    [Fact]
    public void The_revision_date_is_bold_while_its_label_is_not()
    {
        string path = this.thisWorkspace.Path_("revision.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData
        {
            RevisionDate = "2026-May-12",
            HardwareName = "Commodore 64",
            BoardName = "250407",
        });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));
        OfficeOpenXml.ExcelWorksheet sheet =
            package.Workbook.Worksheets[BoardWorkbookSchema.SheetBoardSchematics];

        OfficeOpenXml.ExcelRange cell = sheet.Cells[BoardWorkbookStyle.PreambleRevisionDateRow, 1];

        Assert.Equal("# Revision date: 2026-May-12", cell.Text);
        Assert.True(cell.IsRichText);
        Assert.Equal(2, cell.RichText.Count);

        Assert.False(cell.RichText[0].Bold);
        Assert.True(cell.RichText[1].Bold);
        Assert.Equal("2026-May-12", cell.RichText[1].Text);
    }

    // The whole point of the preamble is that it must not break reading - the reader scans for the
    // header row rather than assuming one, and this proves the two agree.
    [Fact]
    public void The_data_still_reads_back_through_the_ordinary_reader()
    {
        string path = this.thisWorkspace.Path_("roundtrip.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData
        {
            RevisionDate = "2026-May-12",
            HardwareName = "Commodore 64",
            BoardName = "250407",
            Components =
            {
                new ComponentEntry { BoardLabel = "U8", PartNumber = "251715-01" },
            },
        });

        BoardData? read = BoardDataReader.ReadWorkbookUncached(path);

        Assert.NotNull(read);
        Assert.Equal("2026-May-12", read!.RevisionDate);
        Assert.Equal("U8", Assert.Single(read.Components).BoardLabel);
        Assert.Equal("251715-01", read.Components[0].PartNumber);

        // And the preamble values round-trip too, so a re-written board keeps its own caption.
        Assert.Equal("Commodore 64", read.HardwareName);
        Assert.Equal("250407", read.BoardName);
    }

    // A board whose preamble names no hardware simply omits those lines rather than writing
    // "# Hardware: " with nothing after it.
    [Fact]
    public void A_board_with_no_names_omits_the_identity_lines_rather_than_writing_empty_ones()
    {
        string path = this.thisWorkspace.Path_("nameless.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData { RevisionDate = "2026-May-12" });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));
        OfficeOpenXml.ExcelWorksheet sheet =
            package.Workbook.Worksheets[BoardWorkbookSchema.SheetComponents];

        Assert.DoesNotContain(
            "# Hardware:",
            sheet.Cells[BoardWorkbookStyle.PreambleHardwareRow, 1].Text ?? string.Empty,
            StringComparison.Ordinal);
    }

    // ###########################################################################################
    // *** THE COLOURED BANDS, which a screenshot comparison showed were missing
    // (owner report, 2026-09-24). ***
    //
    // A generated workbook had the preamble but was otherwise unshaded, while the reference bands
    // its schematics sheet by MEANING: a black title strip, a grey strip over the five highlight
    // columns under one heading, and the CAD-name column picked out in blue on both rows.
    //
    // Asserted as resolved RGB rather than theme references, because that is what the writer emits
    // and what a reader actually sees - the reference stores theme+tint, which renders differently
    // under a different theme.
    // ###########################################################################################
    [Fact]
    public void The_schematics_sheet_is_banded_like_the_reference_board()
    {
        string path = this.thisWorkspace.Path_("banded.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData
        {
            HardwareName = "Commodore 64",
            BoardName = "250407",
        });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));
        OfficeOpenXml.ExcelWorksheet sheet =
            package.Workbook.Worksheets[BoardWorkbookSchema.SheetBoardSchematics];

        int headerRow = BoardWorkbookStyle.HeaderRow;
        int titleRow = headerRow - 1;

        // The three title bands.
        Assert.Equal("#FF000000", sheet.Cells[titleRow, 1].Style.Fill.BackgroundColor.LookupColor());
        Assert.Equal("#FF7F7F7F", sheet.Cells[titleRow, 3].Style.Fill.BackgroundColor.LookupColor());

        // And the heading that spans the highlight columns.
        Assert.Equal(BoardWorkbookStyle.HighlightsBandTitle, sheet.Cells[titleRow, 3].Text);

        // The header row: grey, with the CAD-name column in blue.
        Assert.Equal("#FFD9D9D9", sheet.Cells[headerRow, 1].Style.Fill.BackgroundColor.LookupColor());

        int cad = 0;
        for (int c = 1; c <= BoardWorkbookSchema.BoardSchematics.ColumnOrder.Count; c++)
        {
            if (sheet.Cells[headerRow, c].Text == BoardWorkbookSchema.ColCadName) { cad = c; break; }
        }

        Assert.True(cad > 0, "the CAD name column was not found");
        Assert.Equal("#FF8FAADC", sheet.Cells[headerRow, cad].Style.Fill.BackgroundColor.LookupColor());
    }

    // ###########################################################################################
    // *** EVERY SHEET'S TITLE BAND IS BLACK WITH WHITE TEXT, across all its columns (owner report,
    // 2026-09-30). *** This test used to assert the opposite - "an ordinary sheet has no title
    // banding", the band NOT black - which pinned a misreading of the reference's theme colours:
    // theme 1 is the dark one. Both references, and the project owner's screenshot of production,
    // show "Components" white on black. (The schematics sheet then re-bands its own highlight and
    // CAD-name columns - the test above.) The header row under it stays light grey.
    // ###########################################################################################
    [Fact]
    public void Every_sheet_has_a_black_title_band_with_white_text_across_its_columns()
    {
        string path = this.thisWorkspace.Path_("bands.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData { HardwareName = "C64", BoardName = "250407" });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));

        foreach (BoardWorkbookSchema.SheetDefinition definition in BoardWorkbookSchema.AllSheets)
        {
            OfficeOpenXml.ExcelWorksheet sheet = package.Workbook.Worksheets[definition.SheetName];
            int headerRow = BoardWorkbookStyle.HeaderRow;
            int titleRow = headerRow - 1;

            Assert.Equal(BoardWorkbookStyle.SectionTitleFor(definition.SheetName), sheet.Cells[titleRow, 1].Text);
            Assert.Equal("#FF000000", sheet.Cells[titleRow, 1].Style.Fill.BackgroundColor.LookupColor());
            Assert.Equal("#FFFFFFFF", sheet.Cells[titleRow, 1].Style.Font.Color.LookupColor());
            Assert.Equal("#FFD9D9D9", sheet.Cells[headerRow, 1].Style.Fill.BackgroundColor.LookupColor());

            // Across all its columns - on every sheet but schematics, whose later columns band.
            if (definition.SheetName != BoardWorkbookSchema.SheetBoardSchematics)
            {
                int last = definition.ColumnOrder.Count;
                Assert.Equal("#FF000000", sheet.Cells[titleRow, last].Style.Fill.BackgroundColor.LookupColor());
            }
        }
    }

    // ###########################################################################################
    // A BRAND-NEW BOARD CARRIES ITS OWN NAME (owner report, 2026-09-24).
    //
    // DraftSeeder.CreateNewBoard wrote `new BoardData()`, so the identity lines came out blank and
    // a new board opened in Excel with no caption at all - visible in the project owner's screenshot
    // as an empty row 1 and 2 where every published board names its hardware.
    // ###########################################################################################
    [Fact]
    public void A_new_board_workbook_names_its_hardware_and_board()
    {
        string path = this.thisWorkspace.Path_("newboard.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData
        {
            HardwareName = "Hest Harware",
            BoardName = "Hest Board",
        });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));
        OfficeOpenXml.ExcelWorksheet sheet =
            package.Workbook.Worksheets[BoardWorkbookSchema.SheetComponents];

        Assert.Equal(
            "# Hardware: Hest Harware",
            sheet.Cells[BoardWorkbookStyle.PreambleHardwareRow, 1].Text);

        Assert.Equal(
            "# Board: Hest Board",
            sheet.Cells[BoardWorkbookStyle.PreambleBoardRow, 1].Text);
    }

    // ###########################################################################################
    // *** NOTHING ON THE HEADER ROWS IS BOLD, AND THE HEADERS WRAP (owner report,
    // 2026-09-24). ***
    //
    // The first version bolded both bands and left the headers unwrapped. Side by side with the
    // reference that was immediately visible: bold text the shipped boards do not have, and a
    // sheet far wider than it should be because each column had grown to fit its whole header on
    // one line.
    //
    // The shading is what marks these rows out - the weight is not.
    // ###########################################################################################
    [Fact]
    public void The_header_rows_are_NOT_bold_and_the_headers_wrap()
    {
        string path = this.thisWorkspace.Path_("weights.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData { HardwareName = "C64", BoardName = "250407" });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));
        OfficeOpenXml.ExcelWorksheet sheet =
            package.Workbook.Worksheets[BoardWorkbookSchema.SheetBoardSchematics];

        int headerRow = BoardWorkbookStyle.HeaderRow;

        Assert.False(sheet.Cells[headerRow, 1].Style.Font.Bold);
        Assert.False(sheet.Cells[headerRow - 1, 1].Style.Font.Bold);

        Assert.True(sheet.Cells[headerRow, 1].Style.WrapText);
        Assert.Equal(BoardWorkbookStyle.HeaderRowHeight, sheet.Row(headerRow).Height, 1);
    }

    // The wrap only helps if the columns stay narrow, which means the autofit must not measure the
    // header it is wrapping - that is what widened the sheet in the first place.
    [Fact]
    public void Column_widths_are_fitted_to_the_DATA_not_to_the_wrapped_headers()
    {
        string path = this.thisWorkspace.Path_("widths.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData
        {
            HardwareName = "C64",
            BoardName = "250407",
            Schematics =
            {
                new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = "m.png" },
            },
        });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));
        OfficeOpenXml.ExcelWorksheet sheet =
            package.Workbook.Worksheets[BoardWorkbookSchema.SheetBoardSchematics];

        // "Schematic highlight opacity" is 27 characters; fitted to the data it must be far
        // narrower than that, which is only true when the header is excluded from the measurement.
        int opacityColumn = 0;
        int headerRow = BoardWorkbookStyle.HeaderRow;

        for (int c = 1; c <= BoardWorkbookSchema.BoardSchematics.ColumnOrder.Count; c++)
        {
            if (sheet.Cells[headerRow, c].Text == BoardWorkbookSchema.ColSchematicHighlightOpacity)
            {
                opacityColumn = c;
                break;
            }
        }

        Assert.True(opacityColumn > 0, "the opacity column was not found");
        Assert.True(
            sheet.Column(opacityColumn).Width < 27,
            $"column was {sheet.Column(opacityColumn).Width}, so the header was measured after all");
    }

    // ###########################################################################################
    // *** ALL THREE IDENTITY LINES CARRY A BOLD VALUE, not just the revision date
    // (owner report, 2026-09-24). ***
    //
    // The reference stores "# Hardware: " plus the name in bold as two runs, and the same for the
    // board. Only the revision date was written that way, because it had its own copy of the
    // logic - the other two came out entirely plain. All three now share one helper.
    // ###########################################################################################
    [Theory]
    [InlineData(1, "# Hardware: ", "Test HW2")]
    [InlineData(2, "# Board: ", "Test Board2")]
    public void An_identity_line_is_a_plain_label_with_a_BOLD_value(int row, string label, string value)
    {
        string path = this.thisWorkspace.Path_("identity.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData
        {
            HardwareName = "Test HW2",
            BoardName = "Test Board2",
        });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));
        OfficeOpenXml.ExcelRange cell =
            package.Workbook.Worksheets[BoardWorkbookSchema.SheetComponents].Cells[row, 1];

        Assert.True(cell.IsRichText);
        Assert.Equal(2, cell.RichText.Count);

        Assert.Equal(label, cell.RichText[0].Text);
        Assert.False(cell.RichText[0].Bold);

        Assert.Equal(value, cell.RichText[1].Text);
        Assert.True(cell.RichText[1].Bold);
    }

    // ###########################################################################################
    // *** THE HIGHLIGHTS BAND IS MERGED, which is what makes centring it correct. ***
    //
    // Centring a single unmerged cell put the text in the middle of one column and read as a
    // mistake. The reference merges C8:G8; reproducing the merge makes the centring right.
    // ###########################################################################################
    [Fact]
    public void The_highlights_band_is_merged_across_its_columns()
    {
        string path = this.thisWorkspace.Path_("merged.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData { HardwareName = "C64", BoardName = "250407" });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));
        OfficeOpenXml.ExcelWorksheet sheet =
            package.Workbook.Worksheets[BoardWorkbookSchema.SheetBoardSchematics];

        int titleRow = BoardWorkbookStyle.HeaderRow - 1;

        // Five highlight columns, C through G on this sheet's order.
        Assert.Contains(
            $"C{titleRow}:G{titleRow}",
            sheet.MergedCells.Select(range => range ?? string.Empty));
    }

    // The project owner asked for the reference's own widths rather than fitted ones.
    [Fact]
    public void Column_widths_come_from_the_reference_board()
    {
        string path = this.thisWorkspace.Path_("refwidths.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData { HardwareName = "C64", BoardName = "250407" });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));
        OfficeOpenXml.ExcelWorksheet sheet =
            package.Workbook.Worksheets[BoardWorkbookSchema.SheetBoardSchematics];

        // "Schematic name" is 32.1 in the reference and "Schematic image file" 62.2 - values no
        // autofit would land on, so matching them proves the table is being used.
        Assert.Equal(32.1, sheet.Column(1).Width, 1);
        Assert.Equal(62.2, sheet.Column(2).Width, 1);
    }

    // The workbook's own default font, which a generated package does not inherit.
    [Fact]
    public void The_workbook_font_is_Calibri_11()
    {
        string path = this.thisWorkspace.Path_("font.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData { HardwareName = "C64", BoardName = "250407" });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));

        Assert.Equal(
            BoardWorkbookStyle.FontName,
            package.Workbook.Styles.NamedStyles[0].Style.Font.Name);

        Assert.Equal(
            BoardWorkbookStyle.FontSize,
            package.Workbook.Styles.NamedStyles[0].Style.Font.Size);
    }

    [Fact]
    public void A_written_workbook_has_its_sheets_in_the_published_order_with_Credits_last()
    {
        // The file itself, not just the list it is written from: a hand-made board and one the app
        // wrote must look the same when opened in Excel.
        using var workspace = new TempWorkspace();
        string path = System.IO.Path.Combine(workspace.Root, "board.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData());

        EpplusLicense.Ensure();
        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));
        List<string> names = package.Workbook.Worksheets.Select(sheet => sheet.Name).ToList();

        Assert.Equal(BoardWorkbookSchema.SheetCredits, names[^1]);
        Assert.Equal(BoardWorkbookSchema.SheetKiCadImportantSignals, names[^2]);
        Assert.Equal(BoardWorkbookSchema.AllSheets.Select(sheet => sheet.SheetName), names);
    }

    // ###########################################################################################
    // *** THE HEADER ROW'S HEIGHT AND WRAP ARE PER SHEET (owner report, 2026-09-30). *** 57.6 on
    // every sheet drew a tall grey band over "Components" and all the others. The references have
    // Board schematics at 57.6, Components and Component images at 28.8 (two lines), and every
    // other sheet an ordinary one-line header with no wrap.
    // ###########################################################################################
    [Fact]
    public void The_header_row_is_as_tall_as_its_sheets_headers_need_and_no_taller()
    {
        string path = this.thisWorkspace.Path_("heights.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData { HardwareName = "C64", BoardName = "250407" });

        using var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path));

        double HeightOf(string sheetName, out bool custom, out bool wraps)
        {
            OfficeOpenXml.ExcelWorksheet sheet = package.Workbook.Worksheets[sheetName];
            int row = BoardWorkbookStyle.HeaderRow;
            custom = sheet.Row(row).CustomHeight;
            wraps = sheet.Cells[row, 1].Style.WrapText;
            return sheet.Row(row).Height;
        }

        Assert.Equal(57.6, HeightOf(BoardWorkbookSchema.SheetBoardSchematics, out _, out bool schematicsWrap), 1);
        Assert.True(schematicsWrap);
        Assert.Equal(28.8, HeightOf(BoardWorkbookSchema.SheetComponents, out _, out bool componentsWrap), 1);
        Assert.True(componentsWrap);
        Assert.Equal(28.8, HeightOf(BoardWorkbookSchema.SheetComponentImages, out _, out _), 1);

        foreach (string plain in new[]
                 {
                     BoardWorkbookSchema.SheetComponentLocalFiles, BoardWorkbookSchema.SheetComponentLinks,
                     BoardWorkbookSchema.SheetBoardLocalFiles, BoardWorkbookSchema.SheetBoardLinks,
                     BoardWorkbookSchema.SheetKiCadImportantSignals, BoardWorkbookSchema.SheetCredits
                 })
        {
            HeightOf(plain, out bool custom, out bool wraps);
            Assert.False(custom, $"{plain}'s header row was given a height of its own");
            Assert.False(wraps, $"{plain}'s header wraps");
        }
    }

    // ###########################################################################################
    // *** THE DESCRIPTION HEADER BREAKS OVER TWO LINES, AT THE REFERENCE'S WIDTH (owner report,
    // 2026-09-30). *** It was written on one line, so it ran across a column far wider than the
    // reference's - whose width, keyed on the broken form, was never even looked up. And a board
    // written that way still reads back: the reader makes the break a space.
    // ###########################################################################################
    [Fact]
    public void The_description_header_breaks_over_two_lines_at_its_reference_width_and_reads_back()
    {
        string path = this.thisWorkspace.Path_("description.xlsx");

        BoardWorkbookWriter.Write(path, new BoardData
        {
            HardwareName = "C64",
            BoardName = "250407",
            Components = { new ComponentEntry { BoardLabel = "U8", Category = "IC", Description = "The CIA" } },
        });

        int column = BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColDescription) + 1;
        int headerRow = BoardWorkbookStyle.HeaderRow;

        using (var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(path)))
        {
            OfficeOpenXml.ExcelWorksheet sheet = package.Workbook.Worksheets[BoardWorkbookSchema.SheetComponents];

            Assert.Equal("Short one-liner description\n(one short line only!)", sheet.Cells[headerRow, column].Text);
            Assert.Equal(41.2, sheet.Column(column).Width, 1);
        }

        BoardData? read = BoardDataReader.ReadWorkbookUncached(path);
        Assert.Equal("The CIA", Assert.Single(read!.Components).Description);
    }

    // ###########################################################################################
    // *** THE REFERENCE ITSELF, READ BACK - so the colours cannot be misread a second time. *** The
    // shipped C128 310378 board is the production file the project owner compared against. Its
    // "Components" band is the theme's TEXT colour (dark) with the BACKGROUND colour (light) as text
    // - what the writer's black and white stand for - and its header row is 28.8 points.
    // ###########################################################################################
    [Fact]
    public void The_shipped_reference_has_a_dark_title_band_and_a_two_line_components_header()
    {
        string? folder = AppContext.BaseDirectory;

        while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
            folder = Path.GetDirectoryName(folder);

        Assert.NotNull(folder);

        string reference = Path.Combine(folder!, "Assets", "Data", "Commodore", "C128", "310378", "Data C128 310378 v2.0.0.xlsx");

        EpplusLicense.Ensure();
        using var package = new OfficeOpenXml.ExcelPackage(new FileInfo(reference));
        OfficeOpenXml.ExcelWorksheet sheet = package.Workbook.Worksheets[BoardWorkbookSchema.SheetComponents];

        // Found by its TEXT rather than by row number: the published boards lost their three
        // documentation rows by hand (2026-10-02), so the band is on a different row depending on
        // which copy of the reference this is.
        int titleRow = Enumerable.Range(1, 12).FirstOrDefault(row => sheet.Cells[row, 1].Text == "Components");
        int headerRow = titleRow + 1;

        Assert.True(titleRow > 0, "the Components title band was not found");
        Assert.Equal(OfficeOpenXml.Drawing.eThemeSchemeColor.Text1, sheet.Cells[headerRow - 1, 1].Style.Fill.BackgroundColor.Theme);
        Assert.Equal(OfficeOpenXml.Drawing.eThemeSchemeColor.Background1, sheet.Cells[headerRow - 1, 1].Style.Font.Color.Theme);
        Assert.Equal(BoardWorkbookStyle.HeaderRowHeightFor(BoardWorkbookSchema.SheetComponents)!.Value, sheet.Row(headerRow).Height, 1);
    }
}
