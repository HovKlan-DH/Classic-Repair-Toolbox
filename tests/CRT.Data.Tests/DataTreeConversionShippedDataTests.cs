using System.Reflection;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// THE SHIPPED DATA CONVERTS TO THE CRT 4.0 TREE WITH NOTHING LOST (owner decision, 2026-10-24:
// "I do expect you also will migrate all existing data").
//
// Plans the conversion of the real CRT 3 Excel tree - "tests/CRT 3 Excel data", which was Assets/Data
// until CRT 4.0's installers carried the 4.0 tree there (2026-11-22) - nothing is written - and checks, board by board,
// that the board file holds exactly what CRT reads from the workbook and its JSON file today:
// every value of every row, every highlight, calibration and rename, and every file a row cites
// as the same BYTES under the board's own name. It is the proof the release day's conversion of
// BETA and stable rests on, and it fails the moment a board, a column or a sidecar gains
// something the conversion would not carry. What the conversion ADDS - the pinouts of each board's
// scope baselines and KiCad data (owner decision, 2026-11-06) - may only fill in what was empty, and
// the revision date is the conversion's day (owner request, 2026-11-22).
// ###########################################################################################
[Collection("BoardData")]
public sealed class DataTreeConversionShippedDataTests
{
    // The plan, made once for the run - SubmissionRulesShippedDataTests puts the same converted
    // boards through the server's rules.
    internal static readonly Lazy<DataTreeConversionPlan> Planned = new(() => DataTreeConversion.Plan(DataTreeConversionShippedDataTests.DataRoot(), DataTreeConversionTests.ConvertedOn));

    private static string DataRoot()
    {
        string? folder = AppContext.BaseDirectory;

        while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
            folder = Path.GetDirectoryName(folder);

        Assert.NotNull(folder);

        return Path.Combine(folder!, "tests", "CRT 3 Excel data");
    }

    [Fact]
    public void The_shipped_data_converts_without_a_problem()
    {
        DataTreeConversionPlan plan = DataTreeConversionShippedDataTests.Planned.Value;

        Assert.True(plan.CanWrite, "The shipped data could not be converted:" + Environment.NewLine + string.Join(Environment.NewLine, plan.Problems));
        Assert.True(plan.Boards.Count >= 14, $"Only {plan.Boards.Count} boards were converted.");
    }

    [Fact]
    public void Every_shipped_board_holds_what_CRT_reads_from_its_workbook_today()
    {
        DataTreeConversionPlan plan = DataTreeConversionShippedDataTests.Planned.Value;
        string root = DataTreeConversionShippedDataTests.DataRoot();

        // Each source file's checksum, as the plan took it - hashing the gigabyte again here would
        // double the test's time; DataTreeConversionTests proves a written file's bytes are its source's.
        var shaByPath = plan.SourceFiles.Values.ToDictionary(file => file.RelativePath, file => file.Sha256, StringComparer.OrdinalIgnoreCase);

        Assert.NotEmpty(plan.Boards);

        foreach (ConvertedBoard converted in plan.Boards)
        {
            string workbook = Path.Combine(root, converted.SourceWorkbook);
            BoardData before = BoardDataReader.ReadWorkbookUncached(workbook)!;
            List<KiCadCalibrationEntry> calibrationsBefore = BoardComponentHighlightStorage.LoadAllKiCadCalibrations(workbook);
            List<SchematicRenameRecord> renamesBefore = BoardComponentHighlightStorage.LoadSchematicRenames(workbook);

            // Through the file's bytes, as a CRT 4 reads it.
            BoardDocument document = DataTreeJson.ReadBoard(DataTreeJson.Write(converted.Document));
            BoardContent after = BoardDocumentMapping.ToContent(document);

            var shaByName = document.Files.ToDictionary(file => file.Name, file => file.Sha256, StringComparer.OrdinalIgnoreCase);
            string who = converted.BoardId;

            // A file is the same when the board's name for it names the bytes the row's path held.
            void SameFile(string pathBefore, string nameAfter, string where)
            {
                if (string.IsNullOrWhiteSpace(pathBefore))
                {
                    Assert.True(nameAfter.Length == 0, $"{who} {where}: no file before, [{nameAfter}] after");
                    return;
                }

                Assert.True(shaByName.TryGetValue(nameAfter, out string? sha), $"{who} {where}: [{nameAfter}] is not in the board's files");
                Assert.True(
                    shaByPath.TryGetValue(BoardFileNames.Normalise(pathBefore), out string? shaBefore),
                    $"{who} {where}: [{pathBefore}] was not carried");
                Assert.Equal(shaBefore, sha);
            }

            // A deliberate change: the board is dated the day of the conversion, never the workbook's date
            // (owner request, 2026-11-22: "this is when the JSON files are generated").
            Assert.Equal(BoardWorkbookStyle.FormatRevisionDate(DataTreeConversionTests.ConvertedOn), after.Board.RevisionDate);
            Assert.Equal(before.HardwareName, after.Board.HardwareName);
            Assert.Equal(before.BoardName, after.Board.BoardName);

            DataTreeConversionShippedDataTests.SameRows(before.Schematics, after.Board.Schematics, who, nameof(BoardSchematicEntry.SchematicImageFile), SameFile);

            // The other deliberate change: the pinouts the conversion makes (owner decision, 2026-11-06 -
            // ConversionPinouts). A component's package is filled in only where the workbook had none, and
            // every pin row the workbook had is there as it was; the rest of every row is the workbook's.
            DataTreeConversionShippedDataTests.SameRows(before.Components, after.Board.Components, who, null, SameFile, filledIn: nameof(ComponentEntry.Package));

            var pinsAfter = after.Board.ComponentPins.Select(DataTreeConversionShippedDataTests.PinKey).ToHashSet();

            foreach (ComponentPinEntry pin in before.ComponentPins)
                Assert.True(pinsAfter.Contains(DataTreeConversionShippedDataTests.PinKey(pin)), $"{who}: pin [{DataTreeConversionShippedDataTests.PinKey(pin)}] was lost");

            DataTreeConversionShippedDataTests.SameRows(before.ComponentHighlights, after.Board.ComponentHighlights, who, null, SameFile);
            DataTreeConversionShippedDataTests.SameRows(before.ComponentLocalFiles, after.Board.ComponentLocalFiles, who, nameof(ComponentLocalFileEntry.File), SameFile);
            DataTreeConversionShippedDataTests.SameRows(before.ComponentLinks, after.Board.ComponentLinks, who, null, SameFile);
            DataTreeConversionShippedDataTests.SameRows(before.BoardLocalFiles, after.Board.BoardLocalFiles, who, nameof(BoardLocalFileEntry.File), SameFile);
            DataTreeConversionShippedDataTests.SameRows(before.BoardLinks, after.Board.BoardLinks, who, null, SameFile);
            // The credits exactly as the workbook has them - undated, as credits from before the history are (owner decision,
            // 2026-11-15: "the existing credits should actually have NO date at all").
            DataTreeConversionShippedDataTests.SameRows(before.Credits, after.Board.Credits, who, null, SameFile);
            Assert.All(after.Board.Credits, credit => Assert.True(string.IsNullOrEmpty(credit.Date), $"{who}: a converted credit is dated [{credit.Date}]"));
            DataTreeConversionShippedDataTests.SameRows(before.KiCadImportantSignals, after.Board.KiCadImportantSignals, who, null, SameFile);

            // The deliberate changes: pictures first, then the baselines, each in its own order - and every expected reading in
            // plain words (owner request, 2026-11-17 - ReadingWords: "HIGH" is "High"), the rest of each row the workbook's.
            List<ComponentImageEntry> imagesBefore = before.ComponentImages
                .Where(image => !ComponentImageKinds.IsScopeBaseline(image))
                .Concat(before.ComponentImages.Where(ComponentImageKinds.IsScopeBaseline))
                .Select(image => image.WithExpectedReading(ReadingWords.Of(image.ExpectedOscilloscopeReading)))
                .ToList();

            DataTreeConversionShippedDataTests.SameRows(imagesBefore, after.Board.ComponentImages, who, nameof(ComponentImageEntry.File), SameFile);
            DataTreeConversionShippedDataTests.SameRows(calibrationsBefore, after.Calibrations.ToList(), who, null, SameFile);

            Assert.Equal(renamesBefore.Count, after.Renames.Count);
        }
    }

    [Fact]
    public void Every_file_a_shipped_board_names_is_stored_with_those_bytes()
    {
        DataTreeConversionPlan plan = DataTreeConversionShippedDataTests.Planned.Value;

        foreach (ConvertedBoard board in plan.Boards)
        {
            Assert.NotEmpty(board.Document.Files);

            foreach (BoardDocument.StoredFileRow file in board.Document.Files)
            {
                string stored = DataTreeLayout.StoredFilePath(file.Sha256, file.Name);

                Assert.True(plan.StoredFiles.TryGetValue(stored, out DataTreeSourceFile? source), $"{board.BoardId}: [{file.Name}] is not stored");
                Assert.Equal(file.Sha256, source!.Sha256);
                Assert.Equal(file.Size, source.Size);
            }
        }
    }

    // ###########################################################################################
    // THE CONVERSION MAKES EVERY SHIPPED BOARD'S PINOUTS (owner decision, 2026-11-06: "populate all boards
    // having KiCad data with its pinouts", every board's scope baselines too): each board with a scope
    // baseline gets pins, and each board with KiCad files gets pins for components that have NO baseline -
    // only its KiCad data can have said those.
    // ###########################################################################################
    [Fact]
    public void The_conversion_makes_the_pinouts_of_every_shipped_board_of_its_baselines_and_its_KiCad_data()
    {
        DataTreeConversionPlan plan = DataTreeConversionShippedDataTests.Planned.Value;
        int withKiCad = 0;

        foreach (ConvertedBoard converted in plan.Boards)
        {
            BoardData board = BoardDocumentMapping.ToContent(converted.Document).Board;
            var pinned = board.ComponentPins.Select(pin => pin.BoardLabel.Trim()).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var baselined = board.ComponentImages.Where(ComponentImageKinds.IsScopeBaseline).Select(image => image.BoardLabel.Trim()).ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (baselined.Count > 0)
                Assert.True(pinned.Count > 0, $"{converted.BoardId}: scope baselines, and no pins made of them");

            if (converted.Document.Files.Any(file => file.Name.EndsWith(".kicad_pcb", StringComparison.OrdinalIgnoreCase)))
            {
                withKiCad++;
                Assert.True(pinned.Any(label => !baselined.Contains(label)), $"{converted.BoardId}: KiCad data, and no pins made of it");
                Assert.Contains(plan.Notes, note => note.Contains($"[{converted.BoardId}]: pinouts made", StringComparison.Ordinal) && note.Contains("KiCad data", StringComparison.Ordinal));
            }
        }

        Assert.True(withKiCad >= 5, $"Only {withKiCad} shipped boards with KiCad data were found.");
    }

    // ###########################################################################################
    // THE CONVERSION GIVES EVERY PIN OF A CHIP CRT HOLDS THE DATASHEET OF ITS DATASHEET'S DIRECTION (owner request, 2026-11-17:
    // "convert the data, so it also will show correctly in the text part of it ... make sure it also gets done when we
    // go-live"; and the same day "OK to what you suggest for 1-4" - the datasheet before KiCad, DatasheetPins): on every
    // shipped board each such pin says what the datasheet does - but a "Not connected" the board says, and what the board
    // says of a pin the datasheet leaves unconnected or unsettled. Measured on the 250407, whose KiCad project types its 556's
    // DISCHARGE, its 6510's address lines and its ROMs' data pins as inputs, and on the 250469, whose project types every pin
    // of its 6526s as an input.
    // ###########################################################################################
    [Fact]
    public void The_conversion_gives_every_pin_of_a_chip_CRT_holds_the_datasheet_of_its_datasheets_direction()
    {
        int logicPins = 0;
        int otherPins = 0;

        foreach (ConvertedBoard converted in DataTreeConversionShippedDataTests.Planned.Value.Boards)
        {
            BoardData board = BoardDocumentMapping.ToContent(converted.Document).Board;

            foreach (ComponentEntry row in board.Components)
            {
                List<ComponentPinEntry> pins = ComponentPaneEdits.PinsOf(board, row.BoardLabel, row.Region);

                if (DatasheetPins.Of(row, pins) is not { } datasheet)
                {
                    continue;
                }

                bool logic = LogicChips.For(row, pins) is not null;

                foreach (ComponentPinEntry pin in pins)
                {
                    if (ComponentVocabulary.PinNumberOf(pin.Pin) is not int number || datasheet(number) is not { Length: > 0 } said)
                    {
                        continue;
                    }

                    bool kept = pin.Direction == "Not connected" || said == "Not connected";

                    Assert.True(pin.Direction == said || kept, $"{converted.BoardId} {row.BoardLabel} pin {pin.Pin}: [{pin.Direction}], the datasheet says [{said}]");

                    if (logic)
                    {
                        logicPins++;
                    }
                    else
                    {
                        otherPins++;
                    }
                }
            }
        }

        Assert.True(logicPins > 1500, $"Only {logicPins} pins of logic chips were checked.");
        Assert.True(otherPins > 2000, $"Only {otherPins} pins of other ICs were checked.");

        string DirectionOf(string boardId, string label, string pin) =>
            BoardDocumentMapping.ToContent(DataTreeConversionShippedDataTests.Planned.Value.Boards.Single(item => item.BoardId == boardId).Document).Board
                .ComponentPins.Single(row => row.BoardLabel == label && row.Pin == pin).Direction;

        Assert.Equal("Output", DirectionOf("Commodore/C64/250407", "U20", "1"));
        Assert.Equal("Output", DirectionOf("Commodore/C64/250407", "U20", "5"));
        Assert.Equal("Output", DirectionOf("Commodore/C64/250407", "U7", "7"));
        Assert.Equal("Input/output", DirectionOf("Commodore/C64/250407", "U7", "30"));
        Assert.Equal("Output", DirectionOf("Commodore/C64/250407", "U4", "9"));
        Assert.Equal("Input/output", DirectionOf("Commodore/C64/250469", "U1", "2"));
    }

    // ###########################################################################################
    // THE CONVERSION GIVES EVERY CHIP CRT HOLDS THE DATASHEET OF EVERY PIN AND ITS CASE (owner request, 2026-11-17: "Can you
    // now draw also this IC? (and add its pins in data migration) (when go-live)", and of the rest of the chips not drawn:
    // "investigate ... if they are drawn as pinouts and its data created/migrated at go-live"): on every shipped board each
    // such chip has a pin of every number its datasheet lists, and a DIP of that many pins - so CRT draws it. The C128
    // 250477's chips had neither, nor the CPC 664's: the 4464s come out as the owner's picture names their pins, the 8568 with
    // its own pins (not the 8563's), the Gate Array whole beside the pins its baselines gave it, the CX20185 in its QFP-44,
    // the 64H156 as the 40-pin 251828-01 it is (the owner counted its pins, 2026-11-18).
    // ###########################################################################################
    [Fact]
    public void The_conversion_gives_every_chip_CRT_holds_the_datasheet_of_every_pin_and_its_case()
    {
        int chips = 0;

        foreach (ConvertedBoard converted in DataTreeConversionShippedDataTests.Planned.Value.Boards)
        {
            BoardData board = BoardDocumentMapping.ToContent(converted.Document).Board;

            foreach (ComponentEntry row in board.Components)
            {
                List<ComponentPinEntry> pins = ComponentPaneEdits.PinsOf(board, row.BoardLabel, row.Region);

                if (DatasheetPins.PinoutOf(row, pins) is not { } pinout)
                {
                    continue;
                }

                chips++;

                var numbers = pins.Select(pin => ComponentVocabulary.PinNumberOf(pin.Pin)).ToHashSet();
                string who = $"{converted.BoardId} {row.BoardLabel} {row.Region} ({pinout.Chip})";

                Assert.All(pinout.Pins, pin => Assert.True(numbers.Contains(int.Parse(pin.Pin, System.Globalization.CultureInfo.InvariantCulture)), $"{who}: no pin {pin.Pin}"));

                if (pinout.Package is not null)
                {
                    Assert.True(ComponentVocabulary.PinCountOf(row.Package) == pinout.Pins.Count, $"{who}: package [{row.Package}]");
                }
            }
        }

        // 483 on 2026-11-17.
        Assert.True(chips > 450, $"Only {chips} chips with a datasheet were found.");

        BoardData Board(string boardId) =>
            BoardDocumentMapping.ToContent(DataTreeConversionShippedDataTests.Planned.Value.Boards.Single(item => item.BoardId == boardId).Document).Board;
        string[] PinsOf(BoardData board, string label) =>
            [.. board.ComponentPins.Where(pin => pin.BoardLabel == label).Select(pin => $"{pin.Pin}|{pin.Name}|{pin.Direction}")];
        string PackageOf(BoardData board, string label) => board.Components.First(row => row.BoardLabel == label).Package;

        BoardData dcr = Board("Commodore/C128/250477");

        foreach (string dram in new[] { "U23", "U25", "U38", "U39", "U40", "U41" })
        {
            Assert.Equal(
                ["1|/G|Input", "2|D0|Input/output", "3|D1|Input/output", "4|/W|Input", "5|/RAS|Input", "6|A6|Input", "7|A5|Input", "8|A4|Input", "9|VCC|Power",
                 "10|A7|Input", "11|A3|Input", "12|A2|Input", "13|A1|Input", "14|A0|Input", "15|D2|Input/output", "16|/CAS|Input", "17|D3|Input/output", "18|VSS|Ground"],
                PinsOf(dcr, dram));
            Assert.Equal("DIP-18", PackageOf(dcr, dram));
        }

        Assert.Equal("DIP-48", PackageOf(dcr, "U22"));
        Assert.Contains("9|/INTR|Output", PinsOf(dcr, "U22"));
        Assert.Equal("DIP-48", PackageOf(dcr, "U107"));
        Assert.Equal(44, PinsOf(dcr, "U108").Length);
        Assert.Equal("QFP-44", PackageOf(dcr, "U108"));
        Assert.Equal(40, PinsOf(dcr, "U105").Length);
        Assert.Equal("DIP-40", PackageOf(dcr, "U105"));

        BoardData cpc = Board("Amstrad/CPC 664/MC0005A");

        Assert.Equal(40, PinsOf(cpc, "IC116").Length);
        Assert.Equal("DIP-40", PackageOf(cpc, "IC116"));
        Assert.Equal("DIP-28", PackageOf(cpc, "IC103"));
    }

    // ###########################################################################################
    // AN OSCILLATOR'S CAN FROM ITS KiCad FOOTPRINT (owner request, 2026-11-18 - for the Amiga, proven here on the C128's): the
    // Open128's 16 MHz Y1 stands on KiCad's "Oscillator_DIP-14", pads 1, 7, 8 and 14, so it is converted as that can, every pin
    // of it on a corner - on both boards the Open128 project is the KiCad data of.
    // ###########################################################################################
    [Theory]
    [InlineData("Commodore/C128/310378")]
    [InlineData("Commodore/C128/310378 Open128")]
    public void The_conversion_makes_the_Open128s_oscillator_its_can(string boardId)
    {
        BoardData board = BoardDocumentMapping.ToContent(DataTreeConversionShippedDataTests.Planned.Value.Boards.Single(item => item.BoardId == boardId).Document).Board;
        ComponentEntry oscillator = board.Components.First(row => row.BoardLabel == "Y1");

        Assert.Equal("Oscillator DIP-14", oscillator.Package);
        Assert.NotEmpty(ComponentPaneEdits.PinsOf(board, "Y1", oscillator.Region));
        Assert.All(ComponentPaneEdits.PinsOf(board, "Y1", oscillator.Region), pin => Assert.True(ComponentVocabulary.IsOnPackage(oscillator.Package, pin.Pin), pin.Pin));
    }

    // ###########################################################################################
    // THE CONVERSION WRITES THE EXPECTED READINGS IN PLAIN WORDS (owner request, 2026-11-17: "make sure it is also done on
    // go-live day"): no shipped board's baseline says "HIGH", "LOW", "FLOAT" or "VIDEO" afterwards - thousands said one
    // before - and the 250407's 556 reads "Low" on its pin 1.
    // ###########################################################################################
    [Fact]
    public void The_conversion_writes_every_expected_reading_in_plain_words()
    {
        int readings = 0;

        foreach (ConvertedBoard converted in DataTreeConversionShippedDataTests.Planned.Value.Boards)
        {
            foreach (ComponentImageEntry image in BoardDocumentMapping.ToContent(converted.Document).Board.ComponentImages.Where(image => image.ExpectedOscilloscopeReading.Length > 0))
            {
                Assert.Equal(ReadingWords.Of(image.ExpectedOscilloscopeReading), image.ExpectedOscilloscopeReading);
                Assert.DoesNotMatch(@"\b(HIGH|LOW|FLOAT|VIDEO)\b", image.ExpectedOscilloscopeReading);
                readings++;
            }
        }

        Assert.True(readings > 10000, $"Only {readings} expected readings were checked.");
        Assert.Contains(DataTreeConversionShippedDataTests.Planned.Value.Notes, note => note.Contains("expected oscilloscope reading(s) written in plain words", StringComparison.Ordinal));

        BoardData c64 = BoardDocumentMapping.ToContent(DataTreeConversionShippedDataTests.Planned.Value.Boards.Single(item => item.BoardId == "Commodore/C64/250407").Document).Board;

        Assert.Contains(c64.ComponentImages, image => image.BoardLabel == "U20" && image.Pin == "1" && image.ExpectedOscilloscopeReading == "Low");
    }

    // A board with KiCad data keeps every file of it, under "KiCad data/".
    [Fact]
    public void The_KiCad_files_of_a_shipped_board_go_with_it()
    {
        DataTreeConversionPlan plan = DataTreeConversionShippedDataTests.Planned.Value;

        ConvertedBoard board = plan.Boards.Single(item => item.BoardId == "Commodore/C64/250407");

        Assert.Contains(board.Document.Files, file => file.Name.StartsWith("KiCad data/", StringComparison.Ordinal) && file.Name.EndsWith(".kicad_pcb", StringComparison.Ordinal));
    }

    [Fact]
    public void The_shipped_board_files_read_back_to_the_same_bytes()
    {
        foreach (ConvertedBoard board in DataTreeConversionShippedDataTests.Planned.Value.Boards)
        {
            string written = DataTreeJson.Write(board.Document);

            Assert.Equal(written, DataTreeJson.Write(DataTreeJson.ReadBoard(written)));
        }
    }

    [Fact]
    public void The_main_file_lists_every_shipped_board_and_oscilloscope()
    {
        DataTreeConversionPlan plan = DataTreeConversionShippedDataTests.Planned.Value;
        string master = MasterListing.NewestMasterPath(DataTreeConversionShippedDataTests.DataRoot())!;

        Assert.True(MasterListing.TryRead(master, out IReadOnlyList<MasterListingRow> rows, out _));
        Assert.Equal(rows.Select(row => row.BoardId), plan.Main.Boards.Select(board => board.BoardId));
        Assert.Equal(rows.Select(row => row.HardwareName), plan.Main.Boards.Select(board => board.HardwareName));

        Assert.True(plan.Main.Oscilloscopes.Count >= 20, $"Only {plan.Main.Oscilloscopes.Count} oscilloscopes were converted.");
        Assert.All(plan.Main.Oscilloscopes, scope => Assert.False(string.IsNullOrWhiteSpace(scope.Brand), scope.SeriesOrModel));
        Assert.NotEmpty(plan.Main.BoardsRevisionDate);

        Assert.Equal(DataTreeJson.Write(plan.Main), DataTreeJson.Write(DataTreeJson.ReadMain(DataTreeJson.Write(plan.Main))));
    }

    [Fact]
    public void The_MiniPro_tests_are_carried()
    {
        DataTreeConversionPlan plan = DataTreeConversionShippedDataTests.Planned.Value;

        Assert.Contains(plan.CopiedFiles.Keys, path => path.StartsWith(DataTreeLayout.MiniProTestsFolder + "/", StringComparison.Ordinal));
    }

    // The point of storing each file once: the boards' files add up to more than is stored.
    [Fact]
    public void Storing_each_file_once_takes_less_room_than_the_boards_name()
    {
        DataTreeConversionPlan plan = DataTreeConversionShippedDataTests.Planned.Value;

        Assert.True(plan.StoredBytes < plan.BoardFileBytes, $"Stored {plan.StoredBytes} bytes for {plan.BoardFileBytes} named.");
    }

    // ---- helpers ---------------------------------------------------------------------------

    private static string PinKey(ComponentPinEntry pin) => $"{pin.BoardLabel}|{pin.Region}|{pin.Pin}|{pin.Name}|{pin.Direction}";

    // Two lists of one row type, equal property by property - by reflection, so a property added
    // to a row type later is compared too. `fileProperty` is compared as the same bytes instead;
    // `filledIn` may hold a value after where it was empty before, and only then.
    private static void SameRows<T>(
        List<T> before,
        List<T> after,
        string who,
        string? fileProperty,
        Action<string, string, string> sameFile,
        string? filledIn = null)
    {
        string type = typeof(T).Name;

        Assert.True(before.Count == after.Count, $"{who}: {before.Count} {type} rows before, {after.Count} after");

        PropertyInfo[] properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

        for (int i = 0; i < before.Count; i++)
        {
            foreach (PropertyInfo property in properties)
            {
                object? a = property.GetValue(before[i]);
                object? b = property.GetValue(after[i]);

                if (property.Name == fileProperty)
                    sameFile((string)a!, (string)b!, $"{type} row {i + 1}");
                else if (property.Name == filledIn && string.IsNullOrWhiteSpace(a as string))
                    continue;
                else
                    Assert.True(Equals(a, b), $"{who}: {type} row {i + 1} {property.Name} was [{a}], is [{b}]");
            }
        }
    }

}
