using System.Reflection;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// THE SHIPPED DATA CONVERTS TO THE CRT 4.0 TREE WITH NOTHING LOST (owner decision, 2026-10-24:
// "I do expect you also will migrate all existing data").
//
// Plans the conversion of the real Assets/Data - nothing is written - and checks, board by board,
// that the board file holds exactly what CRT reads from the workbook and its JSON file today:
// every value of every row, every highlight, calibration and rename, and every file a row cites
// as the same BYTES under the board's own name. It is the proof the release day's conversion of
// BETA and stable rests on, and it fails the moment a board, a column or a sidecar gains
// something the conversion would not carry.
// ###########################################################################################
[Collection("BoardData")]
public sealed class DataTreeConversionShippedDataTests
{
    // The plan, made once for the run - SubmissionRulesShippedDataTests puts the same converted
    // boards through the server's rules.
    internal static readonly Lazy<DataTreeConversionPlan> Planned = new(() => DataTreeConversion.Plan(DataTreeConversionShippedDataTests.DataRoot()));

    private static string DataRoot()
    {
        string? folder = AppContext.BaseDirectory;

        while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
            folder = Path.GetDirectoryName(folder);

        Assert.NotNull(folder);

        return Path.Combine(folder!, "Assets", "Data");
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

            Assert.Equal(before.RevisionDate, after.Board.RevisionDate);
            Assert.Equal(before.HardwareName, after.Board.HardwareName);
            Assert.Equal(before.BoardName, after.Board.BoardName);

            DataTreeConversionShippedDataTests.SameRows(before.Schematics, after.Board.Schematics, who, nameof(BoardSchematicEntry.SchematicImageFile), SameFile);
            DataTreeConversionShippedDataTests.SameRows(before.Components, after.Board.Components, who, null, SameFile);
            DataTreeConversionShippedDataTests.SameRows(before.ComponentHighlights, after.Board.ComponentHighlights, who, null, SameFile);
            DataTreeConversionShippedDataTests.SameRows(before.ComponentLocalFiles, after.Board.ComponentLocalFiles, who, nameof(ComponentLocalFileEntry.File), SameFile);
            DataTreeConversionShippedDataTests.SameRows(before.ComponentLinks, after.Board.ComponentLinks, who, null, SameFile);
            DataTreeConversionShippedDataTests.SameRows(before.BoardLocalFiles, after.Board.BoardLocalFiles, who, nameof(BoardLocalFileEntry.File), SameFile);
            DataTreeConversionShippedDataTests.SameRows(before.BoardLinks, after.Board.BoardLinks, who, null, SameFile);
            DataTreeConversionShippedDataTests.SameRows(before.Credits, after.Board.Credits, who, null, SameFile);
            DataTreeConversionShippedDataTests.SameRows(before.KiCadImportantSignals, after.Board.KiCadImportantSignals, who, null, SameFile);

            // The one deliberate change: pictures first, then the baselines, each in its own order.
            List<ComponentImageEntry> imagesBefore = before.ComponentImages
                .Where(image => !ComponentImageKinds.IsScopeBaseline(image))
                .Concat(before.ComponentImages.Where(ComponentImageKinds.IsScopeBaseline))
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

    // Two lists of one row type, equal property by property - by reflection, so a property added
    // to a row type later is compared too. `fileProperty` is compared as the same bytes instead.
    private static void SameRows<T>(
        List<T> before,
        List<T> after,
        string who,
        string? fileProperty,
        Action<string, string, string> sameFile)
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
                else
                    Assert.True(Equals(a, b), $"{who}: {type} row {i + 1} {property.Name} was [{a}], is [{b}]");
            }
        }
    }

}
