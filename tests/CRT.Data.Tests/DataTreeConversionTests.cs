using System.Security.Cryptography;
using System.Text.Json;
using Handlers.DataHandling;
using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Converting an Excel-era data tree into the CRT 4.0 tree, on small trees written by CRT's own
// workbook and sidecar writers. What must hold: every board names its files without the folders
// they were kept in, each file's bytes are stored once, the KiCad files and the MiniPro tests go
// along, anything the readers would quietly drop stops the conversion, and Write puts exactly
// the plan into a folder of its own - and nowhere else.
// ###########################################################################################
public sealed class DataTreeConversionTests : IClassFixture<DataTreeConversionTests.StandardTree>
{
    private const string BoardA = "Commodore/C64/250407";
    private const string BoardB = "Commodore/C64/250425";
    private const string WorkbookA = BoardA + "/Data C64 250407 v3.0.0.xlsx";
    private const string WorkbookB = BoardB + "/Data C64 250425 v3.0.0.xlsx";

    private readonly StandardTree thisTree;

    public DataTreeConversionTests(StandardTree tree)
    {
        this.thisTree = tree;
    }

    // ---- the standard tree ------------------------------------------------------------------

    // Two boards, the shared folders, KiCad data, the MiniPro tests and an older generation -
    // built once for every test that only reads it.
    public sealed class StandardTree : IDisposable
    {
        public ExcelTree Tree { get; } = ExcelTree.Standard();

        public DataTreeConversionPlan Plan { get; }

        public StandardTree()
        {
            this.Plan = DataTreeConversion.Plan(this.Tree.DataRoot);
        }

        public void Dispose() => this.Tree.Dispose();
    }

    private DataTreeConversionPlan Plan => this.thisTree.Plan;

    private ConvertedBoard Board(string id) => this.Plan.Boards.Single(board => board.BoardId == id);

    [Fact]
    public void The_standard_tree_converts_without_a_problem()
    {
        Assert.True(this.Plan.CanWrite, string.Join(" | ", this.Plan.Problems));
        Assert.Equal([BoardA, BoardB], this.Plan.Boards.Select(board => board.BoardId));
        Assert.Equal(new Version(3, 0, 0), this.Plan.Generation);
    }

    [Fact]
    public void A_board_names_its_files_without_the_folders_they_were_kept_in()
    {
        BoardDocument a = this.Board(BoardA).Document;

        Assert.Equal("Board layout.png", a.Schematics[0].ImageFile);
        Assert.Equal("Component images/7406.png", a.ComponentImages.Single(row => row.BoardLabel == "U7").File);
        Assert.Equal("Scope baseline/U1_1_PAL.png", a.ScopeBaselines.Single().File);
        Assert.Equal("Component local files/7406.pdf", a.ComponentFiles.Single().File);
        Assert.Equal("Service manual.pdf", a.BoardFiles.Single().File);

        Assert.Equal(
            ["Board layout.png", "Component images/7406.png", "Component local files/7406.pdf", "KiCad data/250407_.kicad_pcb", "KiCad data/Pages/vic.kicad_sch", "Scope baseline/U1_1_PAL.png", "Service manual.pdf"],
            a.Files.Select(file => file.Name));
    }

    // The datasheet both boards cite, and the manual board A cites from board B's folder: one
    // stored copy each, whichever boards name them.
    [Fact]
    public void Bytes_named_by_several_boards_are_stored_once()
    {
        string datasheet = this.Board(BoardA).Document.Files.Single(file => file.Name == "Component local files/7406.pdf").Sha256;
        string manual = this.Board(BoardA).Document.Files.Single(file => file.Name == "Service manual.pdf").Sha256;

        Assert.Equal(datasheet, this.Board(BoardB).Document.Files.Single(file => file.Name == "Component local files/7406.pdf").Sha256);
        Assert.Equal(manual, this.Board(BoardB).Document.Files.Single(file => file.Name == "Service manual.pdf").Sha256);
        Assert.Single(this.Plan.StoredFiles.Keys, key => key.Contains(datasheet, StringComparison.Ordinal));
        Assert.True(this.Plan.StoredBytes < this.Plan.BoardFileBytes);
    }

    [Fact]
    public void A_file_s_checksum_is_the_SHA_256_of_its_bytes()
    {
        BoardDocument.StoredFileRow layout = this.Board(BoardA).Document.Files.Single(file => file.Name == "Board layout.png");

        Assert.Equal(Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(this.thisTree.Tree.Full(BoardA + "/Board layout.png")))), layout.Sha256);
        Assert.Equal(new FileInfo(this.thisTree.Tree.Full(BoardA + "/Board layout.png")).Length, layout.Size);
    }

    [Fact]
    public void The_rows_highlights_calibration_and_renames_go_with_the_board()
    {
        BoardDocument a = this.Board(BoardA).Document;

        Assert.Equal(["U1", "U7"], a.Components.Select(row => row.BoardLabel));
        Assert.Equal("1uS", a.ScopeBaselines.Single().TimeDiv);
        Assert.Equal(new[] { 100, 200, 30, 40 }, new[] { a.Highlights[0].X, a.Highlights[0].Y, a.Highlights[0].Width, a.Highlights[0].Height });
        Assert.Equal(0.75, a.KiCadCalibrations.Single().ScaleX);
        Assert.Equal("Board layout.png", a.SchematicRenames.Single().File);
        Assert.Equal("2026-October-4", a.RevisionDate);
    }

    // A note-only image row stays a picture with no file.
    [Fact]
    public void A_row_with_a_note_and_no_file_stays_without_one()
    {
        BoardDocument.ComponentImageRow pinout = this.Board(BoardA).Document.ComponentImages.Single(row => row.BoardLabel == "U1" && row.Name == "Pinout");

        Assert.Equal(string.Empty, pinout.File);
        Assert.Equal("Same as 6510", pinout.Note);
    }

    [Fact]
    public void The_main_file_lists_the_boards_in_order_and_every_oscilloscope_with_its_brand()
    {
        MainDocument main = this.Plan.Main;

        Assert.Equal([BoardA, BoardB], main.Boards.Select(board => board.BoardId));
        Assert.Equal("Commodore 64", main.Boards[0].HardwareName);
        Assert.Equal("Long board", main.Boards[0].HardwareNotes);
        Assert.Equal("2026-October-11", main.BoardsRevisionDate);
        Assert.Equal("2026-October-4", main.OscilloscopesRevisionDate);

        // The second row left its brand blank, as the sheet does; CRT carried it down.
        Assert.Equal(["Rigol", "Rigol"], main.Oscilloscopes.Select(scope => scope.Brand));
        Assert.Equal(["DHO800", "DHO900"], main.Oscilloscopes.Select(scope => scope.SeriesOrModel));
        Assert.Equal(":TIMEBASE:SCALE {0}", main.Oscilloscopes[0].SetTimeDiv);
        Assert.Equal("2nS,5nS", main.Oscilloscopes[0].TimeDivList);
    }

    [Fact]
    public void The_MiniPro_tests_go_to_their_own_folder()
    {
        Assert.Equal(["MiniPro/IC tests/Catalogue/7406.json"], this.Plan.CopiedFiles.Keys);
    }

    // The older generation stays behind for older CRTs - and the plan says so, file by file.
    [Fact]
    public void What_only_an_older_generation_uses_is_left_out_and_listed()
    {
        Assert.Contains(BoardA + "/Data C64 250407 v2.0.0.xlsx", this.Plan.LeftOut);
        Assert.Contains(BoardA + "/Old only.png", this.Plan.LeftOut);
        Assert.DoesNotContain(BoardA + "/Board layout.png", this.Plan.LeftOut);
        Assert.DoesNotContain(WorkbookA, this.Plan.LeftOut);
        Assert.DoesNotContain(BoardA + "/Data C64 250407 v3.0.0.json", this.Plan.LeftOut);
    }

    // ---- what stops a conversion ------------------------------------------------------------

    private static IReadOnlyList<string> ProblemsOf(Action<ExcelTree> spoil)
    {
        using ExcelTree tree = ExcelTree.Standard();
        spoil(tree);
        return DataTreeConversion.Plan(tree.DataRoot).Problems;
    }

    [Fact]
    public void A_file_a_row_cites_that_is_not_there_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree => File.Delete(tree.Full(BoardA + "/Scope baseline/U1_1_PAL.png")));

        Assert.Contains(problems, problem => problem.Contains("Scope baseline/U1_1_PAL.png", StringComparison.Ordinal) && problem.Contains("not in the data", StringComparison.Ordinal));
    }

    [Fact]
    public void A_column_CRT_does_not_read_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree => tree.EditWorkbook(WorkbookA, package =>
        {
            ExcelWorksheet sheet = package.Workbook.Worksheets[BoardWorkbookSchema.SheetComponents];
            int header = ExcelTree.HeaderRow(sheet);
            sheet.Cells[header, 20].Value = "Colour of the case";
            sheet.Cells[header + 1, 20].Value = "Beige";
        }));

        Assert.Contains(problems, problem => problem.Contains("[Colour of the case]", StringComparison.Ordinal));
    }

    [Fact]
    public void A_value_under_no_column_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree => tree.EditWorkbook(WorkbookA, package =>
        {
            ExcelWorksheet sheet = package.Workbook.Worksheets[BoardWorkbookSchema.SheetComponents];
            sheet.Cells[ExcelTree.HeaderRow(sheet) + 1, 25].Value = "A note typed beside the table";
        }));

        Assert.Contains(problems, problem => problem.Contains("A note typed beside the table", StringComparison.Ordinal));
    }

    [Fact]
    public void A_repeated_column_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree => tree.EditWorkbook(WorkbookA, package =>
        {
            ExcelWorksheet sheet = package.Workbook.Worksheets[BoardWorkbookSchema.SheetComponents];
            sheet.Cells[ExcelTree.HeaderRow(sheet), 21].Value = BoardWorkbookSchema.ColFriendlyName;
        }));

        Assert.Contains(problems, problem => problem.Contains("twice", StringComparison.Ordinal));
    }

    [Fact]
    public void A_sheet_CRT_does_not_read_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree => tree.EditWorkbook(WorkbookA, package =>
            package.Workbook.Worksheets.Add("My notes").Cells[1, 1].Value = "Bought at a flea market"));

        Assert.Contains(problems, problem => problem.Contains("[My notes]", StringComparison.Ordinal));
    }

    [Fact]
    public void A_sheet_with_values_and_no_header_row_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree => tree.EditWorkbook(WorkbookA, package =>
        {
            ExcelWorksheet sheet = package.Workbook.Worksheets[BoardWorkbookSchema.SheetCredits];
            sheet.Cells[ExcelTree.HeaderRow(sheet), 1].Value = "Kategori";
        }));

        Assert.Contains(problems, problem => problem.Contains("no header row", StringComparison.Ordinal) && problem.Contains(BoardWorkbookSchema.SheetCredits, StringComparison.Ordinal));
    }

    [Fact]
    public void A_sidecar_part_CRT_does_not_read_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree => tree.EditSidecar(WorkbookA, json => json.Insert(1, "\"Notes\": {\"x\": 1},")));

        Assert.Contains(problems, problem => problem.Contains("[Notes]", StringComparison.Ordinal));
    }

    // CRT's reader skips a highlight it cannot read - the converter must not.
    [Fact]
    public void A_highlight_the_reader_would_skip_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree => tree.EditSidecar(WorkbookA, json => json.Replace("\"Board layout\": {", "\"Board layout\": {\"  \": [{\"X\": 1, \"Y\": 1, \"Width\": 1, \"Height\": 1}],")));

        Assert.Contains(problems, problem => problem.Contains("highlight", StringComparison.Ordinal) && problem.Contains("CRT reads", StringComparison.Ordinal));
    }

    [Fact]
    public void A_sidecar_that_does_not_parse_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree => tree.EditSidecar(WorkbookA, json => json[..(json.Length / 2)]));

        Assert.Contains(problems, problem => problem.Contains("could not be read", StringComparison.Ordinal));
    }

    [Fact]
    public void A_listed_board_whose_workbook_is_missing_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree => File.Delete(tree.Full(WorkbookB)));

        Assert.Contains(problems, problem => problem.Contains(WorkbookB, StringComparison.Ordinal) && problem.Contains("not in the data", StringComparison.Ordinal));
    }

    // A board of the generation being converted that nobody listed would vanish.
    [Fact]
    public void A_board_of_the_newest_generation_the_main_workbook_does_not_list_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree =>
            tree.WriteBoard("Commodore/VIC-20/250403/Data VIC20 250403 v3.0.0.xlsx", ExcelTree.SmallBoard("Commodore/VIC-20/250403/Top.png")));

        Assert.Contains(problems, problem => problem.Contains("Data VIC20 250403 v3.0.0.xlsx", StringComparison.Ordinal) && problem.Contains("not listed", StringComparison.Ordinal));
    }

    [Fact]
    public void An_oscilloscope_row_with_values_and_no_model_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree => tree.EditWorkbook(ExcelTree.MasterName, package =>
        {
            ExcelWorksheet sheet = package.Workbook.Worksheets[MasterWorkbookSchema.OscilloscopeSheetName];
            sheet.Cells[20, 3].Value = "5025";
        }));

        Assert.Contains(problems, problem => problem.Contains("no series or model", StringComparison.Ordinal));
    }

    // "Files" and "MiniPro" are the 4.0 tree's own folders, beside the boards' - never a manufacturer.
    [Fact]
    public void A_board_filed_under_a_folder_the_tree_keeps_for_itself_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree =>
        {
            tree.WriteBoard("Files/Disk/Drive/Data Disk Drive v3.0.0.xlsx", ExcelTree.SmallBoard("Files/Disk/Drive/Top.png"));
            tree.WriteFile("Files/Disk/Drive/Top.png", "a drive");
            tree.WriteMaster(
                [
                    ("Commodore 64", "250407", WorkbookA, "Long board"),
                    ("Commodore 64", "250425", WorkbookB, "Short board"),
                    ("Disk", "Drive", "Files/Disk/Drive/Data Disk Drive v3.0.0.xlsx", ""),
                ],
                []);
        });

        Assert.Contains(problems, problem => problem.Contains("[Files]", StringComparison.Ordinal) && problem.Contains("keeps for itself", StringComparison.Ordinal));
    }

    [Fact]
    public void A_tree_with_no_versioned_main_workbook_is_a_problem()
    {
        IReadOnlyList<string> problems = ProblemsOf(tree => File.Delete(tree.Full(ExcelTree.MasterName)));

        Assert.Contains(problems, problem => problem.Contains("no versioned main workbook", StringComparison.Ordinal));
    }

    [Fact]
    public void A_missing_data_folder_is_a_problem()
    {
        DataTreeConversionPlan plan = DataTreeConversion.Plan(Path.Combine(Path.GetTempPath(), "crt-no-such-folder-" + Guid.NewGuid().ToString("N")));

        Assert.False(plan.CanWrite);
        Assert.Contains(plan.Problems, problem => problem.Contains("could not be found", StringComparison.Ordinal));
    }

    // ---- writing ----------------------------------------------------------------------------

    [Fact]
    public void Write_puts_the_plan_into_the_tree_byte_for_byte()
    {
        using var output = new TempWorkspace();
        string target = Path.Combine(output.Root, "app-data-v4");

        DataTreeWriteResult result = DataTreeConversion.Write(this.Plan, target, publicBaseUrl: null);

        string content = Path.Combine(target, DataTreeLayout.ContentFolder);

        foreach ((string relative, DataTreeSourceFile source) in this.Plan.StoredFiles.Concat(this.Plan.CopiedFiles))
            Assert.Equal(File.ReadAllBytes(source.FullPath), File.ReadAllBytes(Path.Combine(content, relative)));

        foreach (ConvertedBoard board in this.Plan.Boards)
            Assert.Equal(DataTreeJson.Write(board.Document), File.ReadAllText(Path.Combine(content, DataTreeLayout.BoardFilePath(board.BoardId))));

        Assert.Equal(DataTreeJson.Write(this.Plan.Main), File.ReadAllText(Path.Combine(content, DataTreeLayout.MainFileName)));
        Assert.Equal(this.Plan.StoredFiles.Count + this.Plan.CopiedFiles.Count + this.Plan.Boards.Count + 1, result.Written);
        Assert.False(File.Exists(Path.Combine(target, DataTreeLayout.ManifestFileName)));
    }

    // The release day's run over a tree written before keeps what is right and changes nothing.
    [Fact]
    public void Writing_the_same_plan_again_keeps_every_file()
    {
        using var output = new TempWorkspace();
        string target = Path.Combine(output.Root, "app-data-v4");

        DataTreeConversion.Write(this.Plan, target, publicBaseUrl: null);
        DataTreeWriteResult again = DataTreeConversion.Write(this.Plan, target, publicBaseUrl: null);

        Assert.Equal(0, again.Written);
        Assert.Equal(0, again.Removed);
    }

    // A file the tree held that the new plan does not - a test board, a replaced picture - goes.
    [Fact]
    public void A_file_the_plan_does_not_hold_is_removed_from_the_tree()
    {
        using var output = new TempWorkspace();
        string target = Path.Combine(output.Root, "app-data-v4");

        DataTreeConversion.Write(this.Plan, target, publicBaseUrl: null);

        string stray = Path.Combine(target, DataTreeLayout.ContentFolder, "Test", "Board", "Only", "board.json");
        Directory.CreateDirectory(Path.GetDirectoryName(stray)!);
        File.WriteAllText(stray, "{}");

        DataTreeWriteResult again = DataTreeConversion.Write(this.Plan, target, publicBaseUrl: null);

        Assert.Equal(1, again.Removed);
        Assert.False(File.Exists(stray));
        Assert.False(Directory.Exists(Path.Combine(target, DataTreeLayout.ContentFolder, "Test")));
    }

    [Fact]
    public void With_its_public_address_the_tree_gets_a_manifest_of_every_file()
    {
        using var output = new TempWorkspace();
        string target = Path.Combine(output.Root, "app-data-v4");

        DataTreeWriteResult result = DataTreeConversion.Write(this.Plan, target, "https://example.org/app-data-v4/Data");

        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(target, DataTreeLayout.ManifestFileName)));
        List<string> files = manifest.RootElement.EnumerateArray().Select(entry => entry.GetProperty("file").GetString()!).ToList();

        Assert.Equal(result.ManifestEntries, files.Count);
        Assert.Contains(DataTreeLayout.MainFileName, files);
        Assert.Contains(DataTreeLayout.BoardFilePath(BoardA), files);
        Assert.Contains(manifest.RootElement.EnumerateArray(), entry => entry.GetProperty("url").GetString() == "https://example.org/app-data-v4/Data/Commodore/C64/250407/board.json");
    }

    // Leaving an old manifest beside new files would send every CRT the wrong checksums.
    [Fact]
    public void A_tree_with_a_manifest_is_not_written_without_its_public_address()
    {
        using var output = new TempWorkspace();
        string target = Path.Combine(output.Root, "app-data-v4");

        DataTreeConversion.Write(this.Plan, target, "https://example.org/app-data-v4/Data");

        Assert.Throws<InvalidOperationException>(() => DataTreeConversion.Write(this.Plan, target, publicBaseUrl: null));
    }

    [Fact]
    public void A_plan_with_a_problem_is_not_written()
    {
        using ExcelTree tree = ExcelTree.Standard();
        File.Delete(tree.Full(WorkbookB));
        DataTreeConversionPlan spoiled = DataTreeConversion.Plan(tree.DataRoot);

        using var output = new TempWorkspace();
        string target = Path.Combine(output.Root, "app-data-v4");

        Assert.Throws<InvalidOperationException>(() => DataTreeConversion.Write(spoiled, target, publicBaseUrl: null));
        Assert.False(Directory.Exists(target));
    }

    // The source is only ever read - writing into it, or around it, is refused before anything.
    [Fact]
    public void Writing_into_or_around_the_source_is_refused()
    {
        string data = this.thisTree.Tree.DataRoot;

        Assert.Throws<InvalidOperationException>(() => DataTreeConversion.Write(this.Plan, Path.Combine(data, "converted"), publicBaseUrl: null));
        Assert.Throws<InvalidOperationException>(() => DataTreeConversion.Write(this.Plan, data, publicBaseUrl: null));
        Assert.Throws<InvalidOperationException>(() => DataTreeConversion.Write(this.Plan, Path.GetDirectoryName(data)!, publicBaseUrl: null));
    }

    // A folder holding anything but a tree this conversion wrote is not cleared to make one.
    [Fact]
    public void A_folder_holding_other_data_is_refused()
    {
        using var output = new TempWorkspace();
        output.WriteFile("app-data-v4/Data/photo.png", "holiday");

        Assert.Throws<InvalidOperationException>(() => DataTreeConversion.Write(this.Plan, Path.Combine(output.Root, "app-data-v4"), publicBaseUrl: null));
        Assert.True(File.Exists(output.Path_("app-data-v4", "Data", "photo.png")));
    }

    // An Excel tree is never mistaken for a converted one, even holding a 4.0 main file.
    [Fact]
    public void A_folder_holding_workbooks_is_refused()
    {
        using var output = new TempWorkspace();
        string target = Path.Combine(output.Root, "app-data-v4");

        DataTreeConversion.Write(this.Plan, target, publicBaseUrl: null);
        output.WriteFile("app-data-v4/Data/Classic-Repair-Toolbox.v3.0.0.xlsx", "not really");

        Assert.Throws<InvalidOperationException>(() => DataTreeConversion.Write(this.Plan, target, publicBaseUrl: null));
    }

    // A source file changed after the plan was made is not copied as if it were the planned one.
    [Fact]
    public void A_source_file_changed_since_the_plan_stops_the_write()
    {
        using ExcelTree tree = ExcelTree.Standard();
        DataTreeConversionPlan plan = DataTreeConversion.Plan(tree.DataRoot);
        File.WriteAllText(tree.Full(BoardA + "/Board layout.png"), "repainted");

        using var output = new TempWorkspace();

        Assert.Throws<IOException>(() => DataTreeConversion.Write(plan, Path.Combine(output.Root, "app-data-v4"), publicBaseUrl: null));
    }

    // ---- the tree builder -------------------------------------------------------------------

    // An Excel-era tree in a temporary folder, written by CRT's own writers.
    public sealed class ExcelTree : IDisposable
    {
        public const string MasterName = "Classic-Repair-Toolbox.v3.0.0.xlsx";

        private readonly TempWorkspace thisSpace = new();

        static ExcelTree()
        {
            ExcelPackage.License.SetNonCommercialPersonal("Classic Repair Toolbox tests");
        }

        public string DataRoot => Path.Combine(this.thisSpace.Root, "Data");

        public string Full(string relative) => this.thisSpace.Path_(["Data", .. relative.Split('/')]);

        public void WriteFile(string relative, string content) => File.WriteAllText(this.Full(relative), content);

        public void WriteBoard(
            string workbook,
            BoardData data,
            IReadOnlyList<KiCadCalibrationEntry>? calibrations = null,
            IReadOnlyList<SchematicRenameRecord>? renames = null)
        {
            string path = this.Full(workbook);

            BoardWorkbookWriter.Write(path, data);
            BoardSidecarWriter.Write(path, data.ComponentHighlights, calibrations ?? [], renames ?? []);
        }

        // `fileName`: another main-format workbook - a 3.x "_UserContribution" one, say.
        public void WriteMaster(
            IEnumerable<(string Hardware, string Board, string Workbook, string Notes)> boards,
            IEnumerable<string[]> scopes,
            string fileName = MasterName)
        {
            using ExcelPackage package = EpplusLicense.NewPackage();

            ExcelWorksheet listing = package.Workbook.Worksheets.Add(MasterWorkbookSchema.SheetName);
            listing.Cells[1, 1].Value = "# Classic Repair Toolbox";
            listing.Cells[3, 1].Value = "# Revision date: 2026-October-11";
            listing.Cells[5, 1].Value = "Hardware and boards";
            listing.Cells[6, 1].Value = MasterWorkbookSchema.ColHardwareName;
            listing.Cells[6, 2].Value = MasterWorkbookSchema.ColBoardName;
            listing.Cells[6, 3].Value = MasterWorkbookSchema.ColExcelDataFile;
            listing.Cells[6, 4].Value = MasterWorkbookSchema.ColHardwareNotes;

            int row = 7;

            foreach ((string hardware, string board, string workbook, string notes) in boards)
            {
                listing.Cells[row, 1].Value = hardware;
                listing.Cells[row, 2].Value = board;
                listing.Cells[row, 3].Value = workbook;
                listing.Cells[row, 4].Value = notes;
                row++;
            }

            ExcelWorksheet oscilloscopes = package.Workbook.Worksheets.Add(MasterWorkbookSchema.OscilloscopeSheetName);
            oscilloscopes.Cells[3, 1].Value = "# Revision date: 2026-October-4";

            for (int column = 0; column < MasterWorkbookSchema.OscilloscopeColumns.Count; column++)
                oscilloscopes.Cells[6, column + 1].Value = MasterWorkbookSchema.OscilloscopeColumns[column];

            row = 7;

            foreach (string[] scope in scopes)
            {
                for (int column = 0; column < scope.Length; column++)
                    oscilloscopes.Cells[row, column + 1].Value = scope[column];

                row++;
            }

            package.SaveAs(new FileInfo(this.Full(fileName)));
        }

        public void EditWorkbook(string relative, Action<ExcelPackage> change)
        {
            string path = this.Full(relative);
            using var memory = new MemoryStream(File.ReadAllBytes(path));
            using ExcelPackage package = EpplusLicense.OpenPackage(memory);

            change(package);
            package.SaveAs(new FileInfo(path));
        }

        public void EditSidecar(string workbook, Func<string, string> change)
        {
            string path = this.Full(Path.ChangeExtension(workbook, ".json"));
            File.WriteAllText(path, change(File.ReadAllText(path)));
        }

        public static int HeaderRow(ExcelWorksheet sheet)
        {
            for (int row = 1; row <= sheet.Dimension.End.Row; row++)
            {
                if (WorkbookCellText.Trimmed(sheet, row, 1) is BoardWorkbookSchema.ColBoardLabel or BoardWorkbookSchema.ColCategory)
                    return row;
            }

            throw new InvalidOperationException($"No header row in [{sheet.Name}].");
        }

        public static BoardData SmallBoard(string image) => new()
        {
            RevisionDate = "2026-October-4",
            Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = image }],
            Components = [new ComponentEntry { BoardLabel = "U7", FriendlyName = "Inverter", TechnicalNameOrValue = "7406", Category = "IC" }],
        };

        // The standard tree - see StandardTree.
        public static ExcelTree Standard()
        {
            var tree = new ExcelTree();

            tree.WriteFile(BoardA + "/Board layout.png", "layout of 250407");
            tree.WriteFile(BoardA + "/Scope baseline/U1_1_PAL.png", "pin 1 of U1");
            tree.WriteFile(BoardA + "/KiCad data/250407_.kicad_pcb", "(kicad_pcb)");
            tree.WriteFile(BoardA + "/KiCad data/Pages/vic.kicad_sch", "(kicad_sch)");
            tree.WriteFile(BoardA + "/Old only.png", "only the old generation shows this");
            tree.WriteFile(BoardB + "/Top.png", "top of 250425");
            tree.WriteFile(BoardB + "/Service manual.pdf", "%PDF manual");
            tree.WriteFile("Commodore/Shared files/Component images/7406.png", "7406 pinout");
            tree.WriteFile("Generic shared files/Component local files/7406.pdf", "%PDF 7406");
            tree.WriteFile("Generic shared files/MiniPro/IC tests/Catalogue/7406.json", "{}");

            tree.WriteBoard(
                WorkbookA,
                new BoardData
                {
                    RevisionDate = "2026-October-4",
                    HardwareName = "Commodore 64",
                    BoardName = "250407",
                    Schematics = [new BoardSchematicEntry { SchematicName = "Board layout", SchematicImageFile = BoardA + "/Board layout.png", CadName = "250407_ - PCB Top" }],
                    Components =
                    [
                        new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", TechnicalNameOrValue = "6510", Category = "IC" },
                        new ComponentEntry { BoardLabel = "U7", FriendlyName = "Inverter", TechnicalNameOrValue = "7406", Category = "IC" },
                    ],
                    ComponentImages =
                    [
                        new ComponentImageEntry { BoardLabel = "U7", Name = "Pinout", File = "Commodore/Shared files/Component images/7406.png" },
                        new ComponentImageEntry { BoardLabel = "U1", Pin = "1", Name = "PHI0", File = BoardA + "/Scope baseline/U1_1_PAL.png", TimeDiv = "1uS", VoltsDiv = "2V", TriggerLevelVolts = "1.4V" },
                        new ComponentImageEntry { BoardLabel = "U1", Name = "Pinout", Note = "Same as 6510" },
                    ],
                    ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Board layout", BoardLabel = "U1", X = "100", Y = "200", Width = "30", Height = "40" }],
                    ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U7", Name = "Datasheet", File = "Generic shared files/Component local files/7406.pdf" }],
                    BoardLocalFiles = [new BoardLocalFileEntry { Category = "Manuals", Name = "Service manual", File = BoardB + "/Service manual.pdf" }],
                },
                [new KiCadCalibrationEntry { SchematicName = "Board layout", CadName = "250407_ - PCB Top", OffsetX = 10, OffsetY = 20, ScaleX = 0.75, ScaleY = 0.5 }],
                [new SchematicRenameRecord("Top (replica)", "Board layout", BoardA + "/Board layout.png", "2026-October-10")]);

            tree.WriteBoard(
                WorkbookB,
                new BoardData
                {
                    RevisionDate = "2026-October-4",
                    Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = BoardB + "/Top.png" }],
                    Components = [new ComponentEntry { BoardLabel = "U7", FriendlyName = "Inverter", TechnicalNameOrValue = "7406", Category = "IC" }],
                    ComponentLocalFiles = [new ComponentLocalFileEntry { BoardLabel = "U7", Name = "Datasheet", File = "Generic shared files/Component local files/7406.pdf" }],
                    BoardLocalFiles = [new BoardLocalFileEntry { Category = "Manuals", Name = "Service manual", File = BoardB + "/Service manual.pdf" }],
                });

            // The older generation, which only older CRTs read.
            tree.WriteBoard(BoardA + "/Data C64 250407 v2.0.0.xlsx", SmallBoard(BoardA + "/Old only.png"));

            tree.WriteMaster(
                [
                    ("Commodore 64", "250407", WorkbookA, "Long board"),
                    ("Commodore 64", "250425", WorkbookB, "Short board"),
                ],
                [
                    ["Rigol", "DHO800", "5555", "100ms", "*IDN?", "", "", "", "", "", "", "", "", "", "", "", ":TIMEBASE:SCALE {0}", "", "", "", "2nS,5nS", "1V,2V"],
                    ["", "DHO900", "5555"],
                ]);

            return tree;
        }

        public void Dispose() => this.thisSpace.Dispose();
    }
}
