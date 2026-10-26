using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What a CRT 4.0 launch takes over from CRT 3.x on the same machine (LocalDataMigration): the
// drafts 3.x made, converted; the boards a 3.x "_UserContribution" workbook listed, carried into
// the drafts folder ONCE for the import to take in; and the oscilloscope models the user added to
// 3.x's main workbook. The 3.x data folder is only ever read.
//
// "BoardData" collection: these write and read board workbooks and board files.
// ###########################################################################################
[Collection("BoardData")]
public sealed class LocalDataMigrationTests : IDisposable
{
    private const string OwnBoard = "Commodore/C128/310378 Open128";
    private const string OwnWorkbook = OwnBoard + "/Data C128 310378 Open128.xlsx";
    private const string ContributionName = "Classic-Repair-Toolbox.v3.0.0_UserContribution.xlsx";

    private readonly TempWorkspace thisSpace = new();

    public void Dispose() => this.thisSpace.Dispose();

    private string Drafts => Path.Combine(this.thisSpace.Root, "Drafts");

    private string DraftFolder(string boardId) => Path.Combine([this.Drafts, .. boardId.Split('/')]);

    // A 3.x data folder holding a board the user made the old way, listed by its
    // "_UserContribution" workbook - as 3.x's "Add a new board" left it before drafts existed.
    private static DataTreeConversionTests.ExcelTree TreeWithOwnBoard(params string[] extraFiles)
    {
        DataTreeConversionTests.ExcelTree tree = DataTreeConversionTests.ExcelTree.Standard();

        tree.WriteFile(OwnBoard + "/Top.png", "the replica's top");
        tree.WriteBoard(OwnWorkbook, DataTreeConversionTests.ExcelTree.SmallBoard(OwnBoard + "/Top.png"));

        foreach (string file in extraFiles)
            tree.WriteFile(file, "extra " + file);

        tree.WriteMaster([("Commodore 128", "310378 Open128", OwnWorkbook, "Replica")], [], ContributionName);

        return tree;
    }

    // ------------------------------------------------------------------ 3.x drafts

    [Fact]
    public void Every_Excel_era_draft_is_converted_once_and_a_4_0_draft_is_left_alone()
    {
        using DataTreeConversionTests.ExcelTree tree = DataTreeConversionTests.ExcelTree.Standard();
        string treeRoot = Path.Combine(this.thisSpace.Root, "app-data-v4");
        DataTreeConversionPlan plan = DataTreeConversion.Plan(tree.DataRoot);
        DataTreeConversion.Write(plan, treeRoot, publicBaseUrl: null);
        string content = Path.Combine(treeRoot, DataTreeLayout.ContentFolder);
        List<string> boardIds = plan.Main.Boards.Select(board => board.BoardId).ToList();

        LegacyDrafts.Seed(this.Drafts, tree.DataRoot, "Commodore/C64/250407/Data C64 250407 v3.0.0.xlsx");
        Assert.True(DraftSeeder.CreateNewBoard(this.Drafts, new NewBoardRegistration
        {
            HardwareName = "VIC-20",
            BoardName = "250403",
            ExcelDataFile = "Commodore/VIC-20/250403/board.json",
        }).Created);

        (string folder, DraftConversionResult result) = Assert.Single(
            LocalDataMigration.ConvertExcelDrafts(this.Drafts, tree.DataRoot, content, boardIds));

        Assert.Equal(this.DraftFolder("Commodore/C64/250407"), folder);
        Assert.True(result.Converted, result.Failure);
        Assert.True(File.Exists(Path.Combine(folder, DataTreeLayout.BoardFileName)));

        // The next launch finds nothing left to do.
        Assert.Empty(LocalDataMigration.ConvertExcelDrafts(this.Drafts, tree.DataRoot, content, boardIds));
    }

    // A draft that cannot be converted yet (its workbook unreadable) is said, and tried again.
    [Fact]
    public void A_draft_that_cannot_be_converted_is_answered_and_tried_again_next_time()
    {
        string folder = this.DraftFolder("Commodore/C64/250407");
        Directory.CreateDirectory(folder);
        File.WriteAllText(Path.Combine(folder, "Data C64 250407.xlsx"), "not a workbook");
        DraftMarkerStore.Save(Path.Combine(folder, DraftFolderLayout.DraftMarkerFileName), new DraftMarker { BoardKey = "Commodore/C64/250407/Data C64 250407.xlsx" });

        DraftConversionResult first = Assert.Single(LocalDataMigration.ConvertExcelDrafts(this.Drafts, string.Empty, string.Empty, [])).Result;
        DraftConversionResult second = Assert.Single(LocalDataMigration.ConvertExcelDrafts(this.Drafts, string.Empty, string.Empty, [])).Result;

        Assert.False(first.Converted);
        Assert.Contains("could not be read", first.Failure);
        Assert.False(second.Converted);
        Assert.True(File.Exists(Path.Combine(folder, "Data C64 250407.xlsx")));
    }

    [Fact]
    public void An_unmarked_folder_is_not_a_draft_and_is_left_to_the_import()
    {
        Directory.CreateDirectory(this.DraftFolder("Commodore/C64/250407"));

        Assert.Empty(LocalDataMigration.ConvertExcelDrafts(this.Drafts, string.Empty, string.Empty, []));
    }

    // ------------------------------------------------------------------ "_UserContribution" boards

    [Fact]
    public void A_board_a_UserContribution_workbook_listed_is_carried_into_the_drafts_folder()
    {
        using DataTreeConversionTests.ExcelTree tree = TreeWithOwnBoard();
        var notes = new List<string>();

        IReadOnlyList<string> carried = LocalDataMigration.CarryUserContributionBoards(this.Drafts, tree.DataRoot, ["Commodore/C64/250407"], notes);

        Assert.Equal([OwnBoard], carried);
        Assert.Empty(notes);

        string folder = this.DraftFolder(OwnBoard);
        Assert.True(File.Exists(Path.Combine(folder, "Data C128 310378 Open128.xlsx")));
        Assert.Equal("the replica's top", File.ReadAllText(Path.Combine(folder, "Top.png")));

        // A folder put there by hand - the import takes it in - and the 3.x folder untouched.
        Assert.False(File.Exists(Path.Combine(folder, DraftFolderLayout.DraftMarkerFileName)));
        Assert.True(File.Exists(tree.Full(OwnWorkbook)));
        Assert.DoesNotContain(Directory.GetDirectories(this.Drafts), path => Path.GetFileName(path).StartsWith(DraftFolderLayout.BookkeepingPrefix, StringComparison.Ordinal));
    }

    // ###########################################################################################
    // *** ONCE EACH. *** A carried board the contributor then discards must not come back at the
    // next start - the 3.x folder still holds it, so only the record says it was done.
    // ###########################################################################################
    [Fact]
    public void A_carried_board_is_not_carried_again_after_its_draft_is_discarded()
    {
        using DataTreeConversionTests.ExcelTree tree = TreeWithOwnBoard();

        Assert.Single(LocalDataMigration.CarryUserContributionBoards(this.Drafts, tree.DataRoot, [], []));

        Directory.Delete(this.DraftFolder(OwnBoard), recursive: true);

        Assert.Empty(LocalDataMigration.CarryUserContributionBoards(this.Drafts, tree.DataRoot, [], []));
        Assert.False(Directory.Exists(this.DraftFolder(OwnBoard)));
        Assert.True(File.Exists(Path.Combine(this.Drafts, LocalDataMigration.CarriedRecordFileName)));
    }

    // A board the 4.0 main file lists has been published - the published board is the one to show.
    [Fact]
    public void A_board_the_main_file_lists_is_not_carried()
    {
        using DataTreeConversionTests.ExcelTree tree = TreeWithOwnBoard();

        Assert.Empty(LocalDataMigration.CarryUserContributionBoards(this.Drafts, tree.DataRoot, [OwnBoard], []));
        Assert.False(Directory.Exists(this.DraftFolder(OwnBoard)));
    }

    [Fact]
    public void A_board_that_has_a_draft_folder_already_is_not_carried_and_says_so()
    {
        using DataTreeConversionTests.ExcelTree tree = TreeWithOwnBoard();
        Directory.CreateDirectory(this.DraftFolder(OwnBoard));
        var notes = new List<string>();

        Assert.Empty(LocalDataMigration.CarryUserContributionBoards(this.Drafts, tree.DataRoot, [], notes));
        Assert.Contains(notes, note => note.Contains("already has a draft folder", StringComparison.Ordinal));
        Assert.Empty(Directory.GetFileSystemEntries(this.DraftFolder(OwnBoard)));
    }

    // Another generation's workbook beside the board's own would make the folder two boards, which
    // the import refuses.
    [Fact]
    public void Another_generations_workbook_beside_it_stays_behind()
    {
        using DataTreeConversionTests.ExcelTree tree = TreeWithOwnBoard(
            OwnBoard + "/Data C128 310378 Open128 v2.0.0.xlsx",
            OwnBoard + "/Data C128 310378 Open128 v2.0.0.json",
            OwnBoard + "/Board local files/pinouts.xlsx");

        LocalDataMigration.CarryUserContributionBoards(this.Drafts, tree.DataRoot, [], []);

        string folder = this.DraftFolder(OwnBoard);
        Assert.Equal(
            ["Data C128 310378 Open128.json", "Data C128 310378 Open128.xlsx", "Top.png"],
            Directory.GetFiles(folder).Select(Path.GetFileName).Order(StringComparer.Ordinal));
        Assert.True(File.Exists(Path.Combine(folder, "Board local files", "pinouts.xlsx")));
    }

    [Fact]
    public void A_carried_board_is_taken_in_by_the_import_as_a_new_4_0_draft()
    {
        using DataTreeConversionTests.ExcelTree tree = TreeWithOwnBoard();
        LocalDataMigration.CarryUserContributionBoards(this.Drafts, tree.DataRoot, [], []);

        DraftFolderImportOutcome outcome = Assert.Single(DraftFolderImport.ImportUnmarkedFolders(
            this.Drafts, [], DateTimeOffset.UtcNow, oldDataRoot: tree.DataRoot, treeContentRoot: string.Empty));

        Assert.Equal(DraftFolderImportKind.NewBoard, outcome.Kind);

        BoardDocument draft = BoardFiles.ReadDocument(Path.Combine(this.DraftFolder(OwnBoard), DataTreeLayout.BoardFileName))!;
        Assert.Equal("Top.png", draft.Schematics.Single().ImageFile);
        Assert.Equal("the replica's top", File.ReadAllText(new StoredFileResolver(draft, [this.DraftFolder(OwnBoard)]).PathOf("Top.png")!));
    }

    [Fact]
    public void Without_a_CRT_3_data_folder_nothing_is_carried()
    {
        Assert.Empty(LocalDataMigration.CarryUserContributionBoards(this.Drafts, string.Empty, [], []));
        Assert.Empty(LocalDataMigration.CarryUserContributionBoards(this.Drafts, Path.Combine(this.thisSpace.Root, "nowhere"), [], []));
    }

    // ------------------------------------------------------------------ own scope models

    // The user's copy of 3.x's main workbook: one model as published, one changed, one added.
    private static (DataTreeConversionTests.ExcelTree Local, IReadOnlyList<MainDocument.OscilloscopeRow> Published) ScopeSituation()
    {
        using DataTreeConversionTests.ExcelTree publishedTree = DataTreeConversionTests.ExcelTree.Standard();
        IReadOnlyList<MainDocument.OscilloscopeRow> published = DataTreeConversion.ReadOscilloscopes(
            publishedTree.Full(DataTreeConversionTests.ExcelTree.MasterName), []);

        var local = new DataTreeConversionTests.ExcelTree();
        local.WriteMaster(
            [("Commodore 64", "250407", "Commodore/C64/250407/Data C64 250407 v3.0.0.xlsx", "")],
            [
                ["Rigol", "DHO800", "5555", "100ms", "*IDN? (my firmware)", "", "", "", "", "", "", "", "", "", "", "", ":TIMEBASE:SCALE {0}", "", "", "", "2nS,5nS", "1V,2V"],
                ["Rigol", "DHO900", "5555"],
                ["Siglent", "SDS1104X-E", "5025"],
            ]);

        return (local, published);
    }

    [Fact]
    public void The_models_the_user_added_or_changed_are_found_and_the_published_ones_are_not_repeated()
    {
        (DataTreeConversionTests.ExcelTree local, IReadOnlyList<MainDocument.OscilloscopeRow> published) = ScopeSituation();
        using (local)
        {
            IReadOnlyList<MainDocument.OscilloscopeRow> own = LocalDataMigration.OwnScopeModels(local.DataRoot, published, []);

            Assert.Equal(
                ["DHO800" + OwnOscilloscopeModels.MyVersionSuffix, "SDS1104X-E"],
                own.Select(model => model.SeriesOrModel));
        }
    }

    // ###########################################################################################
    // *** THE FILE IS MADE ONCE, EVEN EMPTY. *** Its being there says the search was made, so a
    // model the user deletes from it later is not found again in 3.x's workbook.
    // ###########################################################################################
    [Fact]
    public void The_own_models_file_is_made_at_the_first_start_and_read_from_then_on()
    {
        (DataTreeConversionTests.ExcelTree local, IReadOnlyList<MainDocument.OscilloscopeRow> published) = ScopeSituation();
        using (local)
        {
            string path = this.thisSpace.Path_(OwnOscilloscopeFile.FileName);

            Assert.Equal(2, OwnOscilloscopeFile.LoadOrFind(path, local.DataRoot, published, []).Count);
            Assert.True(File.Exists(path));

            OwnOscilloscopeFile.Write(path, []);

            Assert.Empty(OwnOscilloscopeFile.LoadOrFind(path, local.DataRoot, published, []));
        }
    }

    // ###########################################################################################
    // *** NOT WHEN 3.x's WORKBOOK COULD NOT BE READ (code review, 2026-10-25). *** A workbook held
    // open by another program, or broken, read as "no models" - and the empty file made then stopped
    // every later start from looking, so the user's own command sets were never carried over. Now
    // nothing is written, it is said, and the next start finds them once the workbook can be read.
    // ###########################################################################################
    [Fact]
    public void A_CRT_3_main_workbook_that_cannot_be_read_makes_no_file_and_is_read_again_at_the_next_start()
    {
        (DataTreeConversionTests.ExcelTree local, IReadOnlyList<MainDocument.OscilloscopeRow> published) = ScopeSituation();
        using (local)
        {
            string path = this.thisSpace.Path_(OwnOscilloscopeFile.FileName);
            string master = local.Full(DataTreeConversionTests.ExcelTree.MasterName);
            byte[] good = File.ReadAllBytes(master);
            var notes = new List<string>();

            File.WriteAllText(master, "not a workbook");

            Assert.Empty(OwnOscilloscopeFile.LoadOrFind(path, local.DataRoot, published, notes));
            Assert.False(File.Exists(path));
            Assert.Contains(notes, note => note.Contains("looked for again at the next start", StringComparison.Ordinal));

            File.WriteAllBytes(master, good);

            Assert.Equal(2, OwnOscilloscopeFile.LoadOrFind(path, local.DataRoot, published, []).Count);
            Assert.True(File.Exists(path));
        }
    }

    [Fact]
    public void Without_a_path_there_are_no_own_models()
    {
        Assert.Empty(OwnOscilloscopeFile.LoadOrFind(string.Empty, string.Empty, [], []));
    }

    [Fact]
    public void An_own_models_file_that_cannot_be_read_gives_none_and_says_so()
    {
        string path = this.thisSpace.WriteFile(OwnOscilloscopeFile.FileName, "{ not json");
        var notes = new List<string>();

        Assert.Empty(OwnOscilloscopeFile.LoadOrFind(path, string.Empty, [], notes));
        Assert.Single(notes);
        Assert.Equal("{ not json", File.ReadAllText(path));
    }

    [Fact]
    public void Without_a_CRT_3_main_workbook_the_file_is_made_empty()
    {
        string path = this.thisSpace.Path_(OwnOscilloscopeFile.FileName);

        Assert.Empty(OwnOscilloscopeFile.LoadOrFind(path, Path.Combine(this.thisSpace.Root, "nowhere"), [], []));
        Assert.Empty(DataTreeJson.ReadMain(File.ReadAllText(path)).Oscilloscopes);
    }
}
