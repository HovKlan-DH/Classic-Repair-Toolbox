using System.Security.Cryptography;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Converting a user's Excel-era draft into a CRT 4.0 draft, in place. Drafts are started by CRT's
// own seeder from a published Excel tree, which is also converted into a 4.0 tree beside them - the
// situation a 4.0 CRT meets at its first launch. What must hold: the draft's files carry the
// published board's names wherever the bytes or the path say they are the published file, so the
// draft shows the contributor's changes and nothing else; only bytes the 4.0 tree lacks are stored
// in the draft; the Excel-era draft is kept, whole, beside it; and an interrupted conversion picks
// up where it stopped.
// ###########################################################################################
public sealed class DraftConversionTests
{
    private const string BoardA = "Commodore/C64/250407";
    private const string WorkbookA = BoardA + "/Data C64 250407 v3.0.0.xlsx";

    // The published Excel tree, its 4.0 conversion, and a drafts root - as at a 4.0 first launch.
    private sealed class Situation : IDisposable
    {
        public DataTreeConversionTests.ExcelTree Excel { get; } = DataTreeConversionTests.ExcelTree.Standard();

        private readonly TempWorkspace thisSpace = new();

        public string Tree => Path.Combine(this.thisSpace.Root, "app-data-v4");

        public string Content => Path.Combine(this.Tree, DataTreeLayout.ContentFolder);

        public string Drafts => Path.Combine(this.thisSpace.Root, "Drafts");

        public string DraftFolder => Path.Combine(this.Drafts, "Commodore", "C64", "250407");

        public IReadOnlyCollection<string> BoardIds { get; private set; } = [];

        public Situation()
        {
            DataTreeConversionPlan plan = DataTreeConversion.Plan(this.Excel.DataRoot);
            DataTreeConversion.Write(plan, this.Tree, publicBaseUrl: null);
            this.BoardIds = plan.Main.Boards.Select(board => board.BoardId).ToList();
        }

        public void Seed() => LegacyDrafts.Seed(this.Drafts, this.Excel.DataRoot, WorkbookA);

        public void Edit(Func<BoardData, BoardData> change) => LegacyDrafts.Edit(this.Drafts, WorkbookA, change);

        public void WriteInDraft(string relative, string content)
        {
            string path = Path.Combine([this.DraftFolder, .. relative.Split('/')]);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, content);
        }

        public DraftConversionResult Convert() =>
            DraftConversion.Convert(this.DraftFolder, BoardA, this.Excel.DataRoot, this.Content, this.BoardIds);

        public BoardDocument Draft() => DataTreeJson.ReadBoard(File.ReadAllText(Path.Combine(this.DraftFolder, DataTreeLayout.BoardFileName)));

        public BoardDocument Published() => DataTreeJson.ReadBoard(File.ReadAllText(Path.Combine(this.Content, "Commodore", "C64", "250407", "board.json")));

        public void Dispose()
        {
            this.Excel.Dispose();
            this.thisSpace.Dispose();
        }
    }

    private static string Sha(string text) => Convert.ToHexStringLower(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(text)));

    private static BoardData WithImagesAndComponents(BoardData board, List<ComponentImageEntry> images, List<ComponentEntry> components) => new()
    {
        RevisionDate = board.RevisionDate,
        HardwareName = board.HardwareName,
        BoardName = board.BoardName,
        Schematics = board.Schematics,
        Components = components,
        ComponentImages = images,
        ComponentHighlights = board.ComponentHighlights,
        ComponentLocalFiles = board.ComponentLocalFiles,
        ComponentLinks = board.ComponentLinks,
        BoardLocalFiles = board.BoardLocalFiles,
        BoardLinks = board.BoardLinks,
        Credits = board.Credits,
        KiCadImportantSignals = board.KiCadImportantSignals,
    };

    // A draft nobody changed is the published board, file by file.
    [Fact]
    public void An_untouched_draft_names_every_file_as_the_published_board_does()
    {
        using var situation = new Situation();
        situation.Seed();

        DraftConversionResult result = situation.Convert();

        Assert.True(result.Converted, result.Failure);
        Assert.Equal(
            situation.Published().Files.Select(file => (file.Name, file.Sha256)),
            situation.Draft().Files.Select(file => (file.Name, file.Sha256)));
        Assert.Equal(DataTreeJson.Write(situation.Published()).Replace(situation.Published().RevisionDate, "x"), DataTreeJson.Write(situation.Draft()).Replace(situation.Draft().RevisionDate, "x"));
    }

    // Nothing the 4.0 tree holds is copied into the draft.
    [Fact]
    public void Bytes_the_tree_holds_are_not_stored_in_the_draft()
    {
        using var situation = new Situation();
        situation.Seed();

        situation.Convert();

        Assert.False(Directory.Exists(Path.Combine(situation.DraftFolder, DataTreeLayout.FilesFolder)));
    }

    // The contributor replaced the board layout picture: same name, new bytes, kept in the draft.
    [Fact]
    public void A_replaced_file_keeps_its_published_name_and_its_new_bytes_are_kept_in_the_draft()
    {
        using var situation = new Situation();
        situation.Seed();
        situation.WriteInDraft("Board layout.png", "a better photo of 250407");

        situation.Convert();

        BoardDocument.StoredFileRow layout = situation.Draft().Files.Single(file => file.Name == "Board layout.png");
        Assert.Equal(Sha("a better photo of 250407"), layout.Sha256);

        var resolver = new StoredFileResolver(situation.Draft(), [situation.DraftFolder, situation.Content]);
        Assert.Equal("a better photo of 250407", File.ReadAllText(resolver.PathOf("Board layout.png")!));
        Assert.StartsWith(situation.DraftFolder, resolver.PathOf("Board layout.png")!, StringComparison.Ordinal);
    }

    // A new picture filed in a shared folder, and one in the board's own: named afresh, stored in
    // the draft, never under a name the published board already uses for other bytes.
    [Fact]
    public void A_file_the_draft_added_is_named_afresh_and_stored_in_the_draft()
    {
        using var situation = new Situation();
        situation.Seed();
        situation.WriteInDraft("Commodore/Shared files/Component images/HotCPU.png", "a hot CPU");
        situation.WriteInDraft("Component images/7406.png", "the board's own 7406 picture");
        situation.Edit(board => WithImagesAndComponents(
            board,
            [
                .. board.ComponentImages,
                new ComponentImageEntry { BoardLabel = "U1", Name = "Hot", File = "Commodore/Shared files/Component images/HotCPU.png" },
                new ComponentImageEntry { BoardLabel = "U7", Name = "Own", File = BoardA + "/Component images/7406.png" },
            ],
            board.Components));

        situation.Convert();

        BoardDocument draft = situation.Draft();

        Assert.Equal("Component images/HotCPU.png", draft.ComponentImages.Single(row => row.Name == "Hot").File);
        Assert.Equal("Component images/7406 (2).png", draft.ComponentImages.Single(row => row.Name == "Own").File);
        Assert.Equal("Component images/7406.png", draft.ComponentImages.Single(row => row.BoardLabel == "U7" && row.Name == "Pinout").File);

        var resolver = new StoredFileResolver(draft, [situation.DraftFolder, situation.Content]);
        Assert.Equal("a hot CPU", File.ReadAllText(resolver.PathOf("Component images/HotCPU.png")!));
        Assert.Equal("the board's own 7406 picture", File.ReadAllText(resolver.PathOf("Component images/7406 (2).png")!));
        Assert.Equal("7406 pinout", File.ReadAllText(resolver.PathOf("Component images/7406.png")!));
    }

    [Fact]
    public void The_draft_s_own_changes_to_rows_are_carried()
    {
        using var situation = new Situation();
        situation.Seed();
        situation.Edit(board => WithImagesAndComponents(
            board,
            board.ComponentImages,
            [.. board.Components.Where(row => row.BoardLabel != "U7"), new ComponentEntry { BoardLabel = "U9", FriendlyName = "SID", TechnicalNameOrValue = "6581", Category = "IC", Description = "Sound" }]));

        situation.Convert();

        Assert.Equal(["U1", "U9"], situation.Draft().Components.Select(row => row.BoardLabel).Order());
        Assert.Equal("Sound", situation.Draft().Components.Single(row => row.BoardLabel == "U9").Description);
    }

    // The published board's KiCad files are the draft's too, and the draft's own import wins.
    [Fact]
    public void The_published_KiCad_files_are_listed_and_the_draft_s_own_import_wins()
    {
        using var situation = new Situation();
        situation.Seed();
        situation.WriteInDraft("KiCad data/250407_.kicad_pcb", "(kicad_pcb improved)");

        situation.Convert();

        BoardDocument draft = situation.Draft();

        Assert.Equal(Sha("(kicad_pcb improved)"), draft.Files.Single(file => file.Name == "KiCad data/250407_.kicad_pcb").Sha256);
        Assert.Equal(Sha("(kicad_sch)"), draft.Files.Single(file => file.Name == "KiCad data/Pages/vic.kicad_sch").Sha256);
    }

    // The base is what the draft was started from - the published board, by its published names.
    [Fact]
    public void The_base_is_the_published_board_the_draft_started_from()
    {
        using var situation = new Situation();
        situation.Seed();
        situation.WriteInDraft("Board layout.png", "a better photo of 250407");

        situation.Convert();

        BoardDocument draftBase = DataTreeJson.ReadBoard(File.ReadAllText(Path.Combine(situation.DraftFolder, DraftFolderLayout.DraftBaseFileName)));

        Assert.Equal(Sha("layout of 250407"), draftBase.Files.Single(file => file.Name == "Board layout.png").Sha256);
        Assert.Equal(["U1", "U7"], draftBase.Components.Select(row => row.BoardLabel));
        Assert.Equal("Board layout.png", draftBase.Schematics.Single().ImageFile);
    }

    // The 4.0 identity of a board is its board file ("<id>/board.json", DataTreeLayout.BoardFilePath)
    // - what the main file lists it as, and what every draft lookup builds the folder from.
    [Fact]
    public void The_marker_names_the_board_by_its_board_file_and_keeps_what_it_knew()
    {
        using var situation = new Situation();
        situation.Seed();
        DraftMarker before = DraftMarkerStore.Load(Path.Combine(situation.DraftFolder, DraftFolderLayout.DraftMarkerFileName))!;

        situation.Convert();

        DraftMarker after = DraftMarkerStore.Load(Path.Combine(situation.DraftFolder, DraftFolderLayout.DraftMarkerFileName))!;

        Assert.Equal(DataTreeLayout.BoardFilePath(BoardA), after.BoardKey);
        Assert.Equal(situation.DraftFolder, DraftFolderLayout.GetBoardFolder(situation.Drafts, after.BoardKey));
        Assert.Equal(before.BaseRevision, after.BaseRevision);
        Assert.Equal(before.CreatedUtc, after.CreatedUtc);
        Assert.Empty(after.BaseContentHash);
    }

    // Nothing is lost: every file of the Excel-era draft is kept, as it was.
    [Fact]
    public void The_Excel_era_draft_is_kept_whole_beside_the_new_one()
    {
        using var situation = new Situation();
        situation.Seed();
        situation.WriteInDraft("Board layout.png", "a better photo of 250407");

        var before = Directory.EnumerateFiles(situation.DraftFolder, "*", SearchOption.AllDirectories)
            .ToDictionary(path => Path.GetRelativePath(situation.DraftFolder, path), File.ReadAllBytes);

        situation.Convert();

        string kept = Path.Combine(situation.DraftFolder, DraftConversion.OriginalsFolder);

        foreach ((string relative, byte[] bytes) in before)
            Assert.Equal(bytes, File.ReadAllBytes(Path.Combine(kept, relative)));

        Assert.False(File.Exists(Path.Combine(situation.DraftFolder, "Data C64 250407 v3.0.0.xlsx")));
    }

    [Fact]
    public void A_converted_draft_is_left_alone()
    {
        using var situation = new Situation();
        situation.Seed();
        situation.Convert();
        string written = File.ReadAllText(Path.Combine(situation.DraftFolder, DataTreeLayout.BoardFileName));

        DraftConversionResult again = situation.Convert();

        Assert.True(again.AlreadyConverted);
        Assert.False(again.Converted);
        Assert.Equal(written, File.ReadAllText(Path.Combine(situation.DraftFolder, DataTreeLayout.BoardFileName)));
    }

    // Stopped while setting the old draft aside: part of it is set aside, part still beside it.
    [Fact]
    public void A_conversion_stopped_while_setting_the_old_draft_aside_finishes_it()
    {
        using var situation = new Situation();
        situation.Seed();
        situation.WriteInDraft("Board layout.png", "a better photo of 250407");

        string kept = Path.Combine(situation.DraftFolder, DraftConversion.OriginalsFolder);
        Directory.CreateDirectory(kept);
        File.Move(Path.Combine(situation.DraftFolder, "Data C64 250407 v3.0.0.xlsx"), Path.Combine(kept, "Data C64 250407 v3.0.0.xlsx"));

        DraftConversionResult result = situation.Convert();

        Assert.True(result.Converted, result.Failure);
        Assert.True(File.Exists(Path.Combine(kept, "Data C64 250407 v3.0.0.json")));
        Assert.True(File.Exists(Path.Combine(kept, "Board layout.png")));
        Assert.Equal(Sha("a better photo of 250407"), situation.Draft().Files.Single(file => file.Name == "Board layout.png").Sha256);
    }

    // Stopped while moving the new draft in: it is built again from what was set aside.
    [Fact]
    public void A_conversion_stopped_while_moving_the_new_draft_in_builds_it_again()
    {
        using var situation = new Situation();
        situation.Seed();
        situation.Convert();
        string expected = File.ReadAllText(Path.Combine(situation.DraftFolder, DataTreeLayout.BoardFileName));

        File.Delete(Path.Combine(situation.DraftFolder, DataTreeLayout.BoardFileName));
        Directory.CreateDirectory(Path.Combine(situation.DraftFolder, ".crt-converting", "Files"));

        DraftConversionResult result = situation.Convert();

        Assert.True(result.Converted, result.Failure);
        Assert.Equal(expected, File.ReadAllText(Path.Combine(situation.DraftFolder, DataTreeLayout.BoardFileName)));
        Assert.False(Directory.Exists(Path.Combine(situation.DraftFolder, ".crt-converting")));
    }

    // A file nowhere to be found: the draft still converts, keeps the row, and says so.
    [Fact]
    public void A_cited_file_found_nowhere_is_noted_and_the_draft_still_converts()
    {
        using var situation = new Situation();
        situation.Seed();
        situation.Edit(board => WithImagesAndComponents(
            board,
            [.. board.ComponentImages, new ComponentImageEntry { BoardLabel = "U1", Name = "Lost", File = BoardA + "/Component images/Lost.png" }],
            board.Components));

        DraftConversionResult result = situation.Convert();

        Assert.True(result.Converted, result.Failure);
        Assert.Contains(result.Notes, note => note.Contains("Lost.png", StringComparison.Ordinal));
        Assert.Equal("Component images/Lost.png", situation.Draft().ComponentImages.Single(row => row.Name == "Lost").File);
        Assert.DoesNotContain(situation.Draft().Files, file => file.Name.EndsWith("Lost.png", StringComparison.Ordinal));
    }

    // A board that exists only as a draft: no published names to agree with, every byte its own.
    [Fact]
    public void A_new_board_s_draft_names_its_files_afresh_and_keeps_every_byte()
    {
        using var situation = new Situation();
        const string NewBoard = "Commodore/VIC-20/250403";
        const string NewWorkbook = NewBoard + "/Data VIC20 250403 v3.0.0.xlsx";

        LegacyDrafts.CreateNewBoard(situation.Drafts, new NewBoardRegistration
        {
            HardwareName = "VIC-20",
            BoardName = "250403",
            ExcelDataFile = NewWorkbook,
        });

        string folder = Path.Combine(situation.Drafts, "Commodore", "VIC-20", "250403");
        File.WriteAllText(Path.Combine(folder, "Top.png"), "the VIC's top");
        LegacyDrafts.Edit(situation.Drafts, NewWorkbook, board => board.WithSchematics([new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = NewBoard + "/Top.png" }]));

        DraftConversionResult result = DraftConversion.Convert(folder, NewBoard, situation.Excel.DataRoot, situation.Content, situation.BoardIds);

        Assert.True(result.Converted, result.Failure);

        BoardDocument draft = DataTreeJson.ReadBoard(File.ReadAllText(Path.Combine(folder, DataTreeLayout.BoardFileName)));
        Assert.Equal("Top.png", draft.Schematics.Single().ImageFile);
        Assert.Equal("the VIC's top", File.ReadAllText(new StoredFileResolver(draft, [folder, situation.Content]).PathOf("Top.png")!));
        Assert.Equal(DataTreeLayout.BoardFilePath(NewBoard), DraftMarkerStore.Load(Path.Combine(folder, DraftFolderLayout.DraftMarkerFileName))!.BoardKey);
        Assert.NotNull(DraftMarkerStore.Load(Path.Combine(folder, DraftFolderLayout.DraftMarkerFileName))!.NewBoard);
    }

    [Fact]
    public void A_folder_that_is_no_draft_is_left_alone_saying_why()
    {
        using var situation = new Situation();
        situation.WriteInDraft("Notes.txt", "my own notes");

        DraftConversionResult result = situation.Convert();

        Assert.False(result.Converted);
        Assert.Contains("marker", result.Failure);
        Assert.Equal("my own notes", File.ReadAllText(Path.Combine(situation.DraftFolder, "Notes.txt")));
        Assert.False(Directory.Exists(Path.Combine(situation.DraftFolder, DraftConversion.OriginalsFolder)));
    }

    [Fact]
    public void A_draft_with_two_workbooks_is_left_alone_saying_why()
    {
        using var situation = new Situation();
        situation.Seed();
        File.Copy(Path.Combine(situation.DraftFolder, "Data C64 250407 v3.0.0.xlsx"), Path.Combine(situation.DraftFolder, "Copy.xlsx"));

        DraftConversionResult result = situation.Convert();

        Assert.False(result.Converted);
        Assert.Contains("2 workbooks", result.Failure);
        Assert.False(File.Exists(Path.Combine(situation.DraftFolder, DataTreeLayout.BoardFileName)));
    }

    // ---- the resolver ------------------------------------------------------------------------

    [Fact]
    public void The_resolver_tries_each_store_in_order_and_knows_only_the_board_s_names()
    {
        using var space = new TempWorkspace();
        string sha = Sha("bytes");
        var board = new BoardDocument { Files = [new() { Name = "Component images/U1.png", Sha256 = sha, Size = 5 }] };
        string stored = DataTreeLayout.StoredFilePath(sha, "U1.png");

        string draft = space.Path_("draft");
        string tree = space.Path_("tree");
        space.WriteFile("tree/" + stored, "bytes");

        var resolver = new StoredFileResolver(board, [draft, tree]);
        Assert.Equal(Path.Combine(tree, stored.Replace('/', Path.DirectorySeparatorChar)), resolver.PathOf("component images/u1.png"));

        space.WriteFile("draft/" + stored, "bytes");
        Assert.Equal(Path.Combine(draft, stored.Replace('/', Path.DirectorySeparatorChar)), resolver.PathOf("Component images/U1.png"));

        Assert.Null(resolver.PathOf("Component images/U2.png"));
        Assert.Null(resolver.PathOf(null));
        Assert.Null(new StoredFileResolver(board, [space.Path_("empty")]).PathOf("Component images/U1.png"));
    }
}
