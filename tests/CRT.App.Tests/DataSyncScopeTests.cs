using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What a data check downloads follows "Visible hardware, boards and schematics" (owner request, 2026-11-14: "it will not
// download any data for the non-selected entries"): DataSyncScope, and DataManager.BoardsWithoutData, which tells a board
// checked again whose data is not here.
// ###########################################################################################
[Collection("DataManager")]
public sealed class DataSyncScopeTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose()
    {
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.json");
        this.thisWorkspace.Dispose();
    }

    private static HardwareBoardEntry Board(string hardware, string board, bool draftOnly = false) => new()
    {
        HardwareName = hardware,
        BoardName = board,
        ExcelDataFile = $"{hardware}/{board}/board.json".Replace(' ', '_'),
        IsDraftOnly = draftOnly,
    };

    private static string Normalise(string path) => path.Trim().Replace('\\', '/').TrimStart('/');

    private static IReadOnlySet<string> Unchecked(params string[] keys) => new HashSet<string>(keys, CatalogueVisibility.KeyComparer);

    [Fact]
    public void With_nothing_unchecked_every_published_board_is_downloaded()
    {
        DataSyncScope scope = DataSyncScope.Of(
            [Board("Commodore 64", "250407"), Board("Commodore 64", "250469"), Board("Amiga 500", "Rev 6A")],
            Unchecked(),
            Normalise);

        Assert.Equal(3, scope.BoardFiles.Count);
        Assert.Empty(scope.SkippedBoardFiles);
        Assert.False(scope.SkipsAny);
    }

    // A board unchecked, or its hardware, is left out - each board of that hardware with it, its own checkbox as it was.
    [Fact]
    public void A_board_unchecked_or_its_hardware_is_left_out()
    {
        HardwareBoardEntry long64 = Board("Commodore 64", "250407");
        HardwareBoardEntry short64 = Board("Commodore 64", "250469");
        HardwareBoardEntry amiga = Board("Amiga 500", "Rev 6A");

        DataSyncScope boardOff = DataSyncScope.Of([long64, short64, amiga], Unchecked("Commodore 64|250469"), Normalise);

        Assert.True(boardOff.BoardFiles.SetEquals([long64.ExcelDataFile, amiga.ExcelDataFile]));
        Assert.Equal([short64.ExcelDataFile], boardOff.SkippedBoardFiles);

        DataSyncScope hardwareOff = DataSyncScope.Of([long64, short64, amiga], Unchecked("commodore 64"), Normalise);

        Assert.Equal([amiga.ExcelDataFile], hardwareOff.BoardFiles);
        Assert.Equal(2, hardwareOff.SkippedBoardFiles.Count);
    }

    // A board only a draft has is no file on the server, checked or not.
    [Fact]
    public void A_board_only_a_draft_has_is_never_downloaded()
    {
        DataSyncScope scope = DataSyncScope.Of([Board("Commodore 64", "Mine", draftOnly: true)], Unchecked("Commodore 64|Mine"), Normalise);

        Assert.Empty(scope.BoardFiles);
        Assert.Empty(scope.SkippedBoardFiles);
    }

    // ###########################################################################################
    // The background part: the lists are the launch's, a stored file comes only when a checked board or a draft names it,
    // and everything else - the MiniPro tests - always. Null for the stored files wanted is all of them.
    // ###########################################################################################
    [Fact]
    public void In_the_background_a_stored_file_comes_only_when_wanted()
    {
        static bool IsList(string path) => path.EndsWith("board.json", StringComparison.Ordinal) || path.StartsWith("Classic-Repair-Toolbox.v", StringComparison.Ordinal);

        var wanted = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Files/ab/abcd.png" };

        Assert.True(DataSyncScope.DownloadsInBackground("Files/ab/abcd.png", IsList, wanted));
        Assert.False(DataSyncScope.DownloadsInBackground("Files/cd/cdef.png", IsList, wanted));
        Assert.True(DataSyncScope.DownloadsInBackground("MiniPro/IC tests/7400.json", IsList, wanted));
        Assert.False(DataSyncScope.DownloadsInBackground("Commodore/C64/250407/board.json", IsList, wanted));

        Assert.True(DataSyncScope.DownloadsInBackground("Files/cd/cdef.png", IsList, storedWanted: null));
        Assert.True(DataSyncScope.IsStoredFile("files/ab/abcd.png"));
        Assert.False(DataSyncScope.IsStoredFile("MiniPro/IC tests/Files/x.json"));
    }

    // ###########################################################################################
    // A board checked again has its data here only with its board file AND every stored file it names - a picture of it
    // left out while it was unchecked is data missing too. A board only a draft has has none to miss.
    // ###########################################################################################
    [Fact]
    public void A_board_without_its_board_file_or_a_stored_file_has_no_data_here()
    {
        string dataRoot = this.thisWorkspace.Path_("Data");
        var whole = new MainFiles.Board("Commodore/C64/250407", "Commodore 64", "250407");
        var noPicture = new MainFiles.Board("Commodore/C64/250469", "Commodore 64", "250469");
        var noFile = new MainFiles.Board("Commodore/C128/310378", "Commodore 128", "310378");
        DataManager.LoadFrom(dataRoot, MainFiles.Write(dataRoot, whole, noPicture, noFile));

        BoardDocument.StoredFileRow top = StoredFiles.Put(dataRoot, "Images/Top.png", "top");
        StoredFiles.WriteBoard(DraftBoardSource.PublishedPathOf(dataRoot, whole.Identity), new BoardData(), top);

        BoardDocument.StoredFileRow bottom = StoredFiles.Put(dataRoot, "Images/Bottom.png", "bottom");
        StoredFiles.WriteBoard(DraftBoardSource.PublishedPathOf(dataRoot, noPicture.Identity), new BoardData(), bottom);
        File.Delete(StoredFiles.PathOf(dataRoot, bottom));

        List<HardwareBoardEntry> missing = DataManager.BoardsWithoutData(DataManager.HardwareBoards);

        Assert.Equal(["250469", "310378"], missing.Select(entry => entry.BoardName).Order(StringComparer.Ordinal));
    }
}
