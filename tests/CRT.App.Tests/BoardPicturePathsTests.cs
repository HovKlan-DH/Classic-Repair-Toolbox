using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardPicturePaths - where a board's pictures lie on disk, the Schematics tab's rule for every
// surface that draws or exports them. In CRT 4.0 a picture row names a file of its board, its
// bytes stored under their checksum; an Excel-era path under the data root still resolves.
// ###########################################################################################
public sealed class BoardPicturePathsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private const string Identity = "Commodore/C64/250407/board.json";

    // The board file at its place in the tree, listing `pictures`; returns where each is stored.
    private Dictionary<string, string> Publish(string dataRoot, BoardData board, params string[] pictures)
    {
        BoardDocument.StoredFileRow[] rows = pictures.Select(name => StoredFiles.Put(dataRoot, name, "bytes of " + name)).ToArray();
        StoredFiles.WriteBoard(DraftBoardSource.PublishedPathOf(dataRoot, Identity), board, rows);
        return rows.ToDictionary(row => row.Name, row => StoredFiles.PathOf(dataRoot, row));
    }

    private static BoardData Board(params (string Name, string File)[] pictures) => new()
    {
        Schematics = pictures.Select(picture => new BoardSchematicEntry { SchematicName = picture.Name, SchematicImageFile = picture.File }).ToList(),
    };

    // A 4.0 row's name: where its bytes are stored, through the board file that lists it.
    [Fact]
    public void A_pictures_name_is_found_where_its_bytes_are_stored()
    {
        string dataRoot = this.thisWorkspace.Path_("Data");
        BoardData board = Board(("Top", "Images/top.png"));
        Dictionary<string, string> stored = this.Publish(dataRoot, board, "Images/top.png");

        Assert.Equal(stored["Images/top.png"], BoardPicturePaths.PathOf(dataRoot, string.Empty, "Images/top.png", Identity));
    }

    // Not found anywhere: where an Excel-era row put it under the data root - a path the callers
    // then find missing, as before.
    [Fact]
    public void A_picture_found_nowhere_is_under_the_data_root()
    {
        string dataRoot = this.thisWorkspace.Path_("Data");

        Assert.Equal(
            Path.Combine(dataRoot, "Commodore", "Shared files", "top.png"),
            BoardPicturePaths.PathOf(dataRoot, string.Empty, "Commodore/Shared files/top.png", Identity));
    }

    // Every picture with a file by its name, capitals aside - the first row of a name.
    [Fact]
    public void Every_picture_with_a_file_is_found_by_its_name_the_first_of_a_name()
    {
        string dataRoot = this.thisWorkspace.Path_("Data");
        BoardData board = Board(("Top", "Images/top.png"), ("TOP", "Images/other.png"), ("Bottom", "Images/bottom.png"), ("Blank", " "));
        Dictionary<string, string> stored = this.Publish(dataRoot, board, "Images/top.png", "Images/other.png", "Images/bottom.png");

        Dictionary<string, string> byName = BoardPicturePaths.ByName(board, dataRoot, string.Empty, Identity);

        Assert.Equal(["Top", "Bottom"], byName.Keys);
        Assert.Equal(stored["Images/top.png"], byName["top"]);
        Assert.Equal(stored["Images/bottom.png"], byName["Bottom"]);
        Assert.Empty(BoardPicturePaths.ByName(null, dataRoot, string.Empty, Identity));
    }
}
