using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// LoadedBoardPictures - the pictures of the board Main loaded, where each lies, worked out by that
// load. They answer only for the board they were worked out for: drawn against another board's
// data, a picture would be named by one board and found through another (code reviews, 2026-10-29).
// ###########################################################################################
public sealed class LoadedBoardPicturesTests
{
    private static BoardData Board() => new()
    {
        Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png" }],
    };

    [Fact]
    public void The_pictures_answer_for_their_own_board()
    {
        BoardData board = Board();
        var paths = new Dictionary<string, string> { ["Top"] = "stored/top.png" };

        Assert.Same(paths, new LoadedBoardPictures(board, paths).PathsFor(board));
    }

    // Another board - even one that reads the same, a load that came after - gets no answer.
    [Fact]
    public void The_pictures_give_no_answer_for_another_board()
    {
        var pictures = new LoadedBoardPictures(Board(), new Dictionary<string, string> { ["Top"] = "stored/top.png" });

        Assert.Null(pictures.PathsFor(Board()));
        Assert.Null(pictures.PathsFor(null));
    }
}
