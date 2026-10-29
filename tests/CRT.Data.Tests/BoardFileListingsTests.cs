using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardFileListings - the board files' lists as one lookup reads them. A caller that reads a
// board's whole document (for its rows) keeps its list of files from that same read, so a lookup
// after it reads the file no second time (code review, 2026-10-29: a device's history parsed each
// other board's file twice).
// ###########################################################################################
public sealed class BoardFileListingsTests
{
    [Fact]
    public void A_document_read_whole_keeps_its_list_for_the_lookups_after_it()
    {
        using var workspace = new TempWorkspace();
        string folder = workspace.Path_("Data", "Commodore", "C64", "250407");
        BoardDocument.StoredFileRow top = StoredFiles.Put(folder, "Images/top.png", "top");
        string boardFile = StoredFiles.WriteBoard(Path.Combine(folder, DataTreeLayout.BoardFileName), [top]);

        var listings = new BoardFileListings();

        BoardDocument document = Assert.IsType<BoardDocument>(listings.ReadDocument(boardFile));
        Assert.Equal(["Images/top.png"], document.Files.Select(file => file.Name));

        Assert.Equal(top.Sha256, listings.ListedIn(boardFile)!["images/TOP.png"].Sha256);
        Assert.Equal(1, listings.ReadsForTests);
    }

    // No board file there, or one that cannot be read: no document, and nothing listed.
    [Fact]
    public void No_board_file_is_no_document()
    {
        using var workspace = new TempWorkspace();
        var listings = new BoardFileListings();

        Assert.Null(listings.ReadDocument(workspace.Path_("missing.json")));
        Assert.Null(listings.ReadDocument(workspace.WriteFile("broken.json", "{ not json")));
        Assert.Null(listings.ListedIn(workspace.Path_("missing.json")));
    }
}
