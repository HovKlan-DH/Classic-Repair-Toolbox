using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The "Resources" tab's files: those of the region shown, found as every file of a board is. Asked again
// at every PAL/NTSC click, on the UI thread - so one lookup's memory of the board file serves them all
// (code review, 2026-11-02: DraftFiles.ListedIn does not trust its cache in the two seconds after a save,
// and each of 30 files parsed the whole board file again before the click finished).
// ###########################################################################################
public sealed class MainResourcesTests : IDisposable
{
    private const string Identity = "Commodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string DraftFolder => Path.Combine(this.thisWorkspace.Root, "Drafts", "Commodore", "C64", "250407");

    private BoardDocument.StoredFileRow Stored(string name, byte[] bytes)
    {
        string sha = Convert.ToHexStringLower(System.Security.Cryptography.SHA256.HashData(bytes));
        string path = Path.Combine(this.DraftFolder, DataTreeLayout.StoredFilePath(sha, name).Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, bytes);

        return new BoardDocument.StoredFileRow { Name = name, Sha256 = sha, Size = bytes.Length };
    }

    [Fact]
    public void The_files_of_the_region_shown_are_found_with_one_read_of_the_board_file()
    {
        var board = new BoardData
        {
            BoardLocalFiles =
            [
                new BoardLocalFileEntry { Category = "Manuals", Name = "Every region", File = "Board local files/every.pdf" },
                new BoardLocalFileEntry { Category = "Manuals", Name = "PAL's", File = "Board local files/pal.pdf", Region = "PAL" },
                new BoardLocalFileEntry { Category = "Manuals", Name = "NTSC's", File = "Board local files/ntsc.pdf", Region = "NTSC" },
            ],
        };

        Directory.CreateDirectory(this.DraftFolder);
        BoardFiles.Write(Path.Combine(this.DraftFolder, DataTreeLayout.BoardFileName), BoardDocumentMapping.ToDocument(
            new BoardContent(board, [], []),
            name => name,
            [
                this.Stored("Board local files/every.pdf", [1]),
                this.Stored("Board local files/pal.pdf", [2]),
                this.Stored("Board local files/ntsc.pdf", [3]),
            ]));

        var listings = new BoardFileListings();
        string dataRoot = Path.Combine(this.thisWorkspace.Root, "Data");

        // CRT.Main rather than Main: the app also has a "Main" namespace (BoardKeyFullNameTests says why).
        List<ResourceItem> files = CRT.Main.ResourceFilesOf(board, "PAL", dataRoot, this.DraftFolder, MainResourcesTests.Identity, listings);

        Assert.Equal(["Every region", "PAL's"], files.Select(file => file.Name));
        Assert.All(files, file => Assert.True(File.Exists(file.Target), file.Target));
        Assert.Equal(1, listings.ReadsForTests);

        // The popup's files go by the same read - asked at each of its PAL/NTSC clicks, on the UI thread.
        var popupListings = new BoardFileListings();

        List<ComponentLocalFileItem> popupFiles = board.BoardLocalFiles
            .Select(row => ComponentInfoWindow.FoundFile(row.Name, row.File, dataRoot, this.DraftFolder, MainResourcesTests.Identity, popupListings))
            .ToList();

        Assert.All(popupFiles, file => Assert.True(File.Exists(file.FullPath), file.FullPath));
        Assert.All(popupFiles, file => Assert.Equal(this.DraftFolder, file.SourceRoot));
        Assert.Equal(1, popupListings.ReadsForTests);
    }

    // A board file's note (owner request, 2026-11-06) goes with it into the "Resources" tab, its ends
    // trimmed and its lines kept; a file with none has no note to show.
    [Fact]
    public void Each_file_carries_its_note_into_the_resources_tab()
    {
        var board = new BoardData
        {
            BoardLocalFiles =
            [
                new BoardLocalFileEntry { Category = "Manuals", Name = "Service manual", File = "Board local files/manual.pdf", Note = " Page 12 has the PLA.\nThe rest is the C64C. " },
                new BoardLocalFileEntry { Category = "Manuals", Name = "Dead test", File = "Board local files/dead.pdf" },
            ],
        };

        List<ResourceItem> files = CRT.Main.ResourceFilesOf(board, null, Path.Combine(this.thisWorkspace.Root, "Data"), this.DraftFolder, MainResourcesTests.Identity, new BoardFileListings());

        Assert.Equal("Page 12 has the PLA.\nThe rest is the C64C.", files[0].Note);
        Assert.True(files[0].HasNote);
        Assert.Equal(string.Empty, files[1].Note);
        Assert.False(files[1].HasNote);
    }
}
