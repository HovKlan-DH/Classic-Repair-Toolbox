using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftBoardRename - a new board's name changed (owner request, 2026-11-01: "how to edit those details
// again afterwards"). Its names are its identity and its draft's folder, so the folder moves whole,
// the marker and the board file following; anything that fails leaves everything as it was.
// ###########################################################################################
public sealed class DraftBoardRenameTests : IDisposable
{
    private const string Old = "Comodore/C64/250407/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => this.thisWorkspace.Path_("Drafts");

    public void Dispose() => this.thisWorkspace.Dispose();

    private static NewBoardRegistration Registration(string manufacturer, string hardware, string board, string notes = "") => new()
    {
        HardwareName = hardware,
        BoardName = board,
        HardwareNotes = notes,
        ExcelDataFile = NewBoardIdentity.BuildExcelDataFile(manufacturer, hardware, board),
        CreatedUtc = "2026-11-01T10:00:00.0000000Z",
    };

    private string Seed(string identity = Old, string hardware = "C64", string board = "250407")
    {
        DraftSeedResult seeded = DraftSeeder.CreateNewBoard(this.DraftsRoot, new NewBoardRegistration
        {
            HardwareName = hardware,
            BoardName = board,
            ExcelDataFile = identity,
        });

        Assert.True(seeded.Created, seeded.Reason);

        return seeded.BoardFolder;
    }

    // ###########################################################################################
    // *** A FILE HELD FOR A MOMENT IS NO REASON TO REFUSE (found by the Stop hook's run, 2026-11-02: a
    // rename refused at random under load). *** Windows will not move a folder with a file open in it,
    // and CRT's own reads off the UI thread or a virus scanner hold one that way for a moment. A rename
    // pressed then was told "something may have its folder open" - tried again it went through. Let go
    // of at the first pause, so nothing here hangs on how fast the machine is.
    // ###########################################################################################
    [Theory]
    [InlineData("Commodore", "Commodore 64", "KU-14194HB")]
    [InlineData("comodore", "c64", "250407")]
    public void A_file_held_for_a_moment_does_not_stop_the_rename(string manufacturer, string hardware, string board)
    {
        string folder = this.Seed();
        string heldFile = Path.Combine(folder, "picked.png");
        File.WriteAllText(heldFile, "picked");

        using var held = new FileStream(heldFile, FileMode.Open, FileAccess.Read, FileShare.None);

        NewBoardRegistration wanted = Registration(manufacturer, hardware, board);
        DraftRenameResult result = DraftBoardRename.Run(this.DraftsRoot, Old, wanted, _ => held.Dispose());

        Assert.True(result.Renamed, result.Reason);
        Assert.Equal("picked", File.ReadAllText(Path.Combine(DraftFolderLayout.GetBoardFolder(this.DraftsRoot, wanted.ExcelDataFile), "picked.png")));
    }

    // One still held after the tries is refused as before, with nothing changed.
    [Fact]
    public void A_file_held_for_good_refuses_it_and_changes_nothing()
    {
        Assert.SkipUnless(OperatingSystem.IsWindows(), "Only Windows refuses to move a folder with an open file in it.");

        string folder = this.Seed();
        string heldFile = Path.Combine(folder, "picked.png");
        File.WriteAllText(heldFile, "picked");

        int pauses = 0;
        DraftRenameResult result;

        using (new FileStream(heldFile, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            result = DraftBoardRename.Run(this.DraftsRoot, Old, Registration("Commodore", "Commodore 64", "KU-14194HB"), _ => pauses++);
        }

        Assert.False(result.Renamed);
        Assert.Contains("Nothing was changed", result.Reason);
        Assert.Equal(FileReplacer.MoveAttempts - 1, pauses);
        Assert.Equal("C64", BoardFiles.ReadDocument(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, Old))!.HardwareName);
        Assert.Equal(Old, DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(this.DraftsRoot, Old))!.BoardKey);
    }

    // ###########################################################################################
    // The folder moves whole - a stored file and one not taken in yet with it - the marker naming the
    // new identity and registration, the board file its new names, what the board holds kept.
    // ###########################################################################################
    [Fact]
    public void The_draft_moves_whole_and_its_marker_and_board_file_follow()
    {
        string folder = this.Seed();

        File.WriteAllText(Path.Combine(folder, "picked.png"), "picked");
        Directory.CreateDirectory(Path.Combine(folder, "Files", "ab"));
        File.WriteAllText(Path.Combine(folder, "Files", "ab", "stored.png"), "stored");

        BoardDocument document = BoardFiles.ReadDocument(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, Old))!;
        document.Components.Add(new BoardDocument.ComponentRow { BoardLabel = "U1" });
        BoardFiles.Write(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, Old), document);

        NewBoardRegistration wanted = Registration("Commodore", "Commodore 64", "KU-14194HB", notes: "Mine");
        DraftRenameResult result = DraftBoardRename.Run(this.DraftsRoot, Old, wanted);

        Assert.True(result.Renamed, result.Reason);

        string moved = DraftFolderLayout.GetBoardFolder(this.DraftsRoot, wanted.ExcelDataFile);
        Assert.Equal("picked", File.ReadAllText(Path.Combine(moved, "picked.png")));
        Assert.Equal("stored", File.ReadAllText(Path.Combine(moved, "Files", "ab", "stored.png")));
        Assert.False(Directory.Exists(folder));

        DraftMarker marker = DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(this.DraftsRoot, wanted.ExcelDataFile))!;
        Assert.Equal(wanted.ExcelDataFile, marker.BoardKey);
        Assert.Equal("Commodore 64", marker.NewBoard!.HardwareName);
        Assert.Equal("KU-14194HB", marker.NewBoard.BoardName);
        Assert.Equal("Mine", marker.NewBoard.HardwareNotes);

        BoardDocument read = BoardFiles.ReadDocument(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, wanted.ExcelDataFile))!;
        Assert.Equal("Commodore 64", read.HardwareName);
        Assert.Equal("KU-14194HB", read.BoardName);
        Assert.Equal("U1", Assert.Single(read.Components).BoardLabel);
    }

    // The folders the old path leaves empty go; one still holding another draft stays.
    [Fact]
    public void Only_the_folders_left_empty_go()
    {
        const string Sibling = "Comodore/C64/250425/board.json";

        this.Seed();
        string sibling = this.Seed(Sibling, board: "250425");

        Assert.True(DraftBoardRename.Run(this.DraftsRoot, Old, Registration("Commodore", "C64", "250407")).Renamed);

        Assert.True(Directory.Exists(sibling));
        Assert.True(Directory.Exists(Path.Combine(this.DraftsRoot, "Comodore", "C64")));

        Assert.True(DraftBoardRename.Run(this.DraftsRoot, Sibling, Registration("Commodore", "C64", "250425")).Renamed);

        Assert.False(Directory.Exists(Path.Combine(this.DraftsRoot, "Comodore")));
        Assert.True(Directory.Exists(this.DraftsRoot));
    }

    // Notes alone move nothing.
    [Fact]
    public void Notes_alone_stay_in_the_folder()
    {
        string folder = this.Seed();

        NewBoardRegistration wanted = Registration("Comodore", "C64", "250407", notes: "New notes");

        Assert.True(DraftBoardRename.Run(this.DraftsRoot, Old, wanted).Renamed);

        Assert.True(Directory.Exists(folder));
        Assert.Equal("New notes", DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(this.DraftsRoot, Old))!.NewBoard!.HardwareNotes);
    }

    // A change of capitals is a rename too, on a disk that sees the two as one folder and one that does not.
    [Fact]
    public void A_change_of_capitals_is_a_rename()
    {
        this.Seed();

        NewBoardRegistration wanted = Registration("Comodore", "C64", "250407A".ToLowerInvariant().Replace('a', 'A'));
        Assert.True(DraftBoardRename.Run(this.DraftsRoot, Old, wanted).Renamed);

        NewBoardRegistration capitals = Registration("Comodore", "C64", "250407a");
        DraftRenameResult result = DraftBoardRename.Run(this.DraftsRoot, wanted.ExcelDataFile, capitals);

        Assert.True(result.Renamed, result.Reason);

        string folder = DraftFolderLayout.GetBoardFolder(this.DraftsRoot, capitals.ExcelDataFile);
        Assert.Equal("250407a", Path.GetFileName(Directory.GetDirectories(Path.GetDirectoryName(folder)!).Single()));
        Assert.Equal(capitals.ExcelDataFile, DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(this.DraftsRoot, capitals.ExcelDataFile))!.BoardKey);
    }

    // Never over anything: a folder at the new place refuses it, nothing moved.
    [Fact]
    public void A_folder_at_the_new_place_refuses_the_rename()
    {
        string folder = this.Seed();

        NewBoardRegistration wanted = Registration("Commodore", "C64", "250407");
        string taken = DraftFolderLayout.GetBoardFolder(this.DraftsRoot, wanted.ExcelDataFile);
        Directory.CreateDirectory(taken);
        File.WriteAllText(Path.Combine(taken, "somebodys.txt"), "kept");

        DraftRenameResult result = DraftBoardRename.Run(this.DraftsRoot, Old, wanted);

        Assert.False(result.Renamed);
        Assert.Contains("Pick a different name", result.Reason, StringComparison.Ordinal);
        Assert.True(Directory.Exists(folder));
        Assert.Equal("kept", File.ReadAllText(Path.Combine(taken, "somebodys.txt")));
    }

    // A draft of a published board has no registration: its name is not its own to change.
    [Fact]
    public void A_draft_of_a_published_board_is_refused()
    {
        string folder = this.Seed();
        string marker = DraftFolderLayout.GetMarkerPath(this.DraftsRoot, Old);
        DraftMarkerStore.Save(marker, DraftMarkerStore.Load(marker)! with { NewBoard = null });

        DraftRenameResult result = DraftBoardRename.Run(this.DraftsRoot, Old, Registration("Commodore", "C64", "250407"));

        Assert.False(result.Renamed);
        Assert.True(Directory.Exists(folder));
    }

    // ###########################################################################################
    // All or nothing: a step that fails puts everything back and says why. Here the new place cannot
    // be made - a FILE lies where its parent folder would be.
    // ###########################################################################################
    [Fact]
    public void A_move_that_fails_changes_nothing_and_says_why()
    {
        string folder = this.Seed();
        Directory.CreateDirectory(this.DraftsRoot);
        File.WriteAllText(Path.Combine(this.DraftsRoot, "Commodore"), "a file, not a folder");

        NewBoardRegistration wanted = Registration("Commodore", "C64", "250407");
        DraftRenameResult result = DraftBoardRename.Run(this.DraftsRoot, Old, wanted);

        Assert.False(result.Renamed);
        Assert.Contains("Nothing was changed", result.Reason, StringComparison.Ordinal);

        Assert.True(Directory.Exists(folder));
        DraftMarker marker = DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(this.DraftsRoot, Old))!;
        Assert.Equal(Old, marker.BoardKey);
        Assert.Equal("C64", marker.NewBoard!.HardwareName);
        Assert.Equal("C64", BoardFiles.ReadDocument(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, Old))!.HardwareName);
    }
}
