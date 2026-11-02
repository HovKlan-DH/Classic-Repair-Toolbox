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

    // ###########################################################################################
    // *** "NOTHING WAS CHANGED" ONLY WHEN NOTHING WAS (code review, 2026-11-02). *** The marker could not
    // be written after the folder moved, and the move back failed too: the folder at the new place, its
    // marker naming the old one - and the page said nothing had changed. What is said follows what came
    // of putting it back; the next start repairs the marker (RepairMovedDrafts).
    // ###########################################################################################
    [Fact]
    public void A_rename_that_could_not_be_put_back_does_not_claim_nothing_changed()
    {
        string folder = this.Seed();
        NewBoardRegistration wanted = Registration("Commodore", "C64", "250407");

        DraftRenameResult result = DraftBoardRename.Run(this.DraftsRoot, Old, wanted, _ => { }, afterMove: () =>
        {
            // Something takes the old place, so the move back cannot be made, and the marker cannot be written.
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "in the way.txt"), "kept");
            throw new IOException("the marker is held");
        });

        Assert.False(result.Renamed);
        Assert.DoesNotContain("Nothing was changed", result.Reason, StringComparison.Ordinal);
        Assert.Contains("could not be put back", result.Reason, StringComparison.Ordinal);
        Assert.True(File.Exists(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, wanted.ExcelDataFile)));
    }

    // The partner: a failure the put-back undid still says so.
    [Fact]
    public void A_rename_that_was_put_back_says_nothing_changed()
    {
        string folder = this.Seed();

        DraftRenameResult result = DraftBoardRename.Run(
            this.DraftsRoot,
            Old,
            Registration("Commodore", "C64", "250407"),
            _ => { },
            afterMove: () => throw new IOException("the marker is held"));

        Assert.False(result.Renamed);
        Assert.Contains("Nothing was changed", result.Reason, StringComparison.Ordinal);
        Assert.True(File.Exists(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, Old)));
        Assert.True(Directory.Exists(folder));
    }

    // ------------------------------------------------------------------ Repaired at launch

    // A rename's stop between the move and the marker: the folder at its new place, and everything in it
    // as it was - its marker and board file still naming the old one.
    private string MovedWithoutItsMarker(string to)
    {
        string from = this.Seed();
        string moved = DraftFolderLayout.GetBoardFolder(this.DraftsRoot, to);
        Directory.CreateDirectory(Path.GetDirectoryName(moved)!);
        Directory.Move(from, moved);

        return moved;
    }

    // ###########################################################################################
    // *** A MARKER'S IDENTITY IS NEVER LOOKED FOR OUTSIDE THE DRAFTS ROOT (code review, 2026-11-02 -
    // Security). *** The repair asked whether the folder it names exists, at every launch, of whatever the
    // file said: a UNC path reached another machine with the sign-in's NTLM hash. Here "../" names a folder
    // that does exist: found, the draft was taken for a copy made by hand and left naming it. It names no
    // folder of the drafts root, so its marker is repaired to where it lies.
    // ###########################################################################################
    [Fact]
    public void A_marker_naming_a_folder_outside_the_drafts_root_is_repaired_to_the_folder_it_lies_in()
    {
        const string Outside = "../Outside/250407/board.json";

        Directory.CreateDirectory(this.thisWorkspace.Path_("Outside", "250407"));

        string folder = this.Seed();
        string markerPath = Path.Combine(folder, DraftFolderLayout.DraftMarkerFileName);
        DraftMarker before = DraftMarkerStore.Load(markerPath)!;
        DraftMarkerStore.Save(markerPath, before with { BoardKey = Outside });

        Assert.Equal([Old], DraftBoardRename.RepairMovedDrafts(this.DraftsRoot, []));
        Assert.Equal(Old, DraftMarkerStore.Load(markerPath)!.BoardKey);
    }

    // ###########################################################################################
    // *** A MARKER NAMING ANOTHER PLACE IS REPAIRED AT LAUNCH (code review, 2026-11-02). *** Left as it
    // was, the drop-down lists showed the old identity, whose folder is gone, and nothing could save to
    // it. A crash between the move and the marker leaves the same as a put-back that failed.
    // ###########################################################################################
    [Fact]
    public void A_marker_naming_a_folder_that_is_gone_is_repaired_to_the_folder_it_lies_in()
    {
        NewBoardRegistration wanted = Registration("Commodore", "Commodore 64", "KU-14194HB");
        string moved = this.MovedWithoutItsMarker(wanted.ExcelDataFile);

        string markerPath = Path.Combine(moved, DraftFolderLayout.DraftMarkerFileName);
        DraftMarker before = DraftMarkerStore.Load(markerPath)!;
        DraftMarkerStore.Save(markerPath, before with
        {
            NewBoard = new NewBoardRegistration
            {
                HardwareName = before.NewBoard!.HardwareName,
                BoardName = before.NewBoard.BoardName,
                ExcelDataFile = before.NewBoard.ExcelDataFile,
                HardwareNotes = "Mine",
                CreatedUtc = "2026-10-01T00:00:00Z",
            },
        });

        Assert.Equal([wanted.ExcelDataFile], DraftBoardRename.RepairMovedDrafts(this.DraftsRoot, []));

        DraftMarker marker = DraftMarkerStore.Load(markerPath)!;
        Assert.Equal(wanted.ExcelDataFile, marker.BoardKey);
        Assert.Equal(wanted.ExcelDataFile, marker.NewBoard!.ExcelDataFile);
        Assert.Equal("Commodore 64", marker.NewBoard.HardwareName);
        Assert.Equal("KU-14194HB", marker.NewBoard.BoardName);
        Assert.Equal("Mine", marker.NewBoard.HardwareNotes);
        Assert.Equal("2026-10-01T00:00:00Z", marker.NewBoard.CreatedUtc);

        BoardDocument read = BoardFiles.ReadDocument(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, wanted.ExcelDataFile))!;
        Assert.Equal("Commodore 64", read.HardwareName);
        Assert.Equal("KU-14194HB", read.BoardName);
    }

    // A folder copied by hand claims to be the draft it was copied from: it is not that draft, and
    // nothing here may decide what it is.
    [Fact]
    public void A_copy_whose_original_is_still_there_is_left_alone()
    {
        string original = this.Seed();
        string copy = DraftFolderLayout.GetBoardFolder(this.DraftsRoot, "Commodore/C64/250407/board.json");
        Directory.CreateDirectory(copy);

        foreach (string file in Directory.GetFiles(original))
            File.Copy(file, Path.Combine(copy, Path.GetFileName(file)));

        Assert.Empty(DraftBoardRename.RepairMovedDrafts(this.DraftsRoot, []));
        Assert.Equal(Old, DraftMarkerStore.Load(Path.Combine(copy, DraftFolderLayout.DraftMarkerFileName))!.BoardKey);
    }

    // Repaired, it would be a second board of a name a published board has.
    [Fact]
    public void A_folder_a_published_board_is_listed_under_is_left_alone()
    {
        const string Listed = "Commodore/C64/250407/board.json";

        string moved = this.MovedWithoutItsMarker(Listed);

        Assert.Empty(DraftBoardRename.RepairMovedDrafts(this.DraftsRoot, [Listed.ToLowerInvariant()]));
        Assert.Equal(Old, DraftMarkerStore.Load(Path.Combine(moved, DraftFolderLayout.DraftMarkerFileName))!.BoardKey);
    }

    // Another folder of the same name in other capitals - a disk that tells them apart can hold both, and
    // two new boards differing only in capitals are one folder to anybody else.
    [Fact]
    public void A_folder_differing_from_another_only_in_capitals_is_left_alone()
    {
        Assert.SkipWhen(OperatingSystem.IsWindows() || OperatingSystem.IsMacOS(), "This disk sees the two as one folder.");

        string moved = this.MovedWithoutItsMarker("Commodore/C64/250407/board.json");
        Directory.CreateDirectory(DraftFolderLayout.GetBoardFolder(this.DraftsRoot, "Commodore/c64/250407/board.json"));

        Assert.Empty(DraftBoardRename.RepairMovedDrafts(this.DraftsRoot, []));
        Assert.Equal(Old, DraftMarkerStore.Load(Path.Combine(moved, DraftFolderLayout.DraftMarkerFileName))!.BoardKey);
    }

    // The name in between a change of capitals is no board's: a crash there leaves the folder for the
    // contributor to look at, and its marker saying where it belongs.
    [Fact]
    public void A_folder_stopped_in_the_middle_of_a_rename_of_capitals_is_left_alone()
    {
        string moved = this.MovedWithoutItsMarker("Comodore/C64/250407.crt-renaming/board.json");

        Assert.Empty(DraftBoardRename.RepairMovedDrafts(this.DraftsRoot, []));
        Assert.Equal(Old, DraftMarkerStore.Load(Path.Combine(moved, DraftFolderLayout.DraftMarkerFileName))!.BoardKey);
    }

    // A draft of a published board was never renamed: its marker names the board, whatever folder it is in.
    [Fact]
    public void A_draft_of_a_published_board_is_not_repaired()
    {
        string moved = this.MovedWithoutItsMarker("Commodore/C64/250407/board.json");
        string markerPath = Path.Combine(moved, DraftFolderLayout.DraftMarkerFileName);
        DraftMarkerStore.Save(markerPath, DraftMarkerStore.Load(markerPath)! with { NewBoard = null });

        Assert.Empty(DraftBoardRename.RepairMovedDrafts(this.DraftsRoot, []));
        Assert.Equal(Old, DraftMarkerStore.Load(markerPath)!.BoardKey);
    }

    // The ordinary case, every launch: a draft where its marker says it is, which nothing touches.
    [Fact]
    public void A_draft_where_its_marker_says_is_not_written_again()
    {
        string folder = this.Seed();
        string markerPath = Path.Combine(folder, DraftFolderLayout.DraftMarkerFileName);
        DateTime written = File.GetLastWriteTimeUtc(markerPath);

        Assert.Empty(DraftBoardRename.RepairMovedDrafts(this.DraftsRoot, []));
        Assert.Equal(written, File.GetLastWriteTimeUtc(markerPath));
    }
}
