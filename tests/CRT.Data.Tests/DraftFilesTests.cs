using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftFiles - how a new file lying in a CRT 4.0 draft's folder is taken into the draft's store
// and listed, and, above all, that nothing outside the draft folder is ever taken in: a file taken
// in is DELETED from where it lay once the board file lists it (RemoveArrived), so a row naming a
// file outside the folder had that file deleted (code review, 2026-10-25).
// ###########################################################################################
public sealed class DraftFilesTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string DraftFolder => Path.Combine(this.thisWorkspace.Root, "Drafts", "Commodore", "C64", "250407");

    private string BoardFile => Path.Combine(this.DraftFolder, DataTreeLayout.BoardFileName);

    private static BoardData CitingPicture(string file) => new()
    {
        Components = { new ComponentEntry { BoardLabel = "U8", FriendlyName = "CIA" } },
        ComponentImages = { new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = file } },
    };

    [Fact]
    public void A_new_file_lying_in_the_draft_folder_is_stored_listed_and_then_tidied_away()
    {
        string picked = this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "C64", "250407", "Component images", "u8.png"), "picture");

        List<BoardDocument.StoredFileRow> listed = DraftFiles.ListFor(this.BoardFile, CitingPicture("Component images/u8.png"), [], out List<string> arrived);

        BoardDocument.StoredFileRow entry = Assert.Single(listed);
        Assert.Equal("Component images/u8.png", entry.Name);
        Assert.Equal(DataTreeLayout.Sha256Of(picked), entry.Sha256);
        Assert.True(File.Exists(Path.Combine(this.DraftFolder, DataTreeLayout.StoredFilePath(entry.Sha256, entry.Name))));
        Assert.Equal([picked], arrived);

        DraftFiles.RemoveArrived(this.BoardFile, arrived);

        Assert.False(File.Exists(picked));
        Assert.False(Directory.Exists(Path.Combine(this.DraftFolder, "Component images")));
    }

    // A File cell can hold a whole path - typed, or pasted from Explorer. The file it names is the
    // contributor's own, outside the draft, and must be neither listed nor deleted.
    [Fact]
    public void A_name_that_is_a_full_path_outside_the_draft_is_never_taken_in_or_deleted()
    {
        string outside = this.thisWorkspace.WriteFile(Path.Combine("Pictures", "u8.png"), "my own picture");
        Directory.CreateDirectory(this.DraftFolder);

        List<BoardDocument.StoredFileRow> listed = DraftFiles.ListFor(this.BoardFile, CitingPicture(outside), [], out List<string> arrived);
        DraftFiles.RemoveArrived(this.BoardFile, arrived);

        Assert.Empty(listed);
        Assert.Empty(arrived);
        Assert.Equal("my own picture", File.ReadAllText(outside));
        Assert.False(Directory.Exists(Path.Combine(this.DraftFolder, DataTreeLayout.FilesFolder)));
    }

    // "../" climbs into the next board's draft - its board file, or any of its files.
    [Fact]
    public void A_name_climbing_out_of_the_draft_folder_is_never_taken_in_or_deleted()
    {
        string otherBoard = this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "C64", "250469", "board.json"), "{}");
        Directory.CreateDirectory(this.DraftFolder);

        List<BoardDocument.StoredFileRow> listed = DraftFiles.ListFor(this.BoardFile, CitingPicture("../250469/board.json"), [], out List<string> arrived);
        DraftFiles.RemoveArrived(this.BoardFile, arrived);

        Assert.Empty(listed);
        Assert.Empty(arrived);
        Assert.True(File.Exists(otherBoard));
    }

    // A folder in the draft that links elsewhere: a delete through it deletes the file linked to.
    [Fact]
    public void A_name_reached_through_a_linked_folder_is_never_taken_in_or_deleted()
    {
        string outside = this.thisWorkspace.WriteFile(Path.Combine("Pictures", "u8.png"), "my own picture");
        string link = Path.Combine(this.DraftFolder, "Component images");
        Directory.CreateDirectory(this.DraftFolder);

        try
        {
            Directory.CreateSymbolicLink(link, Path.GetDirectoryName(outside)!);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Assert.Skip("This account may not create a folder link (Windows outside Developer Mode).");
        }

        try
        {
            Assert.False(DraftFiles.TryArrivedPath(this.DraftFolder, "Component images/u8.png", out _));

            List<BoardDocument.StoredFileRow> listed = DraftFiles.ListFor(this.BoardFile, CitingPicture("Component images/u8.png"), [], out List<string> arrived);
            DraftFiles.RemoveArrived(this.BoardFile, arrived);

            Assert.Empty(listed);
            Assert.Equal("my own picture", File.ReadAllText(outside));
        }
        finally
        {
            TempWorkspace.RemoveLink(link);
        }
    }

    [Theory]
    [InlineData("Files/ab/abcdef.png")]
    [InlineData("board.json")]
    [InlineData(".crt-base.json")]
    [InlineData(".crt-pre-4.0/Data C64 250407.xlsx")]
    [InlineData("")]
    public void The_store_the_board_file_and_CRTs_bookkeeping_are_never_a_new_file(string name)
    {
        Assert.False(DraftFiles.TryArrivedPath(this.DraftFolder, name, out string path));
        Assert.Equal(string.Empty, path);
    }

    [Fact]
    public void A_name_inside_the_draft_folder_resolves_to_where_it_would_lie()
    {
        Assert.True(DraftFiles.TryArrivedPath(this.DraftFolder, "Component images/u8.png", out string path));
        Assert.Equal(Path.Combine(Path.GetFullPath(this.DraftFolder), "Component images", "u8.png"), path);
    }

    // ###########################################################################################
    // WHAT THE DRAFT HELD IN MEMORY REMEMBERS (code review, 2026-10-26): an undo can bring back a row
    // citing a file the list dropped at the save before. A name a row cites that the list no longer
    // has takes the entry remembered for it - never a KiCad file, which no row cites: one removed
    // stays removed - and the remembered bytes are kept in the store.
    // ###########################################################################################
    [Fact]
    public void A_cited_name_the_list_dropped_takes_its_remembered_entry()
    {
        BoardDocument.StoredFileRow remembered = StoredFiles.EntryFor("Component images/u8.png", "picture");
        BoardDocument.StoredFileRow kiCad = StoredFiles.EntryFor("KiCad data/board.kicad_pcb", "pcb");
        Directory.CreateDirectory(this.DraftFolder);

        List<BoardDocument.StoredFileRow> listed = DraftFiles.ListFor(this.BoardFile, CitingPicture("Component images/u8.png"), [], out _, [remembered, kiCad]);

        Assert.Same(remembered, Assert.Single(listed));
    }

    [Fact]
    public void An_entry_the_list_still_has_wins_over_the_remembered_one()
    {
        BoardDocument.StoredFileRow now = StoredFiles.EntryFor("Component images/u8.png", "a newer picture");
        Directory.CreateDirectory(this.DraftFolder);

        List<BoardDocument.StoredFileRow> listed = DraftFiles.ListFor(
            this.BoardFile, CitingPicture("Component images/u8.png"), [now], out _, [StoredFiles.EntryFor("Component images/u8.png", "picture")]);

        Assert.Same(now, Assert.Single(listed));
    }

    // ###########################################################################################
    // A name the list holds spelt another way - "KiCad data/./main.kicad_pcb", "Images//u8.png", as a
    // hand-placed draft may - is the same file (code review, 2026-10-26): looked up as Normalise spells
    // it but kept as it was spelt, it was never found, so the next save dropped it from the list and
    // the tidying deleted its stored bytes.
    // ###########################################################################################
    [Fact]
    public void A_listed_name_spelt_another_way_is_kept()
    {
        Directory.CreateDirectory(this.DraftFolder);
        BoardDocument.StoredFileRow kiCad = StoredFiles.EntryFor("KiCad data/./main.kicad_pcb", "pcb");
        BoardDocument.StoredFileRow picture = StoredFiles.EntryFor("Component images//u8.png", "picture");

        List<BoardDocument.StoredFileRow> listed = DraftFiles.ListFor(this.BoardFile, CitingPicture("Component images/u8.png"), [kiCad, picture], out _);

        Assert.Equal(2, listed.Count);
        Assert.Contains(kiCad, listed);
        Assert.Contains(picture, listed);
    }

    [Fact]
    public void The_store_keeps_the_bytes_of_what_it_is_told_to_keep_and_tidies_the_rest()
    {
        BoardDocument.StoredFileRow listed = StoredFiles.Put(this.DraftFolder, "Images/a.png", "listed");
        BoardDocument.StoredFileRow remembered = StoredFiles.Put(this.DraftFolder, "Images/b.png", "remembered");
        BoardDocument.StoredFileRow neither = StoredFiles.Put(this.DraftFolder, "Images/c.png", "neither");

        DraftFiles.RemoveUnlisted(this.BoardFile, [listed], alsoKeep: [remembered]);

        Assert.True(File.Exists(this.Stored(listed)));
        Assert.True(File.Exists(this.Stored(remembered)));
        Assert.False(File.Exists(this.Stored(neither)));
    }

    private string Stored(BoardDocument.StoredFileRow file) =>
        Path.Combine(this.DraftFolder, DataTreeLayout.StoredFilePath(file.Sha256, file.Name).Replace('/', Path.DirectorySeparatorChar));

    [Fact]
    public void The_KiCad_files_lying_in_the_draft_are_named_as_the_board_names_them_in_order()
    {
        this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "C64", "250407", "KiCad data", "b.kicad_pcb"), "pcb");
        this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "C64", "250407", "KiCad data", "Pages", "a.kicad_sch"), "sheet");

        Assert.Equal(["KiCad data/Pages/a.kicad_sch", "KiCad data/b.kicad_pcb"], DraftFiles.ArrivedKiCadNames(this.DraftFolder));
        Assert.Empty(DraftFiles.ArrivedKiCadNames(Path.Combine(this.thisWorkspace.Root, "Nowhere")));
    }
}
