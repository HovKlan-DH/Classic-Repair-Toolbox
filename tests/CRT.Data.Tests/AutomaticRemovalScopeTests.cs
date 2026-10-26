using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// AutomaticRemovalScope - a publish, promotion or push-back removes files ONLY inside the board's
// own folder (owner decision, 2026-09-27: "It can delete any files inside its own folder - not
// outside it - its own system main folder").
// ###########################################################################################
public sealed class AutomaticRemovalScopeTests
{
    private const string C64 = "Commodore/C64/250407";

    [Theory]
    [InlineData("Commodore/C64/250407/old-manual.pdf", true)]
    [InlineData("Commodore/C64/250407/KiCad data/board.kicad_pcb", true)]
    [InlineData("Commodore/Shared files/74LS08.png", false)]
    [InlineData("Generic shared files/Component images/7408.jpg", false)]
    [InlineData("Commodore/C128/310378/borrowed.pdf", false)]
    // A sibling whose name merely STARTS with the folder's is another board.
    [InlineData("Commodore/C64/250407 Rev B/x.png", false)]
    // Ordinal - the server's trees are case-sensitive, so this is another folder.
    [InlineData("commodore/c64/250407/x.png", false)]
    [InlineData("", false)]
    public void Only_a_file_under_the_boards_own_folder_may_be_removed(string path, bool inside)
    {
        Assert.Equal(inside, AutomaticRemovalScope.IsInsideBoardFolder(C64, path));
    }

    [Fact]
    public void The_list_keeps_only_the_boards_own_files()
    {
        Assert.Equal(
            ["Commodore/C64/250407/a.png"],
            AutomaticRemovalScope.Within(C64, ["Commodore/C64/250407/a.png", "Commodore/Shared files/b.png", "Generic shared files/c.png"]));

        Assert.Empty(AutomaticRemovalScope.Within("", ["Commodore/C64/250407/a.png"]));
        Assert.Empty(AutomaticRemovalScope.Within(C64, null));
    }

    // ###########################################################################################
    // *** A STORED FILE OF THE CRT 4.0 TREE MAY BE REMOVED (2026-10-25). *** A board's files' bytes
    // lie under their checksum in "Files/", shared by every board naming them - and DataTreeUsage
    // removes one only when no board of the tree lists it, so the rule still holds: nothing another
    // board uses can go.
    // ###########################################################################################
    [Theory]
    [InlineData("Files/3f/3f7a.png", true)]
    [InlineData("Files", false)]
    [InlineData("files/3f/3f7a.png", false)]
    [InlineData("MiniPro/IC tests/Catalogue/74LS00.txt", false)]
    public void A_stored_file_is_one_a_publish_may_remove(string path, bool stored)
    {
        Assert.Equal(stored, AutomaticRemovalScope.IsStoredFile(path));
        Assert.Equal(stored, AutomaticRemovalScope.Within(C64, [path]).Count == 1);
    }
}
