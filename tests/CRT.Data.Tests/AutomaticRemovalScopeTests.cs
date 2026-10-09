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
}
