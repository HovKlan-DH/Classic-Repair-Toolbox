using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// FolderLinks - whether a path is reached through a folder that links elsewhere, which a check
// on the text of the path cannot see (code review, 2026-10-25). DraftFiles and BoardKiCadFolder ask
// it before they delete a file they take for their own.
// ###########################################################################################
public sealed class FolderLinksTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    [Fact]
    public void Plain_folders_and_folders_not_made_yet_are_no_links()
    {
        string root = this.thisWorkspace.Path_("Root");
        string file = this.thisWorkspace.WriteFile(Path.Combine("Root", "Pages", "a.kicad_sch"), "sheet");

        Assert.False(FolderLinks.AnyBetween(root, file));
        Assert.False(FolderLinks.AnyBetween(root, Path.Combine(root, "Not", "made", "yet.png")));
    }

    [Fact]
    public void A_folder_on_the_way_that_links_elsewhere_is_found_and_the_root_itself_is_not_looked_at()
    {
        string outside = this.thisWorkspace.Path_("Outside", "x");
        Directory.CreateDirectory(Path.GetDirectoryName(outside)!);
        string root = this.thisWorkspace.Path_("Root", "x");
        string link = Path.Combine(Path.GetDirectoryName(root)!, "Linked");

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
            Assert.True(FolderLinks.AnyBetween(Path.GetDirectoryName(root)!, Path.Combine(link, "Pages", "a.kicad_sch")));
            Assert.False(FolderLinks.AnyBetween(link, Path.Combine(link, "a.kicad_sch")));
        }
        finally
        {
            TempWorkspace.RemoveLink(link);
        }
    }
}
