using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Where a new data root is seeded from. The macOS release moved the data pack from
// Contents/MacOS/Data to Contents/Resources/Data (2026-10-12) - if CRT stopped finding it, a new
// install would download all ~1 GB instead of copying it, with nothing failing visibly.
public sealed class BundledDataLocationTests
{
    [Fact]
    public void Windows_and_Linux_find_the_data_pack_next_to_the_executable()
    {
        using var workspace = new TempWorkspace();
        string installFolder = workspace.Path_("Classic-Repair-Toolbox", "current");
        Directory.CreateDirectory(Path.Combine(installFolder, "Data"));

        Assert.Equal(Path.Combine(installFolder, "Data"), BundledDataLocation.Find(installFolder, Directory.Exists));
    }

    [Fact]
    public void A_macOS_app_bundle_finds_the_data_pack_in_its_Resources()
    {
        using var workspace = new TempWorkspace();
        string macos = workspace.Path_("Classic Repair Toolbox.app", "Contents", "MacOS");
        string resourcesData = workspace.Path_("Classic Repair Toolbox.app", "Contents", "Resources", "Data");
        Directory.CreateDirectory(macos);
        Directory.CreateDirectory(resourcesData);

        Assert.Equal(resourcesData, BundledDataLocation.Find(macos, Directory.Exists));
    }

    [Fact]
    public void The_base_directory_may_end_in_a_separator_as_AppContext_gives_it()
    {
        using var workspace = new TempWorkspace();
        string macos = workspace.Path_("Classic Repair Toolbox.app", "Contents", "MacOS");
        string resourcesData = workspace.Path_("Classic Repair Toolbox.app", "Contents", "Resources", "Data");
        Directory.CreateDirectory(macos);
        Directory.CreateDirectory(resourcesData);

        Assert.Equal(resourcesData, BundledDataLocation.Find(macos + Path.DirectorySeparatorChar, Directory.Exists));
    }

    [Fact]
    public void A_data_pack_next_to_the_executable_wins_over_the_bundles_Resources()
    {
        // CRT 3.0.0's bundles carried it in Contents/MacOS; whichever layout is installed must read
        // the same as before.
        using var workspace = new TempWorkspace();
        string macos = workspace.Path_("Classic Repair Toolbox.app", "Contents", "MacOS");
        Directory.CreateDirectory(Path.Combine(macos, "Data"));
        Directory.CreateDirectory(workspace.Path_("Classic Repair Toolbox.app", "Contents", "Resources", "Data"));

        Assert.Equal(Path.Combine(macos, "Data"), BundledDataLocation.Find(macos, Directory.Exists));
    }

    [Fact]
    public void A_folder_that_is_not_a_bundles_MacOS_never_reads_a_neighbouring_Resources()
    {
        // A source build's bin folder could sit beside a "Resources" folder of something else.
        using var workspace = new TempWorkspace();
        string bin = workspace.Path_("src", "CRT.App", "bin", "Release", "net10.0");
        Directory.CreateDirectory(bin);
        Directory.CreateDirectory(workspace.Path_("src", "CRT.App", "bin", "Release", "Resources", "Data"));

        Assert.Null(BundledDataLocation.Find(bin, Directory.Exists));
    }

    [Fact]
    public void A_MacOS_folder_outside_Contents_is_not_a_bundle()
    {
        var asked = new List<string>();

        string? found = BundledDataLocation.Find(
            Path.Combine(Path.GetTempPath(), "Somewhere", "MacOS"),
            path => { asked.Add(path); return false; });

        Assert.Null(found);
        Assert.Single(asked);
    }

    [Fact]
    public void A_build_from_source_has_no_data_pack()
    {
        using var workspace = new TempWorkspace();
        string macos = workspace.Path_("Classic Repair Toolbox.app", "Contents", "MacOS");
        Directory.CreateDirectory(macos);

        Assert.Null(BundledDataLocation.Find(macos, Directory.Exists));
    }
}
