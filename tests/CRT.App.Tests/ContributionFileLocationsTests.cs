using System.IO;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ContributionFileLocations - the folders the component editor offers in its "File location"
// drop-down.
//
// *** THE DRAFTS HALF IS THE REPORTED BUG (2026-09-24). *** The list used to be built from the
// data root alone, so a system created with "Add a new system" - which exists ONLY under the
// drafts root - never offered its own "Scope baseline" folder. The drafts side of these tests is
// built by DraftSeeder.CreateNewSystem itself rather than by hand, so if the seeder stops creating
// that folder, or names it differently, the listing test fails with it.
// ###########################################################################################
public sealed class ContributionFileLocationsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    public void Dispose() => this.thisWorkspace.Dispose();

    private void CreateFolders(string root, params string[] relativeFolders)
    {
        foreach (string relative in relativeFolders)
        {
            Directory.CreateDirectory(Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar)));
        }
    }

    private void CreateNewSystem(string manufacturer, string hardware, string board)
    {
        DraftSeedResult result = DraftSeeder.CreateNewSystem(this.DraftsRoot, new NewSystemRegistration
        {
            HardwareName = hardware,
            BoardName = board,
            ExcelDataFile = NewSystemIdentity.BuildExcelDataFile(manufacturer, hardware, board),
        });

        Assert.True(result.Created, result.Reason);
    }

    [Fact]
    public void Only_END_folders_are_listed_relative_to_the_root_and_forward_slashed()
    {
        this.CreateFolders(
            this.DataRoot,
            "Commodore/C64/250407/Scope baseline",
            "Commodore/C64/250407/KiCad data/Pages");

        Assert.Equal(
            new[] { "Commodore/C64/250407/KiCad data/Pages", "Commodore/C64/250407/Scope baseline" },
            ContributionFileLocations.FindEndFolders(this.DataRoot));
    }

    // ###########################################################################################
    // *** THE REGRESSION TEST FOR THE REPORT. *** A new system's "Scope baseline" folder is
    // offered, in the same Manufacturer/Hardware/Board form as a published board's.
    // ###########################################################################################
    [Fact]
    public void A_NEW_systems_Scope_baseline_folder_is_offered_from_the_drafts_tree()
    {
        this.CreateFolders(this.DataRoot, "Commodore/C64/250407/Scope baseline");
        this.CreateNewSystem("Test Manu4", "HW4", "Board4");

        List<string> folders = ContributionFileLocations.FindEndFolders(this.DataRoot, this.DraftsRoot);

        Assert.Contains("Test Manu4/HW4/Board4/Scope baseline", folders);

        // And the data tree's own folders are still there - the drafts tree is ADDED, not swapped in.
        Assert.Contains("Commodore/C64/250407/Scope baseline", folders);
    }

    // A published board with a seeded draft appears in both trees. Its folders must be listed once,
    // whatever the casing on either side.
    [Fact]
    public void A_folder_in_BOTH_trees_is_listed_once()
    {
        this.CreateFolders(this.DataRoot, "Commodore/C64/250407/Scope baseline");
        this.CreateFolders(this.DraftsRoot, "commodore/C64/250407/Scope Baseline");

        Assert.Single(ContributionFileLocations.FindEndFolders(this.DataRoot, this.DraftsRoot));
    }

    [Fact]
    public void The_merged_list_is_sorted_across_both_trees()
    {
        this.CreateFolders(this.DataRoot, "ZX Spectrum/Shared files/Board local files");
        this.CreateFolders(this.DraftsRoot, "Amstrad/CPC464/Z70200/Scope baseline");
        this.CreateFolders(this.DataRoot, "Commodore/C64/250407/Scope baseline");

        Assert.Equal(
            new[]
            {
                "Amstrad/CPC464/Z70200/Scope baseline",
                "Commodore/C64/250407/Scope baseline",
                "ZX Spectrum/Shared files/Board local files",
            },
            ContributionFileLocations.FindEndFolders(this.DataRoot, this.DraftsRoot));
    }

    // The drafts root is empty when DraftManager has not loaded, and a data root can be missing on a
    // first run before the sync. Neither may stop the other tree being listed.
    [Fact]
    public void A_blank_or_missing_root_is_skipped_rather_than_failing_the_list()
    {
        this.CreateFolders(this.DataRoot, "Commodore/C64/250407/Scope baseline");

        Assert.Equal(
            new[] { "Commodore/C64/250407/Scope baseline" },
            ContributionFileLocations.FindEndFolders(
                this.DataRoot,
                string.Empty,
                null,
                Path.Combine(this.thisWorkspace.Root, "does not exist")));
    }

    [Fact]
    public void An_empty_root_lists_nothing()
    {
        Directory.CreateDirectory(this.DataRoot);

        Assert.Empty(ContributionFileLocations.FindEndFolders(this.DataRoot));
    }
}
