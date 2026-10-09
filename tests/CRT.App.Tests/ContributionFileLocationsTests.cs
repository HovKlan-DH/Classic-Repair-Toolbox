using System.IO;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ContributionFileLocations - the folders the component editor offers in its "File location"
// drop-down.
//
// *** THE DRAFTS HALF IS THE REPORTED BUG (2026-09-24). *** The list used to be built from the
// data root alone, so a board created with "Add a new board" - which exists ONLY under the
// drafts root - never offered its own "Scope baseline" folder. The drafts side of these tests is
// built by DraftSeeder.CreateNewBoard itself rather than by hand, so if the seeder stops creating
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

    private void CreateNewBoard(string manufacturer, string hardware, string board)
    {
        DraftSeedResult result = DraftSeeder.CreateNewBoard(this.DraftsRoot, new NewBoardRegistration
        {
            HardwareName = hardware,
            BoardName = board,
            ExcelDataFile = NewBoardIdentity.BuildExcelDataFile(manufacturer, hardware, board),
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
    // *** THE REGRESSION TEST FOR THE REPORT. *** A new board's "Scope baseline" folder is
    // offered, in the same Manufacturer/Hardware/Board form as a published board's.
    // ###########################################################################################
    [Fact]
    public void A_NEW_boards_Scope_baseline_folder_is_offered_from_the_drafts_tree()
    {
        this.CreateFolders(this.DataRoot, "Commodore/C64/250407/Scope baseline");
        this.CreateNewBoard("Test Manu4", "HW4", "Board4");

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

    // ---------------------------------------------------------------------- which folders are offered

    private const string Board = "Commodore/C64/250407";

    // What the data tree and a draft together list for the C64 250407 before narrowing: the board's
    // own folders, other boards of the same maker, another maker, both kinds of shared folder, the
    // MiniPro IC tests, and a draft's copy of a shared folder (a new file filed in a shared folder is
    // kept in the draft under its whole path, so the drafts-tree scan lists that copy as the board's).
    private static readonly string[] EveryFolder =
    {
        "Amstrad/CPC 664/MC0005A/Scope baseline",
        "Amstrad/Shared files/Component images",
        "Commodore/C128/250477",
        "Commodore/C128/310378/Scope baseline",
        "Commodore/C64/250407/Commodore/Shared files/Component images",
        "Commodore/C64/250407/KiCad data",
        "Commodore/C64/250407/Scope baseline",
        "Commodore/C64/250425/Scope baseline",
        "Commodore/Shared files/Board local files",
        "Commodore/Shared files/Component images",
        "Commodore/Shared files/Component local files",
        "Generic shared files/Component images",
        "Generic shared files/Component local files",
        "Generic shared files/MiniPro/IC tests/Catalogue",
        "Generic shared files/MiniPro/IC tests/Vectors",
    };

    // ###########################################################################################
    // *** THE OWNER'S LIST (2026-09-25): "only show folders within the same board + the parent
    // folders Shared files and Generic shared files". *** The drop-down listed every folder of every
    // board, and a new file filed in almost any of them was refused by the server at submit. The
    // board's own folder is offered too although it is not an end folder - its schematic images
    // live directly in it.
    // ###########################################################################################
    [Fact]
    public void Only_this_boards_folders_and_the_shared_ones_are_offered()
    {
        Assert.Equal(
            new[]
            {
                "Commodore/C64/250407",
                "Commodore/C64/250407/KiCad data",
                "Commodore/C64/250407/Scope baseline",
                "Commodore/Shared files/Board local files",
                "Commodore/Shared files/Component images",
                "Commodore/Shared files/Component local files",
                "Generic shared files/Component images",
                "Generic shared files/Component local files",
            },
            ContributionFileLocations.WritableBy(ContributionFileLocationsTests.Board, ContributionFileLocationsTests.EveryFolder));
    }

    // ###########################################################################################
    // *** THE SERVER AGREES. *** Whose folder a path is, is decided on the server by
    // SubmissionFileScopes, and a submission is refused by SubmissionFileRules. Every folder offered
    // here must take a new file the server accepts - a drop-down offering a folder the server refuses
    // is the report all over again. Put through the server's own rule rather than restated.
    // ###########################################################################################
    [Fact]
    public void Every_offered_folder_takes_a_new_file_the_server_accepts_and_other_boards_are_refused()
    {
        foreach (string folder in ContributionFileLocationsTests.EveryFolder)
        {
            bool offered = ContributionFileLocations.IsWritableFolder(ContributionFileLocationsTests.Board, folder);
            IReadOnlyList<string> refusals = ContributionFileLocationsTests.ServerRefusals(folder + "/New.png");

            if (offered)
            {
                Assert.True(refusals.Count == 0, $"[{folder}] is offered, but the server refuses a new file in it: {string.Join(", ", refusals)}");
            }
        }

        // And the refusal the report was about is real, so the loop above is not vacuous.
        Assert.Contains("file.other_board", ContributionFileLocationsTests.ServerRefusals("Commodore/C128/310378/Scope baseline/New.png"));
        Assert.Contains("file.other_board", ContributionFileLocationsTests.ServerRefusals("Amstrad/Shared files/Component images/New.png"));
    }

    // The report itself: a bare file name. The server refuses it, and so does "Save to draft" now.
    [Fact]
    public void A_file_with_no_folder_is_refused_here_as_the_server_refuses_it()
    {
        Assert.Contains("file.other_board", ContributionFileLocationsTests.ServerRefusals("HotCPU.png"));

        Assert.Equal(
            ContributionFileLocations.NewFileProblem.NoFolder,
            ContributionFileLocations.CheckNewFile(ContributionFileLocationsTests.Board, string.Empty, usedInPlace: false));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_new_file_with_no_folder_is_refused(string? location)
    {
        Assert.Equal(
            ContributionFileLocations.NewFileProblem.NoFolder,
            ContributionFileLocations.CheckNewFile(ContributionFileLocationsTests.Board, location, usedInPlace: false));
    }

    // A file at the very top of the data folder is outside every board too - being picked from
    // there does not make it a board's file.
    [Fact]
    public void A_file_picked_from_the_top_of_the_data_folder_is_still_refused()
    {
        Assert.Equal(
            ContributionFileLocations.NewFileProblem.NoFolder,
            ContributionFileLocations.CheckNewFile(ContributionFileLocationsTests.Board, string.Empty, usedInPlace: true));
    }

    [Theory]
    [InlineData("Commodore/C64/250407")]
    [InlineData("Commodore/C64/250407/Scope baseline")]
    [InlineData("Commodore/Shared files/Component images")]
    [InlineData("Generic shared files/Component local files")]
    public void A_new_file_in_this_boards_folders_or_a_shared_one_is_accepted(string location)
    {
        Assert.Equal(
            ContributionFileLocations.NewFileProblem.None,
            ContributionFileLocations.CheckNewFile(ContributionFileLocationsTests.Board, location, usedInPlace: false));
    }

    [Theory]
    [InlineData("Commodore/C128/310378/Scope baseline")]
    [InlineData("Commodore/C64/250425/Scope baseline")]
    [InlineData("Amstrad/Shared files/Component images")]
    [InlineData("Generic shared files/MiniPro/IC tests/Catalogue")]
    public void A_new_file_in_another_boards_folder_is_refused(string location)
    {
        Assert.Equal(
            ContributionFileLocations.NewFileProblem.FolderNotWritable,
            ContributionFileLocations.CheckNewFile(ContributionFileLocationsTests.Board, location, usedInPlace: false));
    }

    // ###########################################################################################
    // A file picked FROM another board's folder and left there is the published copy itself, and
    // citing it unchanged is how real data is shaped - the C128DCR cites the C128's scope baselines,
    // and the server allows a foreign file that matches what is published.
    // ###########################################################################################
    [Fact]
    public void Another_boards_file_used_where_it_already_is_is_accepted()
    {
        Assert.Equal(
            ContributionFileLocations.NewFileProblem.None,
            ContributionFileLocations.CheckNewFile(
                ContributionFileLocationsTests.Board,
                "Commodore/C128/310378/Scope baseline",
                usedInPlace: true));
    }

    // Whose folder is decided case-sensitively, as on the server's Linux tree.
    [Fact]
    public void A_folder_differing_only_by_capitalisation_is_not_this_boards()
    {
        Assert.False(ContributionFileLocations.IsWritableFolder(ContributionFileLocationsTests.Board, "commodore/C64/250407/Scope baseline"));
    }

    // An unrecognisable board filters nothing: the save refuses such a window anyway, and an empty
    // drop-down would only hide why.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Commodore/C64")]
    public void With_no_recognisable_board_nothing_is_filtered(string? boardId)
    {
        Assert.Equal(
            ContributionFileLocationsTests.EveryFolder,
            ContributionFileLocations.WritableBy(boardId, ContributionFileLocationsTests.EveryFolder));
    }

    // A board that exists only as a draft offers its own folders the same way.
    [Fact]
    public void A_new_boards_own_folders_are_offered()
    {
        Assert.Equal(
            new[] { "Generic shared files/Component images", "Test Manu4/HW4/Board4", "Test Manu4/HW4/Board4/Scope baseline" },
            ContributionFileLocations.WritableBy(
                "Test Manu4/HW4/Board4",
                new[] { "Test Manu4/HW4/Board4/Scope baseline", "Generic shared files/Component images", "Commodore/C64/250407/Scope baseline" }));
    }

    // The codes the server's file rules give a submission for C64 250407 carrying one new file at
    // this path, with nothing published yet.
    private static IReadOnlyList<string> ServerRefusals(string path)
    {
        var manifest = new SubmissionManifest
        {
            BoardId = ContributionFileLocationsTests.Board,
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
        };

        manifest.Files.Add(new SubmissionFile { Path = path, Sha256 = new string('a', 64), SizeBytes = 10 });
        manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = path });

        return SubmissionFileRules.ValidateManifestFiles(manifest, PublishedTreeView.Empty)
            .Select(finding => finding.Code)
            .ToList();
    }
}
