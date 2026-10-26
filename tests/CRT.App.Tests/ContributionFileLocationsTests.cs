using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ContributionFileLocations - the folders the component editor offers in its "File location"
// drop-down.
//
// *** IN CRT 4.0 THEY ARE FOLDERS OF THE BOARD'S OWN FILE NAMES (2026-10-25). *** A row names a
// file of its board ("Scope baseline/U19_1_PAL.png"), and its bytes lie in a store laid out by
// checksum - so the list is no longer read off the data and drafts trees, whose folders in a 4.0
// tree are the store's ("Files/ab") and CRT's bookkeeping (".crt-pre-4.0"). A file picked into
// either was copied into the draft and never listed, so it was lost.
// ###########################################################################################
public sealed class ContributionFileLocationsTests
{
    private static readonly string[] Usual =
    {
        "Board local files",
        "Component images",
        "Component local files",
        "Scope baseline",
    };

    // ###########################################################################################
    // A new board, or one with none of its own files yet, offers the four folders every converted
    // board keeps its files in - the place a scope baseline goes is there before the first one is.
    // ###########################################################################################
    [Fact]
    public void A_board_with_no_files_offers_the_folders_every_board_keeps_its_files_in()
    {
        Assert.Equal(ContributionFileLocationsTests.Usual, ContributionFileLocations.ForBoard(new BoardData()));
        Assert.Equal(ContributionFileLocationsTests.Usual, ContributionFileLocations.ForBoard(null));
    }

    // Every folder the board's rows already file a name in is offered too, sorted in among the
    // usual ones; a name at the top of the board (the schematics) offers no folder.
    [Fact]
    public void The_folders_the_boards_own_names_use_are_offered_beside_the_usual_ones()
    {
        var board = new BoardData
        {
            Schematics = { new BoardSchematicEntry { SchematicName = "Board", SchematicImageFile = "Board layout.png" } },
            ComponentImages = { new ComponentImageEntry { BoardLabel = "U1", File = "Truth tables/7400.png" } },
            ComponentLocalFiles = { new ComponentLocalFileEntry { BoardLabel = "U1", File = "Scope baseline/U1/pin 1.png" } },
            BoardLocalFiles = { new BoardLocalFileEntry { File = "Manuals/Service manual.pdf" } },
        };

        Assert.Equal(
            new[]
            {
                "Board local files",
                "Component images",
                "Component local files",
                "Manuals",
                "Scope baseline",
                "Scope baseline/U1",
                "Truth tables",
            },
            ContributionFileLocations.ForBoard(board));
    }

    // A name is matched ignoring capitals, so one folder in two spellings is one folder - listed
    // once, as the usual folders spell it.
    [Fact]
    public void A_folder_is_listed_once_whatever_its_capitals()
    {
        var board = new BoardData
        {
            ComponentImages =
            {
                new ComponentImageEntry { BoardLabel = "U1", File = "component images/U1.png" },
                new ComponentImageEntry { BoardLabel = "U2", File = "Truth Tables/7400.png" },
                new ComponentImageEntry { BoardLabel = "U3", File = "truth tables/7402.png" },
            },
        };

        List<string> folders = ContributionFileLocations.ForBoard(board);

        Assert.Single(folders, folder => string.Equals(folder, "Component images", StringComparison.OrdinalIgnoreCase));
        Assert.Contains("Component images", folders);
        Assert.Single(folders, folder => string.Equals(folder, "Truth tables", StringComparison.OrdinalIgnoreCase));
    }

    // ###########################################################################################
    // *** THE REPORTED KIND OF LOSS, CAUGHT HERE. *** The store and CRT's bookkeeping are never a
    // place for a board's file: a draft write lists no name there (DraftFiles), so a file filed in
    // one would be copied into the draft and then dropped. Not offered even when a row names one.
    // ###########################################################################################
    [Fact]
    public void The_store_and_CRTs_bookkeeping_are_never_offered()
    {
        var board = new BoardData
        {
            ComponentImages =
            {
                new ComponentImageEntry { BoardLabel = "U1", File = "Files/ab/" + new string('a', 64) + ".png" },
                new ComponentImageEntry { BoardLabel = "U2", File = ".crt-pre-4.0/U2.png" },
            },
        };

        Assert.Equal(ContributionFileLocationsTests.Usual, ContributionFileLocations.ForBoard(board));
    }

    [Theory]
    [InlineData("Component images")]
    [InlineData("Scope baseline/U19")]
    [InlineData("Truth tables")]
    [InlineData("/Board local files/")]
    [InlineData("Manuals\\Old")]
    public void A_plain_folder_of_the_board_can_hold_its_file(string folder)
    {
        Assert.True(ContributionFileLocations.IsWritableFolder(folder));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("Files")]
    [InlineData("files/ab")]
    [InlineData(".crt-pre-4.0")]
    [InlineData(".crt-carried-from-3x/Commodore")]
    [InlineData("../Commodore")]
    [InlineData("Component images/../..")]
    [InlineData("./Component images")]
    [InlineData("Component images//U1")]
    [InlineData("C:/Users")]
    // A hidden folder anywhere along the name (code review, 2026-10-25) - the server refuses it.
    [InlineData(".images")]
    [InlineData("Images/.hidden")]
    [InlineData("Manuals/.git/Old")]
    public void A_folder_that_is_not_a_plain_folder_of_the_board_cannot(string? folder)
    {
        Assert.False(ContributionFileLocations.IsWritableFolder(folder));
    }

    // ###########################################################################################
    // *** A FOLDER CRT LETS A FILE BE FILED IN IS ONE THE SERVER TAKES A FILE FROM (code review,
    // 2026-10-25). *** ".images" was accepted here and the picture copied in and listed by the draft
    // - and the submission was then refused for its hidden name. Every folder the two might disagree
    // about: what CRT accepts, the server accepts.
    // ###########################################################################################
    [Theory]
    [InlineData(".images")]
    [InlineData("Images/.hidden")]
    [InlineData(".crt-pre-4.0")]
    [InlineData("Images")]
    [InlineData("Manuals/Old")]
    [InlineData("Scope baseline/U19")]
    public void A_folder_CRT_accepts_for_a_new_file_is_one_the_server_accepts_it_in(string folder)
    {
        if (ContributionFileLocations.IsWritableFolder(folder))
        {
            Assert.Empty(ContributionFileLocationsTests.ServerRefusals(folder + "/New.png"));
        }
        else
        {
            Assert.NotEmpty(ContributionFileLocationsTests.ServerRefusals(folder + "/New.png"));
        }
    }

    // ###########################################################################################
    // *** THE REPORT (2026-09-25), STILL HELD. *** A new file with no folder would sit at the top of
    // the board, beside its schematics, where no board keeps such a file - "Save to draft" refuses
    // it on the spot.
    // ###########################################################################################
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_new_file_with_no_folder_is_refused(string? location)
    {
        Assert.Equal(ContributionFileLocations.NewFileProblem.NoFolder, ContributionFileLocations.CheckNewFile(location));
    }

    [Theory]
    [InlineData("Files/ab")]
    [InlineData(".crt-pre-4.0")]
    [InlineData("../Commodore/C128/310378/Scope baseline")]
    public void A_new_file_in_a_folder_that_cannot_hold_it_is_refused(string location)
    {
        Assert.Equal(ContributionFileLocations.NewFileProblem.FolderNotWritable, ContributionFileLocations.CheckNewFile(location));
    }

    [Theory]
    [InlineData("Component images")]
    [InlineData("Scope baseline")]
    [InlineData("Truth tables")]
    public void A_new_file_in_a_folder_of_the_board_is_accepted(string location)
    {
        Assert.Equal(ContributionFileLocations.NewFileProblem.None, ContributionFileLocations.CheckNewFile(location));
    }

    // ###########################################################################################
    // *** THE SERVER AGREES. *** Every folder offered here must take a new file the server accepts -
    // a drop-down offering a folder the server refuses is the 2026-09-25 report all over again. Put
    // through the server's own rule (SubmissionFileRules) rather than restated, for every folder
    // the usual list and a board's own names offer.
    //
    // And the two refuse together where both can see it: CRT's bookkeeping is a dot-folder, which
    // the server refuses as a hidden name - shown below, so the loop is not vacuous.
    // ###########################################################################################
    [Fact]
    public void Every_offered_folder_takes_a_new_file_the_server_accepts()
    {
        var board = new BoardData
        {
            ComponentImages = { new ComponentImageEntry { BoardLabel = "U1", File = "Truth tables/7400.png" } },
            BoardLocalFiles = { new BoardLocalFileEntry { File = "Manuals/Old/Service manual.pdf" } },
        };

        List<string> offered = ContributionFileLocations.ForBoard(board);

        Assert.Contains("Manuals/Old", offered);

        foreach (string folder in offered)
        {
            IReadOnlyList<string> refusals = ContributionFileLocationsTests.ServerRefusals(folder + "/New.png");

            Assert.True(refusals.Count == 0, $"[{folder}] is offered, but the server refuses a new file in it: {string.Join(", ", refusals)}");
        }

        Assert.False(ContributionFileLocations.IsWritableFolder(".crt-pre-4.0"));
        Assert.Contains("file.hidden_name", ContributionFileLocationsTests.ServerRefusals(".crt-pre-4.0/New.png"));
        Assert.Contains("file.type_not_allowed", ContributionFileLocationsTests.ServerRefusals("Scope baseline/New.exe"));
    }

    // The codes the server's file rules give a submission for C64 250407 carrying one new file under
    // this name, with nothing published yet.
    private static IReadOnlyList<string> ServerRefusals(string name)
    {
        var manifest = new SubmissionManifest
        {
            BoardId = "Commodore/C64/250407",
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
        };

        manifest.Files.Add(new SubmissionFile { Path = name, Sha256 = new string('a', 64), SizeBytes = 10 });
        manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = name });

        return SubmissionFileRules.ValidateManifestFiles(manifest, PublishedTreeView.Empty)
            .Select(finding => finding.Code)
            .ToList();
    }
}
