using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ContributionFileLocations - the folder a picked file goes in, of the board's own file names.
// ###########################################################################################
public sealed class ContributionFileLocationsTests
{
    // Where every converted board keeps that kind of file.
    [Fact]
    public void Each_kind_of_file_goes_in_the_folder_every_board_keeps_it_in()
    {
        Assert.Equal("Component images", ContributionFileLocations.DefaultFolderFor(ContributionFileLocations.FileRowKind.ComponentPicture));
        Assert.Equal("Scope baseline", ContributionFileLocations.DefaultFolderFor(ContributionFileLocations.FileRowKind.ScopeBaseline));
        Assert.Equal("Component local files", ContributionFileLocations.DefaultFolderFor(ContributionFileLocations.FileRowKind.ComponentFile));
        Assert.Equal("Board local files", ContributionFileLocations.DefaultFolderFor(ContributionFileLocations.FileRowKind.BoardFile));
    }

    // ###########################################################################################
    // *** THE SERVER AGREES. *** A folder files are put in that the server refuses is the 2026-09-25
    // report all over again: copied in, listed by the draft, and refused at Submit. Put through the
    // server's own rule (SubmissionFileRules) rather than restated, for every kind - and the rule
    // shown to refuse, so the loop is not vacuous.
    // ###########################################################################################
    [Fact]
    public void Every_kinds_folder_takes_a_new_file_the_server_accepts()
    {
        foreach (ContributionFileLocations.FileRowKind kind in Enum.GetValues<ContributionFileLocations.FileRowKind>())
        {
            string folder = ContributionFileLocations.DefaultFolderFor(kind);
            IReadOnlyList<string> refusals = ContributionFileLocationsTests.ServerRefusals(folder + "/New.png");

            Assert.True(refusals.Count == 0, $"[{folder}] is used, but the server refuses a new file in it: {string.Join(", ", refusals)}");
        }

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
