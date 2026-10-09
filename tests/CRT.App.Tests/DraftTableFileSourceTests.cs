using System.IO;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftTableFileSource - where the Drafts tab's table reads a file cell's file for its hover card
// (owner request, 2026-09-26). The published side is the local data folder; the draft side is the
// draft's own copy when it has one, else the published file - the board on screen's own rule.
//
// The cell's text is whatever was typed into it, so a path that is not a safe relative path must
// reach no file at all.
// ###########################################################################################
public sealed class DraftTableFileSourceTests : IDisposable
{
    private const string ExcelDataFile = "Commodore/C64/250407/Data C64 250407.xlsx";
    private const string Picture = "Commodore/C64/250407/Board.png";

    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private string DraftBoardFolder =>
        DraftFolderLayout.GetBoardFolder(Path.Combine(this.thisWorkspace.Root, "Drafts"), DraftTableFileSourceTests.ExcelDataFile);

    private DraftTableFileSource Source() => new(this.DataRoot, this.DraftBoardFolder);

    private string WritePublished(string relative, string content)
    {
        string full = Path.Combine(this.DataRoot, relative.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        File.WriteAllText(full, content);
        return full;
    }

    private string WriteDrafted(string relative, string content)
    {
        string full = DraftFileResolver.BuildDraftFileDestination(this.DraftBoardFolder, relative);
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        File.WriteAllText(full, content);
        return full;
    }

    // The two sides of a picture the draft replaced: the synced copy, and the draft's own.
    [Fact]
    public async Task The_published_side_reads_the_data_folder_and_the_draft_side_the_drafts_copy()
    {
        string published = this.WritePublished(DraftTableFileSourceTests.Picture, "published");
        string drafted = this.WriteDrafted(DraftTableFileSourceTests.Picture, "drafted");

        DraftTableFileSource source = this.Source();

        Assert.Equal(Path.GetFullPath(published), Path.GetFullPath(source.Resolve(DraftTableFileSourceTests.Picture, BoardTableFileSide.Published)!.FullPath));
        Assert.Equal(Path.GetFullPath(drafted), Path.GetFullPath(source.Resolve(DraftTableFileSourceTests.Picture, BoardTableFileSide.Current)!.FullPath));

        Assert.Equal("published"u8.ToArray(), await source.ReadAsync(DraftTableFileSourceTests.Picture, BoardTableFileSide.Published));
        Assert.Equal("drafted"u8.ToArray(), await source.ReadAsync(DraftTableFileSourceTests.Picture, BoardTableFileSide.Current));
    }

    // The card names the published side after the source the data was downloaded from, as the
    // text cells' tooltips do (owner request, 2026-10-05).
    [Fact]
    public void The_published_side_is_named_after_the_data_source()
    {
        Assert.Equal("BETA source", new DraftTableFileSource(this.DataRoot, this.DraftBoardFolder, betaSource: true).PublishedLabel);
        Assert.Equal("Stable source", new DraftTableFileSource(this.DataRoot, this.DraftBoardFolder, betaSource: false).PublishedLabel);
        Assert.Equal("Your draft", this.Source().CurrentLabel);
    }

    // A file the draft has not replaced: the draft side IS the published file, as on the board.
    [Fact]
    public async Task With_no_drafted_copy_the_draft_side_is_the_published_file()
    {
        this.WritePublished(DraftTableFileSourceTests.Picture, "published");

        Assert.Equal(
            "published"u8.ToArray(),
            await this.Source().ReadAsync(DraftTableFileSourceTests.Picture, BoardTableFileSide.Current));
    }

    [Fact]
    public async Task A_path_with_no_file_reads_as_nothing_and_opens_with_a_reason()
    {
        DraftTableFileSource source = this.Source();

        Assert.Null(await source.ReadAsync("Commodore/C64/250407/missing.png", BoardTableFileSide.Published));
        Assert.Null(await source.ReadAsync("Commodore/C64/250407/missing.png", BoardTableFileSide.Current));
        Assert.Equal("There is no file at this path.", await source.OpenAsync("Commodore/C64/250407/missing.pdf", BoardTableFileSide.Current));
    }

    // ###########################################################################################
    // *** THE CELL IS TYPED TEXT. *** A path climbing out of the data folder, or an absolute one,
    // reaches no file - even one that exists where it points.
    // ###########################################################################################
    [Fact]
    public async Task A_path_that_is_not_a_safe_relative_path_reaches_no_file()
    {
        File.WriteAllText(Path.Combine(this.thisWorkspace.Root, "outside.png"), "secret");
        Directory.CreateDirectory(this.DataRoot);

        DraftTableFileSource source = this.Source();
        string absolute = Path.Combine(this.thisWorkspace.Root, "outside.png");

        foreach (string path in new[] { "../outside.png", absolute, string.Empty })
        {
            Assert.Null(await source.ReadAsync(path, BoardTableFileSide.Published));
            Assert.Null(await source.ReadAsync(path, BoardTableFileSide.Current));
        }
    }
}
