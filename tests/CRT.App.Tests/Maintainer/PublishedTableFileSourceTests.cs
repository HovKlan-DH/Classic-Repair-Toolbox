using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Where the Boards screen's table reads a file cell's two pictures (PublishedTableFileSource). The
// requests go to AnsweringHttpHandler, which answers from a function - no network (test rule 6).
//
// *** COMPARED, THE "BEFORE" SIDE IS THE OTHER SOURCE'S FILE (owner request, 2026-10-09 - "Compare
// sources"). *** Read from BETA's address it would show BETA's own picture on both sides of a cell
// changed between BETA and the stable source - or "no file at this path" for a file only the
// stable source still cites.
// ###########################################################################################
public sealed class PublishedTableFileSourceTests
{
    private const string BetaUrl = "https://example.org/beta";

    private const string StableUrl = "https://example.org/stable";

    private static (PublishedTableFileSource Source, List<string> Asked) Source(
        PublishedTableBaseline? baseline,
        IReadOnlyDictionary<string, string>? places = null)
    {
        var asked = new List<string>();
        var handler = new AnsweringHttpHandler(request =>
        {
            asked.Add(request.RequestUri!.AbsoluteUri);
            return new HttpResponseMessage(System.Net.HttpStatusCode.OK) { Content = new ByteArrayContent([1, 2, 3]) };
        });

        var client = new ReviewApiClient("https://review.invalid", new HttpClient(handler));
        var source = new PublishedTableFileSource(client, PublishedTableFileSourceTests.BetaUrl, _ => Task.FromResult(true), places: places)
        {
            Baseline = baseline
        };

        return (source, asked);
    }

    // ###########################################################################################
    // *** A CRT 4.0 BOARD'S FILE IS FETCHED FROM ITS STORED PLACE (2026-10-25). *** A cell names the
    // file as the board does ("Images/Top.png"); its bytes lie under their checksum, which the
    // table's answer says (BoardTableAnswer.FilePlaces) - for each SIDE from its own source, since
    // the same name can hold other bytes in BETA and in the stable source. A name the answer does
    // not place (an older server's path) is fetched as it is.
    // ###########################################################################################
    [Fact]
    public async Task A_named_file_is_fetched_from_the_place_its_own_sides_answer_gives()
    {
        var baseline = new PublishedTableBaseline(
            PublishedTableFileSourceTests.StableUrl, BoardSections.StableTreeName, BoardSections.StableSourceSide, BoardSections.BetaSourceSide,
            FilePlaces: new Dictionary<string, string> { ["Images/Top.png"] = "Files/9c/9c01.png" });

        var (source, asked) = PublishedTableFileSourceTests.Source(
            baseline,
            places: new Dictionary<string, string> { ["Images/Top.png"] = "Files/3f/3f7a.png" });

        Assert.Equal("Files/3f/3f7a.png", source.PlaceFor(CRT.BoardTableFileSide.Current, " Images/Top.png "));
        Assert.Equal("Files/9c/9c01.png", source.PlaceFor(CRT.BoardTableFileSide.Published, "Images/Top.png"));
        Assert.Equal("Commodore/C64/250407/Old.png", source.PlaceFor(CRT.BoardTableFileSide.Current, "Commodore/C64/250407/Old.png"));

        await source.ReadAsync("Images/Top.png", CRT.BoardTableFileSide.Current);
        await source.ReadAsync("Images/Top.png", CRT.BoardTableFileSide.Published);

        Assert.Equal(2, asked.Count);
        Assert.EndsWith("Files/3f/3f7a.png", asked[0], StringComparison.Ordinal);
        Assert.StartsWith(PublishedTableFileSourceTests.StableUrl, asked[1], StringComparison.Ordinal);
        Assert.EndsWith("Files/9c/9c01.png", asked[1], StringComparison.Ordinal);
    }

    // ###########################################################################################
    // A cell naming the file in other capitals is the same file (code review, 2026-10-26): the server
    // builds the places ignoring capitals, but JSON hands them over as an ordinary dictionary, so the
    // second spelling was fetched by its name - which a 4.0 tree never holds - and the card said the
    // file was not in BETA.
    // ###########################################################################################
    [Fact]
    public void A_name_in_other_capitals_is_placed_as_the_board_names_it_on_both_sides()
    {
        var baseline = new PublishedTableBaseline(
            PublishedTableFileSourceTests.StableUrl, BoardSections.StableTreeName, BoardSections.StableSourceSide, BoardSections.BetaSourceSide,
            FilePlaces: new Dictionary<string, string> { ["Component images/U8.png"] = "Files/9c/9c01.png" });

        var (source, _) = PublishedTableFileSourceTests.Source(
            baseline,
            places: new Dictionary<string, string> { ["Component images/U8.png"] = "Files/3f/3f7a.png" });

        Assert.Equal("Files/3f/3f7a.png", source.PlaceFor(CRT.BoardTableFileSide.Current, "component images/u8.png"));
        Assert.Equal("Files/9c/9c01.png", source.PlaceFor(CRT.BoardTableFileSide.Published, "COMPONENT IMAGES/U8.PNG"));
    }

    // The ordinary table: both sides from its own source, headed "BETA now" / "Your change".
    [Fact]
    public async Task Uncompared_both_sides_come_from_the_tables_own_source()
    {
        var (source, asked) = PublishedTableFileSourceTests.Source(baseline: null);

        await source.ReadAsync("Commodore/C64/250407/Top.png", CRT.BoardTableFileSide.Published);
        await source.ReadAsync("Commodore/C64/250407/Top.png", CRT.BoardTableFileSide.Current);

        Assert.All(asked, url => Assert.StartsWith(PublishedTableFileSourceTests.BetaUrl, url, StringComparison.Ordinal));
        Assert.Equal(BoardSections.BaselineLabel, source.PublishedLabel);
        Assert.Equal(BoardSections.ChangeLabel, source.CurrentLabel);
    }

    // Compared with the stable source: the side compared with is read from the stable source's
    // address - a request of its own, not BETA's file cached under the same path - and each picture
    // is headed by its source.
    [Fact]
    public async Task Compared_the_side_compared_with_comes_from_the_other_source()
    {
        var baseline = new PublishedTableBaseline(
            PublishedTableFileSourceTests.StableUrl, BoardSections.StableTreeName, BoardSections.StableSourceSide, BoardSections.BetaSourceSide);

        var (source, asked) = PublishedTableFileSourceTests.Source(baseline);

        await source.ReadAsync("Commodore/C64/250407/Top.png", CRT.BoardTableFileSide.Current);
        await source.ReadAsync("Commodore/C64/250407/Top.png", CRT.BoardTableFileSide.Published);

        Assert.Equal(2, asked.Count);
        Assert.StartsWith(PublishedTableFileSourceTests.BetaUrl, asked[0], StringComparison.Ordinal);
        Assert.StartsWith(PublishedTableFileSourceTests.StableUrl, asked[1], StringComparison.Ordinal);
        Assert.Equal("Stable source", source.PublishedLabel);
        Assert.Equal("BETA source", source.CurrentLabel);
    }
}
