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

    private static (PublishedTableFileSource Source, List<string> Asked) Source(PublishedTableBaseline? baseline)
    {
        var asked = new List<string>();
        var handler = new AnsweringHttpHandler(request =>
        {
            asked.Add(request.RequestUri!.AbsoluteUri);
            return new HttpResponseMessage(System.Net.HttpStatusCode.OK) { Content = new ByteArrayContent([1, 2, 3]) };
        });

        var client = new ReviewApiClient("https://review.invalid", new HttpClient(handler));
        var source = new PublishedTableFileSource(client, PublishedTableFileSourceTests.BetaUrl, _ => Task.FromResult(true))
        {
            Baseline = baseline
        };

        return (source, asked);
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
