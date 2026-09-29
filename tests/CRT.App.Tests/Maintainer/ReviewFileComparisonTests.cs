using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// Covers ReviewFileComparison - the complete written list of files a submission changes on the
// server (security review, 2026-09-25).
//
// The review screen used to show images and nothing else. A PDF, a text file, a file no row used
// or a file belonging to another board appeared nowhere, so an administrator approved it without
// once seeing its name. Each test below is one thing that list must say.
public sealed class ReviewFileComparisonTests
{
    private static SubmittedFileFact Fact(
        string path,
        string sha = "new",
        string? published = null,
        SubmissionFileScope scope = SubmissionFileScope.Own,
        bool referenced = true) =>
        new(path, sha, 10, scope, referenced, published);

    [Fact]
    public void A_non_image_file_is_listed_like_any_other()
    {
        // THE gap this closes: a PDF never reached the picture panel, so it reached no panel.
        IReadOnlyList<ReviewFileLine> lines = ReviewFileComparison.Plan(
            [ReviewFileComparisonTests.Fact("Commodore/C64/250407/manual.pdf")],
            []);

        ReviewFileLine line = Assert.Single(lines);

        Assert.Equal(ReviewFileChange.Added, line.Change);
        Assert.Equal("Added: Commodore/C64/250407/manual.pdf", ReviewFileComparison.Describe(line));
    }

    [Fact]
    public void Something_already_published_at_the_path_is_a_replacement_and_an_identical_file_is_no_change()
    {
        IReadOnlyList<ReviewFileLine> lines = ReviewFileComparison.Plan(
            [
                ReviewFileComparisonTests.Fact("Commodore/C64/250407/a.png", sha: "new", published: "old"),
                ReviewFileComparisonTests.Fact("Commodore/C64/250407/b.png", sha: "same", published: "same")
            ],
            ["Commodore/C64/250407/a.png", "Commodore/C64/250407/b.png"]);

        ReviewFileLine line = Assert.Single(lines);

        Assert.Equal("Commodore/C64/250407/a.png", line.Path);
        Assert.Equal(ReviewFileChange.Replaced, line.Change);
    }

    // A file the board drops that something else still uses stays on the server. The line must
    // say exactly that rather than "deleted", and it comes first - it is the change with no
    // picture to catch the eye. (One nothing uses is REMOVED - see the tests below.)
    [Fact]
    public void A_file_the_old_board_cited_and_the_submission_drops_is_NO_LONGER_USED_and_listed_first()
    {
        IReadOnlyList<ReviewFileLine> lines = ReviewFileComparison.Plan(
            [ReviewFileComparisonTests.Fact("Commodore/C64/250407/a.png")],
            ["Commodore/C64/250407/old.pdf"]);

        Assert.Equal(ReviewFileChange.NoLongerUsed, lines[0].Change);
        Assert.Equal("No longer used (stays on the server): Commodore/C64/250407/old.pdf", ReviewFileComparison.Describe(lines[0]));
        Assert.Empty(ReviewFileComparison.Warnings(lines[0]));
    }

    // With no facts at all - an older server, or a payload that failed to load - there is nothing
    // to measure "no longer used" against, and calling every published file gone would be a loud,
    // false alarm.
    [Fact]
    public void Without_any_submitted_facts_nothing_is_claimed_to_have_gone()
    {
        Assert.Empty(ReviewFileComparison.Plan([], ["Commodore/C64/250407/a.png"]));
        Assert.Empty(ReviewFileComparison.Plan(null, ["Commodore/C64/250407/a.png"]));
    }

    // The server refuses these now. The screen flags them anyway, so it does not depend on the
    // server having been right.
    [Fact]
    public void A_file_belonging_to_ANOTHER_BOARD_or_used_by_NO_ROW_is_flagged_most_serious_first()
    {
        ReviewFileLine line = Assert.Single(ReviewFileComparison.Plan(
            [ReviewFileComparisonTests.Fact("Commodore/C64/250425/x.png", scope: SubmissionFileScope.Foreign, referenced: false)],
            []));

        IReadOnlyList<string> warnings = ReviewFileComparison.Warnings(line);

        Assert.Equal(2, warnings.Count);
        Assert.StartsWith("BELONGS TO ANOTHER BOARD", warnings[0], StringComparison.Ordinal);
        Assert.StartsWith("NO ROW USES THIS FILE", warnings[1], StringComparison.Ordinal);
    }

    // A shared file is legitimate, but a change there reaches every board that cites it - the one
    // thing a maintainer looking at one board would not otherwise think about.
    [Theory]
    [InlineData(SubmissionFileScope.ManufacturerShared, "every board of this manufacturer")]
    [InlineData(SubmissionFileScope.GenericShared, "of any manufacturer")]
    public void A_file_in_a_SHARED_folder_says_how_far_the_change_reaches(SubmissionFileScope scope, string expected)
    {
        ReviewFileLine line = Assert.Single(ReviewFileComparison.Plan(
            [ReviewFileComparisonTests.Fact("Commodore/Shared files/x.png", scope: scope)],
            []));

        Assert.Contains(expected, Assert.Single(ReviewFileComparison.Warnings(line)), StringComparison.Ordinal);
    }

    // An ordinary file in the board's own folder that a row uses needs no warning - warnings on
    // everything would teach the maintainer to skip them.
    [Fact]
    public void An_ordinary_own_file_carries_no_warning()
    {
        ReviewFileLine line = Assert.Single(ReviewFileComparison.Plan(
            [ReviewFileComparisonTests.Fact("Commodore/C64/250407/a.png")],
            []));

        Assert.Empty(ReviewFileComparison.Warnings(line));
    }
    // ###########################################################################################
    // A file nothing uses any more is REMOVED by the publish (2026-09-25), and the list says so -
    // from the server's own list, the one the approval sends back.
    // ###########################################################################################
    [Fact]
    public void A_dropped_file_on_the_servers_removal_list_is_REMOVED_and_the_others_stay()
    {
        IReadOnlyList<ReviewFileLine> lines = ReviewFileComparison.Plan(
            [ReviewFileComparisonTests.Fact("Commodore/C64/250407/a.png")],
            ["Commodore/C64/250407/old.pdf", "Commodore/Shared files/Board local files/manual.pdf"],
            new FileRemovalPreview(["Commodore/C64/250407/old.pdf"], null));

        ReviewFileLine removed = lines.Single(line => line.Path == "Commodore/C64/250407/old.pdf");
        ReviewFileLine kept = lines.Single(line => line.Path == "Commodore/Shared files/Board local files/manual.pdf");

        Assert.Equal(ReviewFileChange.Removed, removed.Change);
        Assert.Equal("REMOVED from the server (nothing uses it any more): Commodore/C64/250407/old.pdf", ReviewFileComparison.Describe(removed));
        Assert.Equal(ReviewFileChange.NoLongerUsed, kept.Change);
        Assert.Empty(ReviewFileComparison.Warnings(removed));
    }

    // The list on screen must be the WHOLE list the server removes - a removal the published file
    // list did not name in the same spelling still gets its own line.
    [Fact]
    public void Every_file_the_server_removes_gets_a_line()
    {
        IReadOnlyList<ReviewFileLine> lines = ReviewFileComparison.Plan(
            [ReviewFileComparisonTests.Fact("Commodore/C64/250407/a.png")],
            [],
            new FileRemovalPreview(["Commodore/C64/250407/Old.PDF"], null));

        ReviewFileLine line = Assert.Single(lines, candidate => candidate.Change == ReviewFileChange.Removed);
        Assert.Equal("Commodore/C64/250407/Old.PDF", line.Path);
    }
}
