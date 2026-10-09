using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// Covers ReviewImageComparison - deciding WHICH images a maintainer is shown side by side, and how
// each pair is labelled (Phase 5, task 4).
//
// This is the logic that makes the asset endpoint useful. Fetching bytes is the easy half; the
// half that can be wrong is picking the right pair, and getting it wrong is silent - a maintainer
// comparing the wrong two images approves on the strength of a comparison that never happened.
//
// The three cases are deliberately NOT collapsed into one "changed" idea. A replaced image, a
// brand-new one and a deleted one are different decisions for the person looking, and the
// riskiest of the three is the one easiest to skim past.
public sealed class ReviewImageComparisonTests
{
    private static string Hash(char fill) => new(fill, 64);

    private static ReviewSubmissionAssets Assets(params (string Path, string Hash)[] files) =>
        new(files.Select(file => new ReviewSubmittedFile(file.Path, file.Hash)).ToList());

    [Fact]
    public void A_REPLACED_image_is_paired_before_and_after()
    {
        // The core case: the file name is unchanged, so only the bytes moved. Both sides exist
        // and the maintainer needs to see them together - this is the comparison that cannot be
        // done by reading a row diff, and the whole reason this is a desktop app.
        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            ReviewImageComparisonTests.Assets(("Images/U8.png", ReviewImageComparisonTests.Hash('a'))),
            publishedImagePaths: ["Images/U8.png"]);

        ReviewImagePair pair = Assert.Single(pairs);

        Assert.Equal("Images/U8.png", pair.Path);
        Assert.Equal(ReviewImageChange.Replaced, pair.Change);
        Assert.Equal(ReviewImageComparisonTests.Hash('a'), pair.SubmittedHash);
        Assert.True(pair.HasBefore);
        Assert.True(pair.HasAfter);
    }

    [Fact]
    public void An_ADDED_image_has_no_before_side()
    {
        // Common and low-risk, but it must not be drawn as a comparison against a blank panel -
        // an empty "before" reads as "the old one failed to load", which sends a maintainer looking
        // for a fault that is not there.
        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            ReviewImageComparisonTests.Assets(("Images/new.png", ReviewImageComparisonTests.Hash('b'))),
            publishedImagePaths: []);

        ReviewImagePair pair = Assert.Single(pairs);

        Assert.Equal(ReviewImageChange.Added, pair.Change);
        Assert.False(pair.HasBefore);
        Assert.True(pair.HasAfter);
    }

    [Fact]
    public void A_REMOVED_image_is_shown_even_though_the_submission_carries_no_bytes_for_it()
    {
        // *** THE CASE MOST WORTH GETTING RIGHT. *** A deletion is the least recoverable thing a
        // submission can do, and it is invisible in a list of what was uploaded - there is no
        // blob for it, so anything driven purely off the manifest's file list misses it entirely.
        // It is found by looking at what the published board HAS and the submission does not.
        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            ReviewImageComparisonTests.Assets(),
            publishedImagePaths: ["Images/gone.png"]);

        ReviewImagePair pair = Assert.Single(pairs);

        Assert.Equal("Images/gone.png", pair.Path);
        Assert.Equal(ReviewImageChange.Removed, pair.Change);
        Assert.True(pair.HasBefore);
        Assert.False(pair.HasAfter);
        Assert.Equal(string.Empty, pair.SubmittedHash);
    }

    [Fact]
    public void REMOVALS_are_listed_FIRST()
    {
        // The same rule ReviewSummaryPresenter already follows, and for the same reason: a
        // maintainer scanning for "what did they add" is not looking for a removal, so it has to be
        // where the eye lands rather than below a list of additions.
        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            ReviewImageComparisonTests.Assets(
                ("Images/added.png", ReviewImageComparisonTests.Hash('a')),
                ("Images/replaced.png", ReviewImageComparisonTests.Hash('b'))),
            publishedImagePaths: ["Images/replaced.png", "Images/gone.png"]);

        Assert.Equal(3, pairs.Count);
        Assert.Equal(ReviewImageChange.Removed, pairs[0].Change);
    }

    [Fact]
    public void A_NON_IMAGE_file_is_not_offered_for_visual_comparison()
    {
        // A datasheet is a real part of a submission and a maintainer may well want to open it, but
        // it is not something to draw side by side. Including it would put two blank panels on
        // screen captioned as a comparison.
        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            ReviewImageComparisonTests.Assets(
                ("Files/datasheet.pdf", ReviewImageComparisonTests.Hash('a')),
                ("Files/notes.txt", ReviewImageComparisonTests.Hash('b')),
                ("Images/real.png", ReviewImageComparisonTests.Hash('c'))),
            publishedImagePaths: []);

        ReviewImagePair pair = Assert.Single(pairs);
        Assert.Equal("Images/real.png", pair.Path);
    }

    [Fact]
    public void An_UNCHANGED_image_is_not_shown_at_all()
    {
        // The submission's file list is the COMPLETE intended state, not a list of changes, so an
        // untouched board re-lists every image it already had. Showing all of them would bury the
        // one that moved - the same failure as opening on the whole board.
        //
        // Sameness is decided by the HASH matching what is published, which the caller supplies.
        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            ReviewImageComparisonTests.Assets(("Images/same.png", ReviewImageComparisonTests.Hash('a'))),
            publishedImagePaths: ["Images/same.png"],
            publishedHashesByPath: new Dictionary<string, string>
            {
                ["Images/same.png"] = ReviewImageComparisonTests.Hash('a')
            });

        Assert.Empty(pairs);
    }

    [Fact]
    public void A_DIFFERENT_hash_at_the_same_path_IS_shown()
    {
        // The anti-vacuity half of the test above: with hashes supplied, a genuine replacement
        // must still come through. Otherwise "unchanged images are hidden" could be hiding
        // everything.
        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            ReviewImageComparisonTests.Assets(("Images/same.png", ReviewImageComparisonTests.Hash('b'))),
            publishedImagePaths: ["Images/same.png"],
            publishedHashesByPath: new Dictionary<string, string>
            {
                ["Images/same.png"] = ReviewImageComparisonTests.Hash('a')
            });

        Assert.Equal(ReviewImageChange.Replaced, Assert.Single(pairs).Change);
    }

    [Fact]
    public void Paths_are_compared_CASE_SENSITIVELY()
    {
        // The server's filesystem is Linux and is case-sensitive, and the data tree is already
        // documented as case-sensitive from Phase 3 onward. Treating "U8.png" and "u8.png" as one
        // path here would pair an added file against an unrelated published one and show the
        // maintainer a comparison between two different images.
        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            ReviewImageComparisonTests.Assets(("Images/U8.png", ReviewImageComparisonTests.Hash('a'))),
            publishedImagePaths: ["Images/u8.png"]);

        Assert.Equal(2, pairs.Count);
        Assert.Contains(pairs, pair => pair.Change == ReviewImageChange.Added);
        Assert.Contains(pairs, pair => pair.Change == ReviewImageChange.Removed);
    }

    [Fact]
    public void A_NEW_BOARD_publishes_nothing_before_so_every_image_is_an_addition()
    {
        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            ReviewImageComparisonTests.Assets(
                ("a.png", ReviewImageComparisonTests.Hash('a')),
                ("b.png", ReviewImageComparisonTests.Hash('b'))),
            publishedImagePaths: []);

        Assert.Equal(2, pairs.Count);
        Assert.All(pairs, pair => Assert.Equal(ReviewImageChange.Added, pair.Change));
    }

    [Fact]
    public void Null_inputs_produce_an_empty_plan_rather_than_throwing()
    {
        // This is called while drawing a panel, from a submission whose payload may not have
        // loaded. Throwing would take the whole review screen down over a missing list.
        Assert.Empty(ReviewImageComparison.Plan(null, null));
    }

    [Fact]
    public void Each_pair_appears_ONCE_even_when_a_path_is_listed_twice()
    {
        // A manifest is contributor-supplied and can carry a duplicate row. Two identical panels
        // side by side would read as two different images that happen to look the same.
        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            ReviewImageComparisonTests.Assets(
                ("Images/U8.png", ReviewImageComparisonTests.Hash('a')),
                ("Images/U8.png", ReviewImageComparisonTests.Hash('a'))),
            publishedImagePaths: []);

        Assert.Single(pairs);
    }

    [Theory]
    [InlineData("a.png")]
    [InlineData("a.PNG")]
    [InlineData("a.jpg")]
    [InlineData("a.jpeg")]
    [InlineData("a.gif")]
    [InlineData("a.bmp")]
    [InlineData("a.webp")]
    public void The_image_extensions_match_what_the_SERVER_will_actually_serve(string path)
    {
        // A contract with ReviewAssetLocator's allowlist. If the Maintainer tab offers a comparison for a
        // type the server refuses to serve as an image, the panel shows a download instead of a
        // picture - and the mismatch is invisible until somebody opens that exact submission.
        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            ReviewImageComparisonTests.Assets((path, ReviewImageComparisonTests.Hash('a'))),
            publishedImagePaths: []);

        Assert.Single(pairs);
    }

    // -----------------------------------------------------------------------------------------
    // How each pair is captioned. Wording is tested rather than eyeballed for the same reason the
    // rest of this screen is: a maintainer decides from these lines, and approving cannot be undone.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void Each_of_the_three_outcomes_is_NAMED_in_its_caption()
    {
        Assert.Equal("Added: a.png", ReviewSummaryPresenter.DescribeImagePair(
            new ReviewImagePair("a.png", ReviewImageChange.Added, "x")));

        Assert.Equal("Replaced: a.png", ReviewSummaryPresenter.DescribeImagePair(
            new ReviewImagePair("a.png", ReviewImageChange.Replaced, "x")));

        // *** SHOUTED, and deliberately not softened to "Changed" - but TRUTHFUL. *** The board
        // stops citing the file; whether the publish then REMOVES it (nothing else uses it) or it
        // stays is the file list's to say, from the server's own list (2026-09-25). The caption
        // claims neither, because it cannot know.
        Assert.Equal("NO LONGER USED: a.png", ReviewSummaryPresenter.DescribeImagePair(
            new ReviewImagePair("a.png", ReviewImageChange.Removed, "")));
    }

    // ###########################################################################################
    // *** SOMETHING ALREADY PUBLISHED AT THE PATH MAKES IT A REPLACEMENT, whether or not the old
    // board cited it (security review, 2026-09-25). *** Judging only by the old board's file list
    // called a file that OVERWRITES an existing one "Added" - backwards for exactly the case a
    // maintainer most needs to see. The server now sends a published hash for every submitted path
    // that exists on disk.
    // ###########################################################################################
    [Fact]
    public void A_path_the_old_board_never_cited_but_that_EXISTS_on_the_server_is_a_replacement()
    {
        var submitted = new ReviewSubmissionAssets([new ReviewSubmittedFile("Commodore/C64/250425/x.png", "new")]);

        IReadOnlyList<ReviewImagePair> pairs = ReviewImageComparison.Plan(
            submitted,
            publishedImagePaths: [],
            publishedHashesByPath: new Dictionary<string, string> { ["Commodore/C64/250425/x.png"] = "old" });

        Assert.Equal(ReviewImageChange.Replaced, Assert.Single(pairs).Change);
    }

    [Fact]
    public void A_MISSING_side_is_captioned_rather_than_left_blank()
    {
        // An empty panel beside a full one reads as an image that failed to load, which sends a
        // maintainer looking for a fault that is not there.
        Assert.Equal(
            "Not in the published board",
            ReviewSummaryPresenter.DescribeMissingSide(ReviewImageChange.Added, isBeforeSide: true));

        // Neither "deleted" nor "stays": since 2026-09-25 a file nothing else uses IS removed, and
        // the file list above says which. The caption says only what is true of both.
        Assert.Equal(
            "No longer used by the board",
            ReviewSummaryPresenter.DescribeMissingSide(ReviewImageChange.Removed, isBeforeSide: false));
    }

    [Fact]
    public void A_side_that_HAS_a_picture_gets_no_placeholder_text()
    {
        // The caption is for an absent side only - writing one under a real image would caption
        // the picture with a sentence about it not being there.
        Assert.Equal(string.Empty,
            ReviewSummaryPresenter.DescribeMissingSide(ReviewImageChange.Replaced, isBeforeSide: true));

        Assert.Equal(string.Empty,
            ReviewSummaryPresenter.DescribeMissingSide(ReviewImageChange.Replaced, isBeforeSide: false));

        Assert.Equal(string.Empty,
            ReviewSummaryPresenter.DescribeMissingSide(ReviewImageChange.Added, isBeforeSide: false));

        Assert.Equal(string.Empty,
            ReviewSummaryPresenter.DescribeMissingSide(ReviewImageChange.Removed, isBeforeSide: true));
    }

    [Fact]
    public void A_null_pair_is_refused_rather_than_captioned_blank()
    {
        Assert.Throws<ArgumentNullException>(
            () => ReviewSummaryPresenter.DescribeImagePair(null!));
    }

    [Fact]
    public void An_SVG_is_NOT_offered_as_an_image()
    {
        // The server deliberately refuses to serve SVG as an image - it is a scriptable XML
        // document. Offering it here would produce a comparison panel the server answers with
        // octet-stream, which cannot be decoded and draws as a failure.
        Assert.Empty(ReviewImageComparison.Plan(
            ReviewImageComparisonTests.Assets(("a.svg", ReviewImageComparisonTests.Hash('a'))),
            publishedImagePaths: []));
    }
}
