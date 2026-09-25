using System.Text.Json;
using CRT.Review.Handlers;
using Handlers.DataHandling;

namespace CRT.Review.Tests;

// Covers ReviewSummaryPresenter - what the reviewer's landing view actually says.
//
// WHY THE WORDING IS TESTED AND NOT JUST EYEBALLED: a reviewer decides whether to look closer
// from these lines, and approving is irreversible (the maintainer ruled out retained revisions).
// A line that buries a removal, or that omits a section, is how a change gets approved unseen.
//
// The comparison logic itself lives in CRT.Data and is covered by ReviewSummaryTests; what is
// tested here is presentation - order, omission and phrasing.
public sealed class ReviewSummaryPresenterTests
{
    // ###########################################################################################
    // Runs a real comparison through the REAL WIRE FORMAT - serialise exactly as ASP.NET Core
    // does, then parse with the app's own parser - rather than handing the presenter a view
    // object built by hand.
    //
    // *** THIS IS THE STRONGEST TEST AVAILABLE WITHOUT A LIVE SERVER, and it is worth the extra
    // step. *** The two sides do not share a type: CRT.Data computes ReviewChangeSummary and the
    // app draws ReviewChangeSummaryView. A hand-built view would let the pair drift silently -
    // exactly the defect SubmissionContract.cs exists to prevent elsewhere - whereas this fails
    // the moment a field is renamed on either side.
    // ###########################################################################################
    private static ReviewChangeSummaryView Wire(ReviewChangeSummary summary)
    {
        string json = JsonSerializer.Serialize(
            new { submission = new { id = 1 }, changes = summary, findings = Array.Empty<object>() },
            new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });

        ReviewSubmissionDetail? detail = ReviewApiParser.ParseSubmission(json);

        Assert.NotNull(detail);
        Assert.NotNull(detail!.Changes);

        return detail.Changes!;
    }

    private static ComponentEntry Component(string label, string part = "906114") => new()
    {
        BoardLabel = label,
        FriendlyName = "PLA",
        TechnicalNameOrValue = "906114-01",
        PartNumber = part
    };

    private static BoardData Board(params ComponentEntry[] components) => new()
    {
        RevisionDate = "2026-August-21",
        Components = [.. components]
    };

    [Fact]
    public void A_section_with_no_changes_is_omitted_entirely()
    {
        // Showing "0 changed" for ten sections buries the one that matters in nine lines of
        // nothing - the same failure as opening on the whole board, just smaller.
        BoardData published = ReviewSummaryPresenterTests.Board(ReviewSummaryPresenterTests.Component("U8"));
        BoardData submitted = ReviewSummaryPresenterTests.Board(
            ReviewSummaryPresenterTests.Component("U8"),
            ReviewSummaryPresenterTests.Component("U9"));

        IReadOnlyList<ReviewSummaryLine> lines =
            ReviewSummaryPresenter.BuildLines(ReviewSummaryPresenterTests.Wire(ReviewSummary.Compare(published, submitted)));

        ReviewSummaryLine line = Assert.Single(lines);
        Assert.Equal(BoardWorkbookSchema.SheetComponents, line.Section);
    }

    [Fact]
    public void An_unchanged_submission_produces_no_lines_at_all()
    {
        BoardData board = ReviewSummaryPresenterTests.Board(ReviewSummaryPresenterTests.Component("U8"));

        Assert.Empty(ReviewSummaryPresenter.BuildLines(ReviewSummaryPresenterTests.Wire(ReviewSummary.Compare(board, board))));
    }

    [Fact]
    public void Removals_are_listed_FIRST_within_a_section()
    {
        // *** THE ORDERING THAT MATTERS. *** A removal is the least recoverable thing a
        // submission can do and the easiest to skim past, because a reviewer scanning for "what
        // did they add" is not looking for it. So it leads.
        BoardData published = ReviewSummaryPresenterTests.Board(
            ReviewSummaryPresenterTests.Component("U8"),
            ReviewSummaryPresenterTests.Component("U9"));

        BoardData submitted = ReviewSummaryPresenterTests.Board(
            ReviewSummaryPresenterTests.Component("U8", part: "changed"),
            ReviewSummaryPresenterTests.Component("U10"));

        ReviewSummaryLine line = Assert.Single(
            ReviewSummaryPresenter.BuildLines(ReviewSummaryPresenterTests.Wire(ReviewSummary.Compare(published, submitted))));

        Assert.Equal(ReviewChangeKind.Removed, line.Parts[0].Kind);
        Assert.Equal(ReviewChangeKind.Changed, line.Parts[1].Kind);
        Assert.Equal(ReviewChangeKind.Added, line.Parts[2].Kind);
    }

    [Fact]
    public void A_field_level_diff_SURVIVES_THE_WIRE()
    {
        // *** END TO END FOR TASK 4. *** The field diff is computed in CRT.Data, serialised by the
        // server, parsed by the app and worded by the presenter - four places, no shared type.
        // This runs a real edit through the whole chain, so a rename on any of them fails here
        // rather than showing a reviewer a row with no detail under it.
        BoardData published = ReviewSummaryPresenterTests.Board(
            ReviewSummaryPresenterTests.Component("U8", part: "906114"));

        BoardData submitted = ReviewSummaryPresenterTests.Board(
            ReviewSummaryPresenterTests.Component("U8", part: "251715-01"));

        ReviewChangeSummaryView view =
            ReviewSummaryPresenterTests.Wire(ReviewSummary.Compare(published, submitted));

        ReviewSectionView section = Assert.Single(view.Sections);

        Assert.Equal(["U8"], section.Changed);

        ReviewFieldChangeView field = Assert.Single(section.FieldChanges["U8"]);

        Assert.Equal("Part-number", field.Field);
        Assert.Equal(
            "Part-number: 906114 -> 251715-01",
            ReviewSummaryPresenter.DescribeFieldChange(field));
    }

    [Fact]
    public void A_part_describes_itself_with_the_count_LEADING()
    {
        // Matches how the rest of the project draws a counted pill: the number first, then the
        // word. Consistency across the two apps is worth more than it looks - the maintainer uses
        // both.
        var part = new ReviewSummaryPart(ReviewChangeKind.Removed, 2, ["U8", "U9"]);

        Assert.Equal("2 removed", part.Describe());
    }

    [Fact]
    public void A_part_carries_the_keys_so_a_drill_down_needs_no_second_comparison()
    {
        BoardData published = ReviewSummaryPresenterTests.Board(ReviewSummaryPresenterTests.Component("U8"));
        BoardData submitted = ReviewSummaryPresenterTests.Board(
            ReviewSummaryPresenterTests.Component("U8"),
            ReviewSummaryPresenterTests.Component("U9"));

        ReviewSummaryLine line = Assert.Single(
            ReviewSummaryPresenter.BuildLines(ReviewSummaryPresenterTests.Wire(ReviewSummary.Compare(published, submitted))));

        Assert.Equal(["U9"], Assert.Single(line.Parts).Keys);
    }

    // -----------------------------------------------------------------------------------
    // Renames
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_clean_rename_reads_as_an_arrow_and_nothing_more()
    {
        string described = ReviewSummaryPresenter.DescribeRename(new ReviewRenameView("U8", "U9", AlsoChanged: false));

        Assert.Equal("U8 -> U9", described);
    }

    [Fact]
    public void A_rename_that_also_edited_something_SAYS_SO()
    {
        // Without this, "renamed" invites a reviewer not to look further - which is exactly when
        // an edit hidden behind a rename gets through.
        string described = ReviewSummaryPresenter.DescribeRename(new ReviewRenameView("U8", "U9", AlsoChanged: true));

        Assert.Equal("U8 -> U9 (and edited)", described);
    }

    [Fact]
    public void A_renamed_row_appears_in_the_line_as_a_rename()
    {
        BoardData published = ReviewSummaryPresenterTests.Board(ReviewSummaryPresenterTests.Component("U8"));
        BoardData submitted = ReviewSummaryPresenterTests.Board(ReviewSummaryPresenterTests.Component("U9"));

        ReviewChangeSummary summary = ReviewSummary.Compare(
            published,
            submitted,
            [new SubmissionRename { Section = BoardWorkbookSchema.SheetComponents, From = "U8", To = "U9" }]);

        ReviewSummaryLine line = Assert.Single(
            ReviewSummaryPresenter.BuildLines(ReviewSummaryPresenterTests.Wire(summary)));
        ReviewSummaryPart part = Assert.Single(line.Parts);

        Assert.Equal(ReviewChangeKind.Renamed, part.Kind);
        Assert.Equal(["U8 -> U9"], part.Keys);
    }

    // -----------------------------------------------------------------------------------
    // The headline
    // -----------------------------------------------------------------------------------

    [Fact]
    public void A_new_system_is_called_out_in_the_headline()
    {
        // The highest-risk submission there is (Phase 6 task 3). A reviewer must never have to
        // infer it from row counts.
        BoardData submitted = ReviewSummaryPresenterTests.Board(ReviewSummaryPresenterTests.Component("U8"));

        string headline = ReviewSummaryPresenter.BuildHeadline(ReviewSummaryPresenterTests.Wire(ReviewSummary.Compare(published: null, submitted)));

        Assert.Contains("New system", headline);
    }

    [Fact]
    public void An_unchanged_submission_says_so_rather_than_showing_a_blank_headline()
    {
        BoardData board = ReviewSummaryPresenterTests.Board(ReviewSummaryPresenterTests.Component("U8"));

        Assert.Equal("No changes", ReviewSummaryPresenter.BuildHeadline(ReviewSummaryPresenterTests.Wire(ReviewSummary.Compare(board, board))));
    }

    [Fact]
    public void A_null_summary_is_refused_rather_than_drawn_as_empty()
    {
        Assert.Throws<ArgumentNullException>(() => ReviewSummaryPresenter.BuildLines(null!));
        Assert.Throws<ArgumentNullException>(() => ReviewSummaryPresenter.BuildHeadline(null!));
    }
}
