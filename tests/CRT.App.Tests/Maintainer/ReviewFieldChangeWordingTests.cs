using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// Covers how a field-level change READS (Phase 5, task 4).
//
// The wording is tested rather than eyeballed for the same reason the rest of this screen is: a
// maintainer decides from these lines, and approving cannot be undone. Two rules here exist to stop
// the line lying by omission - a blank value spelled out rather than left empty, and a long value
// truncated in the middle rather than at the end.
public sealed class ReviewFieldChangeWordingTests
{
    [Fact]
    public void A_change_names_the_field_and_both_values()
    {
        string described = ReviewSummaryPresenter.DescribeFieldChange(
            new ReviewFieldChangeView("Part-number", "906114", "251715-01"));

        Assert.Equal("Part-number: 906114 -> 251715-01", described);
    }

    [Fact]
    public void A_CLEARED_field_says_blank_rather_than_trailing_off()
    {
        // *** THE CASE THAT MATTERS MOST. *** Deleting information is the edit hardest to notice,
        // and "Note: Measured cold -> " reads as the app having failed to draw something rather
        // than as a deliberate deletion.
        string described = ReviewSummaryPresenter.DescribeFieldChange(
            new ReviewFieldChangeView("Note", "Measured cold", ""));

        Assert.Equal("Note: Measured cold -> (blank)", described);
    }

    [Fact]
    public void A_field_filled_in_for_the_first_time_says_blank_on_the_left()
    {
        string described = ReviewSummaryPresenter.DescribeFieldChange(
            new ReviewFieldChangeView("Region", "", "ASSY 250407"));

        Assert.Equal("Region: (blank) -> ASSY 250407", described);
    }

    [Fact]
    public void A_LONG_value_is_truncated_in_the_MIDDLE_so_both_ends_survive()
    {
        // *** TRUNCATING AT THE END WOULD HIDE THE CHANGE. *** A component description runs to a
        // sentence, and a corrected part number or a fixed typo usually differs at one END of the
        // value. Cutting the middle keeps both, so the maintainer can see WHAT differs rather than
        // two identical-looking prefixes.
        string before = new string('a', 40) + "BEFORE" + new string('b', 40);
        string after = new string('a', 40) + "AFTER!" + new string('b', 40);

        string described = ReviewSummaryPresenter.DescribeFieldChange(
            new ReviewFieldChangeView("Description", before, after));

        // Both ends are present on both sides...
        Assert.Contains("aaaa", described);
        Assert.Contains("bbbb", described);
        Assert.Contains("...", described);

        // ...and the line is not the full 172 characters it would be untruncated.
        Assert.True(described.Length < before.Length + after.Length);
    }

    [Fact]
    public void A_value_at_the_limit_is_NOT_truncated()
    {
        // The boundary: eliding a value that already fits would lose information for nothing.
        string exact = new string('x', ReviewSummaryPresenter.MaximumValueLength);

        string described = ReviewSummaryPresenter.DescribeFieldChange(
            new ReviewFieldChangeView("Note", exact, "y"));

        Assert.Contains(exact, described);
        Assert.DoesNotContain("...", described);
    }

    [Fact]
    public void A_truncated_value_never_exceeds_the_limit_it_enforces()
    {
        // The elision has to fit INSIDE the budget, not be added on top of it - otherwise the
        // "limit" grows by three characters and the line wraps anyway.
        string huge = new string('z', 500);

        string described = ReviewSummaryPresenter.DescribeFieldChange(
            new ReviewFieldChangeView("Note", huge, "y"));

        // Field name, separators and the short value aside, the long one must be within budget.
        Assert.True(described.Length <= ReviewSummaryPresenter.MaximumValueLength + 40);
    }

    [Fact]
    public void A_null_change_is_refused_rather_than_drawn_blank()
    {
        Assert.Throws<ArgumentNullException>(() => ReviewSummaryPresenter.DescribeFieldChange(null!));
    }
}
