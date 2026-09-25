using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Tests for DraftRevisionComparer - reading the free-text "# Revision date:" values CRT scans off a
// board workbook, and deciding what may honestly be said about two of them
// (NewContributeStrategy.md Phase 2, session 2d - the drift warning).
//
// The values covered here are REAL. They were read out of the shipped board workbooks in
// Assets/Data rather than invented: "2026-May-12", "2026-July-19", "2026-August-21",
// "2026-September-20". The test fixtures (BoardWorkbookBuilder) use "2026-01-15" instead, so both
// shapes exist in practice and both are pinned below.
//
// The single most important test in this file is the lexical-trap one: comparing these strings
// directly gives the WRONG ANSWER, and nothing about the values makes that obvious at a glance.
public sealed class DraftRevisionComparerTests
{
    // ------------------------------------------------------------- The real-world format

    [Fact]
    public void A_later_official_revision_in_the_real_world_format_is_newer()
    {
        Assert.Equal(
            DraftDriftState.OfficialIsNewer,
            DraftRevisionComparer.Compare("2026-May-12", "2026-August-21"));
    }

    // ###########################################################################################
    // THE LEXICAL TRAP, and the reason this class exists at all.
    //
    // As strings, "2026-August-21" sorts BEFORE "2026-May-14" - "A" comes before "M". So a plain
    // string comparison reports the official data as OLDER than the base, which is exactly
    // backwards, and the user would be told nothing had changed when in fact four months of
    // official edits had landed underneath their draft.
    //
    // Both values here are real, taken from shipped boards.
    // ###########################################################################################
    [Fact]
    public void Ordering_is_by_date_not_by_string_so_August_is_later_than_May()
    {
        // Proving the trap is real before proving the code avoids it.
        Assert.True(string.CompareOrdinal("2026-August-21", "2026-May-14") < 0);

        Assert.Equal(
            DraftDriftState.OfficialIsNewer,
            DraftRevisionComparer.Compare("2026-May-14", "2026-August-21"));
    }

    [Theory]
    [InlineData("2026-May-12", "2026-May-16")]
    [InlineData("2026-June-24", "2026-July-19")]
    [InlineData("2026-July-27", "2026-August-23")]
    [InlineData("2026-August-24", "2026-September-20")]
    public void Real_shipped_revision_pairs_order_correctly(string baseRevision, string officialRevision)
    {
        Assert.Equal(
            DraftDriftState.OfficialIsNewer,
            DraftRevisionComparer.Compare(baseRevision, officialRevision));
    }

    [Fact]
    public void An_abbreviated_month_name_parses_too()
    {
        Assert.Equal(
            DraftDriftState.OfficialIsNewer,
            DraftRevisionComparer.Compare("2026-Sep-19", "2026-Sep-20"));
    }

    // ------------------------------------------------------------- The test-fixture format

    // BoardWorkbookBuilder writes "2026-01-15" and BoardDataReaderTests uses "2026-03-01", so the
    // suite's own shape differs from the shipped data's. Both have to work.
    [Fact]
    public void The_numeric_format_used_by_the_test_fixtures_also_parses()
    {
        Assert.Equal(
            DraftDriftState.OfficialIsNewer,
            DraftRevisionComparer.Compare("2026-01-15", "2026-03-01"));
    }

    // A board contributed with one convention and synced against one using the other still has to
    // order correctly - nothing forces the two sides to match.
    [Fact]
    public void The_two_formats_can_be_compared_against_each_other()
    {
        Assert.Equal(
            DraftDriftState.OfficialIsNewer,
            DraftRevisionComparer.Compare("2026-01-15", "2026-August-21"));

        Assert.Equal(
            DraftDriftState.OfficialIsNewer,
            DraftRevisionComparer.Compare("2026-May-12", "2026-12-01"));
    }

    // ------------------------------------------------------------- In sync

    [Fact]
    public void Identical_revisions_are_in_sync()
    {
        Assert.Equal(
            DraftDriftState.InSync,
            DraftRevisionComparer.Compare("2026-August-21", "2026-August-21"));
    }

    // Two boards stamped with the same unparseable text are as much in sync as two stamped with the
    // same date - "did this change" is answerable even when "which is older" is not.
    [Fact]
    public void Identical_unparseable_revisions_are_still_in_sync()
    {
        Assert.Equal(
            DraftDriftState.InSync,
            DraftRevisionComparer.Compare("see changelog", "see changelog"));
    }

    [Fact]
    public void Surrounding_whitespace_does_not_make_two_equal_revisions_differ()
    {
        Assert.Equal(
            DraftDriftState.InSync,
            DraftRevisionComparer.Compare("  2026-May-12 ", "2026-May-12"));
    }

    // ------------------------------------------------------------- Changed, but not orderable

    // The honest fallback: the strings differ, so something changed, but no ordering claim can be
    // made. The UI must say "changed", never "newer", on this result.
    [Fact]
    public void An_unparseable_official_revision_is_changed_rather_than_newer()
    {
        Assert.Equal(
            DraftDriftState.Changed,
            DraftRevisionComparer.Compare("2026-May-12", "revision two"));
    }

    [Fact]
    public void An_unparseable_base_revision_is_changed_rather_than_newer()
    {
        Assert.Equal(
            DraftDriftState.Changed,
            DraftRevisionComparer.Compare("revision one", "2026-August-21"));
    }

    // Pins the InvariantCulture choice: the month names in the real data are English, so a machine
    // with a Danish locale must not start parsing (or failing to parse) them differently.
    [Fact]
    public void A_non_english_month_name_does_not_parse_and_degrades_to_changed()
    {
        Assert.Equal(
            DraftDriftState.Changed,
            DraftRevisionComparer.Compare("2026-Maj-12", "2026-August-21"));
    }

    // Two different spellings of the same date: it genuinely changed as text, but neither is later,
    // so there is no ordering to claim.
    [Fact]
    public void Two_spellings_of_the_same_date_are_changed_not_newer()
    {
        Assert.Equal(
            DraftDriftState.Changed,
            DraftRevisionComparer.Compare("2026-05-12", "2026-May-12"));
    }

    // A rollback, or a hand-edited workbook. It is still "changed underneath you" - folded in here
    // rather than becoming a fourth state nobody could act on differently.
    [Fact]
    public void An_official_revision_older_than_the_base_is_reported_as_changed()
    {
        Assert.Equal(
            DraftDriftState.Changed,
            DraftRevisionComparer.Compare("2026-August-21", "2026-May-12"));
    }

    [Fact]
    public void A_missing_official_revision_against_a_recorded_base_is_changed()
    {
        Assert.Equal(
            DraftDriftState.Changed,
            DraftRevisionComparer.Compare("2026-May-12", ""));
    }

    // ------------------------------------------------------------- Unknown

    // A blank base means the base was never recorded, which is NOT evidence that anything drifted.
    // Warning on it would be a claim the data cannot support.
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    public void A_blank_base_revision_is_unknown_rather_than_changed(string? baseRevision)
    {
        Assert.Equal(
            DraftDriftState.Unknown,
            DraftRevisionComparer.Compare(baseRevision, "2026-August-21"));
    }

    [Fact]
    public void Both_sides_blank_is_unknown()
    {
        Assert.Equal(DraftDriftState.Unknown, DraftRevisionComparer.Compare(null, null));
    }

    // ------------------------------------------------------------- TryParse itself

    [Theory]
    [InlineData("2026-May-12")]
    [InlineData("2026-September-20")]
    [InlineData("2026-Sep-20")]
    [InlineData("2026-01-15")]
    [InlineData("2026-1-5")]
    public void Every_accepted_shape_parses(string revision)
    {
        Assert.True(DraftRevisionComparer.TryParse(revision, out _));
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("garbage")]
    [InlineData("2026-Maj-12")]
    [InlineData("12-05-2026")]
    public void Nothing_else_parses(string? revision)
    {
        Assert.False(DraftRevisionComparer.TryParse(revision, out _));
    }

    // Verified empirically: TryParseExact rejects untrimmed input, so the trim inside TryParse is
    // load-bearing rather than tidiness - a spreadsheet cell easily carries a trailing space.
    [Fact]
    public void Untrimmed_input_still_parses_because_the_comparer_trims_first()
    {
        Assert.True(DraftRevisionComparer.TryParse("  2026-May-12  ", out DateTime parsed));
        Assert.Equal(new DateTime(2026, 5, 12), parsed);
    }
}
