using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Tests for KiCadReferenceMatcher - the matched/unmatched report an imported KiCad folder produces
// (NewContributeStrategy.md Phase 2, session 2c, task 9).
//
// The rule under test is not this class's own invention: it reproduces TabSchematics.KiCad.cs's
// BuildKiCadNormalizedNetNamesForReferences, the only place a board label becomes copper (first PCB
// file only, footprint reference, trimmed, case-insensitive). If these two ever disagree the report
// becomes a lie about what will actually light up, which is worse than no report at all.
public sealed class KiCadReferenceMatcherTests
{
    private static KiCadReferenceMatchReport Build(
        string[] labels,
        string[] references,
        int pcbFileCount = 1,
        string[]? views = null) =>
        KiCadReferenceMatcher.Build(labels, references, pcbFileCount, views);

    // ------------------------------------------------------------------------ Matching

    [Fact]
    public void A_label_with_a_footprint_of_the_same_name_matches()
    {
        var report = Build(new[] { "U8" }, new[] { "U8" });

        Assert.Equal(new[] { "U8" }, report.MatchedLabels);
        Assert.Empty(report.UnmatchedLabels);
    }

    // The render path compares with OrdinalIgnoreCase, so this must too - otherwise the report
    // would warn about a component that works perfectly well.
    [Fact]
    public void Matching_ignores_casing_exactly_as_the_overlay_does()
    {
        var report = Build(new[] { "u8" }, new[] { "U8" });

        Assert.Equal(new[] { "u8" }, report.MatchedLabels);
    }

    // The render path trims the footprint reference, so a stray space in either place is still a
    // match on screen and must be reported as one here.
    [Fact]
    public void Matching_ignores_surrounding_whitespace_on_both_sides()
    {
        var report = Build(new[] { " U8 " }, new[] { "U8  " });

        Assert.Single(report.MatchedLabels);
        Assert.Empty(report.UnmatchedLabels);
    }

    // THE failure this whole report exists for, named after the Wiki's own example: a component
    // labelled "PLA/U17" rather than "U17" highlights on the image and lights up no copper, with no
    // error anywhere in the app today.
    [Fact]
    public void A_label_that_is_not_the_bare_reference_designator_is_reported_as_unmatched()
    {
        var report = Build(new[] { "PLA/U17" }, new[] { "U17" });

        Assert.Empty(report.MatchedLabels);
        Assert.Equal(new[] { "PLA/U17" }, report.UnmatchedLabels);
    }

    [Fact]
    public void A_footprint_with_no_board_label_yet_is_listed_as_unused()
    {
        var report = Build(new[] { "U8" }, new[] { "U8", "C15" });

        Assert.Equal(new[] { "C15" }, report.UnusedReferences);
    }

    [Fact]
    public void Matched_unmatched_and_unused_are_all_reported_together()
    {
        var report = Build(new[] { "U8", "PLA/U17" }, new[] { "U8", "C15" });

        Assert.Equal(new[] { "U8" }, report.MatchedLabels);
        Assert.Equal(new[] { "PLA/U17" }, report.UnmatchedLabels);
        Assert.Equal(new[] { "C15" }, report.UnusedReferences);
    }

    // A component with one row per region is two rows but one component as far as "does this find
    // copper" is concerned - counting it twice would overstate both sides of the report.
    [Fact]
    public void A_label_listed_twice_is_counted_once()
    {
        var report = Build(new[] { "U8", "U8" }, new[] { "U8" });

        Assert.Single(report.MatchedLabels);
    }

    [Fact]
    public void Blank_labels_and_references_are_ignored_rather_than_reported()
    {
        var report = Build(new[] { "U8", "", "   " }, new[] { "U8", "", "  " });

        Assert.Single(report.MatchedLabels);
        Assert.Empty(report.UnmatchedLabels);
        Assert.Empty(report.UnusedReferences);
    }

    // Sorted so the report reads the same way twice and a long list can be scanned.
    [Fact]
    public void Each_list_is_sorted_case_insensitively()
    {
        var report = Build(new[] { "U10", "C15", "R2" }, new[] { "Q1", "D4" });

        Assert.Equal(new[] { "C15", "R2", "U10" }, report.UnmatchedLabels);
        Assert.Equal(new[] { "D4", "Q1" }, report.UnusedReferences);
    }

    // ###########################################################################################
    // *** NATURAL ORDER, for human eyes (maintainer request, 2026-09-24). ***
    //
    // Every list is shown in full as badges, and a board runs to hundreds of references. Plain
    // string order read "C1, C10, C100, C101, ..., C2", which nobody can scan for C7. All three
    // lists are checked, since each is sorted in its own place in Build.
    // ###########################################################################################
    [Fact]
    public void Each_list_is_in_NATURAL_order_so_C2_comes_before_C10()
    {
        var report = Build(
            new[] { "C100", "C10", "C2", "PLA/U17", "IC12", "IC3" },
            new[] { "C100", "C10", "C2", "R105", "R11", "R3" });

        Assert.Equal(new[] { "C2", "C10", "C100" }, report.MatchedLabels);
        Assert.Equal(new[] { "IC3", "IC12", "PLA/U17" }, report.UnmatchedLabels);
        Assert.Equal(new[] { "R3", "R11", "R105" }, report.UnusedReferences);
    }

    // ------------------------------------------------------------------------ Empty cases

    [Fact]
    public void A_board_with_no_labels_yet_reports_every_footprint_as_unused()
    {
        var report = Build(Array.Empty<string>(), new[] { "U8", "C15" });

        Assert.Empty(report.MatchedLabels);
        Assert.Empty(report.UnmatchedLabels);
        Assert.Equal(2, report.UnusedReferences.Count);
    }

    [Fact]
    public void A_project_with_no_footprints_reports_every_label_as_unmatched()
    {
        var report = Build(new[] { "U8", "C15" }, Array.Empty<string>());

        Assert.Empty(report.MatchedLabels);
        Assert.Equal(2, report.UnmatchedLabels.Count);
    }

    // ------------------------------------------------------------------------ PCB file count

    // Only the FIRST PCB file is ever used for trace matching (see the render path). A project with
    // several is a silent surprise today, so the report has to say so.
    [Fact]
    public void More_than_one_pcb_file_is_surfaced_as_a_problem()
    {
        var report = Build(new[] { "U8" }, new[] { "U8" }, pcbFileCount: 2);

        Assert.Equal(2, report.PcbFileCount);
        Assert.True(report.HasProblems);
    }

    [Fact]
    public void No_pcb_file_at_all_is_called_out_in_the_summary()
    {
        var report = Build(new[] { "U8" }, Array.Empty<string>(), pcbFileCount: 0);

        Assert.True(report.HasProblems);
        Assert.Contains("No PCB file", report.Summary);
    }

    [Fact]
    public void A_single_pcb_file_with_everything_matched_is_not_a_problem()
    {
        var report = Build(new[] { "U8" }, new[] { "U8" });

        Assert.False(report.HasProblems);
    }

    [Fact]
    public void Any_unmatched_label_makes_it_a_problem()
    {
        var report = Build(new[] { "U8", "PLA/U17" }, new[] { "U8" });

        Assert.True(report.HasProblems);
    }

    // ------------------------------------------------------------------------ The summary

    // Leads with the problem when there is one - a summary opening with the good news would bury
    // the only line that needs acting on.
    [Fact]
    public void The_summary_leads_with_the_unmatched_count_when_there_is_one()
    {
        var report = Build(new[] { "U8", "PLA/U17" }, new[] { "U8" });

        Assert.StartsWith("1 of 2", report.Summary);
        Assert.Contains("will not light up", report.Summary);
    }

    [Fact]
    public void The_summary_says_so_plainly_when_everything_matches()
    {
        var report = Build(new[] { "U8", "C15" }, new[] { "U8", "C15" });

        Assert.Contains("All 2 labelled components match", report.Summary);
    }

    [Fact]
    public void The_summary_covers_a_board_with_nothing_labelled_yet()
    {
        var report = Build(Array.Empty<string>(), new[] { "U8", "C15" });

        Assert.Contains("No components are labelled yet", report.Summary);
        Assert.Contains("2 components", report.Summary);
    }

    [Theory]
    [InlineData(1, "component")]
    [InlineData(2, "components")]
    public void The_summary_gets_its_singular_and_plural_right(int unmatchedCount, string expectedWord)
    {
        var labels = Enumerable.Range(1, unmatchedCount).Select(i => $"X{i}").Append("U8").ToArray();

        var report = Build(labels, new[] { "U8" });

        Assert.Contains($"labelled {expectedWord} found no match", report.Summary);
    }

    // ------------------------------------------------------------------------ View names

    // These are what the "CAD name" column has to carry, and the Wiki currently tells contributors
    // to read them out of the logfile by hand - surfacing them here removes that whole step.
    [Fact]
    public void The_generated_view_names_are_carried_through()
    {
        var report = Build(
            new[] { "U8" },
            new[] { "U8" },
            views: new[] { "250407_ - PCB Top", "250407_ - PCB Bottom" });

        Assert.Equal(new[] { "250407_ - PCB Top", "250407_ - PCB Bottom" }, report.ViewDisplayNames);
    }

    [Fact]
    public void Blank_view_names_are_dropped()
    {
        var report = Build(new[] { "U8" }, new[] { "U8" }, views: new[] { "PCB Top", "", "  " });

        Assert.Single(report.ViewDisplayNames);
    }

    [Fact]
    public void No_view_names_given_is_an_empty_list_rather_than_null()
    {
        var report = KiCadReferenceMatcher.Build(new[] { "U8" }, new[] { "U8" }, 1);

        Assert.NotNull(report.ViewDisplayNames);
        Assert.Empty(report.ViewDisplayNames);
    }
}
