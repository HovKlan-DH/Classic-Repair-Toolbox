using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Tests for WorklogSearchQuery - the "Find a previous repair" box's query language.
//
// The grammar came from the feature request as four worked examples, and each of them is pinned
// down here by name so a future change cannot quietly redefine one:
//   "full text"  a quoted run is one term, and still matches inside a larger word
//   cpu super    space = AND
//   p c u        single characters are just short terms, so this finds "CPU"
//   -cpu         a leading minus excludes
public class WorklogSearchQueryTests
{
    // -------------------------------------------------------------- Parse

    [Fact]
    public void A_blank_query_parses_to_nothing_and_matches_everything()
    {
        // An empty box is not a filter. Returning "matches nothing" here would blank the tab the
        // moment the user cleared the field, which reads as the data having been lost.
        foreach (var raw in new[] { null, "", "   ", "\t" })
        {
            var query = WorklogSearchQuery.Parse(raw);

            Assert.True(query.IsEmpty);
            Assert.Empty(query.Terms);
            Assert.True(query.Matches("anything at all"));
        }
    }

    [Fact]
    public void Spaces_split_a_query_into_separate_terms()
    {
        var query = WorklogSearchQuery.Parse("cpu super");

        Assert.Equal(new[] { "cpu", "super" }, query.Terms.Select(t => t.Text).ToArray());
        Assert.All(query.Terms, t => Assert.False(t.IsExcluded));
    }

    [Fact]
    public void A_quoted_run_is_one_term_including_its_spaces()
    {
        var query = WorklogSearchQuery.Parse("\"full text\"");

        var term = Assert.Single(query.Terms);
        Assert.Equal("full text", term.Text);
        Assert.False(term.IsExcluded);
    }

    [Fact]
    public void Quoted_and_bare_terms_mix_in_one_query()
    {
        var query = WorklogSearchQuery.Parse("\"black screen\" cpu -\"ruled out\" -pla");

        Assert.Equal(
            new[] { "black screen", "cpu", "ruled out", "pla" },
            query.Terms.Select(t => t.Text).ToArray());

        Assert.Equal(
            new[] { false, false, true, true },
            query.Terms.Select(t => t.IsExcluded).ToArray());
    }

    [Fact]
    public void A_leading_minus_marks_a_term_excluded_and_is_not_part_of_the_text()
    {
        var query = WorklogSearchQuery.Parse("-cpu");

        var term = Assert.Single(query.Terms);
        Assert.Equal("cpu", term.Text);
        Assert.True(term.IsExcluded);
    }

    // A minus INSIDE a term is just a character - part numbers and board labels are full of them
    // ("74LS-04"), so only a LEADING minus can mean exclusion.
    [Fact]
    public void A_minus_inside_a_term_is_an_ordinary_character()
    {
        var query = WorklogSearchQuery.Parse("74ls-04");

        var term = Assert.Single(query.Terms);
        Assert.Equal("74ls-04", term.Text);
        Assert.False(term.IsExcluded);
    }

    // Parsing runs on every keystroke, so half-typed input must degrade gracefully rather than
    // throw or stop filtering.
    [Fact]
    public void Malformed_input_parses_to_the_terms_it_can_make_sense_of()
    {
        // A lone minus carries no term.
        Assert.Empty(WorklogSearchQuery.Parse("-").Terms);
        Assert.Empty(WorklogSearchQuery.Parse("- ").Terms);

        // An empty quoted run carries no constraint - dropped rather than kept as a term that
        // would trivially match (or, negated, match nothing).
        Assert.Empty(WorklogSearchQuery.Parse("\"\"").Terms);

        // An unclosed quote runs to the end of the input, so typing a phrase filters on what has
        // been typed so far instead of on nothing.
        var unclosed = WorklogSearchQuery.Parse("\"black scr");
        Assert.Equal("black scr", Assert.Single(unclosed.Terms).Text);
    }

    // -------------------------------------------------------------- Matches

    [Fact]
    public void A_full_text_term_matches_inside_a_larger_string()
    {
        // Straight from the request: searching "full text" also finds "Afull textB".
        var query = WorklogSearchQuery.Parse("\"full text\"");

        Assert.True(query.Matches("Afull textB"));
        Assert.False(query.Matches("full-text"));
    }

    [Fact]
    public void Every_term_must_be_found_because_space_means_and()
    {
        var query = WorklogSearchQuery.Parse("cpu super");

        // Order does not matter - "Super Cpu" satisfies both terms, per the request.
        Assert.True(query.Matches("Super Cpu"));
        Assert.False(query.Matches("Super only"));
        Assert.False(query.Matches("Cpu only"));
    }

    [Fact]
    public void Single_character_terms_match_the_characters_anywhere_in_the_text()
    {
        // "p c u" finds "CPU": each character is present somewhere.
        var query = WorklogSearchQuery.Parse("p c u");

        Assert.True(query.Matches("CPU"));
        Assert.False(query.Matches("CP"));
    }

    [Fact]
    public void An_excluded_term_rules_a_record_out_even_when_the_others_match()
    {
        // "-cpu super" must NOT find "Super CPU", per the request.
        var query = WorklogSearchQuery.Parse("-cpu super");

        Assert.False(query.Matches("Super CPU"));
        Assert.True(query.Matches("Super PLA"));
    }

    [Fact]
    public void A_query_of_only_exclusions_matches_anything_not_containing_them()
    {
        var query = WorklogSearchQuery.Parse("-cpu");

        Assert.True(query.Matches("dead PLA"));
        Assert.False(query.Matches("dead CPU"));
    }

    [Fact]
    public void Matching_ignores_case_in_both_the_query_and_the_text()
    {
        Assert.True(WorklogSearchQuery.Parse("CPU").Matches("cpu"));
        Assert.True(WorklogSearchQuery.Parse("cpu").Matches("CPU"));
        Assert.True(WorklogSearchQuery.Parse("\"BLACK Screen\"").Matches("black screen"));
    }

    [Fact]
    public void Terms_are_anded_across_all_of_a_records_fields_not_within_one()
    {
        // The fields are one record's searchable text taken together, so a title carrying one term
        // and a comment carrying the other is a match - that is the whole point of passing them as
        // a set rather than searching each field separately.
        var query = WorklogSearchQuery.Parse("cpu super");

        Assert.True(query.Matches("CPU replaced", "super glue used"));
        Assert.False(query.Matches("CPU replaced", "nothing else"));
    }

    [Fact]
    public void Blank_and_null_fields_are_skipped_rather_than_matched_against()
    {
        var query = WorklogSearchQuery.Parse("cpu");

        Assert.False(query.Matches(null, "", "   "));
        Assert.True(query.Matches(null, "", "cpu"));
    }

    // -------------------------------------------------------------- FindHits

    [Fact]
    public void Hits_report_where_each_required_term_was_found()
    {
        var query = WorklogSearchQuery.Parse("this text");

        // The request's own highlighting example: "This is a text" highlights "This" and "text".
        var hits = query.FindHits("This is a text");

        Assert.Equal(2, hits.Count);
        Assert.Equal(0, hits[0].Start);
        Assert.Equal(4, hits[0].Length);
        Assert.Equal(10, hits[1].Start);
        Assert.Equal(4, hits[1].Length);
    }

    [Fact]
    public void Hits_index_the_original_text_so_the_original_casing_can_be_rendered()
    {
        var query = WorklogSearchQuery.Parse("cpu");

        var hit = Assert.Single(query.FindHits("The CPU is dead"));

        Assert.Equal(4, hit.Start);
        Assert.Equal("CPU", "The CPU is dead".Substring(hit.Start, hit.Length));
    }

    [Fact]
    public void Every_occurrence_of_a_term_is_reported_not_just_the_first()
    {
        var query = WorklogSearchQuery.Parse("cpu");

        var hits = query.FindHits("CPU and another cpu");

        Assert.Equal(2, hits.Count);
        Assert.Equal(0, hits[0].Start);
        Assert.Equal(16, hits[1].Start);
    }

    [Fact]
    public void Overlapping_hits_are_merged_into_one_run()
    {
        // Two terms covering overlapping text would otherwise produce two spans that double-draw
        // on top of each other when rendered.
        var query = WorklogSearchQuery.Parse("cpu pu");

        var hit = Assert.Single(query.FindHits("CPU"));

        Assert.Equal(0, hit.Start);
        Assert.Equal(3, hit.Length);
    }

    [Fact]
    public void Excluded_terms_are_never_highlighted()
    {
        // An excluded term cannot appear in a record that matched, so there is nothing of it to
        // highlight - and highlighting it in a NON-matching record would be actively misleading.
        var query = WorklogSearchQuery.Parse("dead -cpu");

        var hit = Assert.Single(query.FindHits("dead CPU"));

        Assert.Equal(0, hit.Start);
        Assert.Equal(4, hit.Length);
    }

    [Fact]
    public void An_empty_query_or_blank_text_yields_no_hits()
    {
        Assert.Empty(WorklogSearchQuery.Parse("").FindHits("anything"));
        Assert.Empty(WorklogSearchQuery.Parse("cpu").FindHits(""));
        Assert.Empty(WorklogSearchQuery.Parse("cpu").FindHits(null));
        Assert.Empty(WorklogSearchQuery.Parse("cpu").FindHits("no match here"));
    }

    // -------------------------------------------------------------- SplitIntoSegments

    [Fact]
    public void Segments_alternate_plain_and_matched_text_in_order()
    {
        // The request's own example: searching "this text" over "This is a text" marks two runs.
        var query = WorklogSearchQuery.Parse("this text");

        var segments = query.SplitIntoSegments("This is a text");

        Assert.Equal(
            new[] { ("This", true), (" is a ", false), ("text", true) },
            segments.ToArray());
    }

    // The single most important property: the segments must reassemble into exactly the input.
    // Any off-by-one in the split shows up on screen as a dropped or doubled character.
    [Theory]
    [InlineData("this text", "This is a text")]
    [InlineData("cpu", "CPU")]
    [InlineData("cpu", "the CPU and the cpu again")]
    [InlineData("cpu pu", "CPU")]
    [InlineData("a", "aaa")]
    [InlineData("dead -cpu", "dead CPU")]
    [InlineData("nomatch", "nothing here")]
    public void Segments_always_reassemble_into_the_original_text(string rawQuery, string text)
    {
        var segments = WorklogSearchQuery.Parse(rawQuery).SplitIntoSegments(text);

        Assert.Equal(text, string.Concat(segments.Select(s => s.Text)));
    }

    [Fact]
    public void Text_with_no_hits_is_one_unmatched_segment_rather_than_nothing()
    {
        // Callers render the result unconditionally, so "no hits" has to still carry the text.
        var segments = WorklogSearchQuery.Parse("cpu").SplitIntoSegments("dead PLA");

        var only = Assert.Single(segments);
        Assert.Equal("dead PLA", only.Text);
        Assert.False(only.IsMatch);
    }

    [Fact]
    public void An_empty_query_leaves_the_text_as_one_unmatched_segment()
    {
        var segments = WorklogSearchQuery.Parse("").SplitIntoSegments("dead PLA");

        var only = Assert.Single(segments);
        Assert.Equal("dead PLA", only.Text);
        Assert.False(only.IsMatch);
    }

    [Fact]
    public void Blank_text_yields_no_segments()
    {
        Assert.Empty(WorklogSearchQuery.Parse("cpu").SplitIntoSegments(""));
        Assert.Empty(WorklogSearchQuery.Parse("cpu").SplitIntoSegments(null));
    }

    [Fact]
    public void A_match_at_the_very_start_and_end_produces_no_empty_segments()
    {
        // A hit flush against either end must not emit a zero-length plain segment beside it.
        var segments = WorklogSearchQuery.Parse("cpu").SplitIntoSegments("CPU");

        var only = Assert.Single(segments);
        Assert.Equal("CPU", only.Text);
        Assert.True(only.IsMatch);
        Assert.All(segments, s => Assert.NotEqual(0, s.Text.Length));
    }
}
