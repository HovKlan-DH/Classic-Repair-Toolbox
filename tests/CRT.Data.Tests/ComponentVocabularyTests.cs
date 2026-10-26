using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentVocabulary - the words CRT 4.0's new component data is written in (step 2,
// 2026-10-25). The words ARE the data (a board file stores "Datasheet", not a code), so what a
// check calls known is pinned here, with the rules every value of a board follows: empty is fine,
// capitals and surrounding spaces do not matter.
// ###########################################################################################
public sealed class ComponentVocabularyTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("Datasheet")]
    [InlineData("  service MANUAL ")]
    public void Empty_or_a_known_word_in_any_capitals_is_known(string? value)
    {
        Assert.True(ComponentVocabulary.IsKnownOrEmpty(ComponentVocabulary.BaselineSources, value));
    }

    [Theory]
    [InlineData("My cousin")]
    [InlineData("Datasheets")]
    public void Anything_else_is_not(string value)
    {
        Assert.False(ComponentVocabulary.IsKnownOrEmpty(ComponentVocabulary.BaselineSources, value));
    }

    // The plan's words (2026-10-24), each list in the order the component pane offers it - with
    // "Other" last wherever it is offered.
    [Fact]
    public void The_lists_hold_the_plans_words()
    {
        Assert.Equal(["Known-good board", "Datasheet", "Service manual", "Other"], ComponentVocabulary.BaselineSources);
        Assert.Equal(["Drop-in", "Needs changes"], ComponentVocabulary.ReplacementKinds);
        Assert.Equal("Other", ComponentVocabulary.PinDirections[^1]);
        Assert.Equal("Other", ComponentVocabulary.Packages[^1]);
        Assert.Contains("DIP-40", ComponentVocabulary.Packages);
        Assert.Contains("TO-220", ComponentVocabulary.Packages);
    }

    // Plain ASCII and no word twice - they are shown in drop-down lists and written into files.
    [Fact]
    public void Every_word_is_plain_ASCII_and_listed_once()
    {
        foreach (IReadOnlyList<string> words in new[]
                 {
                     ComponentVocabulary.PinDirections,
                     ComponentVocabulary.BaselineSources,
                     ComponentVocabulary.ReplacementKinds,
                     ComponentVocabulary.Packages,
                 })
        {
            Assert.All(words, word => Assert.True(word.All(character => character is >= ' ' and <= '~'), word));
            Assert.Equal(words.Count, words.Distinct(StringComparer.OrdinalIgnoreCase).Count());
        }
    }

    // ###########################################################################################
    // A case's pin count, by its name - what the pane numbers the pins and draws the case from. A
    // DIP's is in its name, however it is typed; TO-92 and TO-220 have three; a SIP's, "Other"'s
    // and anything else's cannot be told.
    // ###########################################################################################
    [Theory]
    [InlineData("DIP-40", 40)]
    [InlineData("DIP-24 wide", 24)]
    [InlineData(" dip 28 ", 28)]
    [InlineData("DIP16", 16)]
    [InlineData("TO-92", 3)]
    [InlineData("to-220", 3)]
    [InlineData("SIP", null)]
    [InlineData("Other", null)]
    [InlineData("DIP", null)]
    [InlineData("DIP-7", null)]
    [InlineData("DIP-0", null)]
    [InlineData("PLCC-44", null)]
    [InlineData("", null)]
    [InlineData(null, null)]
    public void A_cases_pin_count_is_read_off_its_name(string? package, int? count)
    {
        Assert.Equal(count, ComponentVocabulary.PinCountOf(package));
    }

    // Every DIP the pane offers can be counted and drawn.
    [Fact]
    public void Every_offered_DIP_is_a_dual_in_line_case_with_a_count()
    {
        foreach (string package in ComponentVocabulary.Packages.Where(word => word.StartsWith("DIP", StringComparison.Ordinal)))
        {
            Assert.True(ComponentVocabulary.IsDualInLine(package), package);
            Assert.NotNull(ComponentVocabulary.PinCountOf(package));
        }

        Assert.False(ComponentVocabulary.IsDualInLine("TO-92"));
        Assert.False(ComponentVocabulary.IsDualInLine("DIP"));
    }
}
