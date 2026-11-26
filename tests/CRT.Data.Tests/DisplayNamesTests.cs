using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DisplayNames - the name others see (owner requests, 2026-11-25: "it needs to be unique for all users", "allow only normal
// A-Z + 0-9 and then some common special charatcers like: +-/@.[]() ... Make the max length 30 characters"; 2026-11-26:
// spaces and "_" too). The ONE rule
// the server holds and CRT's account window says as a name is typed; uniqueness is the server's database's
// (AccountFlowTests and AccountSelfServiceFlowsTests hold a name taken in other capitals to it, through FakeAccountStore).
// ###########################################################################################
public sealed class DisplayNamesTests
{
    [Theory]
    [InlineData("Dennis")]
    [InlineData("Dennis@CRT")]
    [InlineData("C64-fixer")]
    [InlineData("a.b+c/d[e](f)")]
    [InlineData("AB")]
    [InlineData("  Padded  ")]
    [InlineData("123456789012345678901234567890")]

    // Spaces and "_" too (owner, 2026-11-26: "allow those two also").
    [InlineData("Dennis Hansen")]
    [InlineData("under_score")]
    [InlineData("C64 fixer_2 (DK)")]
    public void A_name_of_the_allowed_characters_and_length_is_taken(string name)
    {
        Assert.Null(DisplayNames.Problem(name));
    }

    // The point of the rule: nothing outside A-Z, 0-9, single spaces and the marks - so no name in lookalike letters of
    // another alphabet.
    [Theory]
    [InlineData("Dеnnis", "A display name cannot hold \"е\".")] // a Cyrillic "е" (U+0435) in place of the Latin "e"
    [InlineData("Søren", "A display name cannot hold \"ø\".")]
    [InlineData("Dennis#CRT", "A display name cannot hold \"#\".")]
    [InlineData("Fix😀er", "A display name cannot hold \"😀\".")] // said whole, never as half of a surrogate pair
    [InlineData("Tab\there", "A display name cannot hold \"U+0009\".")] // a character nobody can see is said by its code
    [InlineData("Zero​width", "A display name cannot hold \"U+200B\".")]
    public void A_character_outside_the_rule_is_named_with_what_is_allowed(string name, string expectedStart)
    {
        string? problem = DisplayNames.Problem(name);

        Assert.NotNull(problem);
        Assert.StartsWith(expectedStart, problem, StringComparison.Ordinal);
        Assert.EndsWith(DisplayNames.AllowedLine, problem, StringComparison.Ordinal);
    }

    // Where the rule stands beside the box already (CRT's account window, "My account"), the sentence leaves it out -
    // said twice, it read as a wall of text; a limit, with nothing after it, reads the same either way.
    // Two spaces in a row would let "Dennis  Hansen" pass for another "Dennis Hansen" - uniqueness ignores only the spaces
    // around a name.
    [Fact]
    public void Two_spaces_in_a_row_are_refused()
    {
        Assert.Equal("A display name cannot hold two spaces in a row.", DisplayNames.Problem("Dennis  Hansen"));
        Assert.Null(DisplayNames.Problem("  Dennis Hansen  "));
    }

    [Fact]
    public void Beside_a_box_showing_the_rule_the_sentence_leaves_the_rule_out()
    {
        Assert.Equal("A display name cannot hold \"#\".", DisplayNames.Problem("Dennis#CRT", sayRule: false));
        Assert.Equal(DisplayNames.Problem("D"), DisplayNames.Problem("D", sayRule: false));
        Assert.Null(DisplayNames.Problem("Dennis", sayRule: false));
    }

    [Fact]
    public void A_name_too_short_or_too_long_says_its_limit()
    {
        Assert.Equal("A display name needs at least 2 characters.", DisplayNames.Problem("D"));
        Assert.Equal(
            "A display name can have at most 30 characters - this one has 31.",
            DisplayNames.Problem(new string('x', 31)));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_blank_name_asks_for_one(string? name)
    {
        Assert.Equal(DisplayNames.Missing, DisplayNames.Problem(name));
    }

    // The words a person reads beside the box are held to the numbers the rule uses - a limit changed alone would leave
    // the sentence saying another.
    [Fact]
    public void The_sentence_beside_the_box_names_the_limits_and_every_mark()
    {
        Assert.StartsWith($"{DisplayNames.MinimumLength} to {DisplayNames.MaximumLength} characters", DisplayNames.AllowedLine, StringComparison.Ordinal);

        foreach (char mark in DisplayNames.AllowedMarks)
        {
            Assert.Contains($" {mark}", DisplayNames.AllowedLine, StringComparison.Ordinal);
        }

        Assert.Contains("single spaces", DisplayNames.AllowedLine, StringComparison.Ordinal);
    }
}
