using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Tests for NaturalLabelComparer - the order component labels are shown in wherever a person has to
// scan a list of them (the board highlight JSON, and the KiCad match report's badges).
//
// Lifted out of BoardComponentHighlightStorage unchanged, so these pin down what it already did;
// BoardComponentHighlightStorageTests still covers the JSON ordering end to end.
public sealed class NaturalLabelComparerTests
{
    private static string[] Sorted(params string[] labels) =>
        labels.OrderBy(label => label, NaturalLabelComparer.Instance).ToArray();

    // The whole point: plain string ordering puts C10 and C100 before C2.
    [Fact]
    public void Numbers_inside_a_label_sort_by_VALUE_not_by_character()
    {
        Assert.Equal(
            new[] { "C1", "C2", "C3", "C10", "C11", "C100", "C105" },
            Sorted("C100", "C10", "C3", "C105", "C1", "C11", "C2"));
    }

    // A board mixes prefixes; within each prefix the numbers must still count up.
    [Fact]
    public void Labels_group_by_their_letters_and_count_up_inside_each_group()
    {
        Assert.Equal(
            new[] { "C2", "C10", "R1", "R20", "U8", "U17" },
            Sorted("U17", "R20", "C10", "U8", "C2", "R1"));
    }

    [Fact]
    public void The_letters_compare_case_insensitively()
    {
        Assert.Equal(0, NaturalLabelComparer.Compare("u8", "U8"));
        Assert.True(NaturalLabelComparer.Compare("c10", "U2") < 0);
    }

    // C02 and C2 are the same number. The shorter one goes first, so the order is still repeatable
    // rather than depending on which one happened to arrive first.
    [Fact]
    public void Leading_zeros_do_not_change_a_numbers_value_but_break_a_tie()
    {
        Assert.True(NaturalLabelComparer.Compare("C02", "C10") < 0);
        Assert.True(NaturalLabelComparer.Compare("C2", "C02") < 0);
    }

    // A label with more than one number in it (a pin, a sub-part) compares each number in turn.
    [Fact]
    public void Every_number_in_a_label_is_compared_as_a_number()
    {
        Assert.Equal(
            new[] { "U1.2", "U1.10", "U2.1" },
            Sorted("U2.1", "U1.10", "U1.2"));
    }

    [Fact]
    public void A_label_that_is_a_prefix_of_another_sorts_first()
    {
        Assert.Equal(new[] { "U", "U1", "UA" }, Sorted("UA", "U1", "U"));
    }

    [Fact]
    public void Surrounding_whitespace_and_nulls_are_harmless()
    {
        Assert.Equal(0, NaturalLabelComparer.Compare(" C2 ", "C2"));
        Assert.Equal(0, NaturalLabelComparer.Compare(null, string.Empty));
        Assert.True(NaturalLabelComparer.Compare(null, "C1") < 0);
    }
}
