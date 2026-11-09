using CRT;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The values listed under a field the components picked differ in while it has the focus (ComponentPaneDifferingValue; owner
// requests, 2026-11-09) - each once, with whose it is, narrowed to what is typed.
// ###########################################################################################
public sealed class ComponentPaneDifferingValueTests
{
    [Fact]
    public void Each_value_is_listed_once_with_every_component_having_it_in_the_order_first_seen()
    {
        IReadOnlyList<ComponentPaneDifferingValue> values = ComponentPaneDifferingValue.Of(
            ["R21", "R22", "R23", "R24"],
            ["Pull-up", "Pull-down", "Pull-up", ""],
            isMarkup: false);

        Assert.Equal(["Pull-up", "Pull-down", ""], values.Select(value => value.Value).ToArray());
        Assert.Equal(["R21, R23", "R22", "R24"], values.Select(value => value.Whose).ToArray());
    }

    // An empty value is a value - it is what picking it sets on all of them - and reads as such.
    [Fact]
    public void An_empty_value_reads_as_empty()
    {
        ComponentPaneDifferingValue empty = ComponentPaneDifferingValue.Of(["R21"], ["  "], isMarkup: false)[0];

        Assert.Equal("(empty)", empty.Shown);
    }

    // A note with formatting is listed as its words, on one line - the value itself kept as stored, to be set as it is.
    [Fact]
    public void A_note_with_formatting_reads_as_its_words_and_keeps_its_markup_as_the_value()
    {
        ComponentPaneDifferingValue note = ComponentPaneDifferingValue.Of(["C12"], ["**Decoupling** for U8\n- near pin 40"], isMarkup: true)[0];

        Assert.StartsWith("Decoupling for U8 ", note.Shown, StringComparison.Ordinal);
        Assert.EndsWith("near pin 40", note.Shown, StringComparison.Ordinal);
        Assert.DoesNotContain('\n', note.Shown);
        Assert.DoesNotContain("**", note.Shown, StringComparison.Ordinal);
        Assert.Equal("**Decoupling** for U8\n- near pin 40", note.Value);
    }

    // ###########################################################################################
    // WHAT IS TYPED NARROWS THE LIST (owner request, 2026-11-09: "where I can start typing also and then it will filter the list
    // based on that" - as "Board: Links"' "Category"): the values holding the text, anywhere and capitals aside; nothing typed
    // (or only spaces), all of them; "(empty)" is no text, so typing never finds it.
    // ###########################################################################################
    [Fact]
    public void What_is_typed_narrows_the_list_to_the_values_holding_it()
    {
        IReadOnlyList<ComponentPaneDifferingValue> values = ComponentPaneDifferingValue.Of(
            ["R21", "R22", "R23"],
            ["Pull-up", "Pull-down", ""],
            isMarkup: false);

        Assert.Same(values, ComponentPaneDifferingValue.Matching(values, ""));
        Assert.Same(values, ComponentPaneDifferingValue.Matching(values, "  "));
        Assert.Same(values, ComponentPaneDifferingValue.Matching(values, null));
        Assert.Equal(["Pull-up", "Pull-down"], ComponentPaneDifferingValue.Matching(values, "pull").Select(value => value.Value).ToArray());
        Assert.Equal(["Pull-down"], ComponentPaneDifferingValue.Matching(values, "DOWN ").Select(value => value.Value).ToArray());
        Assert.Empty(ComponentPaneDifferingValue.Matching(values, "empty"));
        Assert.Empty(ComponentPaneDifferingValue.Matching(values, "Resistor"));

        // A note is found by its words, not its marks.
        IReadOnlyList<ComponentPaneDifferingValue> notes = ComponentPaneDifferingValue.Of(["C1", "C2"], ["**Decoupling** U8", "Filter"], isMarkup: true);
        Assert.Equal(["**Decoupling** U8"], ComponentPaneDifferingValue.Matching(notes, "decoupling u8").Select(value => value.Value).ToArray());

        // ... all of them, past what the list shows of a long one (second code review, 2026-11-09: a word after the 160th
        // character found nothing).
        string longNote = new string('x', ComponentPaneDifferingValue.ShownAtMost + 20) + " tantalum";
        IReadOnlyList<ComponentPaneDifferingValue> longNotes = ComponentPaneDifferingValue.Of(["C1", "C2"], [longNote, "Filter"], isMarkup: true);
        Assert.DoesNotContain("tantalum", longNotes[0].Shown, StringComparison.Ordinal);
        Assert.Equal([longNote], ComponentPaneDifferingValue.Matching(longNotes, "tantalum").Select(value => value.Value).ToArray());
    }

    [Fact]
    public void A_long_value_is_cut_short_in_the_list()
    {
        string longer = new('x', ComponentPaneDifferingValue.ShownAtMost + 40);

        ComponentPaneDifferingValue value = ComponentPaneDifferingValue.Of(["U1"], [longer], isMarkup: false)[0];

        Assert.Equal(ComponentPaneDifferingValue.ShownAtMost + 3, value.Shown.Length);
        Assert.EndsWith("...", value.Shown, StringComparison.Ordinal);
        Assert.Equal(longer, value.Value);
    }
}
