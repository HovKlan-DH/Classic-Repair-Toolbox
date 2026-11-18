using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What color a pin's direction is drawn in (PinDirectionKinds; owner request, 2026-11-17: "Not connected + Power + Other +
// Unknown can all have the same lightgrey'ish color ... and then probably the Input/output one needs a new color"): every word
// of the pin list's vocabulary, read without regard to capitals or the space round it, and a word it does not have.
// ###########################################################################################
public sealed class PinDirectionKindsTests
{
    [Theory]
    [InlineData("Input", PinDirectionKind.Input)]
    [InlineData(" input ", PinDirectionKind.Input)]
    [InlineData("Output", PinDirectionKind.Output)]
    [InlineData("OUTPUT", PinDirectionKind.Output)]
    [InlineData("Input/output", PinDirectionKind.InputOutput)]
    [InlineData("Input/Output", PinDirectionKind.InputOutput)]
    [InlineData("Not connected", PinDirectionKind.Other)]
    [InlineData("Ground", PinDirectionKind.Other)]
    [InlineData("Power", PinDirectionKind.Other)]
    [InlineData("Other", PinDirectionKind.Other)]
    [InlineData("Unknown", PinDirectionKind.Other)]
    [InlineData("Analog input", PinDirectionKind.Other)]
    [InlineData("", PinDirectionKind.Other)]
    [InlineData(null, PinDirectionKind.Other)]
    public void A_direction_is_colored_as_an_input_an_output_both_or_neither(string? direction, PinDirectionKind kind) =>
        Assert.Equal(kind, PinDirectionKinds.Of(direction));

    // Every word the pin list offers has a color - those that are no signal's way the quiet one.
    [Fact]
    public void Every_word_of_the_vocabulary_is_one_of_the_four()
    {
        Assert.Equal(
            [PinDirectionKind.Other, PinDirectionKind.Input, PinDirectionKind.Output, PinDirectionKind.InputOutput, PinDirectionKind.Other, PinDirectionKind.Other, PinDirectionKind.Other, PinDirectionKind.Other],
            ComponentVocabulary.PinDirections.Select(PinDirectionKinds.Of));
    }

    private static ComponentPinEntry Pin(int number, string direction) =>
        new() { BoardLabel = "U1", Pin = number.ToString(System.Globalization.CultureInfo.InvariantCulture), Direction = direction };

    // ###########################################################################################
    // EACH PIN IN THE COLOR OF THE DIRECTION ITS PIN LIST GIVES IT (owner request, 2026-11-17: "color it for its pin direction,
    // as seen in the pinout" - and, the same day, the data converted rather than the drawing reading past the list: "so it also
    // will show correctly in the text part of it"): "Unknown", "Power" and "Other" grey whatever the datasheet says; the
    // datasheet's only for a pin the list gives no direction or has no row for - as the drawing names such a pin by the
    // datasheet too.
    // ###########################################################################################
    [Fact]
    public void Each_pin_is_the_direction_its_list_gives_it_and_the_datasheets_only_where_the_list_gives_none()
    {
        IReadOnlyDictionary<int, PinDirectionKind> kinds = PinDirectionKinds.ByPin(
            14,
            [Pin(1, "Input"), Pin(2, "Other"), Pin(3, " unknown "), Pin(4, "Input/output"), Pin(5, ""), Pin(6, " output "), Pin(14, "Power")],
            LogicChips.Named("74LS08"));

        Assert.Equal(PinDirectionKind.Input, kinds[1]);
        Assert.Equal(PinDirectionKind.Other, kinds[2]);
        Assert.Equal(PinDirectionKind.Other, kinds[3]);
        Assert.Equal(PinDirectionKind.InputOutput, kinds[4]);
        Assert.Equal(PinDirectionKind.Input, kinds[5]);
        Assert.Equal(PinDirectionKind.Output, kinds[6]);
        Assert.Equal(PinDirectionKind.Other, kinds[7]);
        Assert.Equal(PinDirectionKind.Output, kinds[8]);
        Assert.Equal(PinDirectionKind.Other, kinds[14]);
        Assert.Equal(14, kinds.Count);
    }

    // A chip CRT knows nothing inside of - a CPU in its empty case - is what its list says, every other pin grey.
    [Fact]
    public void A_chip_with_no_datasheet_is_what_its_list_says_and_grey_elsewhere()
    {
        IReadOnlyDictionary<int, PinDirectionKind> kinds = PinDirectionKinds.ByPin(40, [Pin(2, "Input"), Pin(7, "Output"), Pin(30, "Input/output")], chip: null);

        Assert.Equal(40, kinds.Count);
        Assert.Equal(PinDirectionKind.Input, kinds[2]);
        Assert.Equal(PinDirectionKind.Output, kinds[7]);
        Assert.Equal(PinDirectionKind.InputOutput, kinds[30]);
        Assert.Equal(37, kinds.Values.Count(kind => kind == PinDirectionKind.Other));
    }

    // ###########################################################################################
    // THE ONE RULE for a pin's direction, which the pinout's drawing and the "Schematics" tab's numbered pins both ask (code
    // review, 2026-11-18: written twice): the list's word, else the datasheet's for a pin 1 to the chip's last - its pin 0 is
    // the wires inside a unit - else none.
    // ###########################################################################################
    [Fact]
    public void A_pins_word_is_its_lists_else_the_datasheets_for_a_pin_of_the_chip()
    {
        LogicChip chip = LogicChips.Named("74LS08")!;

        Assert.Equal("Input/output", PinDirectionKinds.WordOf(" Input/output ", 3, chip));
        Assert.Equal("Output", PinDirectionKinds.WordOf(null, 3, chip));
        Assert.Equal("Output", PinDirectionKinds.WordOf("  ", 3, chip));
        Assert.Equal(string.Empty, PinDirectionKinds.WordOf(null, 0, chip));
        Assert.Equal(string.Empty, PinDirectionKinds.WordOf(null, 15, chip));
        Assert.Equal(string.Empty, PinDirectionKinds.WordOf(null, null, chip));
        Assert.Equal(string.Empty, PinDirectionKinds.WordOf(null, 3, chip: null));
    }
}
