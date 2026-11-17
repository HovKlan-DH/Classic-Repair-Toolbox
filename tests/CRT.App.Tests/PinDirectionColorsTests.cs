using Handlers.DataHandling;
using Handlers.Theming;
using System.Linq;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The colors a chip's insides are painted in (PinDirectionColors; by direction since 2026-11-17): each direction has its own
// ink, tint, number and wire color - theme resources, whose contrast ThemeContrastTests holds in all three themes - and its
// word in the legend, the pin list's own.
// ###########################################################################################
public sealed class PinDirectionColorsTests
{
    private static readonly PinDirectionKind[] Kinds = [PinDirectionKind.Input, PinDirectionKind.Output, PinDirectionKind.InputOutput, PinDirectionKind.Other];

    [Fact]
    public void Each_direction_has_its_own_ink_tint_number_and_wire_and_its_word()
    {
        Assert.Equal(["PinoutChip_Input", "PinoutChip_Output", "PinoutChip_InputOutput", "PinoutChip_Other"], Kinds.Select(PinDirectionColors.InkKeyOf));
        Assert.Equal(["PinoutChip_Input_Bg", "PinoutChip_Output_Bg", "PinoutChip_InputOutput_Bg", "PinoutChip_Other_Bg"], Kinds.Select(PinDirectionColors.FillKeyOf));
        Assert.Equal(["PinoutChip_Input_Fg", "PinoutChip_Output_Fg", "PinoutChip_InputOutput_Fg", "PinoutChip_Other_Fg"], Kinds.Select(PinDirectionColors.NumberKeyOf));
        Assert.Equal(["PinoutChip_Input_Wire", "PinoutChip_Output_Wire", "PinoutChip_InputOutput_Wire", "PinoutChip_Other_Wire"], Kinds.Select(PinDirectionColors.WireKeyOf));
        Assert.Equal(["Input", "Output", "Input/output", "Other"], Kinds.Select(PinDirectionColors.WordOf));

        // The legend's words are the pin list's.
        Assert.All(Kinds.Select(PinDirectionColors.WordOf), word => Assert.Contains(word, ComponentVocabulary.PinDirections));
    }
}
