using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// KiCadPinTable - a component's pins from a board's loaded KiCad data: the FIRST PCB file's pads
// (only it lights up copper - KiCadReferenceMatcher's rule) and every sheet's symbol pins, read
// end to end from KiCad files on disk.
// ###########################################################################################
public sealed class KiCadPinTableTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    [Fact]
    public async Task The_pins_come_from_the_pads_and_the_symbols_of_the_component()
    {
        string pcb = this.thisWorkspace.WriteFile("board.kicad_pcb", KiCadFixtures.PcbWithPinFunctions);
        string schematic = this.thisWorkspace.WriteFile("gates.kicad_sch", KiCadFixtures.SchematicWithLibrarySymbols);

        KiCadProjectRoot? root = await KiCadRawProjectLoader.LoadAsync([pcb, schematic]);

        Assert.Equal(
            [new KiCadPin("1", "", "", ""), new KiCadPin("40", "~RES", "Input", "~RESET")],
            KiCadPinTable.PinsOf(root, " u8 "));

        Assert.Equal(
            ["1:A:Input", "3:~Y:Output", "4:A:Input", "7:GND:Ground", "14:VCC:Power"],
            KiCadPinTable.PinsOf(root, "U5").Select(pin => $"{pin.Pin}:{pin.Name}:{pin.Direction}"));

        Assert.Empty(KiCadPinTable.PinsOf(root, "U99"));
        Assert.Empty(KiCadPinTable.PinsOf(null, "U8"));
    }

    // ###########################################################################################
    // An older file's unconnected pad is on KiCad's "no net", (net 0 "") - no net on this board, not a
    // net named "0" (code review, 2026-10-26: the pad's net never said so, and the pin showed "0").
    // ###########################################################################################
    [Fact]
    public async Task A_pad_on_KiCads_no_net_has_no_net()
    {
        const string Pcb = """
        (kicad_pcb (version 20221018) (generator pcbnew)
          (net 0 "")
          (net 1 "PHI2")
          (footprint "Package_DIP:DIP-40" (layer "F.Cu")
            (at 10 10)
            (fp_text reference "U8" (at 0 -5) (layer "F.SilkS"))
            (pad "1" thru_hole rect (at 0 0) (size 1.6 1.6) (layers "*.Cu") (net 0 ""))
            (pad "2" thru_hole oval (at 0 2.54) (size 1.6 1.6) (layers "*.Cu") (net 1 "PHI2"))
          )
        )
        """;

        KiCadProjectRoot? root = await KiCadRawProjectLoader.LoadAsync([this.thisWorkspace.WriteFile("old.kicad_pcb", Pcb)]);

        Assert.Equal([("1", ""), ("2", "PHI2")], KiCadPinTable.PinsOf(root, "U8").Select(pin => (pin.Pin, pin.Net)));
    }

    // A second PCB file lights up nothing, so its pads are no pins - as KiCadReferenceMatcher says.
    [Fact]
    public void Only_the_first_PCB_files_pads_count()
    {
        static KiCadPcb Pcb(string net) => new()
        {
            Footprints = [new KiCadPcbFootprint { Reference = "U8", Pads = [new KiCadPcbPad { Number = "1", Net = new KiCadNetRef { Name = net } }] }],
        };

        var root = new KiCadProjectRoot { Pcb = [Pcb("PHI2"), Pcb("OTHER")] };

        Assert.Equal("PHI2", Assert.Single(KiCadPinTable.PinsOf(root, "U8")).Net);
    }
}
