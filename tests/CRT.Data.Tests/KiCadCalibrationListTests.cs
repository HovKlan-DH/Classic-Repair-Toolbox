using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// KiCadCalibrationList - a board's KiCad calibrations with one saved: one per schematic image, by
// its name (trimmed, capitals ignored), in name order. The board file's writer and the Schematics
// tab's calibration through the draft held in memory (2026-10-25) both use it.
// ###########################################################################################
public sealed class KiCadCalibrationListTests
{
    private static KiCadCalibrationEntry Calibration(string schematic, double scaleX = 1) =>
        new() { SchematicName = schematic, CadName = "PCB Top", ScaleX = scaleX, ScaleY = 1 };

    [Fact]
    public void A_calibration_saved_for_a_schematic_replaces_the_one_it_had_whatever_its_capitals()
    {
        List<KiCadCalibrationEntry> saved = KiCadCalibrationList.With(
            [Calibration("Top", scaleX: 1), Calibration("Bottom")],
            Calibration(" top ", scaleX: 2));

        Assert.Equal(["Bottom", "top"], saved.Select(entry => entry.SchematicName));
        Assert.Equal(2, saved.Single(entry => entry.SchematicName == "top").ScaleX);
    }

    // Every field is carried, trimmed where it is a name - and the list is a new one, the old left as it was.
    [Fact]
    public void A_calibration_for_a_schematic_without_one_is_added_in_name_order_with_every_field()
    {
        List<KiCadCalibrationEntry> before = [Calibration("Top")];

        List<KiCadCalibrationEntry> saved = KiCadCalibrationList.With(before, new KiCadCalibrationEntry
        {
            SchematicName = "Bottom",
            CadName = " PCB Bottom ",
            OffsetX = 10,
            OffsetY = 20,
            ScaleX = 0.5,
            ScaleY = 0.25,
            MirrorX = true,
            MirrorY = true,
        });

        Assert.Equal(["Bottom", "Top"], saved.Select(entry => entry.SchematicName));

        KiCadCalibrationEntry bottom = saved[0];
        Assert.Equal("PCB Bottom", bottom.CadName);
        Assert.Equal((10.0, 20.0, 0.5, 0.25, true, true), (bottom.OffsetX, bottom.OffsetY, bottom.ScaleX, bottom.ScaleY, bottom.MirrorX, bottom.MirrorY));

        Assert.Single(before);
    }

    [Fact]
    public void A_calibration_with_no_schematic_name_is_refused()
    {
        Assert.Throws<InvalidOperationException>(() => KiCadCalibrationList.With([], Calibration("  ")));
    }
}
