using System.Reflection;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// The copies the "Schematic images" window writes with (2026-10-10): BoardSchematicEntry.WithCadName
// and the three WithSchematicName a rename uses - the "Board schematics" row, and the highlight and
// calibration filed under its name in the sidecar. A copy that forgot a property would erase it from
// the draft without a word, as a hand-written BoardData copy once dropped the "# Hardware:" caption.
// Kept together because they are the same promise; ComponentHighlightEntry and KiCadCalibrationEntry
// have no other copy to test.
public sealed class BoardSchematicEntryTests
{
    // ###########################################################################################
    // Every property is given its own value, so one the copy forgets - including one added after
    // this test was written - comes back as its default and fails here. Only the one property the
    // copy is MEANT to change differs.
    // ###########################################################################################
    private static void AssertCopiesEverythingBut<T>(Func<T, T> copy, string changed, object expected)
        where T : class, new()
    {
        var original = new T();
        PropertyInfo[] properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

        for (int i = 0; i < properties.Length; i++)
        {
            PropertyInfo property = properties[i];

            object value = property.PropertyType == typeof(string) ? "value of " + property.Name
                : property.PropertyType == typeof(double) ? i + 0.5
                : property.PropertyType == typeof(bool) ? true
                : throw new InvalidOperationException($"{typeof(T).Name}.{property.Name} is a {property.PropertyType.Name} - give this test a value for it.");

            property.SetValue(original, value);
        }

        T result = copy(original);

        Assert.NotSame(original, result);
        Assert.Equal(expected, typeof(T).GetProperty(changed)!.GetValue(result));

        foreach (PropertyInfo property in properties.Where(property => property.Name != changed))
        {
            Assert.Equal(property.GetValue(original), property.GetValue(result));
        }
    }

    [Fact]
    public void WithCadName_changes_the_CAD_name_and_keeps_every_other_column()
    {
        AssertCopiesEverythingBut<BoardSchematicEntry>(
            entry => entry.WithCadName("Open128 - PCB Top"),
            nameof(BoardSchematicEntry.CadName),
            "Open128 - PCB Top");
    }

    [Fact]
    public void A_renamed_schematic_row_keeps_every_other_column()
    {
        AssertCopiesEverythingBut<BoardSchematicEntry>(
            entry => entry.WithSchematicName("PCB; Top"),
            nameof(BoardSchematicEntry.SchematicName),
            "PCB; Top");
    }

    [Fact]
    public void A_highlight_on_a_renamed_schematic_keeps_its_rectangle_and_label()
    {
        AssertCopiesEverythingBut<ComponentHighlightEntry>(
            highlight => highlight.WithSchematicName("PCB; Top"),
            nameof(ComponentHighlightEntry.SchematicName),
            "PCB; Top");
    }

    [Fact]
    public void A_calibration_on_a_renamed_schematic_keeps_its_box_and_mirroring()
    {
        AssertCopiesEverythingBut<KiCadCalibrationEntry>(
            calibration => calibration.WithSchematicName("PCB; Top"),
            nameof(KiCadCalibrationEntry.SchematicName),
            "PCB; Top");
    }
}
