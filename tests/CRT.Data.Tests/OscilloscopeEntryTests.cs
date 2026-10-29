using System.Reflection;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// OscilloscopeEntry - a model of the main file as the Oscilloscope tab runs it. Its values come from
// the ONE map of a model's values (OscilloscopeModelValues; code review, 2026-10-29: they were mapped
// four times over, and a value added to the row was sent and published while the tab, Try it and
// the editor never saw it). These fail on a value of the row the entry does not carry.
// ###########################################################################################
public sealed class OscilloscopeEntryTests
{
    // A model whose every value is told apart from every other: "<name> value".
    private static MainDocument.OscilloscopeRow EveryValueItsOwn()
    {
        var row = new MainDocument.OscilloscopeRow { Brand = "Rigol", SeriesOrModel = "DS1054Z" };

        foreach (string name in OscilloscopeModelValues.Names)
            OscilloscopeModelValues.Set(row, name, $"{name} value");

        return row;
    }

    [Fact]
    public void Every_value_of_the_row_reaches_the_entry_and_is_read_back_by_its_name()
    {
        MainDocument.OscilloscopeRow row = EveryValueItsOwn();

        OscilloscopeEntry entry = OscilloscopeEntry.From(row, isOwn: true);

        Assert.Equal("Rigol", entry.Brand);
        Assert.Equal("DS1054Z", entry.SeriesOrModel);
        Assert.True(entry.IsOwn);
        Assert.Same(row, entry.Row);

        foreach (string name in OscilloscopeModelValues.Names)
            Assert.Equal($"{name} value", entry.ValueOf(name));
    }

    // Every text value of the row has the entry's property of its name - else From would leave it out.
    [Fact]
    public void Every_text_value_of_the_row_has_a_property_of_its_name_on_the_entry()
    {
        IEnumerable<string> missing = typeof(MainDocument.OscilloscopeRow)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(property => property.PropertyType == typeof(string))
            .Where(property => typeof(OscilloscopeEntry).GetProperty(property.Name)?.PropertyType != typeof(string))
            .Select(property => property.Name);

        Assert.Empty(missing);
    }

    [Fact]
    public void A_name_no_value_goes_by_is_refused()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OscilloscopeEntry().ValueOf("launchMissiles"));
        Assert.Throws<ArgumentOutOfRangeException>(() => OscilloscopeModelValues.Get(new MainDocument.OscilloscopeRow(), "brand"));
    }

    // Get and Set are the values Of and ToModel carry - one map, both ways.
    [Fact]
    public void Get_and_Set_are_the_values_a_model_is_sent_with()
    {
        MainDocument.OscilloscopeRow row = EveryValueItsOwn();

        IReadOnlyDictionary<string, string> sent = OscilloscopeModelValues.Of(row);

        foreach (string name in OscilloscopeModelValues.Names)
            Assert.Equal(sent[name], OscilloscopeModelValues.Get(row, name));

        OscilloscopeModelValues.Set(row, "dumpImage", null);
        Assert.Equal(string.Empty, row.DumpImage);
    }
}
