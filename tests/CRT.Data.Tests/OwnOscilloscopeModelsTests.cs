using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// The scope models a user added to their own copy of the main Excel file - the pre-4.0 way to add
// a scope CRT did not know - must survive the move to the 4.0 main file. None may be lost; none
// may be repeated when it is just the published model.
public sealed class OwnOscilloscopeModelsTests
{
    private static MainDocument.OscilloscopeRow Model(string brand, string series, string dumpImage = ":DISPLAY:DATA?") => new()
    {
        Brand = brand,
        SeriesOrModel = series,
        Port = "5555",
        DumpImage = dumpImage,
    };

    private static readonly IReadOnlyList<MainDocument.OscilloscopeRow> Published =
        [Model("Rigol", "DHO800"), Model("Keysight", "InfiniiVision 2000 X")];

    [Fact]
    public void A_model_the_published_list_lacks_is_the_user_s_own()
    {
        IReadOnlyList<MainDocument.OscilloscopeRow> own = OwnOscilloscopeModels.Find(
            [Model("Rigol", "DHO800"), Model("Voltcraft", "DOV702")],
            Published);

        Assert.Equal("DOV702", Assert.Single(own).SeriesOrModel);
    }

    // The user may have changed a command to suit their firmware: kept as their own version.
    [Fact]
    public void A_published_model_with_a_value_changed_is_kept_as_the_user_s_version()
    {
        IReadOnlyList<MainDocument.OscilloscopeRow> own = OwnOscilloscopeModels.Find(
            [Model("Rigol", "DHO800", ":DISPLAY:DATA? PNG")],
            Published);

        MainDocument.OscilloscopeRow mine = Assert.Single(own);
        Assert.Equal("DHO800 (my version)", mine.SeriesOrModel);
        Assert.Equal(":DISPLAY:DATA? PNG", mine.DumpImage);
        Assert.Equal("Rigol", mine.Brand);
    }

    [Fact]
    public void A_model_exactly_as_published_is_not_repeated()
    {
        Assert.Empty(OwnOscilloscopeModels.Find([Model("rigol", " DHO800 "), Model("Keysight", "InfiniiVision 2000 X")], Published));
    }

    // The rows of a real main workbook, read as the conversion reads them, brand carried down.
    [Fact]
    public void The_models_are_read_from_a_main_workbook_as_the_conversion_reads_them()
    {
        using DataTreeConversionTests.ExcelTree tree = DataTreeConversionTests.ExcelTree.Standard();
        var problems = new List<string>();

        IReadOnlyList<MainDocument.OscilloscopeRow> models = DataTreeConversion.ReadOscilloscopes(tree.Full(DataTreeConversionTests.ExcelTree.MasterName), problems);

        Assert.Empty(problems);
        Assert.Equal(["DHO800", "DHO900"], models.Select(model => model.SeriesOrModel));
        Assert.Equal(["Rigol", "Rigol"], models.Select(model => model.Brand));
    }

    [Fact]
    public void A_main_workbook_that_is_not_there_is_a_problem_and_no_models()
    {
        var problems = new List<string>();

        Assert.Empty(DataTreeConversion.ReadOscilloscopes(Path.Combine(Path.GetTempPath(), "crt-none-" + Guid.NewGuid().ToString("N") + ".xlsx"), problems));
        Assert.Single(problems);
    }
}
