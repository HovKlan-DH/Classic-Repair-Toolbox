using System.Text.Json;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// A model of the user's own, made and edited on the Oscilloscope tab (CRT 4.0 step 5): no two models
// in the list may share a name, a copy takes every command of the model it starts from, and a
// model published later under an own model's name takes it over or names it apart - nothing the
// user made is lost, nothing is listed twice.
public sealed class OscilloscopeModelRulesTests
{
    private static MainDocument.OscilloscopeRow Model(string brand, string series, string dumpImage = ":DISP:DATA?") => new()
    {
        Brand = brand,
        SeriesOrModel = series,
        Port = "5555",
        Identify = "*IDN?",
        SetTimeDiv = ":TIM:SCAL {0}",
        DumpImage = dumpImage,
        TimeDivList = "1ms,2ms",
    };

    [Theory]
    [InlineData("", "DS1054Z", OscilloscopeModelRules.NoVendor)]
    [InlineData("  ", "DS1054Z", OscilloscopeModelRules.NoVendor)]
    [InlineData("Rigol", "", OscilloscopeModelRules.NoModel)]
    [InlineData(null, "DS1054Z", OscilloscopeModelRules.NoVendor)]
    public void A_model_needs_a_vendor_and_a_name(string? vendor, string model, string problem)
    {
        Assert.Equal(problem, OscilloscopeModelRules.NameProblem(vendor, model, []));
    }

    // The drop-downs list a name once and pick a model by it: a second model of the same name,
    // capitals and spaces aside, would never be reachable.
    [Theory]
    [InlineData("Rigol", "DS1000Z")]
    [InlineData("rigol", " ds1000z ")]
    public void A_name_another_model_has_is_refused(string vendor, string model)
    {
        string? problem = OscilloscopeModelRules.NameProblem(vendor, model, [Model("Rigol", "DS1000Z")]);

        Assert.Equal(OscilloscopeModelRules.Taken(vendor, model), problem);
        Assert.Contains("already", problem);
    }

    [Fact]
    public void A_name_no_other_model_has_is_fine()
    {
        Assert.Null(OscilloscopeModelRules.NameProblem("Rigol", "DS1054Z", [Model("Rigol", "DS1000Z"), Model("Siglent", "DS1054Z")]));
    }

    [Fact]
    public void A_copy_takes_every_value_of_the_model_it_starts_from_under_its_own_name()
    {
        MainDocument.OscilloscopeRow source = Model("Rigol", "DS1000Z");
        source.Unknown = new Dictionary<string, JsonElement> { ["laterField"] = JsonDocument.Parse("1").RootElement };

        MainDocument.OscilloscopeRow copy = OscilloscopeModelRules.CopyOf(source, " Rigol ", " DS1054Z ");

        Assert.Equal("Rigol", copy.Brand);
        Assert.Equal("DS1054Z", copy.SeriesOrModel);
        Assert.Equal("5555", copy.Port);
        Assert.Equal(":TIM:SCAL {0}", copy.SetTimeDiv);
        Assert.Equal("1ms,2ms", copy.TimeDivList);
        Assert.Equal(":DISP:DATA?", copy.DumpImage);

        // A field a later CRT wrote comes along, in a list of the copy's own.
        Assert.True(copy.Unknown!.ContainsKey("laterField"));
        Assert.NotSame(source.Unknown, copy.Unknown);
        Assert.Equal("DS1000Z", source.SeriesOrModel);
    }

    [Fact]
    public void A_change_is_made_to_a_copy_never_to_the_model_held()
    {
        MainDocument.OscilloscopeRow held = Model("Rigol", "DS1054Z");

        MainDocument.OscilloscopeRow changed = OscilloscopeModelRules.With(held, model => model.Stop = ":STOP");

        Assert.Equal(":STOP", changed.Stop);
        Assert.Equal(string.Empty, held.Stop);
    }

    // ------------------------------------------------------------------ the launch's reconcile

    [Fact]
    public void Own_models_no_published_model_names_are_kept_as_they_are()
    {
        IReadOnlyList<MainDocument.OscilloscopeRow> own = [Model("Voltcraft", "DOV702")];

        (IReadOnlyList<MainDocument.OscilloscopeRow> kept, bool changed) = OscilloscopeModelRules.Reconciled(own, [Model("Rigol", "DS1000Z")]);

        Assert.False(changed);
        Assert.Same(own[0], Assert.Single(kept));
    }

    // The user's model sent in and published as it was: the own copy is no longer needed.
    [Fact]
    public void An_own_model_published_exactly_as_it_is_goes()
    {
        (IReadOnlyList<MainDocument.OscilloscopeRow> kept, bool changed) = OscilloscopeModelRules.Reconciled(
            [Model("voltcraft", " DOV702 "), Model("Owon", "ADS800")],
            [Model("Voltcraft", "DOV702")]);

        Assert.True(changed);
        Assert.Equal("ADS800", Assert.Single(kept).SeriesOrModel);
    }

    // Published with a command changed (the administrator fixed one, say): the user's stays beside
    // it, named apart, so neither hides the other.
    [Fact]
    public void An_own_model_published_differently_is_kept_as_the_user_s_version()
    {
        (IReadOnlyList<MainDocument.OscilloscopeRow> kept, bool changed) = OscilloscopeModelRules.Reconciled(
            [Model("Voltcraft", "DOV702", dumpImage: ":DISP:DATA? BMP")],
            [Model("Voltcraft", "DOV702")]);

        Assert.True(changed);
        MainDocument.OscilloscopeRow mine = Assert.Single(kept);
        Assert.Equal("DOV702" + OwnOscilloscopeModels.MyVersionSuffix, mine.SeriesOrModel);
        Assert.Equal(":DISP:DATA? BMP", mine.DumpImage);
    }

    // Each rename is told - the name it had and the name it has - so the Oscilloscope tab follows
    // that very model, never one guessed from its name's start (code review, 2026-10-29).
    [Fact]
    public void Every_own_model_named_apart_is_told_with_its_old_and_new_name()
    {
        (IReadOnlyList<MainDocument.OscilloscopeRow> kept, _) = OscilloscopeModelRules.Reconciled(
            [Model("Voltcraft", "DOV702" + OwnOscilloscopeModels.MyVersionSuffix), Model("Voltcraft", "DOV702", dumpImage: ":DISP:DATA? BMP"), Model("Owon", "ADS800")],
            [Model("Voltcraft", "DOV702")],
            out IReadOnlyList<(string Vendor, string From, string To)> renamed);

        Assert.Equal([("Voltcraft", "DOV702", "DOV702 (my version) 2")], renamed);
        Assert.Equal(["DOV702 (my version)", "DOV702 (my version) 2", "ADS800"], kept.Select(model => model.SeriesOrModel));
    }

    // ###########################################################################################
    // Two names are one model only trimmed and with capitals aside - an accent counts, and "ss" is
    // not "ß" (code review, 2026-10-29: the server's database matched by its collation, which ignores
    // both, and took two models the main file keeps apart for one). The store asks this rule.
    // ###########################################################################################
    [Theory]
    [InlineData("Rigol", "DS1054Z", " rigol ", "ds1054z ", true)]
    [InlineData("Picó", "2204A", "Pico", "2204A", false)]
    [InlineData("Groß", "X1", "Gross", "X1", false)]
    [InlineData("Rigol", "DS1054Z", "Rigol", "DS1054Z+", false)]
    public void Two_names_are_one_model_only_with_capitals_and_spaces_aside(string vendor, string model, string otherVendor, string otherModel, bool same)
    {
        Assert.Equal(same, OscilloscopeModelRules.SameName(vendor, model, otherVendor, otherModel));
    }

    [Fact]
    public void A_my_version_name_that_is_taken_too_gets_a_number()
    {
        (IReadOnlyList<MainDocument.OscilloscopeRow> kept, _) = OscilloscopeModelRules.Reconciled(
            [Model("Voltcraft", "DOV702", dumpImage: "A")],
            [Model("Voltcraft", "DOV702"), Model("Voltcraft", "DOV702 (my version)")]);

        Assert.Equal("DOV702 (my version) 2", Assert.Single(kept).SeriesOrModel);
    }

    // A hand-edited file can hold a name twice; the second was unreachable in the drop-downs.
    [Fact]
    public void An_own_model_named_like_one_above_it_is_named_apart()
    {
        (IReadOnlyList<MainDocument.OscilloscopeRow> kept, bool changed) = OscilloscopeModelRules.Reconciled(
            [Model("Voltcraft", "DOV702", dumpImage: "A"), Model("Voltcraft", "DOV702", dumpImage: "B")],
            []);

        Assert.True(changed);
        Assert.Equal(["DOV702", "DOV702 (my version)"], kept.Select(model => model.SeriesOrModel));
        Assert.Equal("B", kept[1].DumpImage);
    }
}
