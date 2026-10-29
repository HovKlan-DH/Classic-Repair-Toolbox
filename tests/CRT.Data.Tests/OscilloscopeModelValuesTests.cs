using System.Text;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// How an oscilloscope model sent in travels and is checked (CRT 4.0 step 5): its values by the main
// file's own names (OscilloscopeModelValues) and its fingerprint, the checks CRT and the server share
// (OscilloscopeModelRules.SubmissionProblems), and its write into a main file (MasterListing
// .PutOscilloscope) - in place of the model of its name, never a second one.
public sealed class OscilloscopeModelValuesTests
{
    private static MainDocument.OscilloscopeRow Model(string vendor = "Voltcraft", string model = "DOV702", string dumpImage = ":DISP:DATA?") => new()
    {
        Brand = vendor,
        SeriesOrModel = model,
        Port = "3000",
        Identify = "*IDN?",
        DumpImage = dumpImage,
        VoltsDivList = "1V,2V",
    };

    // ------------------------------------------------------------------ the values

    [Fact]
    public void A_model_s_values_go_by_the_main_file_s_names_and_come_back_the_same()
    {
        IReadOnlyDictionary<string, string> values = OscilloscopeModelValues.Of(Model());

        Assert.Equal("3000", values["port"]);
        Assert.Equal(":DISP:DATA?", values["dumpImage"]);
        Assert.Equal(string.Empty, values["stop"]);
        Assert.DoesNotContain("brand", values.Keys);
        Assert.DoesNotContain("seriesOrModel", values.Keys);
        Assert.Equal(OscilloscopeModelValues.Names, values.Keys);

        MainDocument.OscilloscopeRow back = OscilloscopeModelValues.ToModel(" Voltcraft ", "DOV702 ", values);

        Assert.Equal("Voltcraft", back.Brand);
        Assert.Equal("DOV702", back.SeriesOrModel);
        Assert.Equal(DataTreeJson.Write(new MainDocument { Oscilloscopes = [Model()] }), DataTreeJson.Write(new MainDocument { Oscilloscopes = [back] }));
    }

    // A name the row does not know is left out - the checks refuse it before it gets this far.
    [Fact]
    public void A_value_the_main_file_does_not_know_is_left_out()
    {
        MainDocument.OscilloscopeRow model = OscilloscopeModelValues.ToModel("Voltcraft", "DOV702", new Dictionary<string, string>
        {
            ["identify"] = "*IDN?",
            ["brand"] = "Owon",
            ["launchMissiles"] = "*RST",
        });

        Assert.Equal("Voltcraft", model.Brand);
        Assert.Null(model.Unknown);
    }

    // ###########################################################################################
    // EVERY value of the model, and its name, decides whether the suite went through on it as it is -
    // the set the model is sent with (code reviews, 2026-10-29: a list of the tab's own would miss a
    // value added to the model, and the administrator was told of a suite never run on it).
    // ###########################################################################################
    [Fact]
    public void Every_value_of_the_model_and_its_name_make_its_fingerprint()
    {
        MainDocument.OscilloscopeRow model = Model();
        string ran = OscilloscopeModelValues.Fingerprint(model);
        IReadOnlyDictionary<string, string> values = OscilloscopeModelValues.Of(model);

        foreach (string name in OscilloscopeModelValues.Names)
        {
            var changed = values.ToDictionary(value => value.Key, value => value.Value);
            changed[name] += " changed";

            Assert.NotEqual(ran, OscilloscopeModelValues.Fingerprint(OscilloscopeModelValues.ToModel("Voltcraft", "DOV702", changed)));
        }

        Assert.NotEqual(ran, OscilloscopeModelValues.Fingerprint(OscilloscopeModelValues.ToModel("Voltcraft", "DOV702-2", values)));
        Assert.NotEqual(ran, OscilloscopeModelValues.Fingerprint(OscilloscopeModelValues.ToModel("Voltcraft2", "DOV702", values)));
        Assert.Equal(ran, OscilloscopeModelValues.Fingerprint(OscilloscopeModelValues.ToModel(" Voltcraft ", "DOV702", values)));
    }

    // ------------------------------------------------------------------ the checks

    [Fact]
    public void A_whole_model_has_no_problem()
    {
        Assert.Empty(OscilloscopeModelRules.SubmissionProblems("Voltcraft", "DOV702", OscilloscopeModelValues.Of(Model()), "Firmware 1.2"));
    }

    [Fact]
    public void A_model_is_refused_for_each_thing_wrong_with_it()
    {
        var values = new Dictionary<string, string>(OscilloscopeModelValues.Of(Model()))
        {
            ["identify"] = "",
            ["stop"] = ":STOP\r\n:RUN",
            ["run"] = new string('x', OscilloscopeModelRules.ValueLimit + 1),
            ["launchMissiles"] = "*RST",
        };

        IReadOnlyList<string> problems = OscilloscopeModelRules.SubmissionProblems(
            "",
            new string('M', OscilloscopeModelRules.NameLimit + 1),
            values,
            new string('m', OscilloscopeModelRules.MessageLimit + 1));

        Assert.Contains(OscilloscopeModelRules.NoVendor, problems);
        Assert.Contains(problems, problem => problem.StartsWith("The series or model is longer", StringComparison.Ordinal));
        Assert.Contains(OscilloscopeModelRules.NoIdentify, problems);
        Assert.Contains(problems, problem => problem.Contains("\"stop\"") && problem.Contains("line break"));
        Assert.Contains(problems, problem => problem.Contains("\"run\"") && problem.Contains("longer"));
        Assert.Contains(problems, problem => problem.Contains("launchMissiles"));
        Assert.Contains(problems, problem => problem.StartsWith("The message is longer", StringComparison.Ordinal));
    }

    [Fact]
    public void A_model_with_no_values_or_a_name_holding_a_line_break_is_refused()
    {
        Assert.Contains("The model came without its commands.", OscilloscopeModelRules.SubmissionProblems("Voltcraft", "DOV702", null, null));
        Assert.Contains(
            OscilloscopeModelRules.SubmissionProblems("Volt\ncraft", "DOV702", OscilloscopeModelValues.Of(Model()), null),
            problem => problem.Contains("line break"));
    }

    // ------------------------------------------------------------------ into a main file

    private sealed class MainFile : IDisposable
    {
        public MainFile(params MainDocument.OscilloscopeRow[] models)
        {
            Directory.CreateDirectory(this.Folder);
            File.WriteAllText(this.Path, DataTreeJson.Write(new MainDocument
            {
                Boards = [new MainDocument.BoardListingRow { BoardId = "Commodore/C64/250407", HardwareName = "C64", BoardName = "250407" }],
                Oscilloscopes = models.ToList(),
            }), new UTF8Encoding(false));
        }

        public string Folder { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "crt-put-scope", Guid.NewGuid().ToString("N"));

        public string Path => System.IO.Path.Combine(this.Folder, DataTreeLayout.MainFileName);

        public MainDocument Read() => DataTreeJson.ReadMain(File.ReadAllText(this.Path));

        public void Dispose() => Directory.Delete(this.Folder, recursive: true);
    }

    [Fact]
    public void A_new_model_goes_after_the_last_and_dates_the_list()
    {
        using var main = new MainFile(Model("Rigol", "DS1000Z"));

        MasterListingEdit edit = MasterListing.PutOscilloscope(main.Path, Model(), new DateTimeOffset(2026, 10, 29, 12, 0, 0, TimeSpan.Zero));

        Assert.True(edit.IsDone);
        Assert.True(edit.Changed);

        MainDocument written = main.Read();
        Assert.Equal(["DS1000Z", "DOV702"], written.Oscilloscopes.Select(model => model.SeriesOrModel));
        Assert.Equal("C64", Assert.Single(written.Boards).HardwareName);
        Assert.Equal(BoardWorkbookStyle.FormatRevisionDate(new DateTimeOffset(2026, 10, 29, 12, 0, 0, TimeSpan.Zero)), written.OscilloscopesRevisionDate);
    }

    [Fact]
    public void A_model_of_a_name_there_takes_its_place_and_one_as_it_is_writes_nothing()
    {
        using var main = new MainFile(Model("Rigol", "DS1000Z"), Model(), Model("Siglent", "SDS1000X-E"));
        DateTime before = File.GetLastWriteTimeUtc(main.Path);

        MasterListingEdit same = MasterListing.PutOscilloscope(main.Path, Model());

        Assert.True(same.IsDone);
        Assert.False(same.Changed);
        Assert.Equal(before, File.GetLastWriteTimeUtc(main.Path));

        MasterListingEdit update = MasterListing.PutOscilloscope(main.Path, Model("voltcraft", "dov702", ":DISP:DATA? BMP"));

        Assert.True(update.Changed);
        MainDocument written = main.Read();
        Assert.Equal(3, written.Oscilloscopes.Count);
        Assert.Equal(":DISP:DATA? BMP", written.Oscilloscopes[1].DumpImage);
        Assert.Equal("dov702", written.Oscilloscopes[1].SeriesOrModel);
    }

    // The Excel generations are frozen: nothing is written into one.
    [Fact]
    public void An_Excel_main_file_takes_no_model()
    {
        MasterListingEdit edit = MasterListing.PutOscilloscope(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "Classic-Repair-Toolbox.v3.0.0.xlsx"), Model());

        Assert.False(edit.IsDone);
        Assert.Contains("frozen", edit.Failure);
    }
}
