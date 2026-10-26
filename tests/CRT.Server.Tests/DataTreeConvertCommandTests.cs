using CRT.Server.Handlers.DataTree;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // `CRT.Server convert-data-tree` - the command the release day's conversion of BETA and stable
    // is run with. What must hold: a command line it does not understand converts nothing (a
    // mistyped option must never write into the wrong place), a conversion with a problem writes
    // nothing and says why, and the report says what was converted and what was left behind.
    // The conversion itself is CRT.Data's, tested there (DataTreeConversionTests).
    // ###########################################################################################
    public class DataTreeConvertCommandTests
    {
        [Fact]
        public void The_command_is_recognised_by_its_name_only_as_the_first_argument()
        {
            Assert.True(DataTreeConvertCommand.IsCommand(["convert-data-tree", "--source=x"]));
            Assert.False(DataTreeConvertCommand.IsCommand(["--urls=http://127.0.0.1:5199", "convert-data-tree"]));
            Assert.False(DataTreeConvertCommand.IsCommand([]));
        }

        [Fact]
        public void Every_option_is_read()
        {
            DataTreeConvertCommand.Options? options = DataTreeConvertCommand.Parse(
                ["--source=/srv/app-data/Data", "--target=\"/srv/app-data-v4\"", "--public-url=https://classic-repair-toolbox.dk/app-data-v4/Data"],
                out string why);

            Assert.NotNull(options);
            Assert.Equal("/srv/app-data/Data", options!.Source);
            Assert.Equal("/srv/app-data-v4", options.Target);
            Assert.Equal("https://classic-repair-toolbox.dk/app-data-v4/Data", options.PublicUrl);
            Assert.False(options.PlanOnly);
            Assert.Empty(why);
        }

        [Fact]
        public void A_plan_needs_no_target()
        {
            DataTreeConvertCommand.Options? options = DataTreeConvertCommand.Parse(["--source=/srv/app-data/Data", "--plan-only"], out _);

            Assert.NotNull(options);
            Assert.True(options!.PlanOnly);
        }

        [Theory]
        [InlineData(new[] { "--target=/srv/app-data-v4" }, "--source is required")]
        [InlineData(new[] { "--source=/srv/app-data/Data" }, "--target is required")]
        [InlineData(new[] { "--source=/a", "--target=/b", "--taget=/c" }, "Unknown option [--taget=/c]")]
        [InlineData(new[] { "--source=/a", "--source=/b", "--target=/c" }, "given twice")]
        [InlineData(new[] { "--source=", "--target=/c" }, "needs a value")]
        [InlineData(new[] { "--source=/a", "--target=/c", "--public-url=app-data-v4/Data" }, "not a web address")]
        [InlineData(new[] { "--source=/a", "--target=/c", "--public-url=ftp://example.org/Data" }, "not a web address")]
        public void A_command_line_it_does_not_understand_is_refused_saying_why(string[] args, string expected)
        {
            Assert.Null(DataTreeConvertCommand.Parse(args, out string why));
            Assert.Contains(expected, why);
        }

        [Fact]
        public void A_wrong_command_line_converts_nothing_and_shows_the_usage()
        {
            var output = new StringWriter();

            int exit = DataTreeConvertCommand.Run(["convert-data-tree", "--source=/a"], output);

            Assert.Equal(DataTreeConvertCommand.WrongCommandLine, exit);
            Assert.Contains("Usage:", output.ToString());
        }

        // A source that is not there is a problem, and a problem writes nothing.
        [Fact]
        public void A_conversion_with_a_problem_writes_nothing_and_says_why()
        {
            string missing = Path.Combine(Path.GetTempPath(), "crt-no-data-" + Guid.NewGuid().ToString("N"));
            string target = Path.Combine(Path.GetTempPath(), "crt-no-target-" + Guid.NewGuid().ToString("N"));
            var output = new StringWriter();

            int exit = DataTreeConvertCommand.Run(["convert-data-tree", $"--source={missing}", $"--target={target}"], output);

            Assert.Equal(DataTreeConvertCommand.Refused, exit);
            Assert.Contains("PROBLEM: The data folder", output.ToString());
            Assert.Contains("Nothing was written", output.ToString());
            Assert.False(Directory.Exists(target));
        }

        [Fact]
        public void The_report_says_what_was_converted_and_what_was_left_behind()
        {
            var plan = new DataTreeConversionPlan
            {
                SourceDataRoot = "/srv/app-data/Data",
                Generation = new Version(3, 0, 0),
                Main = new MainDocument { Oscilloscopes = [new() { Brand = "Rigol", SeriesOrModel = "DHO800" }] },
                Boards =
                [
                    new ConvertedBoard("Commodore/C64/250407", "Commodore/C64/250407/Data C64 250407 v3.0.0.xlsx",
                        new BoardDocument { Files = [new() { Name = "a.png", Sha256 = new string('a', 64), Size = 1_500_000 }] },
                        new Dictionary<string, string>()),
                ],
                StoredFiles = new Dictionary<string, DataTreeSourceFile>
                {
                    ["Files/aa/" + new string('a', 64) + ".png"] = new("Commodore/C64/250407/a.png", "/srv/a.png", new string('a', 64), 1_500_000),
                },
                CopiedFiles = new Dictionary<string, DataTreeSourceFile>(),
                SourceFiles = new Dictionary<string, DataTreeSourceFile>(),
                LeftOut = Enumerable.Range(1, 505).Select(i => $"Old/{i}.xlsx").ToList(),
                Problems = ["Board [X]: something"],
                Notes = ["Board [Y]: spelled differently"],
            };

            IReadOnlyList<string> report = DataTreeConvertCommand.Report(plan);

            Assert.Contains("Generation converted: 3.0.0", report);
            Assert.Contains("Boards: 1", report);
            Assert.Contains("Oscilloscopes: 1", report);
            Assert.Contains("Files the boards name: 1 (1.5 MB)", report);
            Assert.Contains("Stored once: 1 file(s) (1.5 MB)", report);
            Assert.Contains("Left out - not used by the generation converted (older workbooks, what only they cite, !-notes): 505 file(s)", report);
            Assert.Contains("  Old/500.xlsx", report);
            Assert.DoesNotContain("  Old/501.xlsx", report);
            Assert.Contains("  ... and 5 more", report);
            Assert.Contains("Note: Board [Y]: spelled differently", report);
            Assert.Contains("PROBLEM: Board [X]: something", report);
        }
    }
}
