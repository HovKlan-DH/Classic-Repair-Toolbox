using System.Reflection;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests;

// The values the Oscilloscope tab's model editor shows (ScopeModelFields): every value of a main
// file's oscilloscope entry but its name, each read and written where the main file keeps it, and
// the settings CRT could not use told apart from what is the scope's business.
public sealed class ScopeModelFieldsTests
{
    // A value the main file gains must get a box in the editor - or a decision that it has none.
    [Fact]
    public void Every_value_of_an_oscilloscope_entry_but_its_name_has_a_field()
    {
        string[] values = typeof(MainDocument.OscilloscopeRow)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(property => property.PropertyType == typeof(string))
            .Select(property => property.Name)
            .Where(name => name is not (nameof(MainDocument.OscilloscopeRow.Brand) or nameof(MainDocument.OscilloscopeRow.SeriesOrModel)))
            .ToArray();

        Assert.Equal(values.Length, ScopeModelFields.All.Count);

        // Each field writes ITS value and reads it back - no two share a property.
        foreach (ScopeModelField field in ScopeModelFields.All)
        {
            var row = new MainDocument.OscilloscopeRow();
            field.Set(row, "value of " + field.Key);

            Assert.Equal("value of " + field.Key, field.Get(row));

            string[] set = values.Where(name => (string?)typeof(MainDocument.OscilloscopeRow).GetProperty(name)!.GetValue(row) == "value of " + field.Key).ToArray();
            Assert.Single(set);
        }
    }

    // The key is the main file's own name for the value, so a field is found by what the file says.
    [Fact]
    public void A_field_s_key_is_the_main_file_s_name_for_its_value()
    {
        foreach (ScopeModelField field in ScopeModelFields.All)
        {
            var row = new MainDocument.OscilloscopeRow();
            field.Set(row, "x");

            PropertyInfo written = typeof(MainDocument.OscilloscopeRow)
                .GetProperties()
                .Single(property => property.PropertyType == typeof(string) && (string?)property.GetValue(row) == "x");

            Assert.Equal(field.Key, written.GetCustomAttribute<System.Text.Json.Serialization.JsonPropertyNameAttribute>()!.Name);
        }
    }

    [Fact]
    public void Every_command_has_one_field_with_Try_it()
    {
        foreach (ScopeCommand command in Enum.GetValues<ScopeCommand>())
        {
            Assert.Equal(ScopeModelFieldKind.Command, ScopeModelFields.Of(command).Kind);
        }

        Assert.Equal(Enum.GetValues<ScopeCommand>().Length, ScopeModelFields.All.Count(field => field.Command is not null));
    }

    [Theory]
    [InlineData("5025", null)]
    [InlineData("1", null)]
    [InlineData("65535", null)]
    [InlineData("0", ScopeModelFields.PortProblem)]
    [InlineData("65536", ScopeModelFields.PortProblem)]
    [InlineData("50x5", ScopeModelFields.PortProblem)]
    [InlineData("", ScopeModelFields.PortProblem)]
    public void A_network_port_is_a_number_from_1_to_65535(string value, string? problem)
    {
        ScopeModelField port = ScopeModelFields.All.Single(field => field.Key == ScopeModelFields.PortKey);

        Assert.Equal(problem, ScopeModelFields.ProblemOf(port, value));
        Assert.True(ScopeModelFields.IsHeldBackWhenWrong(port));
    }

    // Empty is the tab's own 250 ms (ReadOscilloscopeSelectionInputs).
    [Theory]
    [InlineData("", null)]
    [InlineData("100ms", null)]
    [InlineData("1s", null)]
    [InlineData("100", ScopeModelFields.WaitProblem)]
    [InlineData("soon", ScopeModelFields.WaitProblem)]
    public void The_wait_between_commands_is_a_time_with_its_unit(string value, string? problem)
    {
        ScopeModelField wait = ScopeModelFields.All.Single(field => field.Key == ScopeModelFields.WaitKey);

        Assert.Equal(problem, ScopeModelFields.ProblemOf(wait, value));
    }

    // A list of steps is saved anyway - the steps CRT can read still work - and the ones it cannot
    // are named, so the user can find them.
    [Fact]
    public void A_list_of_steps_names_the_steps_CRT_cannot_read()
    {
        ScopeModelField time = ScopeModelFields.All.Single(field => field.Key == ScopeModelFields.TimeStepsKey);
        ScopeModelField volts = ScopeModelFields.All.Single(field => field.Key == ScopeModelFields.VoltsStepsKey);

        Assert.Null(ScopeModelFields.ProblemOf(time, "1us, 2us,5us"));
        Assert.Null(ScopeModelFields.ProblemOf(volts, "500mV,1V,2V"));
        Assert.Equal(ScopeModelFields.NoSteps, ScopeModelFields.ProblemOf(time, " "));

        string? problem = ScopeModelFields.ProblemOf(time, "1us,2,5 parsecs");
        Assert.Contains("\"2\"", problem);
        Assert.Contains("\"5 parsecs\"", problem);
        Assert.DoesNotContain("\"1us\"", problem);

        Assert.Contains("\"1V0\"", ScopeModelFields.ProblemOf(volts, "1V0,2V"));
        Assert.False(ScopeModelFields.IsHeldBackWhenWrong(time));
    }

    // A command is the scope's business: Try it says whether it works.
    [Fact]
    public void A_command_has_no_problem_CRT_can_tell()
    {
        Assert.Null(ScopeModelFields.ProblemOf(ScopeModelFields.Of(ScopeCommand.Stop), "anything at all"));
        Assert.False(ScopeModelFields.IsHeldBackWhenWrong(ScopeModelFields.Of(ScopeCommand.Stop)));
    }

    // Words a person reads are plain ASCII (CLAUDE.md) - the {0} the help names included.
    [Fact]
    public void The_words_are_plain_ASCII()
    {
        foreach (ScopeModelField field in ScopeModelFields.All)
        {
            Assert.All((field.Label + field.Help).ToCharArray(), letter => Assert.True(letter < 128, $"{field.Key}: '{letter}'"));
        }
    }
}
