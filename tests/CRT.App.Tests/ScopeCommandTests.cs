using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for the scope command palette and resolver.
//
// The most useful thing this file does is COMPLETENESS checking: both classes switch over
// enums, and adding a new ScopeCommand or ScopeCommandPalette without wiring it up throws
// only when a user happens to click that button with a scope connected. These tests turn
// that into a build-time failure instead.
public class ScopeCommandTests
{
    private static IEnumerable<ScopeCommand> AllCommands =>
        Enum.GetValues<ScopeCommand>();

    private static IEnumerable<ScopeCommandPalette> AllPalettes =>
        Enum.GetValues<ScopeCommandPalette>();

    // -------------------------------------------------------- ScopeCommandResolver

    [Fact]
    public void GetCommandText_handles_every_ScopeCommand_value()
    {
        // If this throws, a new ScopeCommand was added without a case in the resolver.
        var entry = new OscilloscopeEntry();

        foreach (ScopeCommand command in AllCommands)
        {
            Exception? thrown = Record.Exception(() => ScopeCommandResolver.GetCommandText(entry, command));
            Assert.True(thrown is null, $"ScopeCommandResolver.GetCommandText has no case for {command}");
        }
    }

    [Fact]
    public void GetCommandText_returns_the_text_defined_by_the_scope_entry()
    {
        var entry = new OscilloscopeEntry
        {
            Identify = "*IDN?",
            Stop = ":STOP",
            SetTimeDiv = ":TIM:SCAL {0}"
        };

        Assert.Equal("*IDN?", ScopeCommandResolver.GetCommandText(entry, ScopeCommand.Identify));
        Assert.Equal(":STOP", ScopeCommandResolver.GetCommandText(entry, ScopeCommand.Stop));
        Assert.Equal(":TIM:SCAL {0}", ScopeCommandResolver.GetCommandText(entry, ScopeCommand.SetTimeDiv));
    }

    [Fact]
    public void GetCommandText_returns_empty_for_a_command_the_scope_does_not_define()
    {
        // Scope entries come from the master workbook and are frequently partial.
        Assert.Equal(string.Empty, ScopeCommandResolver.GetCommandText(new OscilloscopeEntry(), ScopeCommand.DumpImage));
    }

    [Fact]
    public void GetCommandText_throws_for_an_undefined_enum_value()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => ScopeCommandResolver.GetCommandText(new OscilloscopeEntry(), (ScopeCommand)9999));
    }

    [Theory]
    [InlineData(ScopeCommand.Identify)]
    [InlineData(ScopeCommand.DrainErrorQueue)]
    [InlineData(ScopeCommand.OperationComplete)]
    [InlineData(ScopeCommand.QueryActiveTrigger)]
    [InlineData(ScopeCommand.QueryTriggerMode)]
    [InlineData(ScopeCommand.QueryTriggerLevel)]
    [InlineData(ScopeCommand.QueryTimeDiv)]
    [InlineData(ScopeCommand.QueryVoltsDiv)]
    public void Query_style_commands_expect_a_text_response(ScopeCommand command)
    {
        Assert.True(ScopeCommandResolver.ExpectsTextResponse(command));
    }

    [Theory]
    [InlineData(ScopeCommand.ClearStatistics)]
    [InlineData(ScopeCommand.Stop)]
    [InlineData(ScopeCommand.Single)]
    [InlineData(ScopeCommand.Run)]
    [InlineData(ScopeCommand.SetTriggerLevel)]
    [InlineData(ScopeCommand.SetTimeDiv)]
    [InlineData(ScopeCommand.SetVoltsDiv)]
    [InlineData(ScopeCommand.DumpImage)]
    public void Action_style_commands_do_not_expect_a_text_response(ScopeCommand command)
    {
        // Waiting for a reply that never comes is how the scope client hangs.
        Assert.False(ScopeCommandResolver.ExpectsTextResponse(command));
    }

    [Fact]
    public void Every_ScopeCommand_is_classified_by_one_of_the_two_response_tests_above()
    {
        // Guard against a new command being added and silently defaulting to "no response".
        var covered = new HashSet<ScopeCommand>
        {
            ScopeCommand.Identify, ScopeCommand.DrainErrorQueue, ScopeCommand.OperationComplete,
            ScopeCommand.QueryActiveTrigger, ScopeCommand.QueryTriggerMode, ScopeCommand.QueryTriggerLevel,
            ScopeCommand.QueryTimeDiv, ScopeCommand.QueryVoltsDiv,
            ScopeCommand.ClearStatistics, ScopeCommand.Stop, ScopeCommand.Single, ScopeCommand.Run,
            ScopeCommand.SetTriggerLevel, ScopeCommand.SetTimeDiv, ScopeCommand.SetVoltsDiv,
            ScopeCommand.DumpImage
        };

        var missing = AllCommands.Where(c => !covered.Contains(c)).ToList();

        Assert.True(missing.Count == 0,
            "New ScopeCommand value(s) not covered by the response-expectation tests: " +
            string.Join(", ", missing));
    }

    // ------------------------------------------------- ScopeCommandPaletteDefinitions

    [Fact]
    public void GetCommands_is_defined_for_every_palette()
    {
        // If this throws, a new ScopeCommandPalette was added without a definition.
        foreach (ScopeCommandPalette palette in AllPalettes)
        {
            Exception? thrown = Record.Exception(() => ScopeCommandPaletteDefinitions.GetCommands(palette));
            Assert.True(thrown is null, $"ScopeCommandPaletteDefinitions has no entry for {palette}");
        }
    }

    [Fact]
    public void Every_palette_resolves_to_at_least_one_command()
    {
        foreach (ScopeCommandPalette palette in AllPalettes)
        {
            Assert.NotEmpty(ScopeCommandPaletteDefinitions.GetCommands(palette));
        }
    }

    [Fact]
    public void GetCommands_throws_for_an_undefined_palette_value()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => ScopeCommandPaletteDefinitions.GetCommands((ScopeCommandPalette)9999));
    }

    [Fact]
    public void A_state_changing_palette_settles_the_scope_before_returning()
    {
        // Every palette that changes scope state ends with OperationComplete + DrainErrorQueue,
        // so the next command is not issued while the scope is still busy or in an error state.
        var stateChanging = new[]
        {
            ScopeCommandPalette.ClearStatistics, ScopeCommandPalette.Stop, ScopeCommandPalette.Single,
            ScopeCommandPalette.Run, ScopeCommandPalette.SetTriggerLevel, ScopeCommandPalette.SetTimeDiv,
            ScopeCommandPalette.SetVoltsDiv
        };

        foreach (ScopeCommandPalette palette in stateChanging)
        {
            var commands = ScopeCommandPaletteDefinitions.GetCommands(palette);

            Assert.True(commands.Count >= 3, $"{palette} should issue its command then settle the scope");
            Assert.Equal(ScopeCommand.OperationComplete, commands[^2]);
            Assert.Equal(ScopeCommand.DrainErrorQueue, commands[^1]);
        }
    }

    [Fact]
    public void A_query_only_palette_issues_exactly_one_command()
    {
        var queryOnly = new[]
        {
            ScopeCommandPalette.Identify, ScopeCommandPalette.DrainErrorQueue,
            ScopeCommandPalette.OperationComplete, ScopeCommandPalette.QueryActiveTrigger,
            ScopeCommandPalette.QueryTriggerMode, ScopeCommandPalette.QueryTriggerLevel,
            ScopeCommandPalette.QueryTimeDiv, ScopeCommandPalette.QueryVoltsDiv,
            ScopeCommandPalette.DumpImage
        };

        foreach (ScopeCommandPalette palette in queryOnly)
        {
            Assert.Single(ScopeCommandPaletteDefinitions.GetCommands(palette));
        }
    }

    [Fact]
    public void The_full_execution_order_contains_every_palette_exactly_once()
    {
        var order = ScopeCommandPaletteDefinitions.GetFullCommandPaletteExecutionOrder();

        Assert.Equal(AllPalettes.Count(), order.Count);
        Assert.Equal(AllPalettes.OrderBy(p => p), order.OrderBy(p => p));
        Assert.Equal(order.Count, order.Distinct().Count());
    }

    [Fact]
    public void The_full_execution_order_identifies_the_scope_first()
    {
        var order = ScopeCommandPaletteDefinitions.GetFullCommandPaletteExecutionOrder();
        Assert.Equal(ScopeCommandPalette.Identify, order[0]);
    }

    // --------------------------------------------------------------- enum agreement

    [Fact]
    public void ScopeCommand_and_ScopeCommandPalette_declare_the_same_names()
    {
        // The two enums mirror each other by design. Adding to one and not the other is the
        // mistake this catches.
        var commandNames = Enum.GetNames<ScopeCommand>().OrderBy(n => n, StringComparer.Ordinal);
        var paletteNames = Enum.GetNames<ScopeCommandPalette>().OrderBy(n => n, StringComparer.Ordinal);

        Assert.Equal(commandNames, paletteNames);
    }
}
