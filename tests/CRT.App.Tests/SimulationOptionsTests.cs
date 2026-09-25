using CRT;
using Handlers.DataHandling;
using System;
using Xunit;

// ###########################################################################################
// SimulationOptions replaced the "#if DEBUG" blocks that used to decide whether the application
// faked an update. The point of the class is that the build configuration no longer changes
// behaviour, so what these tests pin down is the argument grammar - the only thing that decides
// it now. Parse is pure, so almost everything here runs without touching any static state.
//
// Tests that assign SimulationOptions.Current live at the bottom and are the only ones that need
// serialising; see the note above them.
// ###########################################################################################
public class SimulationOptionsTests
{
    // A launch with no arguments must be the real behaviour. This is the guarantee that makes it
    // safe for the switch to exist in a RELEASE build at all.
    [Fact]
    public void No_arguments_means_no_simulation()
    {
        var options = SimulationOptions.Parse(Array.Empty<string>());

        Assert.False(options.SimulateUpdate);
        Assert.False(options.IsAnyActive);
    }

    [Fact]
    public void A_null_argument_array_means_no_simulation()
    {
        var options = SimulationOptions.Parse(null);

        Assert.False(options.SimulateUpdate);
    }

    [Fact]
    public void The_bare_switch_simulates_an_update_at_the_configured_default_version()
    {
        var options = SimulationOptions.Parse(new[] { "--simulate-update" });

        Assert.True(options.SimulateUpdate);
        Assert.True(options.IsAnyActive);
        Assert.Equal(AppConfig.SimulatedUpdateVersion, options.SimulatedUpdateVersion);
    }

    [Fact]
    public void An_explicit_version_overrides_the_default()
    {
        var options = SimulationOptions.Parse(new[] { "--simulate-update=2.7.0" });

        Assert.True(options.SimulateUpdate);
        Assert.Equal("2.7.0", options.SimulatedUpdateVersion);
    }

    // An empty value is a typo, not a request for a nameless update - fall back rather than put an
    // empty string in the banner.
    [Theory]
    [InlineData("--simulate-update=")]
    [InlineData("--simulate-update=   ")]
    public void An_empty_version_value_falls_back_to_the_default(string arg)
    {
        var options = SimulationOptions.Parse(new[] { arg });

        Assert.True(options.SimulateUpdate);
        Assert.Equal(AppConfig.SimulatedUpdateVersion, options.SimulatedUpdateVersion);
    }

    // Matches how --data-root= behaves, so the two switches cannot surprise each other.
    [Theory]
    [InlineData("--SIMULATE-UPDATE")]
    [InlineData("--Simulate-Update=3.1.4")]
    public void Switch_matching_is_case_insensitive(string arg)
    {
        Assert.True(SimulationOptions.Parse(new[] { arg }).SimulateUpdate);
    }

    // Shells hand quoted values through with the quotes attached; ResolveDataRoot strips them and
    // so does this, or the banner would read: Version ["2.7.0"] is available.
    [Theory]
    [InlineData("--simulate-update=\"2.7.0\"")]
    [InlineData("--simulate-update='2.7.0'")]
    public void Quotes_around_a_version_value_are_stripped(string arg)
    {
        Assert.Equal("2.7.0", SimulationOptions.Parse(new[] { arg }).SimulatedUpdateVersion);
    }

    // THE important one. The switch is matched by exact equality, not StartsWith - if someone
    // "simplifies" the parser into a single prefix test, this is what catches it. A prefix match
    // would turn every one of these into an active simulation.
    [Theory]
    [InlineData("--simulate-updates")]
    [InlineData("--simulate-update-now")]
    [InlineData("--simulate-updatefoo")]
    public void A_switch_that_merely_starts_with_the_name_does_not_activate_it(string arg)
    {
        Assert.False(SimulationOptions.Parse(new[] { arg }).SimulateUpdate);
    }

    // The application already has --data-root=; unknown arguments must pass through untouched so
    // that adding a switch later cannot retroactively change what an existing one does.
    [Fact]
    public void Unrelated_arguments_are_ignored()
    {
        var options = SimulationOptions.Parse(new[] { "--data-root=C:\\Temp\\Data", "--verbose", "nonsense" });

        Assert.False(options.SimulateUpdate);
    }

    [Fact]
    public void The_switch_is_found_alongside_other_arguments_in_any_position()
    {
        var options = SimulationOptions.Parse(new[] { "--data-root=C:\\Temp\\Data", "--simulate-update=1.2.3" });

        Assert.True(options.SimulateUpdate);
        Assert.Equal("1.2.3", options.SimulatedUpdateVersion);
    }

    // First wins, matching DataManager.ResolveDataRoot's documented behaviour for duplicates.
    [Fact]
    public void The_first_occurrence_wins_when_the_switch_is_repeated()
    {
        var options = SimulationOptions.Parse(new[] { "--simulate-update=1.0.0", "--simulate-update=2.0.0" });

        Assert.Equal("1.0.0", options.SimulatedUpdateVersion);
    }

    [Fact]
    public void Blank_and_whitespace_arguments_are_skipped_without_throwing()
    {
        var options = SimulationOptions.Parse(new[] { string.Empty, "   ", "--simulate-update" });

        Assert.True(options.SimulateUpdate);
    }

    // Shells and launch profiles occasionally deliver a padded argument; treat it as the switch.
    [Fact]
    public void Surrounding_whitespace_on_the_switch_itself_is_tolerated()
    {
        Assert.True(SimulationOptions.Parse(new[] { "  --simulate-update  " }).SimulateUpdate);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void The_default_options_describe_nothing_for_the_log(bool checkVersionOnLaunch)
    {
        // No simulation means nothing to warn about, whatever the launch version check is set to -
        // the caveat below must never appear on an ordinary run.
        Assert.Empty(SimulationOptions.None.DescribeForLog(checkVersionOnLaunch));
    }

    // These lines are the record that an update was faked. If they ever stop naming the switch and
    // the version, a screenshot of the banner becomes indistinguishable from a real update report.
    [Fact]
    public void An_active_simulation_names_its_switch_and_version_for_the_log()
    {
        var lines = SimulationOptions.Parse(new[] { "--simulate-update=4.5.6" })
            .DescribeForLog(checkVersionOnLaunchEnabled: true);

        string description = Assert.Single(lines);
        Assert.Contains(SimulationOptions.SimulateUpdateArg, description);
        Assert.Contains("4.5.6", description);
    }

    // The dead end that cost a real debugging session: the switch is accepted and announced, then
    // nothing happens, because Main only runs the update check when this setting is on. The log has
    // to say so in the same block, or the next person concludes the switch is broken.
    [Fact]
    public void A_simulated_update_warns_when_the_launch_version_check_is_disabled()
    {
        var lines = SimulationOptions.Parse(new[] { "--simulate-update=4.1.0" })
            .DescribeForLog(checkVersionOnLaunchEnabled: false);

        Assert.Equal(2, lines.Count);

        // Named exactly as the Configuration tab labels it. "data" and "version" are two different
        // checkboxes three rows apart, and naming the wrong one sends the reader to the wrong place.
        Assert.Contains("Check for new version at application launch", lines[1]);
    }

    [Fact]
    public void A_simulated_update_stays_quiet_when_the_launch_version_check_is_enabled()
    {
        var lines = SimulationOptions.Parse(new[] { "--simulate-update" })
            .DescribeForLog(checkVersionOnLaunchEnabled: true);

        Assert.Single(lines);
    }

    // ###########################################################################################
    // Below here the process-wide SimulationOptions.Current is assigned, so these belong to the
    // "SimulationOptions" collection - alone they pass either way, but in a full parallel run an
    // unserialised assignment here could be read by another test's code under test.
    // ###########################################################################################
    [Collection("SimulationOptions")]
    public class CurrentTests : IDisposable
    {
        // Always hand the process back its real behaviour, whatever the test did.
        public void Dispose() => SimulationOptions.Initialize(Array.Empty<string>());

        [Fact]
        public void Current_defaults_to_no_simulation()
        {
            SimulationOptions.Initialize(Array.Empty<string>());

            Assert.False(SimulationOptions.Current.SimulateUpdate);
        }

        [Fact]
        public void Initialize_publishes_the_parsed_options_as_Current()
        {
            SimulationOptions.Initialize(new[] { "--simulate-update=9.8.7" });

            Assert.True(SimulationOptions.Current.SimulateUpdate);
            Assert.Equal("9.8.7", SimulationOptions.Current.SimulatedUpdateVersion);
        }
    }
}
