using CRT;
using System;
using System.Collections.Generic;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Runtime simulation switches, parsed once from the command line at startup.
    //
    // These replace the "#if DEBUG" blocks that used to make a DEBUG build fake an application
    // update while a RELEASE build could not, and vice versa. The build configuration no longer
    // changes what the application does: a DEBUG build and a RELEASE build given the same
    // arguments behave identically, which is the whole point of this class. That also means the
    // simulation is reachable from a RELEASE build - deliberately, since a simulation only a
    // DEBUG build can reach is a simulation CI and the Stop hook can never exercise.
    //
    // Nothing here is persisted. A launch with no arguments is always the real behaviour, so a
    // simulation can never be left switched on by accident the way an edited constant could.
    //
    // Parsing follows the conventions DataManager.ResolveDataRoot already established for
    // "--data-root=": case-insensitive, surrounding quotes stripped from a value, first match
    // wins, unrecognised arguments ignored.
    //
    // Parse() is pure and takes its arguments rather than reading the environment, so the whole
    // grammar is testable; Current is the single process-wide value that the rest of the app reads.
    // ###########################################################################################
    public sealed class SimulationOptions
    {
        // Offers a fake application update instead of asking GitHub. Accepted bare, or with an
        // explicit version as "--simulate-update=2.7.0".
        public const string SimulateUpdateArg = "--simulate-update";

        private const string SimulateUpdateValueArg = SimulateUpdateArg + "=";

        // ###########################################################################################
        // The "no simulation" value - what an ordinary launch gets, and the starting value of Current
        // so that anything constructed before Initialize (or in a test) sees real behaviour.
        // ###########################################################################################
        public static readonly SimulationOptions None =
            new(simulateUpdate: false, simulatedUpdateVersion: AppConfig.SimulatedUpdateVersion);

        // ###########################################################################################
        // The options this process was started with. Set once by Program.Main.
        // ###########################################################################################
        public static SimulationOptions Current { get; private set; } = None;

        // True when the update check should report a fake update instead of querying GitHub.
        public bool SimulateUpdate { get; }

        // The version the fake update claims to be. Only meaningful while SimulateUpdate is true.
        public string SimulatedUpdateVersion { get; }

        // True when any simulation at all is active, so callers can log the warning block once.
        public bool IsAnyActive => this.SimulateUpdate;

        private SimulationOptions(bool simulateUpdate, string simulatedUpdateVersion)
        {
            this.SimulateUpdate = simulateUpdate;
            this.SimulatedUpdateVersion = simulatedUpdateVersion;
        }

        // ###########################################################################################
        // Parses the simulation switches out of a command line. Pure - no statics are touched and
        // nothing is read from the environment, so tests can drive the whole grammar directly.
        //
        // "--simulate-update" alone uses the default simulated version; "--simulate-update=2.7.0"
        // overrides it; "--simulate-update=" falls back to the default rather than claiming an
        // empty version. Unrecognised arguments are ignored so that "--data-root=" and anything
        // added later cannot accidentally trip a switch.
        // ###########################################################################################
        public static SimulationOptions Parse(IEnumerable<string>? args)
        {
            if (args == null)
                return SimulationOptions.None;

            foreach (var arg in args)
            {
                if (string.IsNullOrWhiteSpace(arg))
                    continue;

                string candidate = arg.Trim();

                // Exact match, not StartsWith - otherwise a future "--simulate-updates" or a typo
                // like "--simulate-update-now" would silently switch the simulation on.
                if (string.Equals(candidate, SimulationOptions.SimulateUpdateArg, StringComparison.OrdinalIgnoreCase))
                {
                    return new SimulationOptions(true, AppConfig.SimulatedUpdateVersion);
                }

                if (candidate.StartsWith(SimulationOptions.SimulateUpdateValueArg, StringComparison.OrdinalIgnoreCase))
                {
                    string version = candidate[SimulationOptions.SimulateUpdateValueArg.Length..]
                        .Trim('"', '\'')
                        .Trim();

                    return new SimulationOptions(
                        true,
                        string.IsNullOrWhiteSpace(version) ? AppConfig.SimulatedUpdateVersion : version);
                }
            }

            return SimulationOptions.None;
        }

        // ###########################################################################################
        // The lines describing every active simulation, for the startup log. Returns nothing when no
        // simulation is active. Kept here rather than in App so the wording is testable - these lines
        // are the only thing standing between a faked update and a bug report about a real one.
        //
        // checkVersionOnLaunchEnabled is passed in rather than read from UserSettings, both to keep
        // this pure and because a simulated update that no one will ever see is the single most
        // confusing state this switch can produce: the argument is accepted, announced, and then
        // nothing happens, because the launch version check the simulation would have answered is
        // switched off in settings. Saying so here puts the cause in the same block as the effect.
        // ###########################################################################################
        public IReadOnlyList<string> DescribeForLog(bool checkVersionOnLaunchEnabled)
        {
            var lines = new List<string>();

            if (this.SimulateUpdate)
            {
                lines.Add(
                    $"[{SimulationOptions.SimulateUpdateArg}] a fake update [{this.SimulatedUpdateVersion}] is offered " +
                    "where the download is faked");

                if (!checkVersionOnLaunchEnabled)
                {
                    lines.Add(
                        "Simulated update requested, but [Check for new version at application launch] is not checked " +
                        "in \"Configuration\" tab - enable that for the simulation to run");
                }
            }

            return lines;
        }

        // ###########################################################################################
        // Sets the options for this process from its command line. Called once, from Program.Main,
        // before anything can read Current. Internal so that only the entry point and the test suite
        // can reach it - the same seam discipline as UserSettings.LoadFrom and DataManager.LoadFrom.
        // ###########################################################################################
        internal static void Initialize(IEnumerable<string>? args)
        {
            SimulationOptions.Current = SimulationOptions.Parse(args);
        }
    }
}
