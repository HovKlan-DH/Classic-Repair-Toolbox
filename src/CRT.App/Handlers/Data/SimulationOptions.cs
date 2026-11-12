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

        // Plays every notice once after launch, in the real window (Main.NoticeDemo.cs), so their
        // look and behaviour can be seen without waiting for each real cause (owner request,
        // 2026-10-11). Bare only.
        public const string DemoNoticesArg = "--demo-notices";

        // ###########################################################################################
        // Shows each tab's introduction tour the first time the tab is shown in this run, whatever was seen before, and
        // remembers nothing (owner request, 2026-11-11: "I need a possibility to enforce this every time when in
        // implementation mode - maybe via a commandline?"; Main.Tours.cs). Bare: the introduction. With a version,
        // "--show-tours=3.1.1": as somebody who last saw the tours in that version - the tour of what came since.
        // ###########################################################################################
        public const string ShowToursArg = "--show-tours";

        private const string ShowToursValueArg = ShowToursArg + "=";

        // ###########################################################################################
        // The "no simulation" value - what an ordinary launch gets, and the starting value of Current
        // so that anything constructed before Initialize (or in a test) sees real behaviour.
        // ###########################################################################################
        public static readonly SimulationOptions None =
            new(simulateUpdate: false, simulatedUpdateVersion: AppConfig.SimulatedUpdateVersion, demoNotices: false, showTours: false, toursSeenAs: null);

        // ###########################################################################################
        // The options this process was started with. Set once by Program.Main.
        // ###########################################################################################
        public static SimulationOptions Current { get; private set; } = None;

        // True when the update check should report a fake update instead of querying GitHub.
        public bool SimulateUpdate { get; }

        // The version the fake update claims to be. Only meaningful while SimulateUpdate is true.
        public string SimulatedUpdateVersion { get; }

        // True when every notice is to be played once after launch (--demo-notices).
        public bool DemoNotices { get; }

        // True when each tab's introduction tour is shown once a run, whatever was seen (--show-tours).
        public bool ShowTours { get; }

        // With ShowTours, the version the tours count as last seen in ("--show-tours=3.1.1") - null for none, which
        // shows the introduction.
        public string? ToursSeenAs { get; }

        // True when any simulation at all is active, so callers can log the warning block once.
        public bool IsAnyActive => this.SimulateUpdate || this.DemoNotices || this.ShowTours;

        private SimulationOptions(bool simulateUpdate, string simulatedUpdateVersion, bool demoNotices, bool showTours, string? toursSeenAs)
        {
            this.SimulateUpdate = simulateUpdate;
            this.SimulatedUpdateVersion = simulatedUpdateVersion;
            this.DemoNotices = demoNotices;
            this.ShowTours = showTours;
            this.ToursSeenAs = toursSeenAs;
        }

        // ###########################################################################################
        // Parses the simulation switches out of a command line. Pure - no statics are touched and
        // nothing is read from the environment, so tests can drive the whole grammar directly.
        //
        // "--simulate-update" alone uses the default simulated version; "--simulate-update=2.7.0"
        // overrides it; "--simulate-update=" falls back to the default rather than claiming an
        // empty version. Unrecognised arguments are ignored so that "--data-root=" and anything
        // added later cannot accidentally trip a switch. "--demo-notices" is found anywhere, beside
        // the update switch or not.
        // ###########################################################################################
        public static SimulationOptions Parse(IEnumerable<string>? args)
        {
            if (args == null)
                return SimulationOptions.None;

            bool simulateUpdate = false;
            string simulatedVersion = AppConfig.SimulatedUpdateVersion;
            bool demoNotices = false;
            bool showTours = false;
            string? toursSeenAs = null;

            foreach (var arg in args)
            {
                if (string.IsNullOrWhiteSpace(arg))
                    continue;

                string candidate = arg.Trim();

                // Exact match, like the update switch below.
                if (string.Equals(candidate, SimulationOptions.DemoNoticesArg, StringComparison.OrdinalIgnoreCase))
                {
                    demoNotices = true;
                    continue;
                }

                // The tours: bare, or with the version they count as seen in - the first occurrence wins, an empty
                // version is the bare switch.
                if (string.Equals(candidate, SimulationOptions.ShowToursArg, StringComparison.OrdinalIgnoreCase))
                {
                    showTours = true;
                    continue;
                }

                if (candidate.StartsWith(SimulationOptions.ShowToursValueArg, StringComparison.OrdinalIgnoreCase))
                {
                    if (!showTours)
                    {
                        string seen = candidate[SimulationOptions.ShowToursValueArg.Length..].Trim('"', '\'').Trim();
                        showTours = true;
                        toursSeenAs = seen.Length == 0 ? null : seen;
                    }

                    continue;
                }

                // The first occurrence of the update switch wins.
                if (simulateUpdate)
                    continue;

                // Exact match, not StartsWith - otherwise a future "--simulate-updates" or a typo
                // like "--simulate-update-now" would silently switch the simulation on.
                if (string.Equals(candidate, SimulationOptions.SimulateUpdateArg, StringComparison.OrdinalIgnoreCase))
                {
                    simulateUpdate = true;
                    continue;
                }

                if (candidate.StartsWith(SimulationOptions.SimulateUpdateValueArg, StringComparison.OrdinalIgnoreCase))
                {
                    string version = candidate[SimulationOptions.SimulateUpdateValueArg.Length..]
                        .Trim('"', '\'')
                        .Trim();

                    simulateUpdate = true;
                    simulatedVersion = string.IsNullOrWhiteSpace(version) ? AppConfig.SimulatedUpdateVersion : version;
                }
            }

            return simulateUpdate || demoNotices || showTours
                ? new SimulationOptions(simulateUpdate, simulatedVersion, demoNotices, showTours, toursSeenAs)
                : SimulationOptions.None;
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

            if (this.DemoNotices)
            {
                lines.Add(
                    $"[{SimulationOptions.DemoNoticesArg}] every notice is played once after launch, with example words - " +
                    "none of them is real");
            }

            // A value that is no version ("4.0", "four") counts as nothing seen - TourPlan's rule - so the line says the
            // introduction, and why (code review, 2026-11-12: it claimed a tour of what is new since the typo). A version
            // shows only what is NEWER - one as new as every step shows nothing, which the line says too.
            if (this.ShowTours)
            {
                bool sinceAVersion = CrtVersion.TryParse(this.ToursSeenAs, out CrtVersion? seen) && seen is not null;

                lines.Add(this.ToursSeenAs is null
                    ? $"[{SimulationOptions.ShowToursArg}] each tab's introduction tour is shown the first time the tab is shown, and nothing is remembered"
                    : sinceAVersion
                        ? $"[{SimulationOptions.ShowToursArg}={this.ToursSeenAs}] each tab's tour of what is new since [{this.ToursSeenAs}] is shown the first time the tab is shown - a tab with nothing newer shows none - and nothing is remembered"
                        : $"[{SimulationOptions.ShowToursArg}={this.ToursSeenAs}] is no version, so each tab's introduction tour is shown the first time the tab is shown, and nothing is remembered");
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
