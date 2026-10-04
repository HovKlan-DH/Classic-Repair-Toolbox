using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // "Account" > API USAGE (owner request, 2026-10-04: "how about tracking the API
    // end-points, to see if it is possible to retire any, if almost no versions uses it any more").
    // Every word ApiUsageView shows, grouped and ordered here so a test can read it.
    //
    // *** THE ORDER IS THE QUESTION BEING ASKED. *** Within each group the LEAST-called routes come
    // first - a route nobody called heads its group - because "can anything be retired?" is answered
    // at the top, not by scrolling past the busy ones. The groups are the server's areas: what the
    // Drafts tab's Submit uses, what this tab uses, and what every CRT ever released sends, which is
    // never retired (owner decision, 2026-10-04) and comes last.
    // ###########################################################################################
    public static class ApiUsageDisplay
    {
        public const string Heading = "API usage";

        public const string Explanation =
            "Which CRT versions called each route of the server in the chosen days, counted by the server - no " +
            "address and no account is kept. A route that only old versions call, or nobody calls, can be retired: " +
            "it then answers those versions \"please update CRT\" instead of working. The routes every CRT ever " +
            "released sends are never retired.";

        // The windows offered, and the one shown first.
        public static IReadOnlyList<int> DayChoices { get; } = [30, 90, 365];

        public const int DefaultDays = 90;

        public const string SubmissionsHeading = "Submit, in the Drafts tab";

        public const string MaintainerHeading = "The Maintainer tab";

        public const string ForeverHeading = "Sent by every CRT ever released - never retired";

        public const string NotCrtLabel = "Not CRT (a browser or a check of the server)";

        public const string OtherLabel = "Other versions (too many different ones that day)";

        public static string DayChoiceLabel(int days) =>
            days switch
            {
                30 => "The last 30 days",
                90 => "The last 90 days",
                365 => "The last year",
                _ => string.Create(CultureInfo.InvariantCulture, $"The last {days} days")
            };

        // ###########################################################################################
        // The groups, in order, each with its routes least-called first (then by route). An area this
        // version does not know is shown under its own name after the known ones, so a newer server's
        // routes are still listed.
        // ###########################################################################################
        public static IReadOnlyList<ApiUsageGroup> Groups(ApiUsageAnswer answer)
        {
            ArgumentNullException.ThrowIfNull(answer);

            return answer.Routes
                .GroupBy(route => route.NeverRetired ? "Forever" : route.Area ?? string.Empty, StringComparer.Ordinal)
                .OrderBy(group => ApiUsageDisplay.AreaOrder(group.Key))
                .ThenBy(group => group.Key, StringComparer.Ordinal)
                .Select(group => new ApiUsageGroup(
                    ApiUsageDisplay.AreaHeading(group.Key),
                    group
                        .OrderBy(route => route.Calls)
                        .ThenBy(route => route.Route, StringComparer.Ordinal)
                        .ThenBy(route => route.Method, StringComparer.Ordinal)
                        .Select(route => new ApiUsageRouteLines(
                            $"{route.Method} {route.Route}",
                            ApiUsageDisplay.Summary(route, answer.Days),
                            route.Versions.Select(ApiUsageDisplay.VersionLine).ToList()))
                        .ToList()))
                .ToList();
        }

        // "No calls in the last [90] days", or "[120] calls, the last on 2026-October-4".
        public static IReadOnlyList<ReviewNoteRun> Summary(ApiUsageRoute route, int days)
        {
            ArgumentNullException.ThrowIfNull(route);

            if (route.Calls <= 0 || route.LastUtc is null)
            {
                return
                [
                    new ReviewNoteRun("No calls in the last [", IsCount: false),
                    new ReviewNoteRun(days.ToString(CultureInfo.InvariantCulture), IsCount: true),
                    new ReviewNoteRun("] days", IsCount: false)
                ];
            }

            return ApiUsageDisplay.Calls(route.Calls, route.LastUtc.Value);
        }

        // "3.0.0: [100] calls, the last on 2026-October-4".
        public static IReadOnlyList<ReviewNoteRun> VersionLine(ApiUsageVersion version)
        {
            ArgumentNullException.ThrowIfNull(version);

            var runs = new List<ReviewNoteRun> { new(ApiUsageDisplay.VersionLabel(version.Version) + ": ", IsCount: false) };
            runs.AddRange(ApiUsageDisplay.Calls(version.Calls, version.LastUtc));

            return runs;
        }

        // ###########################################################################################
        // The launches per version from the check-ins - "3.0.0: [42] installations, [305] launches" -
        // or, with none, a line saying so for the days shown.
        // ###########################################################################################
        public static IReadOnlyList<IReadOnlyList<ReviewNoteRun>> InstallationLines(ApiUsageAnswer answer)
        {
            ArgumentNullException.ThrowIfNull(answer);

            if (answer.Installations.Count == 0)
            {
                return
                [
                    [
                        new ReviewNoteRun("No CRT checked in at launch in the last [", IsCount: false),
                        new ReviewNoteRun(answer.Days.ToString(CultureInfo.InvariantCulture), IsCount: true),
                        new ReviewNoteRun("] days", IsCount: false)
                    ]
                ];
            }

            return answer.Installations
                .Select(installation => (IReadOnlyList<ReviewNoteRun>)
                [
                    new ReviewNoteRun(ApiUsageDisplay.VersionLabel(installation.Version) + ": [", IsCount: false),
                    new ReviewNoteRun(installation.Installations.ToString("N0", CultureInfo.InvariantCulture), IsCount: true),
                    new ReviewNoteRun("] " + (installation.Installations == 1 ? "installation" : "installations") + ", [", IsCount: false),
                    new ReviewNoteRun(installation.Launches.ToString("N0", CultureInfo.InvariantCulture), IsCount: true),
                    new ReviewNoteRun("] " + (installation.Launches == 1 ? "launch" : "launches"), IsCount: false)
                ])
                .ToList();
        }

        public const string InstallationsHeading = "CRT versions launched (from the launch check-ins)";

        public static string VersionLabel(string version) =>
            version switch
            {
                ApiUsageVersion.NotCrt => ApiUsageDisplay.NotCrtLabel,
                ApiUsageVersion.Other => ApiUsageDisplay.OtherLabel,
                _ => version
            };

        private static IReadOnlyList<ReviewNoteRun> Calls(long calls, DateTimeOffset lastUtc) =>
        [
            new ReviewNoteRun("[", IsCount: false),
            new ReviewNoteRun(calls.ToString("N0", CultureInfo.InvariantCulture), IsCount: true),
            new ReviewNoteRun("] " + (calls == 1 ? "call" : "calls") + ", the last on " + SubmissionReceiptPresenter.FormatDate(lastUtc), IsCount: false)
        ];

        private static int AreaOrder(string area) =>
            area switch
            {
                "Submissions" => 0,
                "Maintainer" => 1,
                "Forever" => 3,
                _ => 2
            };

        private static string AreaHeading(string area) =>
            area switch
            {
                "Submissions" => ApiUsageDisplay.SubmissionsHeading,
                "Maintainer" => ApiUsageDisplay.MaintainerHeading,
                "Forever" => ApiUsageDisplay.ForeverHeading,
                "" => "Other routes",
                _ => area
            };
    }

    // One group of routes under its heading.
    public sealed record ApiUsageGroup(string Heading, IReadOnlyList<ApiUsageRouteLines> Routes);

    // One route: "POST /api/review/systems/edit", its summary line, and a line per version.
    public sealed record ApiUsageRouteLines(
        string Title,
        IReadOnlyList<ReviewNoteRun> Summary,
        IReadOnlyList<IReadOnlyList<ReviewNoteRun>> Versions);
}
