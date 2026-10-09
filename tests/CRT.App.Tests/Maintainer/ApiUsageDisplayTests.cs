using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// "Account" > API usage (owner request, 2026-10-04: "how about tracking the API
// end-points, to see if it is possible to retire any, if almost no versions uses it any more").
//
// What matters: the groups say who uses a route, the routes every CRT ever released sends are set
// apart as never retired, and within a group the LEAST-called route comes first - the question is
// "can anything be retired?", answered at the top.
// ###########################################################################################
public sealed class ApiUsageDisplayTests
{
    private static readonly DateTimeOffset Last = new(2026, 10, 4, 10, 0, 0, TimeSpan.Zero);

    private static string Text(IReadOnlyList<ReviewNoteRun> line) => string.Concat(line.Select(run => run.Text));

    private static ApiUsageRoute Route(string method, string route, string area, long calls, params ApiUsageVersion[] versions) =>
        new(method, route, area, area == "Forever", calls, calls > 0 ? ApiUsageDisplayTests.Last : null, versions);

    private static ApiUsageAnswer Answer() => new(
        90,
        [
            ApiUsageDisplayTests.Route("POST", "/api/usage/check-in", "Forever", 900, new ApiUsageVersion("2.5.0", 900, ApiUsageDisplayTests.Last)),
            ApiUsageDisplayTests.Route("GET", "/api/review/queue", "Maintainer", 1500, new ApiUsageVersion("3.0.0", 1500, ApiUsageDisplayTests.Last)),
            ApiUsageDisplayTests.Route("POST", "/api/review/boards/edit", "Maintainer", 0),
            ApiUsageDisplayTests.Route("POST", "/api/submissions", "Submissions", 12, new ApiUsageVersion("3.0.0", 12, ApiUsageDisplayTests.Last)),
            ApiUsageDisplayTests.Route("GET", "/api/health", "Forever", 3, new ApiUsageVersion(ApiUsageVersion.NotCrt, 3, ApiUsageDisplayTests.Last)),
        ],
        []);

    // Submit, then this tab, then what is never retired - each least-called first.
    [Fact]
    public void The_groups_come_in_order_with_the_least_called_route_first()
    {
        IReadOnlyList<ApiUsageGroup> groups = ApiUsageDisplay.Groups(ApiUsageDisplayTests.Answer());

        Assert.Equal(
            [ApiUsageDisplay.SubmissionsHeading, ApiUsageDisplay.MaintainerHeading, ApiUsageDisplay.ForeverHeading],
            groups.Select(group => group.Heading));

        Assert.Equal(["POST /api/review/boards/edit", "GET /api/review/queue"], groups[1].Routes.Select(route => route.Title));
        Assert.Equal(["GET /api/health", "POST /api/usage/check-in"], groups[2].Routes.Select(route => route.Title));
    }

    [Fact]
    public void A_route_nobody_called_says_so_for_the_days_shown()
    {
        ApiUsageRouteLines unused = ApiUsageDisplay.Groups(ApiUsageDisplayTests.Answer())[1].Routes[0];

        Assert.Equal("No calls in the last [90] days", ApiUsageDisplayTests.Text(unused.Summary));
        Assert.Empty(unused.Versions);
    }

    // A called route: its total, its last day, and a line per version - the count bold.
    [Fact]
    public void A_called_route_gives_its_total_and_a_line_per_version()
    {
        ApiUsageRouteLines queue = ApiUsageDisplay.Groups(ApiUsageDisplayTests.Answer())[1].Routes[1];
        string day = SubmissionReceiptPresenter.FormatDate(ApiUsageDisplayTests.Last);

        Assert.Equal($"[1,500] calls, the last on {day}", ApiUsageDisplayTests.Text(queue.Summary));
        Assert.Equal($"3.0.0: [1,500] calls, the last on {day}", ApiUsageDisplayTests.Text(Assert.Single(queue.Versions)));
        Assert.Equal("1,500", queue.Summary.Single(run => run.IsCount).Text);
    }

    // The server's labels for "no CRT version" and "too many versions" are never shown raw.
    [Fact]
    public void A_request_naming_no_CRT_reads_as_such()
    {
        ApiUsageRouteLines health = ApiUsageDisplay.Groups(ApiUsageDisplayTests.Answer())[2].Routes[0];

        Assert.StartsWith(ApiUsageDisplay.NotCrtLabel + ": [3] calls", ApiUsageDisplayTests.Text(Assert.Single(health.Versions)), StringComparison.Ordinal);
        Assert.Equal(ApiUsageDisplay.OtherLabel, ApiUsageDisplay.VersionLabel(ApiUsageVersion.Other));
        Assert.Equal("3.1.0-beta.2", ApiUsageDisplay.VersionLabel("3.1.0-beta.2"));
    }

    // A newer server's area this version has never heard of is still listed, under its own name.
    [Fact]
    public void A_route_in_an_area_this_version_does_not_know_is_still_shown()
    {
        var answer = new ApiUsageAnswer(30, [ApiUsageDisplayTests.Route("GET", "/api/future/thing", "Future", 1)], []);

        ApiUsageGroup group = Assert.Single(ApiUsageDisplay.Groups(answer));

        Assert.Equal("Future", group.Heading);
        Assert.Equal("GET /api/future/thing", Assert.Single(group.Routes).Title);
    }

    [Fact]
    public void The_launches_are_a_line_per_version_or_one_line_saying_there_were_none()
    {
        var answer = new ApiUsageAnswer(90, [], [new ApiUsageInstallations("3.0.0", 42, 305), new ApiUsageInstallations("2.5.0", 1, 1)]);

        Assert.Equal(
            ["3.0.0: [42] installations, [305] launches", "2.5.0: [1] installation, [1] launch"],
            ApiUsageDisplay.InstallationLines(answer).Select(ApiUsageDisplayTests.Text));

        Assert.Equal(
            ["No CRT checked in at launch in the last [30] days"],
            ApiUsageDisplay.InstallationLines(new ApiUsageAnswer(30, [], [])).Select(ApiUsageDisplayTests.Text));
    }

    [Fact]
    public void Ninety_days_is_offered_and_shown_first()
    {
        Assert.Contains(ApiUsageDisplay.DefaultDays, ApiUsageDisplay.DayChoices);
        Assert.Equal("The last 90 days", ApiUsageDisplay.DayChoiceLabel(90));
        Assert.Equal("The last year", ApiUsageDisplay.DayChoiceLabel(365));
    }
}
