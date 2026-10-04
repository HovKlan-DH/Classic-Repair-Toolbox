using CRT;
using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The launch check-in against the server that receives it (2026-10-03, when the check-in moved
// from the old app-checkin PHP page to CRT.Server). The address CRT posts to is the route the
// server maps (RequestBodyLimitsTests pins the same path on the server's real route table), and the
// version CRT sends as its User-Agent must be one the server stores (CheckInRules.VersionFrom
// refuses one without "CRT " or with a character outside letters, digits and  ,.#()*[]!:/- - a
// "+commit" suffix, say, would silently stop every check-in counting).
// ###########################################################################################
public sealed class LaunchCheckInTests
{
    [Fact]
    public void The_check_in_goes_to_the_servers_check_in_route()
    {
        Assert.Equal("https://classic-repair-toolbox.dk/api/usage/check-in", AppConfig.CheckInUrl);
    }

    [Fact]
    public void The_version_CRT_checks_in_with_is_one_the_server_stores()
    {
        Assert.Matches(@"^CRT [0-9A-Za-z.\-]+$", OnlineServices.VersionForServer);
    }
}
