using System.Net.Sockets;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests;

// The hint under a failed oscilloscope connect. A macOS user whose CRT is not allowed on the local
// network sees only "No route to host" - which reads like a network fault, while the scope answers
// ping and nc from Terminal (reported 2026-10-12).
public class ScopeConnectionHintTests
{
    private const string ScopeAddress = "192.168.1.50";

    private static SocketException Error(SocketError code) => new((int)code);

    // -------------------------------------------------------- ForFailure

    [Fact]
    public void No_route_to_host_on_macOS_points_at_the_local_network_permission()
    {
        string? hint = ScopeConnectionHint.ForFailure(
            Error(SocketError.HostUnreachable), ScopeAddress, isMacOS: true, runsFromAppBundle: true);

        Assert.Equal(ScopeConnectionHint.MacLocalNetwork, hint);
    }

    [Fact]
    public void The_hint_covers_a_switch_that_is_already_on()
    {
        // The report: "Local Network" was on for CRT and it still failed. A hint that only says
        // "switch it on" leaves that user with nothing to try.
        Assert.Contains("if it is on already, switch it off and on again", ScopeConnectionHint.MacLocalNetwork);
    }

    [Fact]
    public void The_hint_asks_for_the_check_that_tells_a_permission_from_a_network_fault()
    {
        // "No route to host" is also what a scope on another subnet or behind a VPN gives.
        Assert.StartsWith("On macOS, if the scope answers ping from Terminal,", ScopeConnectionHint.MacLocalNetwork);
        Assert.StartsWith("On macOS, if the scope answers ping from Terminal,", ScopeConnectionHint.MacLocalNetworkOutsideAppBundle);
    }

    [Fact]
    public void A_CRT_started_outside_its_app_bundle_is_pointed_at_the_app_that_started_it()
    {
        // A build from source run from Terminal or an editor is not in the Local Network list -
        // the permission is the starting app's.
        string? hint = ScopeConnectionHint.ForFailure(
            Error(SocketError.HostUnreachable), ScopeAddress, isMacOS: true, runsFromAppBundle: false);

        Assert.Equal(ScopeConnectionHint.MacLocalNetworkOutsideAppBundle, hint);
        Assert.DoesNotContain("Classic Repair Toolbox", hint);
    }

    [Theory]
    [InlineData("scope.local")]
    [InlineData("Scope.LOCAL")]
    [InlineData("scope.local.")]
    [InlineData(" scope.local ")]
    public void A_local_name_that_cannot_be_resolved_on_macOS_gets_the_hint(string host)
    {
        // Resolving a ".local" name is a local network lookup too, so a denied permission fails
        // before any connection is tried.
        string? hint = ScopeConnectionHint.ForFailure(
            Error(SocketError.HostNotFound), host, isMacOS: true, runsFromAppBundle: true);

        Assert.Equal(ScopeConnectionHint.MacLocalNetwork, hint);
    }

    [Theory]
    [InlineData("scope.example.com")]
    [InlineData("localscope")]
    [InlineData("scope.locale")]
    [InlineData(ScopeAddress)]
    public void An_unknown_name_that_is_not_a_local_name_gets_no_hint(string host)
    {
        // Ordinary DNS is not behind the permission - a typo in the name is the likely cause.
        Assert.Null(ScopeConnectionHint.ForFailure(
            Error(SocketError.HostNotFound), host, isMacOS: true, runsFromAppBundle: true));
    }

    [Theory]
    [InlineData(SocketError.HostUnreachable, ScopeAddress)]
    [InlineData(SocketError.HostNotFound, "scope.local")]
    public void Windows_and_Linux_get_no_hint(SocketError code, string host)
    {
        // There is no such permission there, so the hint would send the user looking for a
        // setting that does not exist.
        Assert.Null(ScopeConnectionHint.ForFailure(Error(code), host, isMacOS: false, runsFromAppBundle: false));
    }

    [Theory]
    [InlineData(SocketError.ConnectionRefused)]
    [InlineData(SocketError.TimedOut)]
    [InlineData(SocketError.NetworkUnreachable)]
    public void Other_socket_errors_on_macOS_get_no_hint(SocketError code)
    {
        // "Connection refused" in particular means the scope answered - the permission is fine.
        Assert.Null(ScopeConnectionHint.ForFailure(Error(code), ScopeAddress, isMacOS: true, runsFromAppBundle: true));
    }

    [Fact]
    public void A_failure_that_is_not_a_socket_error_gets_no_hint()
    {
        var error = new InvalidOperationException("Unexpected end of SCPI network stream");

        Assert.Null(ScopeConnectionHint.ForFailure(error, ScopeAddress, isMacOS: true, runsFromAppBundle: true));
    }

    [Fact]
    public void Both_hints_are_plain_ASCII_and_name_where_the_setting_is()
    {
        foreach (string hint in new[] { ScopeConnectionHint.MacLocalNetwork, ScopeConnectionHint.MacLocalNetworkOutsideAppBundle })
        {
            Assert.All(hint, c => Assert.True(c < 128, $"Not ASCII: '{c}'"));
            Assert.Contains("System Settings > Privacy & Security > Local Network", hint);
        }
    }

    // -------------------------------------------------------- IsInsideAppBundle

    [Theory]
    [InlineData("/Applications/Classic Repair Toolbox.app/Contents/MacOS/Classic-Repair-Toolbox", true)]
    [InlineData("/Users/me/Applications/Classic Repair Toolbox.APP/Contents/MacOS/Classic-Repair-Toolbox", true)]
    [InlineData("/Users/me/src/CRT.App/bin/Release/net10.0/Classic-Repair-Toolbox", false)]
    [InlineData("/Users/me/MacOS/Classic-Repair-Toolbox", false)]
    [InlineData(null, false)]
    public void Only_a_path_inside_an_app_bundle_counts_as_one(string? processPath, bool expected)
    {
        Assert.Equal(expected, ScopeConnectionHint.IsInsideAppBundle(processPath));
    }

    // -------------------------------------------------------- ShouldWrite

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void A_connect_the_user_pressed_always_writes_its_failure(bool hasHint, bool automaticHintWritten)
    {
        Assert.True(ScopeConnectionHint.ShouldWrite(isAutomaticAttempt: false, hasHint, automaticHintWritten));
    }

    [Fact]
    public void An_automatic_attempt_writes_a_failure_with_a_hint_once()
    {
        // With Auto-connect ticked, a macOS user who has not allowed CRT would otherwise only
        // ever see "Automatically connecting ..." while every attempt failed.
        Assert.True(ScopeConnectionHint.ShouldWrite(isAutomaticAttempt: true, hasHint: true, automaticHintWritten: false));
        Assert.False(ScopeConnectionHint.ShouldWrite(isAutomaticAttempt: true, hasHint: true, automaticHintWritten: true));
    }

    [Fact]
    public void An_automatic_attempt_without_a_hint_stays_silent()
    {
        // The retries are silent by design - a scope switched off is not worth a line a second.
        Assert.False(ScopeConnectionHint.ShouldWrite(isAutomaticAttempt: true, hasHint: false, automaticHintWritten: false));
    }
}
