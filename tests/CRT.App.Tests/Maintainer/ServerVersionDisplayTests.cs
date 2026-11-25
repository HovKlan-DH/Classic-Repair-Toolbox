using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// "Account" > Server version (owner requests, 2026-10-04: "I would like to see the server version listed, so it is
// clear to me what has been deployed"; 2026-11-24: "make the "Account" page a little prettier to look at, so it is
// more easy to overview" - of this entry). A panel per thing the server is asked about, each with a line saying how
// it stands, in words and as a state its edge and icon draw, and its facts with each value bold.
// ###########################################################################################
public sealed class ServerVersionDisplayTests
{
    private static string[] Lines(IReadOnlyList<ServerPanel> panels) =>
        panels.SelectMany(ServerVersionDisplay.AsLines).ToArray();

    [Fact]
    public void The_server_panel_says_the_version_and_both_api_revisions()
    {
        IReadOnlyList<ServerPanel> panels = ServerVersionDisplay.Panels(" 4.6.0 ", 1, 1);

        Assert.Equal(
            ["Server: This CRT and the server speak the same API", "Version: 4.6.0", "API revision: 1", "This CRT's API revision: 1"],
            ServerVersionDisplayTests.Lines(panels));
        Assert.Equal(ServerPanelState.Good, panels.Single().State);
    }

    // Only the VALUE is bold - never the label.
    [Fact]
    public void Only_each_value_is_bold()
    {
        ServerPanel panel = ServerVersionDisplay.Panels("4.6.0", 1, 1).Single();

        Assert.All(panel.Rows, row => Assert.All(row.Value, run => Assert.True(run.IsCount)));
        Assert.Equal(["4.6.0", "1", "1"], panel.Rows.Select(row => row.Value.Single().Text));
    }

    // ###########################################################################################
    // What the two API revisions mean together is SAID, so nobody has to know what one is (it was a paragraph over
    // the lines until 2026-11-24): the server's higher turns this CRT away (ClientVersionPolicy) - a problem; this
    // CRT's higher is a server not deployed yet - something to wait for.
    // ###########################################################################################
    [Fact]
    public void A_server_on_a_higher_api_revision_says_this_crt_is_behind()
    {
        ServerPanel panel = ServerVersionDisplay.Panels("12.2.0", 3, 2).Single();

        Assert.Equal(ServerPanelState.Problem, panel.State);
        Assert.Equal(ServerVersionDisplay.CrtBehind, panel.Summary);
    }

    [Fact]
    public void A_server_on_a_lower_api_revision_says_it_is_not_deployed_yet()
    {
        ServerPanel panel = ServerVersionDisplay.Panels("12.2.0", 2, 3).Single();

        Assert.Equal(ServerPanelState.Waiting, panel.State);
        Assert.Equal(ServerVersionDisplay.ServerBehind, panel.Summary);
    }

    // A server older than 4.6.0 reports no API revision - said, never a value with nothing in it.
    [Fact]
    public void A_server_reporting_no_api_revision_is_said_to_be_older()
    {
        IReadOnlyList<ServerPanel> panels = ServerVersionDisplay.Panels("4.5.1", null, 1);

        Assert.Equal(
            ["Server: The server is older than 4.6.0 and does not say its API revision", "Version: 4.5.1", "This CRT's API revision: 1"],
            ServerVersionDisplayTests.Lines(panels));
        Assert.Equal(ServerPanelState.Neutral, panels.Single().State);
    }

    // No answer is SAID, never a blank panel - which would read as nothing deployed. The application's own revision
    // is known either way.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void No_answer_says_the_server_did_not_answer(string? version)
    {
        IReadOnlyList<ServerPanel> panels = ServerVersionDisplay.Panels(version, 1, 1);

        Assert.Equal(["Server: The server did not answer", "This CRT's API revision: 1"], ServerVersionDisplayTests.Lines(panels));
        Assert.Equal(ServerPanelState.Problem, panels.Single().State);
    }

    [Fact]
    public void While_asking_it_says_so()
    {
        IReadOnlyList<ServerPanel> panels = ServerVersionDisplay.Panels(null, null, 1, asking: true);

        Assert.Equal(["Server: Asking the server", "This CRT's API revision: 1"], ServerVersionDisplayTests.Lines(panels));
        Assert.Equal(ServerPanelState.Neutral, panels.Single().State);
    }

    // ###########################################################################################
    // *** MAILS THE SERVER COULD NOT SEND (code review, 2026-10-13). *** A panel when there are any, nothing when
    // there are none or the server did not say - a postfix that is down is something to look at, and "0" every day
    // would be read past.
    // ###########################################################################################
    [Fact]
    public void Mails_the_server_could_not_send_are_a_panel_only_when_there_are_some()
    {
        ServerPanel mail = ServerVersionDisplay.Panels("1.1.0", 1, 1, unsentMails: 1234).Last();

        Assert.Equal(["Mail: Mails the server could not send yet - it keeps trying", "Not sent yet: 1,234"], ServerVersionDisplay.AsLines(mail));
        Assert.Equal(ServerPanelState.Problem, mail.State);

        Assert.Single(ServerVersionDisplay.Panels("1.1.0", 1, 1, unsentMails: 0));
        Assert.Single(ServerVersionDisplay.Panels("1.1.0", 1, 1, unsentMails: null));
    }

    // ###########################################################################################
    // *** HOW THE SERVER KEEPS ITSELF UP TO DATE (owner request, 2026-11-09: "it could show the status in
    // 'Maintainer' => 'Account' => 'Server version'"). *** A panel after the server's: which builds it installs, its
    // last look, the updater's own sentence, and the last build it installed - the values bold, never the sentence.
    // Times in local time, so these are made in local time and read the same on every machine.
    // ###########################################################################################
    private static readonly DateTimeOffset Looked = new(new DateTime(2026, 11, 9, 14, 20, 0, DateTimeKind.Local));

    private static readonly DateTimeOffset Installed = new(new DateTime(2026, 11, 9, 14, 5, 0, DateTimeKind.Local));

    [Fact]
    public void The_updates_say_the_channel_the_last_look_and_the_last_build_installed()
    {
        var updates = new ServerUpdatesAnswer(
            SetUp: true,
            Channel: "releases",
            CheckedUtc: ServerVersionDisplayTests.Looked,
            Outcome: "installed",
            Message: "Installed 10.3.0 (CRT release [4.1.0]) in place of 10.2.1.",
            NewestVersion: "10.3.0",
            InstalledUtc: ServerVersionDisplayTests.Installed,
            InstalledVersion: "10.3.0",
            InstalledFromVersion: "10.2.1");

        ServerPanel panel = ServerVersionDisplay.Panels("10.3.0", 2, 2, updates: updates, now: ServerVersionDisplayTests.Looked.AddMinutes(5))[1];

        Assert.Equal(
            [
                "Automatic updates: Up to date - the newest build was installed",
                "Installs: releases only",
                "Last look: 2026-November-9 at 14:20",
                "Found: Installed 10.3.0 (CRT release [4.1.0]) in place of 10.2.1.",
                "Last installed: 10.3.0 on 2026-November-9 at 14:05, replacing 10.2.1",
            ],
            ServerVersionDisplay.AsLines(panel));
        Assert.Equal(ServerPanelState.Good, panel.State);

        // The values bold, never the updater's sentence.
        Assert.Equal(["releases only"], panel.Rows[0].Value.Where(run => run.IsCount).Select(run => run.Text));
        Assert.Equal(["2026-November-9 at 14:20"], panel.Rows[1].Value.Where(run => run.IsCount).Select(run => run.Text));
        Assert.DoesNotContain(panel.Rows[2].Value, run => run.IsCount);
        Assert.Equal(["10.3.0", "2026-November-9 at 14:05", "10.2.1"], panel.Rows[3].Value.Where(run => run.IsCount).Select(run => run.Text));
    }

    // The test instance takes every commit to main (owner decision, 2026-11-09: "ALPHA/BETA works on commit").
    [Fact]
    public void The_test_instance_says_it_installs_every_commit_to_main() =>
        Assert.Contains(
            "Installs: every commit to main",
            ServerVersionDisplayTests.Lines(ServerVersionDisplay.Panels("10.3.0", 2, 2, updates: new ServerUpdatesAnswer(SetUp: true, Channel: "main"))));

    // Each outcome the updater writes says how it stands in words, and as a state: what needs a look is never good
    // news (ServerUpdateOutcomes.NeedsALook); a word CRT does not know is only "Set up".
    [Theory]
    [InlineData("up-to-date", ServerPanelState.Good, ServerVersionDisplay.UpToDate)]
    [InlineData("installed", ServerPanelState.Good, ServerVersionDisplay.JustInstalled)]
    [InlineData("would-install", ServerPanelState.Waiting, ServerVersionDisplay.WouldInstall)]
    [InlineData("held", ServerPanelState.Waiting, ServerVersionDisplay.Held)]
    [InlineData("failed", ServerPanelState.Problem, ServerVersionDisplay.Failed)]
    [InlineData("refused", ServerPanelState.Problem, ServerVersionDisplay.Refused)]
    [InlineData("error", ServerPanelState.Problem, ServerVersionDisplay.LookFailed)]
    [InlineData("something-new", ServerPanelState.Neutral, ServerVersionDisplay.SetUp)]
    [InlineData(null, ServerPanelState.Neutral, ServerVersionDisplay.SetUp)]
    public void Each_outcome_says_how_the_updates_stand(string? outcome, ServerPanelState state, string summary)
    {
        var updates = new ServerUpdatesAnswer(SetUp: true, Channel: "releases", CheckedUtc: ServerVersionDisplayTests.Looked, Outcome: outcome);

        ServerPanel panel = ServerVersionDisplay.Panels("10.3.0", 2, 2, updates: updates, now: ServerVersionDisplayTests.Looked.AddMinutes(1))[1];

        Assert.Equal((state, summary), (panel.State, panel.Summary));

        if (ServerUpdateOutcomes.NeedsALook(outcome))
            Assert.NotEqual(ServerPanelState.Good, panel.State);
    }

    // The timer looks every minute (owner request, 2026-11-09): a look over ten minutes old means it may have
    // stopped - said on the look itself and as the panel's line, over good news but never over a problem.
    [Fact]
    public void A_look_over_ten_minutes_old_says_the_timer_may_not_be_running()
    {
        var updates = new ServerUpdatesAnswer(SetUp: true, Channel: "releases", CheckedUtc: ServerVersionDisplayTests.Looked, Outcome: "up-to-date", Message: "Up to date.");

        ServerPanel fresh = ServerVersionDisplay.Panels("10.3.0", 2, 2, updates: updates, now: ServerVersionDisplayTests.Looked.AddMinutes(9))[1];
        ServerPanel stale = ServerVersionDisplay.Panels("10.3.0", 2, 2, updates: updates, now: ServerVersionDisplayTests.Looked.AddMinutes(11))[1];

        Assert.Contains("Last look: 2026-November-9 at 14:20", ServerVersionDisplay.AsLines(fresh));
        Assert.Equal(ServerPanelState.Good, fresh.State);

        Assert.Contains("Last look: 2026-November-9 at 14:20 - over ten minutes ago", ServerVersionDisplay.AsLines(stale));
        Assert.Equal((ServerPanelState.Waiting, ServerVersionDisplay.TimerMayHaveStopped), (stale.State, stale.Summary));

        ServerPanel staleFailure = ServerVersionDisplay.Panels(
            "10.3.0", 2, 2,
            updates: updates with { Outcome = "failed" },
            now: ServerVersionDisplayTests.Looked.AddMinutes(11))[1];

        Assert.Equal((ServerPanelState.Problem, ServerVersionDisplay.Failed), (staleFailure.State, staleFailure.Summary));
    }

    [Fact]
    public void A_server_with_no_updater_says_it_is_not_set_up_and_an_older_server_says_nothing()
    {
        ServerPanel notSetUp = ServerVersionDisplay.Panels("10.3.0", 2, 2, updates: new ServerUpdatesAnswer(SetUp: false))[1];

        Assert.Equal(["Automatic updates: Not set up on this server"], ServerVersionDisplay.AsLines(notSetUp));
        Assert.Equal(ServerPanelState.Neutral, notSetUp.State);

        Assert.Single(ServerVersionDisplay.Panels("10.2.1", 2, 2, updates: null));
    }
}
