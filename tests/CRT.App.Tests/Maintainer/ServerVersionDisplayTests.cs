using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// "Account" > Server version (owner requests, 2026-10-04: "I would like to see the
// server version listed, so it is clear to me what has been deployed", then worded as three lines -
// "Server version [4.6.0]", "API server version [1]", "API application version [1]" - each value
// bold).
// ###########################################################################################
public sealed class ServerVersionDisplayTests
{
    private static string[] Texts(IReadOnlyList<IReadOnlyList<ReviewNoteRun>> lines) =>
        lines.Select(line => string.Concat(line.Select(run => run.Text))).ToArray();

    [Fact]
    public void The_three_lines_say_the_server_version_and_both_api_revisions()
    {
        IReadOnlyList<IReadOnlyList<ReviewNoteRun>> lines = ServerVersionDisplay.Lines(" 4.6.0 ", 1, 2);

        Assert.Equal(
            ["Server version [4.6.0]", "API server version [1]", "API application version [2]"],
            ServerVersionDisplayTests.Texts(lines));
    }

    // Only the VALUE is bold - not the label, and not the brackets around it.
    [Fact]
    public void Only_each_value_is_bold()
    {
        IReadOnlyList<IReadOnlyList<ReviewNoteRun>> lines = ServerVersionDisplay.Lines("4.6.0", 1, 1);

        Assert.All(lines, line => Assert.Equal(
            [false, true, false],
            line.Select(run => run.IsCount)));

        Assert.Equal(["4.6.0", "1", "1"], lines.Select(line => line.Single(run => run.IsCount).Text));
    }

    // A server older than 4.6.0 reports no API revision - said, never a bracket with nothing in it.
    [Fact]
    public void A_server_reporting_no_api_revision_is_said_to_be_older()
    {
        Assert.Equal(
            ["Server version [4.5.1]", "API server version: not reported - the server is older than 4.6.0", "API application version [1]"],
            ServerVersionDisplayTests.Texts(ServerVersionDisplay.Lines("4.5.1", null, 1)));
    }

    // No answer is SAID, never a blank line - which would read as nothing deployed. The
    // application's own revision is known either way.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void No_answer_says_the_server_did_not_answer(string? version)
    {
        Assert.Equal(
            ["Server version: the server did not answer", "API application version [1]"],
            ServerVersionDisplayTests.Texts(ServerVersionDisplay.Lines(version, 1, 1)));
    }

    [Fact]
    public void While_asking_it_says_so()
    {
        Assert.Equal(
            ["Server version: asking the server", "API application version [1]"],
            ServerVersionDisplayTests.Texts(ServerVersionDisplay.Lines(null, null, 1, asking: true)));
    }

    // ###########################################################################################
    // *** MAILS THE SERVER COULD NOT SEND (code review, 2026-10-13). *** A last line when there are
    // any, nothing when there are none or the server did not say - a postfix that is down is
    // something to look at, and "0" every day would be read past.
    // ###########################################################################################
    [Fact]
    public void Mails_the_server_could_not_send_are_a_last_line_only_when_there_are_some()
    {
        Assert.Equal(
            ["Server version [1.1.0]", "API server version [1]", "API application version [1]", "Mails the server could not send yet [1,234]"],
            ServerVersionDisplayTests.Texts(ServerVersionDisplay.Lines("1.1.0", 1, 1, unsentMails: 1234)));

        Assert.Equal(3, ServerVersionDisplay.Lines("1.1.0", 1, 1, unsentMails: 0).Count);
        Assert.Equal(3, ServerVersionDisplay.Lines("1.1.0", 1, 1, unsentMails: null).Count);
    }

    // ###########################################################################################
    // *** HOW THE SERVER KEEPS ITSELF UP TO DATE (owner request, 2026-11-09: "it could show the status in
    // 'Maintainer' => 'Account' => 'Server version'"). *** After the version lines: which builds it installs,
    // its last look with the updater's own sentence, and the last build it installed - the values bold.
    // Times in local time, so these are made in local time and read the same on every machine.
    // ###########################################################################################
    private static readonly DateTimeOffset Looked = new(new DateTime(2026, 11, 9, 14, 20, 0, DateTimeKind.Local));

    private static readonly DateTimeOffset Installed = new(new DateTime(2026, 11, 9, 14, 5, 0, DateTimeKind.Local));

    [Fact]
    public void The_updates_say_the_channel_the_last_look_and_the_last_build_installed()
    {
        var updates = new Handlers.DataHandling.ServerUpdatesAnswer(
            SetUp: true,
            Channel: "releases",
            CheckedUtc: ServerVersionDisplayTests.Looked,
            Outcome: "installed",
            Message: "Installed 10.3.0 (CRT release [4.1.0]) in place of 10.2.1.",
            NewestVersion: "10.3.0",
            InstalledUtc: ServerVersionDisplayTests.Installed,
            InstalledVersion: "10.3.0",
            InstalledFromVersion: "10.2.1");

        IReadOnlyList<IReadOnlyList<ReviewNoteRun>> lines =
            ServerVersionDisplay.Lines("10.3.0", 2, 2, updates: updates, now: ServerVersionDisplayTests.Looked.AddMinutes(5));

        Assert.Equal(
            [
                "Server version [10.3.0]",
                "API server version [2]",
                "API application version [2]",
                "Automatic updates [releases only]",
                "Last look [2026-November-9 at 14:20]: Installed 10.3.0 (CRT release [4.1.0]) in place of 10.2.1.",
                "Last installed [10.3.0] on [2026-November-9 at 14:05], replacing [10.2.1]",
            ],
            ServerVersionDisplayTests.Texts(lines));

        // The values bold, never the updater's sentence.
        Assert.Equal(["releases only"], lines[3].Where(run => run.IsCount).Select(run => run.Text));
        Assert.Equal(["2026-November-9 at 14:20"], lines[4].Where(run => run.IsCount).Select(run => run.Text));
        Assert.Equal(["10.3.0", "2026-November-9 at 14:05", "10.2.1"], lines[5].Where(run => run.IsCount).Select(run => run.Text));
    }

    // The test instance takes every commit to main (owner decision, 2026-11-09: "ALPHA/BETA works on commit").
    [Fact]
    public void The_test_instance_says_it_installs_every_commit_to_main() =>
        Assert.Contains(
            "Automatic updates [every commit to main]",
            ServerVersionDisplayTests.Texts(ServerVersionDisplay.Lines(
                "10.3.0", 2, 2,
                updates: new Handlers.DataHandling.ServerUpdatesAnswer(SetUp: true, Channel: "main"))));

    // The timer looks every minute (owner request, 2026-11-09): a look over ten minutes old means it may have stopped.
    [Fact]
    public void A_look_over_ten_minutes_old_says_the_timer_may_not_be_running()
    {
        var updates = new Handlers.DataHandling.ServerUpdatesAnswer(SetUp: true, Channel: "releases", CheckedUtc: ServerVersionDisplayTests.Looked, Message: "Up to date.");

        string[] fresh = ServerVersionDisplayTests.Texts(ServerVersionDisplay.Lines("10.3.0", 2, 2, updates: updates, now: ServerVersionDisplayTests.Looked.AddMinutes(9)));
        string[] stale = ServerVersionDisplayTests.Texts(ServerVersionDisplay.Lines("10.3.0", 2, 2, updates: updates, now: ServerVersionDisplayTests.Looked.AddMinutes(11)));

        Assert.Contains("Last look [2026-November-9 at 14:20]: Up to date.", fresh);
        Assert.Contains("Last look [2026-November-9 at 14:20] - over ten minutes ago, so its timer may not be running: Up to date.", stale);
    }

    [Fact]
    public void A_server_with_no_updater_says_it_is_not_set_up_and_an_older_server_says_nothing()
    {
        Assert.Equal(
            "Automatic updates [not set up on this server]",
            ServerVersionDisplayTests.Texts(ServerVersionDisplay.Lines("10.3.0", 2, 2, updates: new Handlers.DataHandling.ServerUpdatesAnswer(SetUp: false))).Last());

        Assert.Equal(3, ServerVersionDisplay.Lines("10.2.1", 2, 2, updates: null).Count);
    }
}
