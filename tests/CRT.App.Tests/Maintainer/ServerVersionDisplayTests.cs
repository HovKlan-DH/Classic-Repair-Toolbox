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
}
