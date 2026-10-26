using System.Reflection;
using CRT;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Which service a CRT talks to - the contribution service's TEST instance while CRT 4.0 is in
// pre-release, the live one otherwise (owner decision, 2026-10-25: "We need the test server";
// AppConfig.ContributionServerRootUrl, ContributionServer).
//
// *** THE TEST THAT MATTERS MOST is A_release_never_ships_talking_to_the_test_instance. *** On
// release day ContributionServerRootUrl goes back to the live service. Forgotten, every installed
// CRT 4.0.0 would submit to a test database the project owner stops that day - so a release version
// pointing at the test instance fails here, and the release workflow runs this suite before it
// builds anything.
// ###########################################################################################
public sealed class ContributionServerTests
{
    private const string Live = "https://classic-repair-toolbox.dk";

    [Fact]
    public void A_release_never_ships_talking_to_the_test_instance()
    {
        Assert.True(
            ContributionServer.MayShip(AppConfig.AppDisplayVersionString, AppConfig.ContributionServerRootUrl),
            $"CRT {AppConfig.AppDisplayVersionString} is a release, but AppConfig.ContributionServerRootUrl still names the test " +
            $"instance ({AppConfig.ContributionServerRootUrl}). Set it back to AppConfig.CrtServerRootUrl - see INSTALLING.md, " +
            "\"Switching the service to the 4.0 trees (release day)\".");
    }

    [Theory]
    [InlineData("4.0.0", false)]
    [InlineData("4.0.0+7f3a2c1", false)]
    [InlineData("4.0.0-alpha.1", true)]
    [InlineData("4.0.0-beta.2+7f3a2c1", true)]
    public void Only_a_pre_release_may_talk_to_the_test_instance(string version, bool mayShip)
    {
        Assert.Equal(mayShip, ContributionServer.MayShip(version, AppConfig.TestServerRootUrl));
        Assert.Equal(mayShip, ContributionServer.MayShip(version, AppConfig.TestServerRootUrl.ToUpperInvariant() + "/"));
        Assert.True(ContributionServer.MayShip(version, ContributionServerTests.Live));
    }

    // ###########################################################################################
    // *** ONLY THE CONTRIBUTION SERVICE MOVES. *** The launch check-in, feedback and board views go
    // to the live service from a pre-release too: they are the usage numbers (the Fun facts page
    // reads them from the live database) and the project owner's feedback mail.
    // ###########################################################################################
    [Fact]
    public void The_check_in_and_feedback_always_go_to_the_live_service()
    {
        Assert.Equal(ContributionServerTests.Live, AppConfig.CrtServerRootUrl);
        Assert.StartsWith(ContributionServerTests.Live + "/api/", AppConfig.CheckInUrl);
        Assert.StartsWith(ContributionServerTests.Live + "/api/", TabFeedback.FeedbackUrl);
    }

    // The submissions and the Maintainer tab talk to one service, the contribution service.
    [Fact]
    public void Submissions_and_the_Maintainer_tab_talk_to_the_contribution_service()
    {
        Assert.Equal(AppConfig.ContributionServerRootUrl, Handlers.MaintainerHandling.ReviewApiRoutes.DefaultBaseAddress);
        Assert.Equal(AppConfig.ContributionServerRootUrl + "/api", AppConfig.ContributionServerBaseUrl);

        var client = new SubmissionClient();
        string? baseUrl = typeof(SubmissionClient)
            .GetField("thisBaseUrl", BindingFlags.Instance | BindingFlags.NonPublic)!
            .GetValue(client) as string;

        Assert.Equal(AppConfig.ContributionServerBaseUrl, baseUrl);
    }

    [Theory]
    [InlineData("submissions.json", false, "submissions.json")]
    [InlineData("submissions.json", true, "submissions-test.json")]
    [InlineData("maintainer-session.json", true, "maintainer-session-test.json")]
    public void What_CRT_remembers_of_the_test_instance_is_kept_in_files_of_its_own(string fileName, bool testServer, string expected)
    {
        Assert.Equal(expected, ContributionServer.FileNameFor(fileName, testServer));
    }

    [Fact]
    public void The_files_follow_the_service_this_CRT_is_built_for()
    {
        Assert.Equal(
            ContributionServer.FileNameFor("submissions.json", AppConfig.UsesTestServer),
            ContributionServer.FileNameFor("submissions.json"));
    }

    // ###########################################################################################
    // *** A DRAFT IS NOT RETIRED AGAINST THE TEST INSTANCE'S DATA (code review, 2026-10-25). ***
    // A pre-release reads the test instance's stable data as stable; a draft whose work reached it
    // was deleted - and that data is replaced on release day, so the work was gone everywhere.
    // ###########################################################################################
    [Fact]
    public void Published_drafts_are_retired_only_against_the_live_service()
    {
        Assert.True(ContributionServer.RetiresPublishedDrafts(testServer: false));
        Assert.False(ContributionServer.RetiresPublishedDrafts(testServer: true));
    }

    // The Submit dialog says the draft stays, so a tester does not expect it to vanish.
    [Fact]
    public void The_submit_note_says_the_draft_is_kept()
    {
        Assert.Contains("Your draft is kept", ContributionServer.SubmitNote, StringComparison.Ordinal);
    }
}
