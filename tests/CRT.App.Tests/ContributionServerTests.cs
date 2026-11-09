using System.Reflection;
using CRT;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Which service a CRT talks to - decided by its version since 2026-11-09 (owner decision: "if the
// <InformationalVersion> is not a 'real' version, then the application will refer to the
// 'crt-server-test' ... If it is a real version, then it should refer to the 'crt-server'"): an alpha
// or a beta the contribution service's TEST instance, a release the live one (ContributionServer.RootUrlFor,
// AppConfig.ContributionServerRootUrl). It was a constant until then, set back to the live service by hand
// on release day, with a test failing a release that still named the test instance.
// ###########################################################################################
public sealed class ContributionServerTests
{
    private const string Live = "https://classic-repair-toolbox.dk";

    private const string Test = "https://classic-repair-toolbox.dk/crt-test";

    [Theory]
    [InlineData("4.0.0", ContributionServerTests.Live)]
    [InlineData("4.0.0+7f3a2c1", ContributionServerTests.Live)]
    [InlineData(" 10.2.3 ", ContributionServerTests.Live)]
    [InlineData("4.0.0-alpha.1", ContributionServerTests.Test)]
    [InlineData("4.0.0-beta.2+7f3a2c1", ContributionServerTests.Test)]
    [InlineData("4.1.0-rc.1", ContributionServerTests.Test)]
    // No version at all is no release: never the live service.
    [InlineData("", ContributionServerTests.Test)]
    [InlineData(null, ContributionServerTests.Test)]
    // Build metadata may hold a "-" ("+build-7"): only what stands before the "+" counts.
    [InlineData("4.0.0+build-7", ContributionServerTests.Live)]
    public void A_release_talks_to_the_live_service_and_an_alpha_or_beta_to_the_test_instance(string? version, string expected)
    {
        Assert.Equal(expected, ContributionServer.RootUrlFor(version));
        Assert.Equal(expected == ContributionServerTests.Live, ContributionServer.IsRelease(version));
    }

    // ###########################################################################################
    // *** THE DATA FOLLOWS THE SAME RULE (owner decisions, 2026-11-09: "we need to have a new data source for the
    // test server, so it does not mix data with the real server" - an alpha or a beta downloads "The test
    // server's"). *** A release downloads the live 4.0 trees, an alpha or a beta the test instance's own; the
    // BETA source of each beside it. A pre-release never reads the live trees, a release never the test ones.
    // ###########################################################################################
    [Theory]
    [InlineData("4.0.0", false, "https://classic-repair-toolbox.dk/app-data-v4/dataChecksums.json")]
    [InlineData("4.0.0", true, "https://classic-repair-toolbox.dk/app-data-v4-BETA/dataChecksums.json")]
    [InlineData("4.0.0-alpha.5", false, "https://classic-repair-toolbox.dk/app-data-v4-test/dataChecksums.json")]
    [InlineData("4.1.0-beta.1+7f3a2c1", true, "https://classic-repair-toolbox.dk/app-data-v4-test-BETA/dataChecksums.json")]
    [InlineData(null, false, "https://classic-repair-toolbox.dk/app-data-v4-test/dataChecksums.json")]
    public void A_release_downloads_the_live_data_and_an_alpha_or_beta_the_test_servers(string? version, bool betaSource, string expected) =>
        Assert.Equal(expected, ContributionServer.ChecksumsUrlFor(version, betaSource));

    [Fact]
    public void This_build_downloads_the_data_its_own_version_names()
    {
        Assert.Equal(ContributionServer.ChecksumsUrlFor(AppConfig.AppDisplayVersionString, betaSource: false), AppConfig.ChecksumsUrl);
        Assert.Equal(ContributionServer.ChecksumsUrlFor(AppConfig.AppDisplayVersionString, betaSource: true), AppConfig.ChecksumsUrl_test);
    }

    // The address this very build uses follows the rule - read at each use, never fixed before the version is known
    // (a static field declared above AppDisplayVersionString would read it as null: every release on the test instance).
    [Fact]
    public void This_build_talks_to_the_service_its_own_version_names()
    {
        Assert.Equal(ContributionServer.RootUrlFor(AppConfig.AppDisplayVersionString), AppConfig.ContributionServerRootUrl);
        Assert.Equal(!ContributionServer.IsRelease(AppConfig.AppDisplayVersionString), AppConfig.UsesTestServer);
        Assert.Equal(ContributionServerTests.Test, AppConfig.TestServerRootUrl);
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
    // A draft whose work reached the test instance's stable data was deleted - and that data never
    // reaches the live data, so the work was gone everywhere.
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

    // Every pre-release talks to the test instance now, not CRT 4.0's alone: no note names one version.
    [Fact]
    public void The_test_server_notes_name_no_version_of_CRT()
    {
        Assert.DoesNotContain("4.0", ContributionServer.SubmitNote, StringComparison.Ordinal);
        Assert.DoesNotContain("4.0", ContributionServer.SignInNote, StringComparison.Ordinal);
    }
}
