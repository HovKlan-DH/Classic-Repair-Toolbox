using Handlers.DataHandling;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Which of the Drafts and Maintainer tabs says "CRT has to be updated", and in which words (owner
// request, 2026-10-09: "When an API diff requires for the CRT app to be updated, then for both the
// tabs ... put a fullpage modal (or alike) in there, that cannot be closed").
//
// The server's API revision covers BOTH tabs; a 426 from one area covers only that area's tab,
// since the server's minimum-version lever is set per area. Nothing ever uncovers a tab.
// ###########################################################################################
public sealed class AppUpdateRequirementTests
{
    private static readonly CrtVersion Own = CrtVersion.Parse("3.0.0");

    // The "API diff": only a server serving a NEWER revision turns this CRT away. The server serves
    // its own revision and newer, so an equal or lower one is fine - and a server older than 4.6.0
    // reports none, which is not a refusal.
    [Theory]
    [InlineData(3, 2, true)]
    [InlineData(2, 2, false)]
    [InlineData(1, 2, false)]
    [InlineData(null, 2, false)]
    public void Only_a_higher_server_revision_is_behind(int? server, int own, bool behind)
    {
        Assert.Equal(behind, AppUpdateRequirement.IsBehind(server, own));
    }

    // Somebody with no drafts and the Maintainer tab off is asked nothing (code review, 2026-10-09):
    // every launch sent the server a request it counts, for two tabs that installation never shows.
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void The_server_is_asked_only_while_either_tab_is_in_use(bool draftsTabShown, bool maintainerTabEnabled, bool asks)
    {
        Assert.Equal(asks, AppUpdateRequirement.AsksServer(draftsTabShown, maintainerTabEnabled));
    }

    [Fact]
    public void A_higher_server_revision_covers_both_tabs_in_the_servers_own_sentence()
    {
        var requirement = new AppUpdateRequirement();

        Assert.True(requirement.ApplyApiRevision(3, 2, Own));

        string expected = ClientVersionContract.OutdatedApi(Own).Message;

        Assert.Equal(expected, requirement.ReasonFor(AppUpdateArea.Drafts));
        Assert.Equal(expected, requirement.ReasonFor(AppUpdateArea.Maintainer));
        Assert.True(requirement.IsRequiredEverywhere);
        Assert.Contains("[3.0.0]", expected, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(1)]
    [InlineData(null)]
    public void A_server_that_still_serves_this_revision_covers_nothing(int? server)
    {
        var requirement = new AppUpdateRequirement();

        Assert.False(requirement.ApplyApiRevision(server, 2, Own));
        Assert.False(requirement.IsRequiredFor(AppUpdateArea.Drafts));
        Assert.False(requirement.IsRequiredFor(AppUpdateArea.Maintainer));
    }

    // A minimum version on the Maintainer area alone must not take the contributors' Drafts tab
    // with it - the refusal covers its own area only.
    [Fact]
    public void A_refusal_covers_only_the_tab_of_the_area_that_refused()
    {
        var requirement = new AppUpdateRequirement();
        string words = ClientVersionContract.Outdated(Own, CrtVersion.Parse("3.2.0")).Message;

        Assert.True(requirement.ApplyRefusal(AppUpdateArea.Maintainer, words, Own));

        Assert.Equal(words, requirement.ReasonFor(AppUpdateArea.Maintainer));
        Assert.False(requirement.IsRequiredFor(AppUpdateArea.Drafts));
        Assert.False(requirement.IsRequiredEverywhere);
    }

    // The server's words may name the version to update to, so they win over CRT's own copy - and
    // a later health answer does not put CRT's sentence back over them.
    [Fact]
    public void The_servers_words_replace_CRTs_own_and_stay()
    {
        var requirement = new AppUpdateRequirement();
        string words = ClientVersionContract.Outdated(Own, CrtVersion.Parse("3.2.0")).Message;

        requirement.ApplyApiRevision(3, 2, Own);

        Assert.True(requirement.ApplyRefusal(AppUpdateArea.Drafts, "  " + words + "  ", Own));
        Assert.False(requirement.ApplyApiRevision(3, 2, Own));

        Assert.Equal(words, requirement.ReasonFor(AppUpdateArea.Drafts));
        Assert.Equal(ClientVersionContract.OutdatedApi(Own).Message, requirement.ReasonFor(AppUpdateArea.Maintainer));
    }

    [Fact]
    public void The_same_refusal_twice_changes_nothing_the_second_time()
    {
        var requirement = new AppUpdateRequirement();

        Assert.True(requirement.ApplyRefusal(AppUpdateArea.Drafts, "Please update CRT.", Own));
        Assert.False(requirement.ApplyRefusal(AppUpdateArea.Drafts, "Please update CRT.", Own));
    }

    // A refusal with no words still covers the tab - in CRT's own sentence - and never blanks out
    // words already there.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_refusal_without_words_covers_the_tab_in_CRTs_own_sentence(string? words)
    {
        var requirement = new AppUpdateRequirement();

        Assert.True(requirement.ApplyRefusal(AppUpdateArea.Maintainer, words, null));
        Assert.Equal(ClientVersionContract.OutdatedApi(null).Message, requirement.ReasonFor(AppUpdateArea.Maintainer));

        requirement.ApplyRefusal(AppUpdateArea.Drafts, "The server's sentence.", Own);

        Assert.False(requirement.ApplyRefusal(AppUpdateArea.Drafts, words, Own));
        Assert.Equal("The server's sentence.", requirement.ReasonFor(AppUpdateArea.Drafts));
    }

    // Nothing but a newer CRT changes the server's answer: a later "fine" health answer does not
    // uncover a tab.
    [Fact]
    public void A_covered_tab_is_never_uncovered()
    {
        var requirement = new AppUpdateRequirement();

        requirement.ApplyApiRevision(3, 2, Own);

        Assert.False(requirement.ApplyApiRevision(2, 2, Own));
        Assert.False(requirement.ApplyApiRevision(null, 2, Own));
        Assert.True(requirement.IsRequiredEverywhere);
    }
}
