using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// Which release versions the two "Allow notification for ALPHA/BETA versions" checkboxes let
// through. The stage has to be read off the version string because GitHub's own "prerelease" flag
// is all-or-nothing, so getting this classification wrong silently offers a user a channel they
// never asked for - nothing throws, the update banner just appears.
public sealed class UpdateChannelFilterTests
{
    [Theory]
    [InlineData("2.5.0")]
    [InlineData("2.5.0+abc123")]
    [InlineData("10.0.1")]
    public void A_version_with_no_pre_release_suffix_is_stable(string version)
    {
        Assert.Equal(UpdateChannelFilter.ReleaseStage.Stable, UpdateChannelFilter.GetStage(version));
    }

    // Build metadata is introduced by "+", not "-", so a "-" appearing only AFTER it is not a
    // pre-release marker. Velopack renders versions this way.
    [Fact]
    public void A_dash_inside_build_metadata_does_not_make_a_version_a_pre_release()
    {
        Assert.Equal(UpdateChannelFilter.ReleaseStage.Stable, UpdateChannelFilter.GetStage("2.5.0+abc-def"));
    }

    [Theory]
    [InlineData("2.6.0-beta.1")]
    [InlineData("2.6.0-BETA.2")]
    [InlineData("2.6.0-beta")]
    public void A_beta_suffix_is_recognised_case_insensitively(string version)
    {
        Assert.Equal(UpdateChannelFilter.ReleaseStage.Beta, UpdateChannelFilter.GetStage(version));
    }

    [Theory]
    [InlineData("2.6.0-alpha.1")]
    [InlineData("2.6.0-ALPHA.3")]
    public void An_alpha_suffix_is_recognised_case_insensitively(string version)
    {
        Assert.Equal(UpdateChannelFilter.ReleaseStage.Alpha, UpdateChannelFilter.GetStage(version));
    }

    // The whole point of the catch-all, and the defect it fixes. The release pipeline's own stage
    // labeller (.github/workflows/build-and-release.yml) falls through to ALPHA for any suffix it
    // does not recognise, so an unrecognised suffix has to mean ALPHA here too. Reading it as "no
    // stage at all" let a -rc.1 or -preview.1 past the check entirely and offered it to a user who
    // had only ticked BETA.
    [Theory]
    [InlineData("2.6.0-rc.1")]
    [InlineData("2.6.0-rc1")]
    [InlineData("2.6.0-preview.1")]
    [InlineData("2.6.0-dev.2")]
    [InlineData("2.6.0-betaish.1")]
    public void An_unrecognised_pre_release_suffix_is_treated_as_alpha(string version)
    {
        Assert.Equal(UpdateChannelFilter.ReleaseStage.Alpha, UpdateChannelFilter.GetStage(version));

        // And therefore is NOT offered to someone who only ticked BETA.
        Assert.False(UpdateChannelFilter.IsVersionAllowed(version, allowAlpha: false, allowBeta: true));
        Assert.True(UpdateChannelFilter.IsVersionAllowed(version, allowAlpha: true, allowBeta: false));
    }

    // The boxes only ADD pre-release channels. Neither of them, nor both unticked, may ever hide an
    // ordinary release - that would leave a user permanently unable to update.
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void A_stable_release_is_always_allowed(bool allowAlpha, bool allowBeta)
    {
        Assert.True(UpdateChannelFilter.IsVersionAllowed("2.5.0", allowAlpha, allowBeta));
    }

    [Fact]
    public void With_neither_box_ticked_no_pre_release_is_allowed()
    {
        Assert.False(UpdateChannelFilter.IsVersionAllowed("2.6.0-alpha.1", allowAlpha: false, allowBeta: false));
        Assert.False(UpdateChannelFilter.IsVersionAllowed("2.6.0-beta.1", allowAlpha: false, allowBeta: false));
    }

    [Fact]
    public void Each_box_admits_only_its_own_stage()
    {
        Assert.True(UpdateChannelFilter.IsVersionAllowed("2.6.0-beta.1", allowAlpha: false, allowBeta: true));
        Assert.False(UpdateChannelFilter.IsVersionAllowed("2.6.0-alpha.1", allowAlpha: false, allowBeta: true));

        Assert.True(UpdateChannelFilter.IsVersionAllowed("2.6.0-alpha.1", allowAlpha: true, allowBeta: false));
        Assert.False(UpdateChannelFilter.IsVersionAllowed("2.6.0-beta.1", allowAlpha: true, allowBeta: false));
    }

    // The reported defect in miniature: a BETA-only user, with a NEWER alpha sitting alongside an
    // available beta. Filtering must leave the beta standing - the old code discarded the single
    // newest candidate instead, so the alpha permanently hid the beta and the user was told they
    // were up to date forever.
    [Fact]
    public void A_newer_alpha_does_not_hide_an_available_beta_from_a_beta_only_user()
    {
        string[] feed = ["2.6.0-alpha.3", "2.6.0-beta.1", "2.5.0"];

        var allowed = feed
            .Where(version => UpdateChannelFilter.IsVersionAllowed(version, allowAlpha: false, allowBeta: true))
            .ToArray();

        Assert.Equal(["2.6.0-beta.1", "2.5.0"], allowed);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void A_missing_version_is_treated_as_stable_rather_than_throwing(string? version)
    {
        Assert.Equal(UpdateChannelFilter.ReleaseStage.Stable, UpdateChannelFilter.GetStage(version));
        Assert.True(UpdateChannelFilter.IsVersionAllowed(version, allowAlpha: false, allowBeta: false));
    }
}
