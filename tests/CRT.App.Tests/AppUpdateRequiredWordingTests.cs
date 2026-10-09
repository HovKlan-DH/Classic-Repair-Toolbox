using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What a covered Drafts or Maintainer tab says (owner request, 2026-10-09). The overlay cannot be
// closed, so its words carry everything: what is wrong, what it stops, that drafts are safe, and
// the one way out - install the update CRT found, or open the download page.
// ###########################################################################################
public sealed class AppUpdateRequiredWordingTests
{
    private const string Reason = "This version of CRT [3.0.0] was made for an older version of the server - please update CRT to the newest version.";

    [Fact]
    public void The_overlay_leads_with_the_reason_it_was_given()
    {
        AppUpdateRequiredView view = AppUpdateRequiredWording.For(AppUpdateArea.Drafts, Reason, null);

        Assert.Equal(AppUpdateRequiredWording.Heading, view.Heading);
        Assert.Equal(Reason, view.Reason);
        Assert.Equal(AppUpdateRequiredWording.RestOfCrt, view.RestOfCrt);
    }

    // A contributor's first worry, with the tab covered, is the drafts - so the Drafts tab says
    // they are safe. The Maintainer tab says the tab cannot be used.
    [Fact]
    public void Each_tab_says_what_it_cannot_do_until_CRT_is_updated()
    {
        string drafts = AppUpdateRequiredWording.For(AppUpdateArea.Drafts, Reason, null).WhatItStops;
        string maintainer = AppUpdateRequiredWording.For(AppUpdateArea.Maintainer, Reason, null).WhatItStops;

        Assert.Contains("drafts cannot be submitted", drafts, StringComparison.Ordinal);
        Assert.Contains("Your drafts stay on this computer", drafts, StringComparison.Ordinal);
        Assert.Contains("nothing is lost", drafts, StringComparison.Ordinal);
        Assert.Contains("Maintainer tab cannot be used", maintainer, StringComparison.Ordinal);
        Assert.NotEqual(drafts, maintainer);
    }

    // With an update already found, the button installs it and the overlay names the version.
    [Fact]
    public void An_update_found_is_offered_for_installing()
    {
        AppUpdateRequiredView view = AppUpdateRequiredWording.For(AppUpdateArea.Maintainer, Reason, " 3.1.0 ");

        Assert.True(view.ButtonInstalls);
        Assert.Equal(AppUpdateRequiredWording.InstallButton, view.ButtonLabel);
        Assert.Equal("Version [3.1.0] is ready to install.", view.ReadyLine);
    }

    // None found - or the check is off, or CRT is not an installed copy - and the button opens the
    // download page instead, with no "ready" line.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("  ")]
    public void With_no_update_found_the_button_opens_the_download_page(string? pendingVersion)
    {
        AppUpdateRequiredView view = AppUpdateRequiredWording.For(AppUpdateArea.Drafts, Reason, pendingVersion);

        Assert.False(view.ButtonInstalls);
        Assert.Equal(AppUpdateRequiredWording.DownloadPageButton, view.ButtonLabel);
        Assert.Null(view.ReadyLine);
    }

    // CLAUDE.md: plain ASCII in what a user reads, no "..." on a button, and no "their" for a
    // contributor.
    [Fact]
    public void The_words_are_plain_ASCII_with_no_ellipsis_and_no_their()
    {
        AppUpdateRequiredView[] views =
        [
            AppUpdateRequiredWording.For(AppUpdateArea.Drafts, Reason, null),
            AppUpdateRequiredWording.For(AppUpdateArea.Maintainer, Reason, "3.1.0")
        ];

        foreach (AppUpdateRequiredView view in views)
        {
            foreach (string text in new[] { view.Heading, view.WhatItStops, view.RestOfCrt, view.ReadyLine ?? string.Empty, view.ButtonLabel })
            {
                Assert.All(text, character => Assert.True(character < 128, $"[{text}] has a character outside ASCII"));
                Assert.DoesNotContain("...", text, StringComparison.Ordinal);
                Assert.DoesNotMatch(@"\b(their|them|they)\b", text);
            }
        }
    }
}
