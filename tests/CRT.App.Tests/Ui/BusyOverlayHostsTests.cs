using System;
using System.IO;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// *** EVERY WINDOW WITH A WAIT HAS AN OVERLAY TO SHOW IT ON (owner decision, 2026-09-28: "I want
// this method everywhere in the entire project where there is a Wait"). ***
//
// BusyOverlay.RunAsync on a window with no overlay still runs the work under the two-minute limit -
// it cannot hang - but it shows NOTHING, which is exactly the "Working..." problem this replaced. A
// host dropping its overlay would compile and pass every other test, so each one is named here.
//
// The dialogs are built; Main's markup is read. The Maintainer tab (the separate CRT Maintainer
// application until 2026-09-29) has NO overlay of its own - its waits run under Main's - and a
// test says so, since a second one would dim only the tab.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BusyOverlayHostsTests
{
    [Fact]
    public void Every_dialog_that_waits_on_something_has_an_overlay()
    {
        UiTest.Run(() =>
        {
            Assert.NotNull(BusyOverlay.For(new MySubmissionsWindow()));
            Assert.NotNull(BusyOverlay.For(new SystemFilesWindow()));
            Assert.NotNull(BusyOverlay.For(new ComponentContributionWindow()));
        });
    }

    [Theory]
    [InlineData("src/CRT.App/Main/Main.axaml")]

    // The Maintainer tab's file tree window (2026-09-28): opening a file fetches it first. A
    // window of its own, so an overlay of its own.
    [InlineData("src/CRT.App/Tabs/Maintainer/FileTreeWindow.axaml")]
    public void The_main_window_and_the_file_tree_window_carry_the_overlay(string relativePath)
    {
        string markup = File.ReadAllText(BusyOverlayHostsTests.ResolveRepositoryPath(relativePath));

        Assert.Contains("<ui:BusyOverlay", markup, StringComparison.Ordinal);
        Assert.Contains("xmlns:ui=\"clr-namespace:CRT;assembly=CRT.UI\"", markup, StringComparison.Ordinal);
    }

    // ###########################################################################################
    // *** THE MAINTAINER TAB HAS NO OVERLAY OF ITS OWN (2026-09-29). *** Its waits run under CRT's
    // window's one, found up the tree. A second one inside the tab would dim only the tab and block
    // nothing above it - the update banner, the other tabs' headers, the data-sync icon would all
    // stay clickable during a publish. It had one while it was a window, so this is the place a
    // careless copy would put it back.
    // ###########################################################################################
    [Fact]
    public void The_maintainer_tab_has_no_overlay_of_its_own()
    {
        string markup = File.ReadAllText(BusyOverlayHostsTests.ResolveRepositoryPath("src/CRT.App/Tabs/Maintainer/TabMaintainer.axaml"));

        Assert.DoesNotContain("<ui:BusyOverlay", markup, StringComparison.Ordinal);
    }

    // Walks up from the test binary until the repository file is found - the same approach
    // MaintainerReleaseSeparationTests and WikiHelpPageNamesTests use.
    private static string ResolveRepositoryPath(string relativePath)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory != null)
        {
            string candidate = Path.Combine(directory.FullName, relativePath);

            if (File.Exists(candidate))
                return candidate;

            directory = directory.Parent;
        }

        Assert.Fail($"{relativePath} was not found above the test binaries");
        return string.Empty;
    }
}
