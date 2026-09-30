using System;
using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// *** EVERY WINDOW WITH A WAIT HAS AN OVERLAY TO SHOW IT ON (owner decision, 2026-09-28: "I want
// this method everywhere in the entire project where there is a Wait"). ***
//
// BusyOverlay.RunAsync on a window with no overlay still runs the work under the two-minute limit -
// it cannot hang - but it shows NOTHING, which is exactly the "Working..." problem this replaced. A
// host dropping its overlay would compile and pass every other test, so each one is named here.
//
// Every host is BUILT and asked through BusyOverlay.For, the way the code finds it. These used to
// read Main's and the file tree window's markup for "<ui:BusyOverlay" (code review, 2026-09-30):
// that passed with the element commented out, and tied the check to one prefix, so the "no overlay
// of its own" test would have passed vacuously once the prefix changed. The Maintainer tab (the
// separate CRT Maintainer application until 2026-09-29) has NO overlay of its own - its waits run
// under Main's - and a test says so, since a second one would dim only the tab.
//
// Main is built, never shown, with UserSettings and the workbook folder pointed at temp files -
// MainMaintainerTabTests' own setup.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BusyOverlayHostsTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public BusyOverlayHostsTests()
    {
        this.RedirectToTemp();
    }

    public void Dispose()
    {
        this.RedirectToTemp();
        this.thisWorkspace.Dispose();
    }

    private void RedirectToTemp()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
        // An EMPTY file, not a missing one: LoadFrom keeps the settings already in memory when the
        // file does not exist.
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile(Guid.NewGuid().ToString("N") + ".json", "{}"));
    }

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

    // ###########################################################################################
    // Main and the Maintainer tab's file tree window (2026-09-28: opening a file fetches it first -
    // a window of its own, so an overlay of its own). Each overlay must be a direct child of the
    // window's ROOT grid: it fades its siblings when it dims, so one nested deeper would fade only
    // part of the window.
    // ###########################################################################################
    [Fact]
    public void The_main_window_and_the_file_tree_window_carry_the_overlay_over_their_whole_content()
    {
        UiTest.Run(() =>
        {
            foreach (Window window in new Window[] { new CRT.Main(), new FileTreeWindow() })
            {
                BusyOverlay? overlay = BusyOverlay.For(window);

                Assert.NotNull(overlay);
                Assert.Same(window.Content, overlay.Parent);
            }
        });
    }

    // ###########################################################################################
    // *** THE MAINTAINER TAB HAS NO OVERLAY OF ITS OWN (2026-09-29). *** Its waits run under CRT's
    // window's one, found up the tree. A second one inside the tab would dim only the tab and block
    // nothing above it - the update banner, the other tabs' headers, the data-sync icon would all
    // stay clickable during a publish. It had one while it was a window, so this is the place a
    // careless copy would put it back.
    //
    // A tab built on its own has no window above it, so For finds only what is inside the tab.
    // ###########################################################################################
    [Fact]
    public void The_maintainer_tab_has_no_overlay_of_its_own()
    {
        UiTest.Run(() =>
        {
            Assert.Null(BusyOverlay.For(new TabMaintainer()));
        });
    }
}
