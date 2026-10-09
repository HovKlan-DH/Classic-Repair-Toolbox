using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.LogicalTree;
using CRT;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// UpdateRequiredOverlay - "CRT has to be updated" over a whole tab, which cannot be closed (owner
// request, 2026-10-09: "put a fullpage modal (or alike) in there, that cannot be closed, stating
// the app needs to be updated").
//
// What must hold: hidden until shown; shown, the tab beneath it fades, takes no click, no focus and
// no key; there is no way to close it - its one button is the way out, and tells the host.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class UpdateRequiredOverlayTests
{
    private const string Reason = "This version of CRT [3.0.0] was made for an older version of the server - please update CRT to the newest version.";

    // A tab laid out the way both hosts lay theirs out: the content, then the overlay last.
    private static (Window Window, Button TabButton, TextBox TabBox, UpdateRequiredOverlay Overlay) BuildHost()
    {
        var button = new Button { Content = "Submit" };
        var box = new TextBox();
        var content = new StackPanel { Children = { button, box } };
        var overlay = new UpdateRequiredOverlay();
        var window = new Window { Content = new Grid { Children = { content, overlay } }, Width = 800, Height = 600 };

        return (window, button, box, overlay);
    }

    [Fact]
    public void It_is_hidden_until_shown()
    {
        UiTest.Run(() =>
        {
            UpdateRequiredOverlay overlay = BuildHost().Overlay;

            Assert.False(overlay.IsShown);
            Assert.False(overlay.IsVisible);
            Assert.Null(overlay.View);
        });
    }

    // Shown, everything else in the tab fades and is turned off - nothing in it can be clicked or
    // take focus - and the overlay itself is at full strength.
    [Fact]
    public void Shown_it_covers_the_tab_and_turns_off_everything_beneath_it()
    {
        UiTest.Run(() =>
        {
            var (_, tabButton, tabBox, overlay) = BuildHost();

            overlay.Show(AppUpdateRequiredWording.For(AppUpdateArea.Drafts, Reason, null));

            Assert.True(overlay.IsShown);
            Assert.Single(overlay.CoveredForTests);
            Assert.Equal(BusyOverlay.Fade, overlay.CoveredForTests[0].Opacity);
            Assert.False(tabButton.IsEffectivelyEnabled);
            Assert.False(tabBox.IsEffectivelyEnabled);
            Assert.Equal(1, overlay.Opacity);
        });
    }

    [Fact]
    public void It_says_what_it_was_given()
    {
        UiTest.Run(() =>
        {
            UpdateRequiredOverlay overlay = BuildHost().Overlay;
            AppUpdateRequiredView view = AppUpdateRequiredWording.For(AppUpdateArea.Maintainer, Reason, "3.1.0");

            overlay.Show(view);

            Assert.Same(view, overlay.View);
            Assert.Equal(AppUpdateRequiredWording.Heading, overlay.GetControl<TextBlock>("HeadingText").Text);
            Assert.Equal(Reason, overlay.GetControl<TextBlock>("ReasonText").Text);
            Assert.Equal(view.WhatItStops, overlay.GetControl<TextBlock>("WhatItStopsText").Text);
            Assert.Equal(AppUpdateRequiredWording.RestOfCrt, overlay.GetControl<TextBlock>("RestOfCrtText").Text);
            Assert.Equal("Version [3.1.0] is ready to install.", overlay.GetControl<TextBlock>("ReadyText").Text);
            Assert.True(overlay.GetControl<TextBlock>("ReadyText").IsVisible);
            Assert.Equal(AppUpdateRequiredWording.InstallButton, overlay.GetControl<Button>("ActionButton").Content);
        });
    }

    // ###########################################################################################
    // *** IT CANNOT BE CLOSED. *** Its one button is the way out, and showing it again only changes
    // the words - an update found later turns "Open download page" into "Install update" - without
    // covering anything twice.
    // ###########################################################################################
    [Fact]
    public void Shown_again_it_changes_its_words_and_still_has_only_the_one_button()
    {
        UiTest.Run(() =>
        {
            UpdateRequiredOverlay overlay = BuildHost().Overlay;

            overlay.Show(AppUpdateRequiredWording.For(AppUpdateArea.Drafts, Reason, null));

            Assert.False(overlay.GetControl<TextBlock>("ReadyText").IsVisible);
            Assert.Equal(AppUpdateRequiredWording.DownloadPageButton, overlay.GetControl<Button>("ActionButton").Content);

            overlay.Show(AppUpdateRequiredWording.For(AppUpdateArea.Drafts, Reason, "3.1.0"));

            Assert.True(overlay.IsShown);
            Assert.Single(overlay.CoveredForTests);
            Assert.True(overlay.GetControl<TextBlock>("ReadyText").IsVisible);
            Assert.Equal(AppUpdateRequiredWording.InstallButton, overlay.GetControl<Button>("ActionButton").Content);
            Assert.Single(overlay.GetLogicalDescendants().OfType<Button>());
        });
    }

    // ###########################################################################################
    // *** KEYS MEANT FOR THE TAB GO NOWHERE; THE OVERLAY'S BUTTON STILL TAKES ITS OWN. *** A control
    // that had focus when the overlay appeared would otherwise take Enter or typing. Real key presses
    // through the headless input stack.
    // ###########################################################################################
    [Fact]
    public void Keys_meant_for_the_tab_go_nowhere_while_the_overlays_button_takes_Enter()
    {
        UiTest.Run(() =>
        {
            var (window, tabButton, tabBox, overlay) = BuildHost();
            window.Show();

            int tabClicks = 0;
            int actions = 0;
            tabButton.Click += (_, _) => tabClicks++;
            overlay.ActionClicked += (_, _) => actions++;

            tabBox.Focus();
            window.KeyTextInput("a");
            Assert.Equal("a", tabBox.Text);

            tabButton.Focus();
            overlay.Show(AppUpdateRequiredWording.For(AppUpdateArea.Drafts, Reason, null));

            window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);
            window.KeyTextInput("b");

            Assert.Equal(0, tabClicks);
            Assert.Equal("a", tabBox.Text);

            overlay.GetControl<Button>("ActionButton").Focus();
            window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);

            Assert.Equal(1, actions);

            window.Close();
        });
    }

    // The button tells the host, which reads View.ButtonInstalls to know what was offered.
    [Fact]
    public void The_button_tells_the_host_what_it_offered()
    {
        UiTest.Run(() =>
        {
            UpdateRequiredOverlay overlay = BuildHost().Overlay;
            object? from = null;

            overlay.ActionClicked += (sender, _) => from = sender;
            overlay.Show(AppUpdateRequiredWording.For(AppUpdateArea.Drafts, Reason, "3.1.0"));

            overlay.ClickActionForTests();

            Assert.Same(overlay, from);
            Assert.True(overlay.View!.ButtonInstalls);
        });
    }
}
