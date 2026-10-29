using Avalonia.Controls;
using Avalonia.LogicalTree;
using Avalonia.VisualTree;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// The Configuration tab's help beside a setting - two kinds, with two looks (CRT 4.0 step 5, the
// plan: "The remaining help buttons read 'Learn more'"):
//
//   - "Learn more" (the LearnMore class, a link) OPENS a Wiki page - MiniPro, Workbooks, the online
//     source and the Maintainer tab (Click, tested only for existence - see below);
//   - the small "?" (HelpIconButton) shows its explanation as a TOOLTIP and opens nothing - the
//     ALPHA/BETA notifications, which have no Wiki page for a single check box's meaning. Both are worth pinning for the same reason: a button silently dropped
// from the markup, or one left with no icon or no tip text in it, is invisible until someone opens
// the tab and looks.
//
// What is NOT tested here is the wiki-opening click itself: the handler goes through
// ExternalTargetLauncher, whose accept path calls Process.Start, and rule 6 keeps that out of the
// suite (the launcher's own containment predicates are covered by ExternalTargetLauncherTests
// instead). A mis-typed Click handler name fails the XAML parse and so is already caught by
// construction.
[Collection("HeadlessUi")]
public sealed class ConfigurationHelpIconTests
{
    // The Font Awesome "circle-question" glyph of a help icon that shows a tooltip.
    private const string HelpGlyph = "\uf059";

    // A button that opens a Wiki page: a link reading "Learn more".
    private static void AssertLearnMore(Button button)
    {
        Assert.Contains("LearnMore", button.Classes);
        Assert.DoesNotContain("HelpIconButton", button.Classes);
        Assert.Equal("Learn more", ((TextBlock)button.Content!).Text);
        Assert.Contains("Wiki", Assert.IsType<string>(ToolTip.GetTip(button)));
    }

    [Fact]
    public void The_workbooks_setting_has_Learn_more_beside_it()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var helpButton = tab.GetControl<Button>("EnableWorklogHelpButton");

            // The same look the MiniPro help carries, so the two read as one affordance rather than
            // as two different kinds of control.
            ConfigurationHelpIconTests.AssertLearnMore(helpButton);
        });
    }

    // The icon has to sit BESIDE the checkbox, not somewhere else on the tab - it is what says
    // "help about this setting" rather than "help about the section".
    [Fact]
    public void The_workbooks_help_icon_shares_a_row_with_its_checkbox()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var helpButton = tab.GetControl<Button>("EnableWorklogHelpButton");
            var checkBox = tab.GetControl<CheckBox>("EnableWorklogCheckBox");

            Assert.Same(checkBox.GetVisualParent(), helpButton.GetVisualParent());
        });
    }

    // ###########################################################################################
    // The Maintainer tab's setting (2026-09-29) - the same "?" beside the same kind of check box,
    // opening the Maintainer-tab Wiki page (AppConfig.WikiPageMaintainer, which
    // WikiHelpPageNamesTests holds to a real page).
    // ###########################################################################################
    [Fact]
    public void The_maintainer_setting_has_Learn_more_beside_it()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var helpButton = tab.GetControl<Button>("EnableMaintainerTabHelpButton");

            ConfigurationHelpIconTests.AssertLearnMore(helpButton);
        });
    }

    [Fact]
    public void The_maintainer_help_icon_shares_a_row_with_its_checkbox()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var helpButton = tab.GetControl<Button>("EnableMaintainerTabHelpButton");
            var checkBox = tab.GetControl<CheckBox>("EnableMaintainerTabCheckBox");

            Assert.Same(checkBox.GetVisualParent(), helpButton.GetVisualParent());
        });
    }

    // ###########################################################################################
    // "View boards as published online" (owner request, 2026-10-02: "have a
    // Wiki page for this and add a helper icon to it") - the same "?" beside its check box, opening
    // AppConfig.WikiPageViewOnlineSource, which WikiHelpPageNamesTests holds to a real page. The
    // label is pinned too: it is the newest wording (owner decision, 2026-10-13), and the Wiki page
    // names it.
    // ###########################################################################################
    [Fact]
    public void The_online_source_setting_has_Learn_more_beside_it_and_its_new_name()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var helpButton = tab.GetControl<Button>("ViewOfficialPublishedOnlyHelpButton");
            var checkBox = tab.GetControl<CheckBox>("ViewOfficialPublishedOnlyCheckBox");

            ConfigurationHelpIconTests.AssertLearnMore(helpButton);
            Assert.Same(checkBox.GetVisualParent(), helpButton.GetVisualParent());
            Assert.Equal(
                "View boards as published online (hide my draft changes)",
                checkBox.Content);
        });
    }

    // The pattern this one was copied from, asserted alongside it so a change to either is made to
    // both rather than leaving the two help icons looking different.
    [Fact]
    public void The_minipro_setting_has_its_matching_Learn_more()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var helpButton = tab.GetControl<Button>("EnableMiniproExperimentalModeHelpButton");

            ConfigurationHelpIconTests.AssertLearnMore(helpButton);
        });
    }

    [Fact]
    public void The_alpha_notification_setting_has_a_help_icon_with_explanatory_tooltip_text()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var helpButton = tab.GetControl<Button>("AllowAlphaVersionNotificationHelpButton");
            var checkBox = tab.GetControl<CheckBox>("AllowAlphaVersionNotificationCheckBox");

            Assert.Contains("HelpIconButton", helpButton.Classes);
            Assert.Equal(HelpGlyph, ((TextBlock)helpButton.Content!).Text);
            Assert.Same(checkBox.GetVisualParent(), helpButton.GetVisualParent());

            var tip = Assert.IsType<string>(ToolTip.GetTip(helpButton));
            Assert.Contains("ALPHA", tip);
            Assert.Contains("development version", tip);
        });
    }

    [Fact]
    public void The_beta_notification_setting_has_a_help_icon_with_explanatory_tooltip_text()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var helpButton = tab.GetControl<Button>("ShowDevelopmentVersionNotificationHelpButton");
            var checkBox = tab.GetControl<CheckBox>("ShowDevelopmentVersionNotificationCheckBox");

            Assert.Contains("HelpIconButton", helpButton.Classes);
            Assert.Equal(HelpGlyph, ((TextBlock)helpButton.Content!).Text);
            Assert.Same(checkBox.GetVisualParent(), helpButton.GetVisualParent());

            var tip = Assert.IsType<string>(ToolTip.GetTip(helpButton));
            Assert.Contains("BETA", tip);
            Assert.Contains("TEST", tip);
        });
    }

    // Pins the ALPHA row sitting immediately after the BETA row, since that ordering was asked for
    // explicitly rather than being incidental to where the checkbox was added in the markup.
    [Fact]
    public void The_alpha_notification_row_comes_directly_after_the_beta_notification_row()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var betaRow = tab.GetControl<CheckBox>("ShowDevelopmentVersionNotificationCheckBox").GetVisualParent();
            var alphaRow = tab.GetControl<CheckBox>("AllowAlphaVersionNotificationCheckBox").GetVisualParent();

            Assert.NotNull(betaRow);
            Assert.NotNull(alphaRow);

            var siblings = ((Avalonia.Controls.Controls)((StackPanel)betaRow!.GetVisualParent()!).Children);
            int betaIndex = siblings.IndexOf((Control)betaRow);
            int alphaIndex = siblings.IndexOf((Control)alphaRow!);

            Assert.Equal(betaIndex + 1, alphaIndex);
        });
    }

    // ###########################################################################################
    // A check box with a help icon WRAPS in a narrow tab (2026-10-13). Each such row was a
    // horizontal StackPanel, which gives its children unlimited width - so "View boards as
    // published online (hide my draft changes)" was cut off at the window's minimum width instead
    // of wrapping. The row is now a Left-aligned Grid: the check box in a "*" column (the width
    // left over, so it wraps), the icon in an Auto column after it. The suite's headless fonts
    // cannot measure the wrap itself, so the SHAPE that makes it possible is what is pinned.
    // ###########################################################################################
    [Fact]
    public void Every_check_box_with_a_help_icon_can_wrap()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var helpButtons = tab.GetLogicalDescendants()
                .OfType<Button>()
                .Where(button => button.Classes.Contains("HelpIconButton") || button.Classes.Contains("LearnMore"))
                .ToList();

            Assert.NotEmpty(helpButtons);

            foreach (Button helpButton in helpButtons)
            {
                var row = Assert.IsType<Grid>(helpButton.Parent);
                var checkBox = Assert.Single(row.Children.OfType<CheckBox>());

                Assert.True(row.ColumnDefinitions[0].Width.IsStar, $"{helpButton.Name}: the check box's column must be \"*\"");
                Assert.Equal(0, Grid.GetColumn(checkBox));
                Assert.Equal(1, Grid.GetColumn(helpButton));

                // Left, or a wide tab would push the icon to the far edge, away from its label.
                Assert.Equal(Avalonia.Layout.HorizontalAlignment.Left, row.HorizontalAlignment);
            }
        });
    }
}
