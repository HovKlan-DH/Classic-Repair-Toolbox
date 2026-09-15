using Avalonia.Controls;
using Avalonia.VisualTree;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// The Configuration tab's help icons - the small "?" buttons that sit beside a setting.
//
// Two flavours share the same "HelpIconButton" look but do different things: the MiniPro/Workbooks
// ones OPEN a wiki page (Click, tested only for existence - see below), while the ALPHA/BETA
// notification ones show their explanation as a TOOLTIP instead, since there is no wiki page for a
// single checkbox's meaning. Both are worth pinning for the same reason: a button silently dropped
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
    // The Font Awesome "circle-question" glyph the other help buttons on this tab already use.
    private const string HelpGlyph = "\uf059";

    [Fact]
    public void The_workbooks_setting_has_a_help_icon_beside_it()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var helpButton = tab.GetControl<Button>("EnableWorklogHelpButton");

            // The same styling and the same glyph the MiniPro help button carries, so the two read
            // as one affordance rather than as two different kinds of control.
            Assert.Contains("HelpIconButton", helpButton.Classes);
            Assert.Equal(HelpGlyph, ((TextBlock)helpButton.Content!).Text);
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

    // The pattern this one was copied from, asserted alongside it so a change to either is made to
    // both rather than leaving the two help icons looking different.
    [Fact]
    public void The_minipro_setting_still_has_its_matching_help_icon()
    {
        UiTest.Run(() =>
        {
            var tab = new TabConfiguration();

            var helpButton = tab.GetControl<Button>("EnableMiniproExperimentalModeHelpButton");

            Assert.Contains("HelpIconButton", helpButton.Classes);
            Assert.Equal(HelpGlyph, ((TextBlock)helpButton.Content!).Text);
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
}
