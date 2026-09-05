using Avalonia.Controls;
using Avalonia.VisualTree;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// The Configuration tab's help icons - the small "?" buttons that open a wiki page for the setting
// they sit beside.
//
// What is NOT tested here is the click itself: the handler goes through ExternalTargetLauncher,
// whose accept path calls Process.Start, and rule 6 keeps that out of the suite (the launcher's own
// containment predicates are covered by ExternalTargetLauncherTests instead). What IS worth pinning
// is that the button exists at all and carries the glyph: a mis-typed Click handler name fails the
// XAML parse and so is already caught by construction, but a button silently dropped from the
// markup, or one left with no icon in it, is invisible until someone opens the tab and looks.
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
}
