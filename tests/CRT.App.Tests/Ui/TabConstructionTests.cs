using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// Every tab is constructed for real, headlessly, with the app's actual App.axaml styles and
// resource dictionaries loaded.
//
// This is the only automated coverage of Tabs/ - the unit tests underneath cover Handlers/
// and never touch a control. Constructing a tab runs InitializeComponent (so the .axaml must
// parse, and every x:Name the code-behind reaches for must still exist) and then the
// constructor body, which for most tabs wires event handlers to those named controls. Rename
// or delete a control in the .axaml without updating the code-behind and these fail, where
// previously nothing would notice until the tab was opened by hand.
//
// "Does not throw" is the whole assertion, deliberately. Whether the layout LOOKS right still
// needs a human running the app; whether it can be built at all no longer does.
// ###########################################################################################
[Collection("HeadlessUi")]
public class TabConstructionTests
{
    [Fact]
    public void The_about_tab_can_be_constructed()
    {
        UiTest.Run(() => Assert.NotNull(new TabAbout()));
    }

    [Fact]
    public void The_configuration_tab_can_be_constructed()
    {
        // Reads UserSettings statics in its constructor. No seeding here on purpose: the
        // defaults are enough to prove it builds, and reading them cannot throw.
        UiTest.Run(() => Assert.NotNull(new TabConfiguration()));
    }

    [Fact]
    public void The_contribute_tab_can_be_constructed()
    {
        UiTest.Run(() => Assert.NotNull(new TabContribute()));
    }

    [Fact]
    public void The_drafts_tab_can_be_constructed()
    {
        UiTest.Run(() => Assert.NotNull(new TabDrafts()));
    }

    [Fact]
    public void The_feedback_tab_can_be_constructed()
    {
        UiTest.Run(() => Assert.NotNull(new TabFeedback()));
    }

    [Fact]
    public void The_oscilloscope_tab_can_be_constructed()
    {
        UiTest.Run(() => Assert.NotNull(new TabOscilloscope()));
    }

    [Fact]
    public void The_overview_tab_can_be_constructed()
    {
        UiTest.Run(() => Assert.NotNull(new TabOverview()));
    }

    [Fact]
    public void The_resources_tab_can_be_constructed()
    {
        UiTest.Run(() => Assert.NotNull(new TabResources()));
    }

    [Fact]
    public void The_schematics_tab_can_be_constructed()
    {
        // The heaviest one: its constructor wires seventeen handlers to named controls, so it
        // is the tab most likely to break silently when the .axaml is edited.
        UiTest.Run(() => Assert.NotNull(new TabSchematics()));
    }

    [Fact]
    public void The_workbooks_tab_can_be_constructed()
    {
        // Currently a static mockup with an empty constructor, so this proves the .axaml parses
        // rather than that any wiring survives. That is worth having on its own here: the tab is
        // several hundred lines of nested layout referencing a dozen Workbooks_* theme keys, and
        // a malformed element or an unparseable value would otherwise only show up when someone
        // ticked "Enable Worklog" and opened the tab by hand.
        UiTest.Run(() => Assert.NotNull(new TabWorkbooks()));
    }

    [Fact]
    public void The_new_system_window_can_be_constructed()
    {
        // "Add a new system" (session 2c, task 9). Its constructor wires three TextChanged handlers
        // and a tunnelled key handler to named controls, then runs its first validation pass - so
        // this proves both the .axaml parses and that the initial pass does not throw against
        // boxes that are still empty.
        UiTest.Run(() => Assert.NotNull(new NewSystemWindow()));
    }

    [Fact]
    public void The_system_files_window_can_be_constructed()
    {
        // The schematic and KiCad import window (session 2c, task 9).
        UiTest.Run(() => Assert.NotNull(new SystemFilesWindow()));
    }

    [Fact]
    public void The_draft_drift_window_can_be_constructed()
    {
        // The "what changed officially" report (session 2d).
        UiTest.Run(() => Assert.NotNull(new DraftDriftWindow()));
    }
}
