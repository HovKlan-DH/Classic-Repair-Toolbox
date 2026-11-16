using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// ENTER TURNS "Blink selected" ON AND OFF ON THE "Schematics" TAB (owner request, 2026-11-14: "if I have selected one or
// more components, and I hit "ENTER" - should it then not toggle the blink? I think I do remember it has worked like this
// at one point in time"). CRT 2.x did ("ENTER will toggle blinking for selected components"); Main.KeyIsForBlink brings it
// back. Pressed the way a real key goes - through the real Main window, from where the keyboard is - since the component
// filter, which holds the focus on that tab, has no Enter of its own to see it with.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainBlinkKeyTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public MainBlinkKeyTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        UserSettings.BlinkSelected = false;
    }

    public void Dispose()
    {
        UserSettings.BlinkSelected = false;
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private static void PressEnter(CRT.Main window) =>
        window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, keySymbol: null);

    [Fact]
    public void Enter_in_the_component_filter_turns_blinking_on_and_off_and_keeps_it()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { Width = 1300, Height = 800 };
            window.Show();
            window.MainTabControl.SelectedItem = window.SchematicsTabItem;
            Dispatcher.UIThread.RunJobs();

            CheckBox blink = window.GetControl<CheckBox>("BlinkSelectedCheckBox");
            window.ComponentSearchTextBox.Focus();
            Assert.True(window.ComponentSearchTextBox.IsFocused);
            Assert.False(blink.IsChecked);

            PressEnter(window);

            Assert.True(blink.IsChecked);
            Assert.True(UserSettings.BlinkSelected);

            PressEnter(window);

            Assert.False(blink.IsChecked);
            Assert.False(UserSettings.BlinkSelected);

            window.Close();
        });
    }

    // Where Enter means something already, it is left alone: a focused button presses, and another tab has nothing to blink.
    [Fact]
    public void Enter_on_a_button_or_another_tab_leaves_blinking_alone()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { Width = 1300, Height = 800 };
            window.Show();
            window.MainTabControl.SelectedItem = window.SchematicsTabItem;
            Dispatcher.UIThread.RunJobs();

            CheckBox blink = window.GetControl<CheckBox>("BlinkSelectedCheckBox");

            Button clear = window.GetControl<Button>("ClearComponentsButton");
            clear.Focus();
            Assert.True(clear.IsFocused);

            PressEnter(window);

            Assert.False(blink.IsChecked);

            window.MainTabControl.SelectedItem = window.AboutTabItem;
            Dispatcher.UIThread.RunJobs();
            window.ComponentSearchTextBox.Focus();

            PressEnter(window);

            Assert.False(blink.IsChecked);

            window.Close();
        });
    }
}
