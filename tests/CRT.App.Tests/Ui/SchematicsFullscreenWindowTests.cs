using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// The fullscreen schematics window's own keyboard behaviour: Escape has always closed it, and F11
// now does too, so pressing F11 a second time toggles fullscreen back off rather than only Escape
// being able to end it - asked for explicitly so both keys behave the same way. This file covers
// the window's OWN behaviour: its key handling, and that closing it restores the hosted content
// exactly once. Main's side of the same feature - ToggleSchematicsFullscreenWindow deciding whether
// to open or close, and what closing fullscreen does to the detached thumbnails window it owns -
// needs a real Main and lives in ThumbnailsDetachPerBoardTests alongside the rest of that
// interaction.
[Collection("HeadlessUi")]
public sealed class SchematicsFullscreenWindowTests
{
    private static void PressKey(Window window, Key key, PhysicalKey physicalKey) =>
        window.KeyPress(key, RawInputModifiers.None, physicalKey, keySymbol: null);

    private static (SchematicsFullscreenWindow Window, Control Hosted, int RestoreCount) BuildWindow()
    {
        var hosted = new Border();
        int restoreCount = 0;

        var window = new SchematicsFullscreenWindow(hosted, _ => restoreCount++);
        window.Show();

        return (window, hosted, restoreCount);
    }

    [Fact]
    public void Escape_closes_the_fullscreen_window()
    {
        UiTest.Run(() =>
        {
            var (window, _, _) = BuildWindow();

            PressKey(window, Key.Escape, PhysicalKey.Escape);

            Assert.False(window.IsVisible);
        });
    }

    // The behaviour asked for: F11 opened this window, so pressing it again while the window is
    // the active one must close it too - the same toggle Escape already provided.
    [Fact]
    public void F11_also_closes_the_fullscreen_window()
    {
        UiTest.Run(() =>
        {
            var (window, _, _) = BuildWindow();

            PressKey(window, Key.F11, PhysicalKey.F11);

            Assert.False(window.IsVisible);
        });
    }

    [Fact]
    public void Closing_via_F11_restores_the_hosted_content_exactly_once()
    {
        UiTest.Run(() =>
        {
            var hosted = new Border();
            int restoreCount = 0;
            var window = new SchematicsFullscreenWindow(hosted, _ => restoreCount++);
            window.Show();

            PressKey(window, Key.F11, PhysicalKey.F11);

            Assert.Equal(1, restoreCount);

            // Closing again (e.g. the window's own Closed firing a second time) must not restore
            // twice - RestoreHostedContent guards on thisHasRestoredHostedContent for this reason.
            window.Close();
            Assert.Equal(1, restoreCount);
        });
    }

    [Fact]
    public void An_unrelated_key_leaves_the_window_open()
    {
        UiTest.Run(() =>
        {
            var (window, _, _) = BuildWindow();

            PressKey(window, Key.A, PhysicalKey.A);

            Assert.True(window.IsVisible);
        });
    }
}
