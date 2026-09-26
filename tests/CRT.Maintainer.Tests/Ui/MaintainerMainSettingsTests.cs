using Avalonia;
using Avalonia.Controls;
using CRT.Maintainer.Handlers;

namespace CRT.Maintainer.Tests.Ui;

// ###########################################################################################
// What the maintainer window remembers between runs (owner requests, 2026-09-26) - its place,
// maximized or not, and "Show changes only" - applied through MaintainerMain.UseSettings, which
// only the running application calls. The window is BUILT, never shown (see
// MaintainerMainTableTests), and the save callback captures instead of writing a file.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MaintainerMainSettingsTests
{
    [Fact]
    public void A_remembered_window_is_placed_before_it_opens()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();

            main.UseSettings(
                new MaintainerSettings { HasWindowPlacement = true, WindowX = 120, WindowY = 80, WindowWidth = 1300, WindowHeight = 850 },
                _ => { });

            Assert.Equal(WindowStartupLocation.Manual, main.WindowStartupLocation);
            Assert.Equal(new PixelPoint(120, 80), main.Position);
            Assert.Equal(1300, main.Width);
            Assert.Equal(850, main.Height);
            Assert.Equal(WindowState.Normal, main.WindowState);
        });
    }

    // Maximized: placed on its saved screen first, so it maximizes THERE and not on whichever
    // monitor the system picks.
    [Fact]
    public void A_window_maximized_last_time_starts_maximized_on_its_screen()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();

            main.UseSettings(
                new MaintainerSettings { HasWindowPlacement = true, WindowMaximized = true, WindowWidth = 1300, WindowHeight = 850, ScreenX = -1920, ScreenY = 0 },
                _ => { });

            Assert.Equal(WindowState.Maximized, main.WindowState);
            Assert.Equal(new PixelPoint(-1820, 100), main.Position);
        });
    }

    // ###########################################################################################
    // *** RESTORING MAXIMIZED MUST NOT FORGET WHERE THE WINDOW WAS (code review, 2026-09-26). ***
    // The window is placed on its saved screen before it is maximized, and that placement raised
    // PositionChanged while the state was still Normal - so the synthetic corner point overwrote
    // the remembered normal position and was saved back on close. A few maximized-only runs later,
    // un-maximizing landed in the corner instead of where the window last was.
    // ###########################################################################################
    [Fact]
    public void A_window_restored_maximized_still_remembers_where_it_was_windowed()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();

            main.UseSettings(
                new MaintainerSettings
                {
                    HasWindowPlacement = true,
                    WindowMaximized = true,
                    WindowX = 800,
                    WindowY = 300,
                    WindowWidth = 1300,
                    WindowHeight = 850,
                    ScreenX = -1920,
                    ScreenY = 0
                },
                _ => { });

            MaintainerSettings saved = main.CurrentSettings();

            Assert.True(saved.WindowMaximized);
            Assert.Equal(800, saved.WindowX);
            Assert.Equal(300, saved.WindowY);
        });
    }

    // A first run has no place to go back to: the window opens where it always did.
    [Fact]
    public void Without_a_remembered_place_the_window_opens_as_it_always_did()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();
            WindowStartupLocation before = main.WindowStartupLocation;

            main.UseSettings(new MaintainerSettings(), _ => { });

            Assert.Equal(before, main.WindowStartupLocation);
            Assert.False(main.CurrentSettings().HasWindowPlacement);
        });
    }

    // ###########################################################################################
    // "Show changes only" starts as it was left, and what is saved is the maintainer's CHOICE -
    // kept even while a new system's table has the filter off.
    // ###########################################################################################
    [Fact]
    public void Show_changes_only_starts_as_it_was_left_and_is_saved_as_chosen()
    {
        UiTest.Run(() =>
        {
            var main = new MaintainerMain();

            main.UseSettings(new MaintainerSettings { ShowChangesOnly = true }, _ => { });
            Assert.True(main.TableEditorForTests.OnlyChanges);

            main.TableEditorForTests.OnlyChanges = false;
            Assert.False(main.CurrentSettings().ShowChangesOnly);

            main.TableEditorForTests.OnlyChanges = true;
            Assert.True(main.CurrentSettings().ShowChangesOnly);
        });
    }
}
