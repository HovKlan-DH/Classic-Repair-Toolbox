using CRT.Maintainer.Handlers;

namespace CRT.Maintainer.Tests;

// ###########################################################################################
// Covers MaintainerSettings / MaintainerSettingsStore - what the maintainer application remembers
// between runs (owner requests, 2026-09-26): its window's place and "Show changes only" - and
// WindowPlacementRules, which keeps a remembered window from opening off every screen.
//
// A temp folder stands in for the user's; nothing here touches the real settings file.
// ###########################################################################################
public sealed class MaintainerSettingsTests : IDisposable
{
    private readonly string thisFolder =
        Path.Combine(Path.GetTempPath(), "crt-maintainer-settings-" + Guid.NewGuid().ToString("N"));

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(this.thisFolder))
                Directory.Delete(this.thisFolder, recursive: true);
        }
        catch (IOException)
        {
            // A leftover temp folder is harmless; failing a test over cleanup is not.
        }
    }

    private string FilePath => Path.Combine(this.thisFolder, "sub", MaintainerSettingsStore.FileName);

    [Fact]
    public void What_is_saved_is_what_is_read_back()
    {
        var saved = new MaintainerSettings
        {
            HasWindowPlacement = true,
            WindowMaximized = true,
            WindowX = -1800,
            WindowY = 40,
            WindowWidth = 1300,
            WindowHeight = 820,
            ScreenX = -1920,
            ScreenY = 0,
            ShowChangesOnly = true
        };

        // The folder does not exist yet - saving creates it.
        MaintainerSettingsStore.Save(this.FilePath, saved);
        MaintainerSettings read = MaintainerSettingsStore.Load(this.FilePath);

        Assert.True(read.HasWindowPlacement);
        Assert.True(read.WindowMaximized);
        Assert.Equal(-1800, read.WindowX);
        Assert.Equal(40, read.WindowY);
        Assert.Equal(1300, read.WindowWidth);
        Assert.Equal(820, read.WindowHeight);
        Assert.Equal(-1920, read.ScreenX);
        Assert.True(read.ShowChangesOnly);
    }

    // A first run, and a file that cannot be read: defaults - never an error, since it is only a
    // window position.
    [Fact]
    public void A_missing_or_unreadable_file_is_the_defaults()
    {
        MaintainerSettings missing = MaintainerSettingsStore.Load(this.FilePath);
        Assert.False(missing.HasWindowPlacement);
        Assert.False(missing.ShowChangesOnly);

        Directory.CreateDirectory(Path.GetDirectoryName(this.FilePath)!);
        File.WriteAllText(this.FilePath, "{ not json");

        Assert.False(MaintainerSettingsStore.Load(this.FilePath).HasWindowPlacement);
    }

    // ###########################################################################################
    // A remembered window opens where it was only while its centre is on a screen - a monitor
    // unplugged since would otherwise leave it where it cannot be reached.
    // ###########################################################################################
    [Fact]
    public void A_window_is_visible_while_its_centre_is_on_a_screen()
    {
        (int, int, int, int)[] screens = [(0, 0, 1920, 1080), (-1920, 0, 1920, 1080)];

        Assert.True(WindowPlacementRules.IsCentreOnAScreen(100, 100, 1100, 720, 1.0, screens));
        Assert.True(WindowPlacementRules.IsCentreOnAScreen(-1800, 100, 1100, 720, 1.0, screens));

        // The left monitor gone.
        Assert.False(WindowPlacementRules.IsCentreOnAScreen(-1800, 100, 1100, 720, 1.0, [(0, 0, 1920, 1080)]));

        // Mostly off the right edge: the centre is past it.
        Assert.False(WindowPlacementRules.IsCentreOnAScreen(1500, 100, 1100, 720, 1.0, [(0, 0, 1920, 1080)]));
    }

    // The size is in device-independent units, the position in pixels - at 150% the centre is
    // further along than the plain size says.
    [Fact]
    public void The_centre_is_found_at_the_screens_scaling()
    {
        (int, int, int, int)[] screen = [(0, 0, 1920, 1080)];

        Assert.True(WindowPlacementRules.IsCentreOnAScreen(1300, 100, 1100, 720, 1.0, screen));
        Assert.False(WindowPlacementRules.IsCentreOnAScreen(1300, 100, 1100, 720, 1.5, screen));
    }

    // ###########################################################################################
    // When the queue checks itself (QueueRefreshRules): never asked, or asked long enough ago.
    // ###########################################################################################
    [Fact]
    public void A_queue_check_is_due_when_never_asked_or_asked_long_enough_ago()
    {
        var now = new DateTimeOffset(2026, 9, 26, 12, 0, 0, TimeSpan.Zero);

        Assert.True(QueueRefreshRules.IsDue(null, now, QueueRefreshRules.ActivationGap));
        Assert.False(QueueRefreshRules.IsDue(now.AddSeconds(-5), now, QueueRefreshRules.ActivationGap));
        Assert.True(QueueRefreshRules.IsDue(now.AddSeconds(-15), now, QueueRefreshRules.ActivationGap));
        Assert.Equal(TimeSpan.FromMinutes(1), QueueRefreshRules.PollInterval);
    }
}
