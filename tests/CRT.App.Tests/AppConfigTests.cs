using CRT;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// AppConfig - the values whose PLACE matters, not only their spelling.
// ###########################################################################################
public sealed class AppConfigTests
{
    // ###########################################################################################
    // *** THE KiCad COPIES ARE THE USER'S OWN (code review, 2026-10-25). *** BoardKiCadFolder
    // deletes every file in a copy's folder it does not want. Under the temporary folder - the
    // shared /tmp on Linux, at a path anyone can work out - another account could make that folder
    // a link to the user's own files.
    // ###########################################################################################
    [Fact]
    public void The_KiCad_cache_is_in_the_users_own_app_data_folder_not_the_shared_temporary_one()
    {
        string ownFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            AppConfig.AppFolderName);

        Assert.StartsWith(ownFolder + Path.DirectorySeparatorChar, AppConfig.DefaultKiCadCacheFolder, StringComparison.Ordinal);
        Assert.False(AppConfig.DefaultKiCadCacheFolder.StartsWith(Path.GetTempPath(), StringComparison.OrdinalIgnoreCase));
    }

    // ###########################################################################################
    // *** THE SUITE NEVER WRITES INTO THE USER'S KiCad CACHE (2026-11-16, performance walkthrough). ***
    // Every test showing a KiCad board left its copy in the real "KiCad cache" for good. The test
    // processes use a folder of their own (TestKiCadCache), set before any test runs.
    // ###########################################################################################
    [Fact]
    public void The_tests_copy_KiCad_files_into_a_folder_of_their_own()
    {
        Assert.Equal(TestKiCadCache.Folder, AppConfig.KiCadCacheFolder);
        Assert.NotEqual(AppConfig.DefaultKiCadCacheFolder, AppConfig.KiCadCacheFolder);
        Assert.StartsWith(Path.GetTempPath(), AppConfig.KiCadCacheFolder, StringComparison.OrdinalIgnoreCase);
    }

    // ###########################################################################################
    // The renderer may keep a large board's schematic image on the graphics card (2026-11-16, performance
    // walkthrough): some 50 MB decoded, against Avalonia's own room of some 28 MB.
    // ###########################################################################################
    [Fact]
    public void The_renderer_has_room_on_the_graphics_card_for_a_large_schematic_image()
    {
        Avalonia.SkiaOptions options = global::Main.Program.SkiaOptionsForCrt();

        Assert.Equal(AppConfig.GpuPictureBytes, options.MaxGpuResourceSizeBytes);
        Assert.True(options.MaxGpuResourceSizeBytes >= 4 * 50L * 1024 * 1024, "Too little room for a few large schematic images");
    }

    // A scope capture waiting to be accepted too (code review, 2026-10-26): in /tmp another
    // account could swap the screenshot before it went into the draft and the submission. (The
    // test process keeps its own captures elsewhere - TestScopeCaptures - so CRT's own folder is
    // the default's.)
    [Fact]
    public void A_scope_capture_waits_in_the_users_own_app_data_folder_not_the_shared_temporary_one()
    {
        string ownFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            AppConfig.AppFolderName);

        Assert.StartsWith(ownFolder + Path.DirectorySeparatorChar, AppConfig.DefaultScopeCaptureWaitFolder, StringComparison.Ordinal);
        Assert.False(AppConfig.DefaultScopeCaptureWaitFolder.StartsWith(Path.GetTempPath(), StringComparison.OrdinalIgnoreCase));
        Assert.Equal(TestScopeCaptures.Folder, AppConfig.ScopeCaptureWaitFolder);
    }
}
