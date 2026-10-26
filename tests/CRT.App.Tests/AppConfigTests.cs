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

        Assert.StartsWith(ownFolder + Path.DirectorySeparatorChar, AppConfig.KiCadCacheFolder, StringComparison.Ordinal);
        Assert.False(AppConfig.KiCadCacheFolder.StartsWith(Path.GetTempPath(), StringComparison.OrdinalIgnoreCase));
    }

    // A scope capture waiting for "Add this baseline" too (code review, 2026-10-26): in /tmp another
    // account could swap the screenshot before it went into the draft and the submission.
    [Fact]
    public void A_scope_capture_waits_in_the_users_own_app_data_folder_not_the_shared_temporary_one()
    {
        string ownFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            AppConfig.AppFolderName);

        Assert.StartsWith(ownFolder + Path.DirectorySeparatorChar, AppConfig.ScopeCaptureWaitFolder, StringComparison.Ordinal);
        Assert.False(AppConfig.ScopeCaptureWaitFolder.StartsWith(Path.GetTempPath(), StringComparison.OrdinalIgnoreCase));
    }
}
