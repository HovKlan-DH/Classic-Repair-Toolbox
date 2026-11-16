using System.Runtime.CompilerServices;
using CRT;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** EVERY TEST PROCESS COPIES KiCad FILES INTO A CACHE OF ITS OWN (2026-11-16, performance
// walkthrough). *** Tests showing a board through Main or the Contribute tab copy its KiCad files
// for the loaders (BoardKiCadFolder) into AppConfig.KiCadCacheFolder - which was the user's real
// "KiCad cache", where the copies of every test board stayed for good. Set before any test runs,
// as the assembly loads, and taken away as the process ends. Compiled into all three App test
// projects (their csproj), so each process has its own folder - they run at once.
// ###########################################################################################
internal static class TestKiCadCache
{
    internal static string Folder { get; } = Path.Combine(Path.GetTempPath(), "crt-tests-kicad-cache-" + Guid.NewGuid().ToString("N"));

#pragma warning disable CA2255 // A module initializer is meant for libraries' own set-up; here it is the one moment before every test.
    [ModuleInitializer]
#pragma warning restore CA2255
    internal static void UseOwnFolder()
    {
        AppConfig.KiCadCacheFolderForTests = TestKiCadCache.Folder;

        AppDomain.CurrentDomain.ProcessExit += (_, _) =>
        {
            try
            {
                if (Directory.Exists(TestKiCadCache.Folder))
                {
                    Directory.Delete(TestKiCadCache.Folder, recursive: true);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Left in the temporary folder, which the system tidies.
            }
        };
    }
}
