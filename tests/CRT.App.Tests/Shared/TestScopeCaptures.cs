using System.Runtime.CompilerServices;
using CRT;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** EVERY TEST PROCESS KEEPS ITS OSCILLOSCOPE CAPTURES IN A FOLDER OF ITS OWN (code review,
// 2026-11-21). *** A capture waits in AppConfig.ScopeCaptureWaitFolder until it is accepted -
// the user's real "Scope captures waiting" folder, where a screen a failing test left behind
// stayed for good. Set before any test runs, as the assembly loads, and taken away as the process
// ends - as TestKiCadCache does. Compiled into all three App test projects (their csproj).
// ###########################################################################################
internal static class TestScopeCaptures
{
    internal static string Folder { get; } = Path.Combine(Path.GetTempPath(), "crt-tests-scope-captures-" + Guid.NewGuid().ToString("N"));

#pragma warning disable CA2255 // A module initializer is meant for libraries' own set-up; here it is the one moment before every test.
    [ModuleInitializer]
#pragma warning restore CA2255
    internal static void UseOwnFolder()
    {
        AppConfig.ScopeCaptureWaitFolderForTests = TestScopeCaptures.Folder;

        AppDomain.CurrentDomain.ProcessExit += (_, _) =>
        {
            try
            {
                if (Directory.Exists(TestScopeCaptures.Folder))
                {
                    Directory.Delete(TestScopeCaptures.Folder, recursive: true);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Left in the temporary folder, which the system tidies.
            }
        };
    }
}
