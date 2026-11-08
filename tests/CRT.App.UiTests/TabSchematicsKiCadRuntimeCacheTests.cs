using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The runtime cache scope the KiCad overlay's caches go into is the one of the files LOADED - never
// looked up on disk again (2026-11-06, the alignment's slowness). It was asked once per net for every
// net a rebuild builds, and each ask took a full path, File.Exists and File.GetLastWriteTimeUtc of
// every KiCad file of the board: on the Open128 (451 nets, 25 files) a rebuild after the calibration
// box moved took 261 ms, and 30 without. Files deleted after the load show it: asked on disk, the
// scope was gone, and every cache with it.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabSchematicsKiCadRuntimeCacheTests
{
    [Fact]
    public void The_runtime_cache_scope_is_the_loaded_files_one_and_is_not_looked_up_on_disk_again()
    {
        using var workspace = new TempWorkspace();
        string pcb = workspace.WriteFile("board.kicad_pcb", "(kicad_pcb)");
        string sheet = workspace.WriteFile("board.kicad_sch", "(kicad_sch)");

        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            tab.UseKiCadRawPathsForTests([pcb, sheet]);

            object? scope = tab.KiCadRuntimeCacheScopeForTests;
            Assert.NotNull(scope);

            File.Delete(pcb);
            File.Delete(sheet);

            Assert.Same(scope, tab.KiCadRuntimeCacheScopeForTests);
        });
    }

    // No KiCad files loaded, no scope: the caches stay the tab's own, as before.
    [Fact]
    public void A_board_without_KiCad_files_has_no_runtime_cache_scope()
    {
        UiTest.Run(() =>
        {
            var tab = new TabSchematics();
            tab.UseKiCadRawPathsForTests([]);

            Assert.Null(tab.KiCadRuntimeCacheScopeForTests);
        });
    }
}
