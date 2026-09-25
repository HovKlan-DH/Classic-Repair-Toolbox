using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for KiCadProjectLoader - the caching wrapper around
// KiCadRawProjectLoader that also builds the schematic net-path index.
//
// The cache key includes each file's last-write time, so an edited KiCad file must produce a
// fresh parse. That is the behaviour a contributor relies on when re-exporting a board.
public sealed class KiCadProjectLoaderTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string WritePcb(string name = "board.kicad_pcb", string? content = null) =>
        this.thisWorkspace.WriteFile(name, content ?? KiCadFixtures.Pcb);

    private static Task<KiCadProjectBundle?> LoadAsync(params string[] paths) =>
        KiCadProjectLoader.LoadRawAsync(paths);

    [Fact]
    public async Task LoadRawAsync_returns_null_when_given_no_paths()
    {
        Assert.Null(await LoadAsync());
    }

    [Fact]
    public async Task LoadRawAsync_returns_null_when_no_path_exists()
    {
        Assert.Null(await LoadAsync(this.thisWorkspace.Path_("missing.kicad_pcb")));
    }

    [Fact]
    public async Task LoadRawAsync_returns_null_when_the_files_parse_to_nothing()
    {
        Assert.Null(await LoadAsync(this.WritePcb(content: "(not_a_board (version 1))")));
    }

    [Fact]
    public async Task LoadRawAsync_returns_a_bundle_wrapping_the_parsed_root()
    {
        KiCadProjectBundle? bundle = await LoadAsync(this.WritePcb());

        Assert.NotNull(bundle);
        Assert.Single(bundle!.Root.Pcb);
        Assert.True(bundle.Root.Ok);
    }

    [Fact]
    public async Task The_same_unchanged_files_come_back_from_the_cache()
    {
        string pcb = this.WritePcb();

        KiCadProjectBundle? first = await LoadAsync(pcb);
        KiCadProjectBundle? second = await LoadAsync(pcb);

        Assert.Same(first, second);
    }

    [Fact]
    public async Task Editing_a_file_invalidates_the_cache()
    {
        // The cache key embeds each file's last-write time, so a re-exported board re-parses.
        string pcb = this.WritePcb();
        KiCadProjectBundle? first = await LoadAsync(pcb);

        File.WriteAllText(pcb, KiCadFixtures.PcbWithPropertyReference);
        File.SetLastWriteTimeUtc(pcb, DateTime.UtcNow.AddSeconds(5));

        KiCadProjectBundle? second = await LoadAsync(pcb);

        Assert.NotSame(first, second);
        Assert.Equal("R1", second!.Root.Pcb[0].Footprints[0].Reference);
    }

    [Fact]
    public async Task Path_order_does_not_change_the_cache_identity()
    {
        // Paths are sorted before the key is built, so the same set hits the same cache entry.
        string pcb = this.WritePcb();
        string sch = this.thisWorkspace.WriteFile("board.kicad_sch", KiCadFixtures.Schematic);

        KiCadProjectBundle? first = await LoadAsync(pcb, sch);
        KiCadProjectBundle? second = await LoadAsync(sch, pcb);

        Assert.Same(first, second);
    }

    [Fact]
    public async Task Duplicate_paths_are_collapsed()
    {
        string pcb = this.WritePcb();

        KiCadProjectBundle? bundle = await LoadAsync(pcb, pcb, pcb);

        Assert.Single(bundle!.Root.Pcb);
    }

    [Fact]
    public async Task Blank_paths_are_ignored()
    {
        string pcb = this.WritePcb();

        KiCadProjectBundle? bundle = await LoadAsync("", "   ", pcb);

        Assert.NotNull(bundle);
        Assert.Single(bundle!.Root.Pcb);
    }

    [Fact]
    public async Task A_failed_load_is_also_cached_so_a_broken_board_is_not_re_parsed()
    {
        // Documents current behaviour: null results are remembered too.
        string bad = this.WritePcb(name: "bad.kicad_pcb", content: "(nope)");

        Assert.Null(await LoadAsync(bad));
        Assert.Null(await LoadAsync(bad));
    }

    [Fact]
    public async Task A_schematic_net_path_index_is_built_for_each_schematic()
    {
        string sch = this.thisWorkspace.WriteFile("board.kicad_sch", KiCadFixtures.Schematic);

        KiCadProjectBundle? bundle = await LoadAsync(sch);

        Assert.NotNull(bundle);
        Assert.NotNull(bundle!.SchematicNetPathIndexBySchematicIndex);
        // One entry per schematic index that produced any wire path.
        Assert.True(bundle.SchematicNetPathIndexBySchematicIndex.Count <= bundle.Root.Schematics.Count);
    }
}
