using CRT;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// PaneImageCache - the component pane's pictures, read once (code review, 2026-10-26). Kept by where
// the bytes lie, their length and write time and the width decoded to - never by name, which another
// board's file of the same name, or bytes put back under it, share - and a load that started before
// the cache was let go keeps nothing.
// ###########################################################################################
public sealed class PaneImageCacheTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    [Fact]
    public void Bytes_of_the_same_name_elsewhere_or_replaced_are_kept_apart()
    {
        string first = this.thisWorkspace.WriteFile(Path.Combine("A", "Scope baseline", "U8_1.png"), "first board's");
        string second = this.thisWorkspace.WriteFile(Path.Combine("B", "Scope baseline", "U8_1.png"), "second board's");

        string? firstKey = PaneImageCache.KeyOf(first, 336);
        Assert.NotNull(firstKey);
        Assert.NotEqual(firstKey, PaneImageCache.KeyOf(second, 336));
        Assert.NotEqual(firstKey, PaneImageCache.KeyOf(first, 640));

        // The same file written again: other bytes, another key.
        File.WriteAllText(first, "first board's, replaced");
        File.SetLastWriteTimeUtc(first, File.GetLastWriteTimeUtc(first).AddSeconds(5));
        Assert.NotEqual(firstKey, PaneImageCache.KeyOf(first, 336));
    }

    [Fact]
    public void No_file_has_no_key()
    {
        Assert.Null(PaneImageCache.KeyOf(null, 336));
        Assert.Null(PaneImageCache.KeyOf("  ", 336));
        Assert.Null(PaneImageCache.KeyOf(this.thisWorkspace.Path_("missing.png"), 336));
    }

    // A load that started before the cache was let go - another board, another component - keeps nothing.
    [Fact]
    public void A_load_from_before_the_cache_was_let_go_keeps_nothing()
    {
        var cache = new PaneImageCache();
        int before = cache.Generation;

        Assert.True(cache.Keep("kept", PaneImage.None, before));
        Assert.True(cache.TryGet("kept", out _));

        cache.Clear();

        Assert.False(cache.TryGet("kept", out _));
        Assert.False(cache.Keep("late", PaneImage.None, before));
        Assert.False(cache.TryGet("late", out _));
        Assert.Equal(0, cache.CountForTests);
    }
}
