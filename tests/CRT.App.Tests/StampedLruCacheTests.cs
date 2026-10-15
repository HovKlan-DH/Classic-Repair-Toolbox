using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// StampedLruCache - the parsed KiCad projects' cache (code review, 2026-10-13: a static Dictionary
// written from pool threads with no lock, that kept one more whole parsed board for every board
// shown and every time its files changed).
public sealed class StampedLruCacheTests
{
    [Fact]
    public void A_value_answers_only_at_the_stamp_it_was_stored_with()
    {
        var cache = new StampedLruCache<string>(4, StringComparer.OrdinalIgnoreCase);

        cache.Store("board-a", "t1", "parse 1");

        Assert.True(cache.TryGet("BOARD-A", "t1", out string? value));
        Assert.Equal("parse 1", value);
        Assert.False(cache.TryGet("board-a", "t2", out _));
        Assert.False(cache.TryGet("board-b", "t1", out _));
    }

    // Files that changed leave no old parse behind: a new stamp REPLACES the key's entry.
    [Fact]
    public void Storing_a_new_stamp_drops_the_older_one_of_the_same_files()
    {
        var cache = new StampedLruCache<string>(4, StringComparer.OrdinalIgnoreCase);

        cache.Store("board-a", "t1", "parse 1");
        cache.Store("board-a", "t2", "parse 2");

        Assert.Equal(1, cache.Count);
        Assert.False(cache.TryGet("board-a", "t1", out _));
        Assert.True(cache.TryGet("board-a", "t2", out string? value));
        Assert.Equal("parse 2", value);
    }

    [Fact]
    public void Past_the_capacity_the_least_recently_used_goes()
    {
        var cache = new StampedLruCache<string>(2, StringComparer.OrdinalIgnoreCase);

        cache.Store("a", "t", "A");
        cache.Store("b", "t", "B");
        Assert.True(cache.TryGet("a", "t", out _));   // a is now the most recently used

        cache.Store("c", "t", "C");

        Assert.Equal(2, cache.Count);
        Assert.True(cache.TryGet("a", "t", out _));
        Assert.False(cache.TryGet("b", "t", out _));
        Assert.True(cache.TryGet("c", "t", out _));
    }

    // A failed load is remembered too (null), so the same broken files are not parsed again.
    [Fact]
    public void A_null_value_is_a_hit()
    {
        var cache = new StampedLruCache<string?>(4, StringComparer.Ordinal);

        cache.Store("broken", "t", null);

        Assert.True(cache.TryGet("broken", "t", out string? value));
        Assert.Null(value);
    }

    // Used from pool threads: many writers and readers at once must leave it whole and in bounds.
    [Fact]
    public async Task Concurrent_use_keeps_it_within_its_capacity()
    {
        var cache = new StampedLruCache<int>(4, StringComparer.Ordinal);

        await Task.WhenAll(Enumerable.Range(0, 8).Select(worker => Task.Run(() =>
        {
            for (int i = 0; i < 2_000; i++)
            {
                string key = $"k{(i + worker) % 10}";
                cache.Store(key, $"s{i % 3}", i);
                cache.TryGet(key, $"s{i % 3}", out _);
            }
        })));

        Assert.InRange(cache.Count, 1, 4);
    }
}
