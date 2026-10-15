using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// LocalChecksumCache - the hashes of the local data files, kept with each file's stamp so a launch
// hashes only what changed (code review, 2026-10-13: every launch hashed the whole 1.1 GB tree).
// The rule that matters most: the cache must never make CRT trust a file it would otherwise
// download again - so every way the stamp can move, and every way the cache file can be wrong,
// must end in "hash it".
public sealed class LocalChecksumCacheTests : IDisposable
{
    private const string Hash = "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef";
    private const string OtherHash = "fedcba9876543210fedcba9876543210fedcba9876543210fedcba9876543210";

    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string CachePath => this.thisWorkspace.Path_("checksums.json");

    private string DataRoot => this.thisWorkspace.Path_("Data");

    private static readonly LocalChecksumCache.FileStamp Stamp = new(1000, new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc).Ticks, 5);

    // Hashed a minute after the last write: well clear of the racy window.
    private static readonly DateTime HashedAt = new(2026, 1, 1, 0, 1, 0, DateTimeKind.Utc);

    [Fact]
    public void A_hash_answers_while_the_stamp_is_exactly_as_it_was()
    {
        var cache = LocalChecksumCache.Load(this.CachePath, this.DataRoot);

        cache.Remember("C64/a.png", Stamp, Hash, HashedAt);

        Assert.Equal(Hash, cache.TryGet("C64/a.png", Stamp));
        Assert.Null(cache.TryGet("C64/b.png", Stamp));
    }

    // Any part of the stamp moving - a write (length or time), a file replaced (creation time) -
    // means hash it again.
    [Fact]
    public void Any_change_to_the_stamp_means_hashing_again()
    {
        var cache = LocalChecksumCache.Load(this.CachePath, this.DataRoot);
        cache.Remember("a.png", Stamp, Hash, HashedAt);

        Assert.Null(cache.TryGet("a.png", Stamp with { Length = 1001 }));
        Assert.Null(cache.TryGet("a.png", Stamp with { LastWriteUtcTicks = Stamp.LastWriteUtcTicks + 1 }));
        Assert.Null(cache.TryGet("a.png", Stamp with { CreationUtcTicks = 6 }));
    }

    // A file hashed within the racy window of its last write could be written again without its
    // stamp moving - so it is not remembered, and an older memory of it is dropped.
    [Fact]
    public void A_file_hashed_right_after_it_was_written_is_not_remembered()
    {
        var cache = LocalChecksumCache.Load(this.CachePath, this.DataRoot);
        cache.Remember("a.png", Stamp, OtherHash, HashedAt);

        var justWritten = new DateTime(Stamp.LastWriteUtcTicks, DateTimeKind.Utc) + TimeSpan.FromSeconds(1);
        cache.Remember("a.png", Stamp, Hash, justWritten);

        Assert.Null(cache.TryGet("a.png", Stamp));
    }

    [Fact]
    public void A_saved_cache_reads_back_for_the_same_data_folder()
    {
        var cache = LocalChecksumCache.Load(this.CachePath, this.DataRoot);
        cache.Remember("C64/a.png", Stamp, Hash, HashedAt);
        cache.Save();

        var reloaded = LocalChecksumCache.Load(this.CachePath, this.DataRoot + Path.DirectorySeparatorChar);

        Assert.Equal(Hash, reloaded.TryGet("C64/a.png", Stamp));
    }

    // Another data folder's hashes say nothing about this one's files.
    [Fact]
    public void A_cache_made_for_another_data_folder_is_empty()
    {
        var cache = LocalChecksumCache.Load(this.CachePath, this.DataRoot);
        cache.Remember("a.png", Stamp, Hash, HashedAt);
        cache.Save();

        var other = LocalChecksumCache.Load(this.CachePath, this.thisWorkspace.Path_("Other data"));

        Assert.Equal(0, other.Count);
        Assert.Null(other.TryGet("a.png", Stamp));
    }

    [Theory]
    [InlineData("")]
    [InlineData("{ not json")]
    [InlineData("null")]
    [InlineData("""{"dataRoot": 5}""")]
    public void A_missing_or_garbled_cache_file_is_an_empty_cache(string content)
    {
        File.WriteAllText(this.CachePath, content);

        var cache = LocalChecksumCache.Load(this.CachePath, this.DataRoot);

        Assert.Equal(0, cache.Count);
        Assert.Equal(0, LocalChecksumCache.Load(this.thisWorkspace.Path_("none.json"), this.DataRoot).Count);
    }

    // An entry whose hash is not a SHA-256 in lowercase hex is not believed.
    [Fact]
    public void An_entry_without_a_proper_hash_is_ignored()
    {
        string root = Path.TrimEndingDirectorySeparator(Path.GetFullPath(this.DataRoot));
        string json = System.Text.Json.JsonSerializer.Serialize(new
        {
            dataRoot = root,
            files = new Dictionary<string, object>
            {
                ["a.png"] = new { length = Stamp.Length, lastWriteUtcTicks = Stamp.LastWriteUtcTicks, creationUtcTicks = Stamp.CreationUtcTicks, sha256 = "XYZ" },
                ["b.png"] = new { length = Stamp.Length, lastWriteUtcTicks = Stamp.LastWriteUtcTicks, creationUtcTicks = Stamp.CreationUtcTicks, sha256 = Hash.ToUpperInvariant() },
                ["c.png"] = new { length = Stamp.Length, lastWriteUtcTicks = Stamp.LastWriteUtcTicks, creationUtcTicks = Stamp.CreationUtcTicks, sha256 = Hash },
            },
        });
        File.WriteAllText(this.CachePath, json);

        var cache = LocalChecksumCache.Load(this.CachePath, this.DataRoot);

        Assert.Null(cache.TryGet("a.png", Stamp));
        Assert.Null(cache.TryGet("b.png", Stamp));
        Assert.Equal(Hash, cache.TryGet("c.png", Stamp));
    }

    [Fact]
    public void A_forgotten_file_is_hashed_again()
    {
        var cache = LocalChecksumCache.Load(this.CachePath, this.DataRoot);
        cache.Remember("a.png", Stamp, Hash, HashedAt);

        cache.Forget("a.png");

        Assert.Null(cache.TryGet("a.png", Stamp));
    }

    // Nothing learned, nothing written.
    [Fact]
    public void An_unchanged_cache_is_not_written()
    {
        var cache = LocalChecksumCache.Load(this.CachePath, this.DataRoot);

        cache.Save();

        Assert.False(File.Exists(this.CachePath));
    }

    // The stamp is read off a real file the way the sync reads it, and moves when the file does.
    [Fact]
    public void A_real_files_stamp_moves_when_it_is_rewritten()
    {
        string file = this.thisWorkspace.WriteFile("Data/a.bin", "first");
        File.SetLastWriteTimeUtc(file, new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc));
        var before = LocalChecksumCache.FileStamp.Of(new FileInfo(file));

        File.WriteAllText(file, "second!");
        var after = LocalChecksumCache.FileStamp.Of(new FileInfo(file));

        Assert.NotEqual(before, after);
    }
}
