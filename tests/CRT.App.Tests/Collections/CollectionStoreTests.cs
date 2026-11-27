using System.Security.Cryptography;
using System.Text.Json.Nodes;
using Handlers.Collections;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// The oscilloscope collections kept on this computer (CollectionStore) - in a folder of the test's own, never the user's.
// ###########################################################################################
public sealed class CollectionStoreTests : IDisposable
{
    private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

    private readonly TempWorkspace thisWorkspace = new();

    private readonly CollectionStore thisStore;

    public CollectionStoreTests()
    {
        this.thisStore = new CollectionStore(this.thisWorkspace.Root);
    }

    public void Dispose() => this.thisWorkspace.Dispose();

    private static LocalCapture Capture(string label = "U7", string pin = "12") =>
        new() { ComponentLabel = label, Pin = pin, SignalName = "/RAS", TimePerDivision = "50us", CapturedUtc = Now };

    private static string ShaOf(byte[] bytes) => Convert.ToHexStringLower(SHA256.HashData(bytes));

    [Fact]
    public void A_new_collection_is_kept_in_a_folder_of_its_own_and_read_back()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", "PAL", "U7 captures", Now)!;

        Assert.True(CollectionStore.IsLocalId(made.LocalId));
        Assert.True(File.Exists(Path.Combine(this.thisWorkspace.Root, made.LocalId, CollectionStore.FileName)));

        LocalCollection read = this.thisStore.Find(made.LocalId)!;

        Assert.Equal(("Commodore 64", "250407", "PAL", "U7 captures"), (read.HardwareName, read.BoardName, read.Region, read.Title));
        Assert.Equal("1y", read.Lifetime);
        Assert.False(read.IsShared);
        Assert.Empty(read.Captures);
    }

    [Fact]
    public void A_board_lists_its_own_collections_newest_first_capitals_aside()
    {
        LocalCollection older = this.thisStore.Create("Commodore 64", "250407", null, "Older", Now)!;
        LocalCollection newer = this.thisStore.Create("Commodore 64", "250407", null, "Newer", Now.AddHours(1))!;
        _ = this.thisStore.Create("Commodore 64", "250425", null, "Another board", Now);

        IReadOnlyList<LocalCollection> listed = this.thisStore.OfBoard("COMMODORE 64", "250407");

        Assert.Equal([newer.LocalId, older.LocalId], listed.Select(collection => collection.LocalId));
    }

    [Fact]
    public void A_capture_is_kept_by_its_checksum_and_the_same_picture_twice_is_one()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", null, null, Now)!;
        byte[] png = CRT.Data.Tests.TestPng.Rgb(32, 20);

        CaptureAdded once = this.thisStore.AddCapture(made.LocalId, png, Capture());
        CaptureAdded twice = this.thisStore.AddCapture(made.LocalId, png, Capture("U8", "3"));

        LocalCapture kept = Assert.Single(twice.Collection!.Captures);

        Assert.Equal(ShaOf(png), kept.Sha256);
        Assert.Equal("U7", kept.ComponentLabel);
        Assert.Equal(png, File.ReadAllBytes(this.thisStore.CapturePath(made.LocalId, ShaOf(png))));
        Assert.Equal(png, this.thisStore.ReadCapture(made.LocalId, ShaOf(png)));
        Assert.Single(once.Collection!.Captures);

        // The second is said to be in already - the popup offers no "Take out" for it, which would take the first out (code
        // review, 2026-11-27).
        Assert.Equal((CaptureAddedKind.Added, CaptureAddedKind.AlreadyIn), (once.Kind, twice.Kind));
    }

    // A collection takes only what the server shares by (CollectionPng) - said at the capture, not at "Share" (code review,
    // 2026-11-27: a screen too large failed the whole share, naming no capture).
    [Fact]
    public void A_capture_the_server_would_refuse_is_not_kept()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", null, null, Now)!;

        CaptureAdded notAPicture = this.thisStore.AddCapture(made.LocalId, [1, 2, 3], Capture());
        CaptureAdded tooLarge = this.thisStore.AddCapture(made.LocalId, CRT.Data.Tests.TestPng.Rgb(2560, 4), Capture("U8", "3"));

        Assert.Equal((CaptureAddedKind.Refused, Handlers.DataHandling.CollectionPng.NotAPicture), (notAPicture.Kind, notAPicture.Problem));
        Assert.Equal(CaptureAddedKind.Refused, tooLarge.Kind);
        Assert.Contains("2048", tooLarge.Problem, StringComparison.Ordinal);
        Assert.Empty(this.thisStore.Find(made.LocalId)!.Captures);
        Assert.Empty(Directory.EnumerateFiles(Path.Combine(this.thisWorkspace.Root, made.LocalId), "*.png"));
    }

    [Fact]
    public void A_capture_taken_out_goes_with_its_file()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", null, null, Now)!;
        byte[] png = CRT.Data.Tests.TestPng.Rgb(16, 10, seed: 9);

        this.thisStore.AddCapture(made.LocalId, png, Capture());

        LocalCollection changed = this.thisStore.RemoveCapture(made.LocalId, ShaOf(png))!;

        Assert.Empty(changed.Captures);
        Assert.False(File.Exists(this.thisStore.CapturePath(made.LocalId, ShaOf(png))));
        Assert.Null(this.thisStore.ReadCapture(made.LocalId, ShaOf(png)));
    }

    [Fact]
    public void A_collection_deleted_goes_whole()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", null, null, Now)!;
        Assert.Equal(CaptureAddedKind.Added, this.thisStore.AddCapture(made.LocalId, CRT.Data.Tests.TestPng.Rgb(8, 8), Capture()).Kind);

        Assert.True(this.thisStore.Delete(made.LocalId));
        Assert.False(Directory.Exists(Path.Combine(this.thisWorkspace.Root, made.LocalId)));
        Assert.Null(this.thisStore.Find(made.LocalId));
    }

    // Every path is made from an id of the right shape - nothing leads out of the folder.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("..")]
    [InlineData("../../Windows")]
    [InlineData("0123456789ABCDEF0123456789ABCDEF")]
    public void Anything_but_a_local_id_finds_and_deletes_nothing(string? id)
    {
        Assert.False(CollectionStore.IsLocalId(id));
        Assert.Null(this.thisStore.Find(id));
        Assert.False(this.thisStore.Delete(id ?? string.Empty));
    }

    [Fact]
    public void A_collection_whose_file_cannot_be_read_is_left_out_and_the_rest_listed()
    {
        LocalCollection good = this.thisStore.Create("Commodore 64", "250407", null, "Good", Now)!;
        string broken = Path.Combine(this.thisWorkspace.Root, Guid.NewGuid().ToString("N"));

        Directory.CreateDirectory(broken);
        File.WriteAllText(Path.Combine(broken, CollectionStore.FileName), "{ not json");

        Assert.Equal([good.LocalId], this.thisStore.All().Select(collection => collection.LocalId));
    }

    // A field a later CRT writes is kept and written back.
    [Fact]
    public void A_field_this_CRT_does_not_know_is_kept()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", null, null, Now)!;
        string path = Path.Combine(this.thisWorkspace.Root, made.LocalId, CollectionStore.FileName);

        JsonObject json = JsonNode.Parse(File.ReadAllText(path))!.AsObject();
        json["laterField"] = "kept";
        File.WriteAllText(path, json.ToJsonString());

        this.thisStore.Save(this.thisStore.Find(made.LocalId)! with { Title = "Renamed" });

        JsonObject saved = JsonNode.Parse(File.ReadAllText(path))!.AsObject();

        Assert.Equal("kept", (string?)saved["laterField"]);
        Assert.Equal("Renamed", (string?)saved["title"]);
    }

    // The list is read once and kept until the store writes again - every popup asked for it twice a component, each time
    // parsing every collection on the UI thread (code review, 2026-11-27).
    [Fact]
    public void The_list_is_kept_until_the_store_writes_again()
    {
        LocalCollection first = this.thisStore.Create("Commodore 64", "250407", null, "First", Now)!;
        IReadOnlyList<LocalCollection> read = this.thisStore.All();

        // Something the store did not write is not seen until it writes again.
        new CollectionStore(this.thisWorkspace.Root).Create("Commodore 64", "250407", null, "Written elsewhere", Now.AddHours(1));

        Assert.Same(read, this.thisStore.All());

        this.thisStore.Update(first.LocalId, collection => collection with { Title = "Renamed" }, out _);

        Assert.Equal(["Written elsewhere", "Renamed"], this.thisStore.All().Select(collection => collection.Title));

        Assert.True(this.thisStore.Delete(first.LocalId));
        Assert.Equal(["Written elsewhere"], this.thisStore.All().Select(collection => collection.Title));
    }

    // A new board renamed takes its collections along - the popup finds them by the board's names (code review, 2026-11-27).
    [Fact]
    public void A_board_renamed_takes_its_collections_along()
    {
        LocalCollection mine = this.thisStore.Create("Commodore 64", "Prototype", null, "Mine", Now)!;
        LocalCollection another = this.thisStore.Create("Commodore 64", "250407", null, "Another board's", Now)!;

        Assert.Equal(1, this.thisStore.RenameBoard("COMMODORE 64", "prototype", "Commodore 64", "Rev A"));

        Assert.Equal([mine.LocalId], this.thisStore.OfBoard("Commodore 64", "Rev A").Select(collection => collection.LocalId));
        Assert.Empty(this.thisStore.OfBoard("Commodore 64", "Prototype"));
        Assert.Equal("250407", this.thisStore.Find(another.LocalId)?.BoardName);
    }

    // ---- Kept per service (code review, 2026-11-27) ---------------------------------------------

    // Only the 4.1.0 pre-releases ever kept collections - all shared with the test instance - in the folder a release uses; a
    // pre-release of 4.1.0 takes them to its own. Once 4.1.0 is released, a later pre-release would take a release's.
    [Theory]
    [InlineData(true, "4.1.0-alpha.1", true)]
    [InlineData(true, "4.1.0-beta.3", true)]
    [InlineData(true, "4.1.1-alpha.1", false)]
    [InlineData(true, "4.2.0-alpha.1", false)]
    [InlineData(false, "4.1.0", false)]
    [InlineData(false, "4.1.0-alpha.1", false)]
    [InlineData(true, "not a version", false)]
    public void Only_a_4_1_0_pre_release_takes_over_the_earlier_folder(bool testServer, string version, bool takesOver) =>
        Assert.Equal(takesOver, CollectionStore.TakesOverEarlierFolder(testServer, version));

    [Fact]
    public void The_earlier_folder_moves_whole_while_there_is_none_yet()
    {
        string earlier = this.thisWorkspace.Path_("Collections");
        string root = this.thisWorkspace.Path_("Collections-test");
        LocalCollection made = new CollectionStore(earlier).Create("Commodore 64", "250407", null, "Mine", Now)!;

        Assert.True(CollectionStore.TakeOverEarlierFolder(earlier, root));
        Assert.False(Directory.Exists(earlier));
        Assert.Equal("Mine", new CollectionStore(root).Find(made.LocalId)?.Title);
        Assert.False(CollectionStore.TakeOverEarlierFolder(earlier, root));
    }

    // A move refused once, and a collection made in the new folder since: what the earlier one still holds moves one by one,
    // never over a collection there, and the earlier folder goes once empty (code review, 2026-11-27).
    [Fact]
    public void What_the_earlier_folder_still_holds_moves_one_collection_at_a_time()
    {
        string earlier = this.thisWorkspace.Path_("Collections");
        string root = this.thisWorkspace.Path_("Collections-test");
        LocalCollection left = new CollectionStore(earlier).Create("Commodore 64", "250407", null, "Left behind", Now)!;
        LocalCollection madeSince = new CollectionStore(root).Create("Commodore 64", "250407", null, "Made since", Now)!;

        Assert.True(CollectionStore.TakeOverEarlierFolder(earlier, root));

        Assert.Equal(["Left behind", "Made since"], new CollectionStore(root).All().Select(collection => collection.Title).Order());
        Assert.False(Directory.Exists(earlier));
        Assert.NotNull(new CollectionStore(root).Find(left.LocalId));
        Assert.NotNull(new CollectionStore(root).Find(madeSince.LocalId));
    }

    // The window's change and a capture going in from the pool are made one after the other on the collection AS KEPT - neither
    // throws the other's away (code review, 2026-11-27).
    [Fact]
    public void An_update_is_made_on_the_collection_as_kept_now()
    {
        LocalCollection made = this.thisStore.Create("Commodore 64", "250407", null, "Before", Now)!;

        this.thisStore.AddCapture(made.LocalId, CRT.Data.Tests.TestPng.Rgb(8, 8), Capture());

        LocalCollection? updated = this.thisStore.Update(made.LocalId, collection => collection with { Title = "After" }, out bool kept);

        Assert.True(kept);
        Assert.Equal(("After", 1), (updated?.Title, updated?.Captures.Count));
        Assert.Equal(("After", 1), (this.thisStore.Find(made.LocalId)?.Title, this.thisStore.Find(made.LocalId)?.Captures.Count));

        // Handing back the very one it was given writes nothing; no such collection is none.
        string file = Path.Combine(this.thisWorkspace.Root, made.LocalId, CollectionStore.FileName);
        DateTime written = File.GetLastWriteTimeUtc(file);

        Assert.Equal("After", this.thisStore.Update(made.LocalId, collection => collection, out bool keptNothing)?.Title);
        Assert.True(keptNothing);
        Assert.Equal(written, File.GetLastWriteTimeUtc(file));
        Assert.Null(this.thisStore.Update(Guid.NewGuid().ToString("N"), collection => collection, out _));
    }
}
